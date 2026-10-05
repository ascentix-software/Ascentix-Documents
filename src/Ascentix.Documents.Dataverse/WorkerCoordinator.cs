using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Domain;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public sealed class WorkerCoordinator
{
    private readonly IOrganizationService service;
    private readonly DocumentStore store;
    private readonly WorkerCatalog catalog;
    private readonly Func<DateTime> clock;
    private readonly string[]? allowedTables;

    public WorkerCoordinator(
        IOrganizationService service,
        Func<DateTime>? clock = null,
        string[]? allowedTables = null
    )
    {
        this.service = service;
        store = new DocumentStore(service);
        catalog = new WorkerCatalog(service);
        this.clock = clock ?? (() => DateTime.UtcNow);
        this.allowedTables = allowedTables;
    }

    /// <summary>
    /// Dispatches one transactional step of record planning, folder work, or recovery.
    /// </summary>
    /// <param name="request">The worker command, operation identity, and command-specific inputs.</param>
    /// <param name="inTransaction">Whether execution is inside a Dataverse transaction; must be true.</param>
    /// <returns>The command status and any work keys, HTTP intent, or continuation information.</returns>
    public WorkerResult Execute(WorkerRequest request, bool inTransaction)
    {
        if (!inTransaction)
            throw new EvaluationBlockedException(
                "Worker coordination requires an ambient Dataverse transaction."
            );
        if (
            new[]
            {
                "Claim",
                "Renew",
                "Observe",
                "PrepareCreate",
                "CreateResponse",
                "Complete",
            }.Contains(request.Command)
        )
        {
            var stopped = StopUnavailable(request);
            if (stopped != null)
                return stopped;
        }
        switch (request.Command)
        {
            case "Queue":
                return Queue(request);
            case "ListOutbox":
                return new WorkerResult
                {
                    Status = "Page",
                    Keys = store.Pending("asx_outbox", now: clock()),
                };
            case "ListOperations":
                return new WorkerResult
                {
                    Status = "Page",
                    Keys = store.Pending("asx_operation", now: clock()),
                };
            case "Plan":
                return Plan(request.Key);
            case "Claim":
                return Claim(request);
            case "Renew":
                return Renew(request);
            case "Observe":
                return Observe(request);
            case "PrepareCreate":
                return PrepareCreate(request);
            case "CreateResponse":
                return CreateResponse(request);
            case "Complete":
                return Complete(request);
            case "Fail":
                return Fail(request);
            case "Retry":
                return Retry(request);
            case "Cancel":
                return Cancel(request);
            case "Replan":
                return Replan(request);
            case "FailUnclaimed":
                return store.FailUnclaimed<OperationDocument>(request.Key, request, clock());
            case "FailOutbox":
                return FailOutbox(request);
            case "RetryOutbox":
                return RetryOutbox(request);
            default:
                throw new EvaluationBlockedException("Unsupported worker command.");
        }
    }

    private WorkerResult Queue(WorkerRequest request)
    {
        if (
            request.RequestId == Guid.Empty
            || request.RecordId == Guid.Empty
            || request.TemplateId == Guid.Empty
        )
            throw new EvaluationBlockedException(
                "Stable request, template and record IDs required."
            );
        string key = "request:" + request.RequestId.ToString("N");
        var existing = store.Find<OutboxDocument>("asx_outbox", key);
        if (existing != null)
        {
            if (
                existing.Value.TemplateId != request.TemplateId
                || existing.Value.RecordId != request.RecordId
            )
                throw new EvaluationBlockedException(
                    "Idempotency key was reused with different intent."
                );
            return new WorkerResult
            {
                Status = existing.Value.Status,
                Key = key,
                Keys = existing.Value.Operations,
            };
        }
        var template = TemplateLifecycle.Find(service, request.TemplateId);
        if (!TemplateLifecycle.Active(template, clock()))
            return new WorkerResult
            {
                Status = "Inactive",
                Notices = new[] { "Template is deleted, deactivated or outside its schedule." },
            };
        var revision =
            template!.GetAttributeValue<EntityReference>("asx_publishedrevisionid")
            ?? throw new EvaluationBlockedException("Template has no published revision.");
        string table = TemplateStore.Text(template, "asx_table");
        Allowed(table);
        if (Retired(table, request.RecordId))
            throw new EvaluationBlockedException(
                "Deleted record requires decommission review; documents are retained."
            );
        service.Retrieve(table, request.RecordId, new ColumnSet(false));
        store.Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = key,
                TemplateId = template.Id,
                RevisionId = revision.Id,
                RecordId = request.RecordId,
                Table = table,
            }
        );
        return new WorkerResult { Status = "Pending", Key = key };
    }

    private WorkerResult Plan(string key)
    {
        var job = store.Require<OutboxDocument>("asx_outbox", key);
        // Any save of this row below ends an earlier wait after a temporary failure; a Plan that
        // fails again rolls back and keeps it.
        if (job.Value.NextAttemptUtc != null || job.Value.Attempts != 0)
        {
            job.Value.NextAttemptUtc = null;
            job.Value.Attempts = 0;
            if (job.Value.Status == "Pending")
                job.Value.Notices = Array.Empty<string>();
        }
        if (job.Value.RelatedRecordId != Guid.Empty)
            return new TargetedReplan(
                service,
                allowedTables ?? RuntimeProfile.Read(service).Tables
            ).Consume(job);
        if (job.Value.SecurityTeamId != Guid.Empty)
            return new SecurityRefresh(service).Consume(job);
        if (job.Value.Status != "Pending")
            return new WorkerResult
            {
                Status = job.Value.Status,
                Key = key,
                Keys = job.Value.Operations,
                Notices = job.Value.Notices,
            };
        if (Retired(job.Value.Table, job.Value.RecordId))
        {
            job.Value.Status = "DecommissionReview";
            job.Value.Notices = new[]
            {
                "Record deleted; existing documents and receipts are retained.",
            };
            store.Save(job);
            return new WorkerResult
            {
                Status = job.Value.Status,
                Key = key,
                Notices = job.Value.Notices,
            };
        }
        var header = TemplateLifecycle.Find(service, job.Value.TemplateId);
        if (!TemplateLifecycle.Active(header, clock()))
        {
            job.Value.Status = "Cancelled";
            job.Value.Notices = new[]
            {
                "Template is deleted, deactivated or outside its schedule. Existing SharePoint content is unchanged.",
            };
            store.Save(job);
            return new WorkerResult
            {
                Status = "Cancelled",
                Key = key,
                Notices = job.Value.Notices,
            };
        }
        var revision =
            header!.GetAttributeValue<EntityReference>("asx_publishedrevisionid")
            ?? throw new EvaluationBlockedException("Template is unpublished.");
        // Only the template's own table must be enabled. Lookup source tables are read with the
        // worker's Read privilege; their changes queue no replans unless they are enabled too.
        Allowed(job.Value.Table);
        if (
            job.Value.PinnedRevision
            && (
                job.Value.RevisionId != revision.Id
                || job.Value.ReviewedSources.Any(source => !MatchesSource(source))
            )
        )
        {
            job.Value.Status = "NeedsReview";
            job.Value.Notices = new[]
            {
                "Reviewed batch revision or source changed before planning. Request a new impact review.",
            };
            store.Save(job);
            return new WorkerResult
            {
                Status = job.Value.Status,
                Key = key,
                Notices = job.Value.Notices,
            };
        }
        job.Value.RevisionId = revision.Id;
        var template = new TemplateStore(service).Read(revision.Id);
        if (template.Id != job.Value.TemplateId || template.Table != job.Value.Table)
            throw new EvaluationBlockedException("Queued template/table identity changed.");
        var snapshot = new SnapshotReader(service).Read(template, job.Value.RecordId);
        var intents = FolderPlanner.Plan(template, job.Value.RecordId, snapshot.Values);
        for (int i = 0; i < snapshot.Records.Count; i++)
        {
            var record = snapshot.Records[i];
            var current = service.Retrieve(record.LogicalName, record.Id, new ColumnSet(false));
            if (
                string.IsNullOrEmpty(snapshot.Versions[i])
                || current.RowVersion != snapshot.Versions[i]
            )
                throw new EvaluationBlockedException(
                    "Source snapshot changed before binding; queue a fresh evaluation."
                );
        }
        string planKey =
            "recordplan:" + template.Id.ToString("N") + ":" + job.Value.RecordId.ToString("N");
        var selected = store.Find<RecordPlanDocument>("asx_outbox", planKey);
        var included = intents.Select(i => i.Section).Distinct().ToArray();
        var sourceVersions = snapshot
            .Records.Select(
                (record, index) =>
                    new SourceVersion
                    {
                        Table = record.LogicalName,
                        Id = record.Id,
                        Version = snapshot.Versions[index],
                    }
            )
            .ToArray();
        if (selected == null)
            store.Create(
                "asx_outbox",
                new RecordPlanDocument
                {
                    Key = planKey,
                    Table = template.Table,
                    TemplateId = template.Id,
                    RecordId = job.Value.RecordId,
                    Status = "Selection",
                    RevisionId = revision.Id,
                    Sources = sourceVersions,
                    IncludedSections = included,
                    Notices = intents.Notices.ToArray(),
                }
            );
        else
        {
            selected.Value.Table = template.Table;
            selected.Value.TemplateId = template.Id;
            selected.Value.RecordId = job.Value.RecordId;
            selected.Value.RevisionId = revision.Id;
            selected.Value.Sources = sourceVersions;
            selected.Value.IncludedSections = included;
            selected.Value.Notices = intents.Notices.ToArray();
            store.Save(selected);
        }
        var operations = new List<string>();
        foreach (var group in intents.GroupBy(i => i.Section))
        {
            var library = catalog.Read(group.First().LibraryApprovalId);
            string operationKey =
                "folderjob:" + DocumentStore.Hash(job.Value.Key + ":" + group.Key);
            var folders = group
                .Select(intent => new FolderStep
                {
                    Key = intent.BindingKey,
                    TemplateId = template.Id,
                    RecordId = job.Value.RecordId,
                    Table = template.Table,
                    Section = intent.Section,
                    Node = intent.Node,
                    ParentBinding = intent.Parent,
                    LibraryId = library.Id,
                    EntryId = library.EntryId,
                    OriginalName = intent.Name,
                    Candidate = intent.Name,
                })
                .ToArray();
            if (folders.Length == 0 || folders.Length > 100 || folders[0].ParentBinding != null)
                throw new EvaluationBlockedException(
                    "A destination job requires one root and at most 100 folders."
                );
            var location = new NativeLocations(service).Find(
                folders[0],
                library.NativeParentId,
                library.EntryUrl,
                library.NativeSiteId
            );
            if (location != null)
            {
                folders[0].Candidate = location.GetAttributeValue<string>("relativeurl");
                folders[0].LocationId = location.Id;
            }
            if (store.Find<OperationDocument>("asx_operation", operationKey) == null)
                store.Create(
                    "asx_operation",
                    new OperationDocument
                    {
                        Key = operationKey,
                        Folders = folders,
                        RevisionId = revision.Id,
                    }
                );
            operations.Add(operationKey);
        }
        foreach (var priorKey in selected?.Value.Operations ?? Array.Empty<string>())
        {
            var prior = store.Find<OperationDocument>("asx_operation", priorKey);
            if (prior?.Value.HistoryCompacted == true)
            {
                prior.Value.HistoryCompacted = false;
                store.Save(prior);
            }
        }
        var completedSelection = store.Require<RecordPlanDocument>("asx_outbox", planKey);
        completedSelection.Value.Operations = operations.ToArray();
        completedSelection.Value.Status = "Selection";
        store.Save(completedSelection);
        job.Value.Operations = operations.ToArray();
        job.Value.Notices = intents.Notices.ToArray();
        job.Value.Status = "Planned";
        store.Save(job);
        return new WorkerResult
        {
            Status = job.Value.Status,
            Key = key,
            Keys = job.Value.Operations,
            Notices = job.Value.Notices,
        };
    }

    private WorkerResult Claim(WorkerRequest request)
    {
        ValidateRun(request.RunId);
        var operation = store.Require<OperationDocument>("asx_operation", request.Key);
        if (operation.Value.Status == "Applied")
            return Done(operation);
        if (
            operation.Value.NextAttemptUtc.HasValue
            && operation.Value.NextAttemptUtc.Value > clock()
        )
            return new WorkerResult { Status = "RetryWait", Key = request.Key };
        if (
            operation.Value.Status == "Blocked"
            || operation.Value.Status == "Superseded"
            || operation.Value.Status == "Cancelled"
        )
            return new WorkerResult { Status = operation.Value.Status, Key = request.Key };
        var binding = operation.Value.Folder;
        var published = TemplateLifecycle
            .Find(service, binding.TemplateId)
            ?.GetAttributeValue<EntityReference>("asx_publishedrevisionid");
        var existingClaim = store.Find<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        var selection = store.Find<RecordPlanDocument>(
            "asx_outbox",
            "recordplan:" + binding.TemplateId.ToString("N") + ":" + binding.RecordId.ToString("N")
        );
        bool excluded =
            Retired(binding.Table, binding.RecordId)
            || selection != null
                && (
                    selection.Value.RevisionId != operation.Value.RevisionId
                    || !selection.Value.Operations.Contains(operation.Value.Key)
                );
        if (
            (published?.Id != operation.Value.RevisionId || excluded)
            && !operation.Value.ExternalSubmitted
            && existingClaim?.Value.OperationKey != request.Key
        )
        {
            operation.Value.Status = "Superseded";
            store.Save(operation);
            return new WorkerResult { Status = "Superseded", Key = request.Key };
        }
        if (
            !WorkCoordination.HasCapacity(
                service,
                WorkCoordination.Operation(service, request.Key),
                clock()
            )
        )
            return new WorkerResult { Status = "Busy", Key = request.Key };
        var library = Current(operation.Value, binding);
        var dispatcher = store.Find<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        if (dispatcher == null)
        {
            store.Create(
                "asx_claim",
                new DispatcherDocument
                {
                    Key = WorkCoordination.Operation(service, request.Key),
                    Status = "Idle",
                }
            );
            dispatcher = store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(service, request.Key)
            );
        }
        bool recovery = false;
        if (dispatcher.Value.RunId != null)
        {
            if (dispatcher.Value.OperationKey != request.Key)
                return new WorkerResult { Status = "Busy", Key = request.Key };
            bool sameLiveRun =
                dispatcher.Value.RunId == request.RunId
                && dispatcher.Value.Token == request.Token
                && dispatcher.Value.LeaseUntilUtc > clock();
            // A live claim of another run is left alone. Once its lease expires (5 minutes, at
            // least twice the connector timeout) the next claim takes over and re-reads.
            if (
                !sameLiveRun
                && !dispatcher.Value.RecoveryPermitted
                && dispatcher.Value.LeaseUntilUtc > clock()
            )
                return new WorkerResult { Status = "Quarantined", Key = request.Key };
            recovery = !sameLiveRun;
        }
        string parent;
        if (binding.ParentBinding == null)
            parent = library.Target.EntryPath;
        else
        {
            var parentBinding = operation.Value.Folders.Single(f =>
                f.Node == binding.ParentBinding
            );
            if (
                parentBinding.Status != "Applied"
                || parentBinding.PhysicalId == Guid.Empty
                || parentBinding.PhysicalPath == null
            )
                return new WorkerResult { Status = "WaitingForParent", Key = request.Key };
            parent = parentBinding.PhysicalPath;
        }
        library.Target.ValidateParent(parent);
        if (
            operation.Value.ParentPath != null
            && !string.Equals(operation.Value.ParentPath, parent, StringComparison.Ordinal)
        )
            throw new EvaluationBlockedException("Pinned parent path changed.");
        operation.Value.ParentPath = parent;
        operation.Value.AbsenceVerified = false;
        if (recovery)
        {
            // The earlier run's request is no longer awaited; the reads below establish what
            // SharePoint did before any write.
            dispatcher.Value.HttpOutstanding = false;
            if (operation.Value.ExternalSubmitted && !operation.Value.ExternalResponseKnown)
                operation.Value.Reprobe = true;
            operation.Value.ExternalResponseKnown = true;
        }
        dispatcher.Value.OperationKey = request.Key;
        dispatcher.Value.RunId = request.RunId;
        dispatcher.Value.Token = Guid.NewGuid();
        dispatcher.Value.LeaseUntilUtc = clock().AddMinutes(5);
        dispatcher.Value.RecoveryPermitted = false;
        dispatcher.Value.Status = "Claimed";
        store.Save(dispatcher);
        Audit(request.Key, request.RunId, recovery ? "RecoveryClaim" : "Claim");
        return Probe(operation, dispatcher.Value, library, "Library");
    }

    private WorkerResult Renew(WorkerRequest request)
    {
        var dispatcher = Assert(request);
        dispatcher.Value.LeaseUntilUtc = clock().AddMinutes(5);
        store.Save(dispatcher);
        return new WorkerResult
        {
            Status = "Renewed",
            Key = request.Key,
            Token = dispatcher.Value.Token,
        };
    }

    private WorkerResult Observe(WorkerRequest request)
    {
        var dispatcher = Assert(request);
        var operation = store.Require<OperationDocument>("asx_operation", request.Key);
        var binding = operation.Value.Folder;
        var library = Current(operation.Value, binding);
        if (
            request.ProbeId == Guid.Empty
            || request.ProbeId != operation.Value.ProbeId
            || request.ProbeKind != operation.Value.ProbeKind
        )
            throw new EvaluationBlockedException("Observation does not match the issued probe.");
        // Earlier versions read the library's role assignments before and after each folder.
        // Folder work no longer checks permissions, so an operation persisted mid-probe ignores
        // that response, whatever it was, and continues with the parent-folder read.
        if (operation.Value.ProbeKind == "Acl" || operation.Value.ProbeKind == "FinalAcl")
            return Probe(
                operation,
                dispatcher.Value,
                library,
                operation.Value.ProbeKind == "FinalAcl" ? "FinalParent" : "Parent"
            );
        if (
            request.HttpStatus == 429
            || request.HttpStatus >= 500
            || request.HttpStatus == 0
            || request.HttpStatus == 408
        )
            return ScheduleRetry(
                operation,
                dispatcher,
                request.RetryAfter,
                request.HttpStatus,
                null
            );
        try
        {
            switch (operation.Value.ProbeKind)
            {
                case "Library":
                    var observedLibrary = SharePointObservations.Body<LibraryObservation>(request);
                    if (
                        observedLibrary.Id != library.Target.ListId
                        || observedLibrary.Root == null
                        || observedLibrary.Root.Id == Guid.Empty
                        || string.IsNullOrWhiteSpace(observedLibrary.Root.Path)
                        || !(
                            library.Target.EntryPath == observedLibrary.Root.Path
                            || library.Target.EntryPath.StartsWith(
                                observedLibrary.Root.Path + "/",
                                StringComparison.Ordinal
                            )
                        )
                    )
                        throw new EvaluationBlockedException(
                            "Observed library/root differs from approved destination."
                        );
                    operation.Value.LibraryRootPath = observedLibrary.Root.Path;
                    operation.Value.LibraryRootId = observedLibrary.Root.Id;
                    return Probe(operation, dispatcher.Value, library, "Parent");
                case "Parent":
                case "FinalParent":
                    var parent = SharePointObservations.Body<FolderObservation>(request);
                    var expectedParentId =
                        binding.ParentBinding == null
                            ? library.EntryId
                            : operation
                                .Value.Folders.Single(f => f.Node == binding.ParentBinding)
                                .PhysicalId;
                    bool isLibraryRoot =
                        operation.Value.ParentPath == operation.Value.LibraryRootPath
                        && parent.Id == operation.Value.LibraryRootId;
                    if (
                        parent.Id != expectedParentId
                        || !string.Equals(
                            parent.Path,
                            operation.Value.ParentPath,
                            StringComparison.Ordinal
                        )
                    )
                        throw new EvaluationBlockedException("Parent folder identity differs.");
                    bool finalParent = operation.Value.ProbeKind == "FinalParent";
                    if (!isLibraryRoot)
                    {
                        string ancestor = parent.Path.Substring(0, parent.Path.LastIndexOf('/'));
                        if (ancestor != operation.Value.LibraryRootPath)
                        {
                            operation.Value.AncestorPath = ancestor;
                            return Probe(
                                operation,
                                dispatcher.Value,
                                library,
                                finalParent ? "FinalAncestor" : "Ancestor"
                            );
                        }
                    }
                    if (finalParent)
                    {
                        operation.Value.Status = "Verified";
                        operation.Value.ProbeId = Guid.Empty;
                        store.Save(operation);
                        return new WorkerResult
                        {
                            Status = "Verified",
                            Key = request.Key,
                            Token = request.Token,
                            PhysicalId = binding.PhysicalId,
                        };
                    }
                    return Probe(operation, dispatcher.Value, library, "Folder");
                case "Ancestor":
                case "FinalAncestor":
                    var ancestorFolder = SharePointObservations.Body<FolderObservation>(request);
                    bool finalAncestor = operation.Value.ProbeKind == "FinalAncestor";
                    if (
                        ancestorFolder.Id == Guid.Empty
                        || ancestorFolder.Path != operation.Value.AncestorPath
                        || !ancestorFolder.Path.StartsWith(
                            operation.Value.LibraryRootPath + "/",
                            StringComparison.Ordinal
                        )
                    )
                        throw new EvaluationBlockedException(
                            "Intermediate ancestor identity differs."
                        );
                    string nextAncestor = ancestorFolder.Path.Substring(
                        0,
                        ancestorFolder.Path.LastIndexOf('/')
                    );
                    if (nextAncestor != operation.Value.LibraryRootPath)
                    {
                        operation.Value.AncestorPath = nextAncestor;
                        return Probe(
                            operation,
                            dispatcher.Value,
                            library,
                            finalAncestor ? "FinalAncestor" : "Ancestor"
                        );
                    }
                    if (finalAncestor)
                    {
                        operation.Value.Status = "Verified";
                        operation.Value.ProbeId = Guid.Empty;
                        store.Save(operation);
                        return new WorkerResult
                        {
                            Status = "Verified",
                            Key = request.Key,
                            Token = request.Token,
                            PhysicalId = binding.PhysicalId,
                        };
                    }
                    return Probe(operation, dispatcher.Value, library, "Folder");
                case "ConflictFile":
                    if (request.HttpStatus == 404)
                        return Block(operation, dispatcher, "CreateConflictNoFolder");
                    var file = SharePointObservations.Body<FolderObservation>(request);
                    if (
                        file.Id == Guid.Empty
                        || file.Path != operation.Value.ParentPath + "/" + binding.Candidate
                    )
                        throw new EvaluationBlockedException("Conflicting file identity differs.");
                    return Block(operation, dispatcher, "FileAtExpectedFolderPath");
                case "Folder":
                    return ObserveFolder(request, operation, binding, dispatcher, library);
                default:
                    throw new EvaluationBlockedException("Unknown observation phase.");
            }
        }
        catch (EvaluationBlockedException)
        {
            return Block(operation, dispatcher, "ObservationMismatch");
        }
    }

    private WorkerResult ObserveFolder(
        WorkerRequest request,
        StoredRow<OperationDocument> operation,
        FolderStep binding,
        StoredRow<DispatcherDocument> dispatcher,
        WorkerLibrary library
    )
    {
        // A child folder is always below the library root; a root folder may sit at it.
        var item = SharePointObservations.Find(
            request,
            operation.Value.ParentPath + "/" + binding.Candidate,
            binding.ParentBinding == null
        );
        if (item == null)
        {
            if (binding.PhysicalId != Guid.Empty || binding.LocationId != Guid.Empty)
                return Block(operation, dispatcher, "EstablishedFolderMissing");
            if (
                operation.Value.ExternalResponseKnown
                && (
                    operation.Value.NameConflict
                    || operation.Value.ErrorCode == "CreateNameConflict"
                )
            )
                return Probe(operation, dispatcher.Value, library, "ConflictFile");
            if (operation.Value.Reprobe)
            {
                // The re-read after an unknown create found no folder: that create made nothing,
                // so it is no longer outstanding and the folder is created once now.
                operation.Value.Reprobe = false;
                operation.Value.ExternalSubmitted = false;
                operation.Value.ExternalResponseKnown = false;
            }
            if (operation.Value.ExternalSubmitted)
                return new WorkerResult
                {
                    Status = "Quarantined",
                    Key = request.Key,
                    Token = request.Token,
                };
            operation.Value.AbsenceVerified = true;
            operation.Value.Status = "ReadyToCreate";
            operation.Value.ProbeId = Guid.Empty;
            store.Save(operation);
            return new WorkerResult
            {
                Status = "ReadyToCreate",
                Key = request.Key,
                Token = request.Token,
            };
        }
        if (
            binding.PhysicalId != Guid.Empty
            && (binding.PhysicalId != item.Id || binding.PhysicalPath != item.Path)
        )
            return Block(operation, dispatcher, "FolderChangedDuringRequest");
        if (operation.Value.ExternalSubmitted && !operation.Value.ExternalResponseKnown)
            return new WorkerResult
            {
                Status = "Quarantined",
                Key = request.Key,
                Token = request.Token,
            };
        // A folder at the path is adopted, so an unknown create is never repeated.
        operation.Value.Reprobe = false;
        operation.Value.NameConflict = false;
        binding.PhysicalId = item.Id;
        binding.PhysicalPath = item.Path;
        binding.Candidate = item.Name;
        binding.Status = "Verified";
        operation.Value.Status = "NeedsFinalPolicy";
        operation.Value.AbsenceVerified = false;
        return Probe(operation, dispatcher.Value, library, "FinalParent");
    }

    private WorkerResult PrepareCreate(WorkerRequest request)
    {
        var dispatcher = Assert(request);
        var operation = store.Require<OperationDocument>("asx_operation", request.Key);
        var binding = operation.Value.Folder;
        var library = Current(operation.Value, binding);
        if (
            operation.Value.Status != "ReadyToCreate"
            || !operation.Value.AbsenceVerified
            || operation.Value.ExternalSubmitted
            || binding.PhysicalId != Guid.Empty
        )
            throw new EvaluationBlockedException(
                "Create requires current verified absence and no unknown previous write."
            );
        var http = SharePointRequests.CreateFolder(
            library.Target,
            operation.Value.ParentPath!,
            binding.Candidate
        );
        operation.Value.Status = "ExternalUnknown";
        operation.Value.ExternalSubmitted = true;
        operation.Value.ExternalResponseKnown = false;
        operation.Value.AbsenceVerified = false;
        store.Save(operation);
        Audit(request.Key, request.RunId, "ExternalPrepared");
        return new WorkerResult
        {
            Status = "Create",
            Key = request.Key,
            Token = dispatcher.Value.Token,
            SiteUrl = library.Target.Web.AbsoluteUri.TrimEnd('/'),
            Http = http,
        };
    }

    private WorkerResult CreateResponse(WorkerRequest request)
    {
        var dispatcher = Assert(request);
        var operation = store.Require<OperationDocument>("asx_operation", request.Key);
        var binding = operation.Value.Folder;
        var library = Current(operation.Value, binding);
        if (
            !operation.Value.ExternalSubmitted
            || operation.Value.ExternalResponseKnown
            || operation.Value.Status != "ExternalUnknown"
        )
            throw new EvaluationBlockedException(
                "No outstanding prepared create response is expected."
            );
        if (request.HttpStatus == 429)
        {
            // SharePoint does not execute a throttled request, so no folder was created. The
            // write is no longer outstanding; after the wait the job re-reads and creates again.
            operation.Value.ExternalSubmitted = false;
            operation.Value.ExternalResponseKnown = false;
            operation.Value.AbsenceVerified = false;
            Audit(request.Key, request.RunId, "ExternalThrottled");
            return ScheduleRetry(operation, dispatcher, request.RetryAfter, 429, null);
        }
        if (request.HttpStatus == 0 || request.HttpStatus == 408 || request.HttpStatus >= 500)
            return new WorkerResult
            {
                Status = "Quarantined",
                Key = request.Key,
                Token = request.Token,
            };
        operation.Value.ExternalResponseKnown = true;
        if (request.HttpStatus == 409)
        {
            operation.Value.ErrorCode = "CreateNameConflict";
            operation.Value.NameConflict = true;
            return Probe(operation, dispatcher.Value, library, "Folder");
        }
        try
        {
            SharePointObservations.CreateSucceeded(request);
        }
        catch (EvaluationBlockedException)
        {
            return Block(operation, dispatcher, "CreateRejected");
        }
        return Probe(operation, dispatcher.Value, library, "Folder");
    }

    private WorkerResult Complete(WorkerRequest request)
    {
        var operation = store.Require<OperationDocument>("asx_operation", request.Key);
        if (operation.Value.Status == "Applied")
            return Done(operation);
        if (
            operation.Value.NextAttemptUtc.HasValue
            && operation.Value.NextAttemptUtc.Value > clock()
        )
            return new WorkerResult { Status = "RetryWait", Key = request.Key };
        var dispatcher = Assert(request);
        var binding = operation.Value.Folder;
        var library = Current(operation.Value, binding);
        if (
            operation.Value.Status != "Verified"
            || binding.Status != "Verified"
            || binding.PhysicalId == Guid.Empty
            || (operation.Value.ExternalSubmitted && !operation.Value.ExternalResponseKnown)
        )
            throw new EvaluationBlockedException(
                "Independent identity/policy observation required for completion."
            );
        bool retired = Retired(binding.Table, binding.RecordId);
        if (!retired)
            binding.LocationId = new NativeLocations(service).Complete(
                binding,
                library.NativeParentId,
                library.EntryUrl,
                library.NativeSiteId
            );
        binding.Status = retired ? "DecommissionReview" : "Applied";
        if (!retired && operation.Value.Cursor + 1 < operation.Value.Folders.Length)
        {
            operation.Value.Cursor++;
            operation.Value.NameConflict = false;
            operation.Value.ExternalSubmitted = false;
            operation.Value.ExternalResponseKnown = false;
            operation.Value.ParentPath = null;
            operation.Value.Status = "Pending";
            store.Save(operation);
            Release(dispatcher);
            return new WorkerResult { Status = "Pending", Key = request.Key };
        }
        operation.Value.Status = "Applied";
        operation.Value.CompletedUtc = clock();
        store.Save(operation);
        Release(dispatcher);
        Audit(request.Key, request.RunId, "Applied");
        return Done(store.Require<OperationDocument>("asx_operation", request.Key));
    }

    /// <summary>
    /// Waits after a temporary failure (a throttled or failed read, or a temporary Dataverse
    /// failure reported by the flow). There is no attempt cap; the notice shows the attempt
    /// count and the last cause, and an operator can still Retry or Cancel the waiting work.
    /// </summary>
    private WorkerResult ScheduleRetry(
        StoredRow<OperationDocument> operation,
        StoredRow<DispatcherDocument> dispatcher,
        string? retryAfter,
        int? statusCode,
        string? errorCode
    )
    {
        if (operation.Value.ExternalSubmitted && !operation.Value.ExternalResponseKnown)
            return Block(operation, dispatcher, "ReadAfterUnknownWrite");
        DocumentStore.Wait(operation.Value, statusCode, errorCode, retryAfter, clock());
        store.Save(operation);
        Release(dispatcher);
        return new WorkerResult { Status = "RetryWait", Key = operation.Value.Key };
    }

    // The first wait. Dataverse service protection measures a 5-minute window and SharePoint
    // throttling clears in seconds to minutes, so 30 seconds lets a short spike pass without
    // adding visible delay to work that only hit a blip.
    public const int FirstRetrySeconds = 30;

    // The longest single wait, including an honored Retry-After. Re-checking every 15 minutes
    // costs one request and never strands work behind a long or malformed server hint.
    public const int MaxRetrySeconds = 900;

    /// <summary>
    /// The next attempt time: exponential from 30 seconds, each interval capped at 15 minutes.
    /// A usable Retry-After (seconds or an HTTP date) lengthens the wait up to that same cap; an
    /// unparseable, negative or very long one never throws and falls back to the backoff.
    /// </summary>
    public static DateTime RetryAt(DateTime now, int attempt, string? retryAfter)
    {
        // Clamping the exponent only keeps Math.Pow finite; the 15-minute cap applies first.
        double backoff = FirstRetrySeconds * Math.Pow(2, Math.Max(0, Math.Min(attempt, 31) - 1));
        var earliest = now.AddSeconds(Math.Min(MaxRetrySeconds, backoff));
        var cap = now.AddSeconds(MaxRetrySeconds);
        if (string.IsNullOrWhiteSpace(retryAfter))
            return earliest;
        var text = retryAfter!.Trim();
        DateTime requested;
        if (long.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var seconds))
            requested = seconds >= MaxRetrySeconds ? cap : now.AddSeconds(seconds);
        else if (
            DateTimeOffset.TryParseExact(
                text,
                "r",
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal,
                out var date
            )
        )
            requested = date.UtcDateTime > cap ? cap : date.UtcDateTime;
        else
            return earliest;
        return requested > earliest ? requested : earliest;
    }

    private WorkerResult Retry(WorkerRequest request)
    {
        var operation = store.Require<OperationDocument>("asx_operation", request.Key);
        bool expired = store.ReleaseExpired(request.Key, clock(), "Retry");
        if (operation.Value.Status == "Applied")
            return Done(operation);
        if (
            !expired
            && operation.Value.Status != "Blocked"
            && operation.Value.Status != "RetryWait"
            && !(operation.Value.ExternalSubmitted && !operation.Value.ExternalResponseKnown)
        )
            throw new EvaluationBlockedException(
                "Only blocked, waiting or interrupted work can be explicitly retried."
            );
        if (operation.Value.Status == "Superseded" || operation.Value.Status == "Cancelled")
            throw new EvaluationBlockedException("Cancelled or superseded work cannot retry.");
        if (operation.Value.ExternalSubmitted && !operation.Value.ExternalResponseKnown)
        {
            // The next claim reads back before writing: it adopts a folder the earlier create
            // made, or creates it once.
            operation.Value.Reprobe = true;
            operation.Value.ExternalResponseKnown = true;
        }
        operation.Value.Status = "Pending";
        operation.Value.NextAttemptUtc = null;
        operation.Value.RetryCount = 0;
        operation.Value.ProbeId = Guid.Empty;
        operation.Value.ErrorCode = null;
        store.Save(operation);
        return new WorkerResult { Status = "Pending", Key = request.Key };
    }

    private WorkerResult Cancel(WorkerRequest request)
    {
        var operation = store.Require<OperationDocument>("asx_operation", request.Key);
        // Cancel works once no run holds a live claim. A folder an earlier create may have
        // made stays in SharePoint; Cancel never deletes content.
        store.ReleaseExpired(request.Key, clock(), "Cancel");
        if (operation.Value.Status == "Applied")
            return Done(operation);
        operation.Value.Status = "Cancelled";
        operation.Value.NextAttemptUtc = null;
        operation.Value.ProbeId = Guid.Empty;
        store.Save(operation);
        Audit(request.Key, "operator", "Cancel");
        return new WorkerResult { Status = "Cancelled", Key = request.Key };
    }

    private WorkerResult Replan(WorkerRequest request)
    {
        var result = Queue(request);
        if (result.Status == "Inactive")
            return result;
        var job = store.Require<OutboxDocument>("asx_outbox", result.Key);
        if (job.Value.Status != "Pending")
            return result;
        job.Value.Replan = true;
        store.Save(job);
        return result;
    }

    public const string PlanningFailedNotice =
        "Planning failed. Check that the site and library are approved, then use Retry in Blocked records.";

    /// <summary>
    /// Returns a Blocked outbox row to Pending so the dispatcher plans it again. Replanning
    /// re-reads the current template, record and catalog, so a row blocked for a reason that
    /// still holds simply blocks again.
    /// </summary>
    /// <param name="request">The outbox key to retry.</param>
    /// <returns>Pending when the row was retried; otherwise its unchanged status.</returns>
    private WorkerResult RetryOutbox(WorkerRequest request)
    {
        var job = store.Require<OutboxDocument>("asx_outbox", request.Key);
        // A row waiting after a temporary failure is Pending with a next attempt; Retry ends
        // the wait so it plans on the next dispatch.
        if (
            job.Value.Status == "Blocked"
            || (job.Value.Status == "Pending" && job.Value.NextAttemptUtc != null)
        )
        {
            job.Value.Status = "Pending";
            job.Value.Notices = Array.Empty<string>();
            job.Value.NextAttemptUtc = null;
            job.Value.Attempts = 0;
            store.Save(job);
        }
        return new WorkerResult
        {
            Status = job.Value.Status,
            Key = request.Key,
            Notices = job.Value.Notices,
        };
    }

    private WorkerResult FailOutbox(WorkerRequest request)
    {
        var job = store.Require<OutboxDocument>("asx_outbox", request.Key);
        if (job.Value.Status == "Pending")
        {
            if (TransientFailure.Is(request))
            {
                // Stays Pending; ListOutbox skips it until the next attempt is due.
                job.Value.Attempts++;
                job.Value.NextAttemptUtc = RetryAt(clock(), job.Value.Attempts, null);
                job.Value.Notices = new[] { TransientFailure.Notice(request, job.Value.Attempts) };
            }
            else
            {
                job.Value.Status = "Blocked";
                job.Value.NextAttemptUtc = null;
                job.Value.Notices = new[] { PlanningFailedNotice };
            }
            store.Save(job);
        }
        return new WorkerResult
        {
            Status = job.Value.Status,
            Key = request.Key,
            Notices = job.Value.Notices,
        };
    }

    private WorkerResult Fail(WorkerRequest request)
    {
        var dispatcher = Assert(request);
        var operation = store.Require<OperationDocument>("asx_operation", request.Key);
        if (TransientFailure.Is(request))
            return ScheduleRetry(
                operation,
                dispatcher,
                null,
                request.StatusCode,
                request.ErrorCode
            );
        return Block(operation, dispatcher, "WorkerFailed");
    }

    public WorkerResult PermitRecovery(WorkerRequest request, bool inTransaction)
    {
        if (!inTransaction)
            throw new EvaluationBlockedException("Recovery requires a transaction.");
        var dispatcher = store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        if (
            dispatcher.Value.RunId == null
            || dispatcher.Value.OperationKey != request.Key
            || dispatcher.Value.RunId != request.RunId
            || dispatcher.Value.Token != request.Token
            || dispatcher.Value.LeaseUntilUtc > clock()
            || string.IsNullOrWhiteSpace(request.Evidence)
            || request.Evidence!.Length > 500
        )
            throw new EvaluationBlockedException(
                "Expired exact writer identity and operator-confirmed termination/outstanding-call evidence required."
            );
        if (request.Key.StartsWith("librarycreate:", StringComparison.Ordinal))
        {
            var setup = store.Require<LibrarySetup>("asx_operation", request.Key);
            if (setup.Value.Mutation == "CreateLibrary" && setup.Value.ListId == Guid.Empty)
            {
                if (
                    string.IsNullOrWhiteSpace(request.ResponseBody)
                    || request.ResponseBody!.Length > 16000
                )
                    throw new EvaluationBlockedException(
                        "Recover the successful original create response from the terminated flow run. A same-name lookup is insufficient."
                    );
                var created = JsonWire
                    .Read<ODataEnvelope<CreatedLibrary>>(request.ResponseBody)
                    .Data;
                if (
                    created == null
                    || created.Id == Guid.Empty
                    || created.Title != setup.Value.Name
                )
                    throw new EvaluationBlockedException(
                        "Original creation response must identify the requested library by ID and title."
                    );
                setup.Value.ListId = created.Id;
                store.Save(setup);
                Audit(
                    request.Key,
                    request.RunId,
                    "RecoveredCreateResponse:" + DocumentStore.Hash(request.ResponseBody)
                );
                setup = store.Require<LibrarySetup>("asx_operation", request.Key);
            }
            // List the permitted setup again so the dispatcher resumes it.
            if (setup.Value.Status == "RecoveryRequired")
            {
                setup.Value.Status = "Pending";
                store.Save(setup);
            }
        }
        dispatcher.Value.HttpOutstanding = false;
        dispatcher.Value.RecoveryPermitted = true;
        dispatcher.Value.TerminationEvidence = request.Evidence;
        store.Save(dispatcher);
        Audit(request.Key, request.RunId, "OperatorRecoveryPermit");
        return new WorkerResult { Status = "RecoveryPermitted", Key = request.Key };
    }

    private WorkerResult? StopUnavailable(WorkerRequest request)
    {
        var operation = store.Require<OperationDocument>("asx_operation", request.Key);
        if (
            operation.Value.HistoryCompacted
            || operation.Value.ExternalSubmitted
            || operation.Value.Status == "Applied"
        )
            return null;
        var binding = operation.Value.Folder;
        if (TemplateLifecycle.Active(TemplateLifecycle.Find(service, binding.TemplateId), clock()))
            return null;
        var claim = store.Find<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        if (claim?.Value.OperationKey == request.Key && claim.Value.RunId != null)
        {
            // Only the holder may release an active writer. A different run cannot take it over.
            if (request.Command == "Claim" && !claim.Value.RecoveryPermitted)
                return new WorkerResult { Status = "Busy", Key = request.Key };
            if (request.Command != "Claim")
                claim = Assert(request);
            Release(claim);
        }
        operation.Value.Status = "Cancelled";
        operation.Value.ErrorCode = "TemplateUnavailable";
        store.Save(operation);
        return new WorkerResult
        {
            Status = "Cancelled",
            Key = request.Key,
            Notices = new[]
            {
                "Template is deleted, deactivated or outside its schedule. Existing SharePoint content is unchanged.",
            },
        };
    }

    private StoredRow<DispatcherDocument> Assert(WorkerRequest request)
    {
        var row = store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        if (
            request.Token == Guid.Empty
            || row.Value.Token != request.Token
            || row.Value.RunId != request.RunId
            || row.Value.OperationKey != request.Key
            || row.Value.LeaseUntilUtc <= clock()
            || row.Value.RecoveryPermitted
        )
            throw new EvaluationBlockedException("Stale or mismatched worker claim.");
        return row;
    }

    private bool MatchesSource(SourceVersion source)
    {
        return !string.IsNullOrEmpty(source.Version)
            && service.Retrieve(source.Table, source.Id, new ColumnSet(false)).RowVersion
                == source.Version;
    }

    public static string RetirementKey(string table, Guid id) =>
        "record-retired:" + table + ":" + id.ToString("N");

    private bool Retired(string table, Guid id) =>
        store.Find<OutboxDocument>("asx_outbox", RetirementKey(table, id)) != null;

    private WorkerLibrary Current(OperationDocument operation, FolderStep binding)
    {
        // A create already sent to SharePoint may finish (response, reads, completion) after its
        // table is removed, so its writer slot is released. Unsent work for the table stops here.
        if (!operation.ExternalSubmitted)
            Allowed(binding.Table);
        if (!operation.ExternalSubmitted && Retired(binding.Table, binding.RecordId))
            throw new EvaluationBlockedException(
                "Deleted record requires decommission review; no new folder write."
            );
        var selection = store.Find<RecordPlanDocument>(
            "asx_outbox",
            "recordplan:" + binding.TemplateId.ToString("N") + ":" + binding.RecordId.ToString("N")
        );
        if (
            !operation.ExternalSubmitted
            && selection != null
            && (
                selection.Value.RevisionId != operation.RevisionId
                || !selection.Value.Operations.Contains(operation.Key)
            )
        )
            throw new EvaluationBlockedException(
                "A newer record evaluation no longer selects this unsubmitted operation."
            );
        // Re-read on every step so a suspension stops work before the next write. The library's
        // policy generation is deliberately not compared: folders inherit the library's access.
        // A create already sent may finish after a suspension; nothing unsent proceeds.
        var library = catalog.Read(
            binding.LibraryId,
            requireApproved: !operation.ExternalSubmitted
        );
        if (library.EntryId != binding.EntryId)
            throw new EvaluationBlockedException("Approved destination entry changed.");
        return library;
    }

    private WorkerResult Probe(
        StoredRow<OperationDocument> operation,
        DispatcherDocument dispatcher,
        WorkerLibrary library,
        string kind
    )
    {
        var binding = operation.Value.Folder;
        HttpIntent http;
        switch (kind)
        {
            case "Library":
                http = new HttpIntent
                {
                    RelativeUri =
                        "_api/web/lists(guid'"
                        + library.Target.ListId
                        + "')?$select=Id,RootFolder/UniqueId,RootFolder/ServerRelativeUrl&$expand=RootFolder",
                };
                break;
            case "Ancestor":
            case "FinalAncestor":
                http = new HttpIntent
                {
                    RelativeUri =
                        "_api/web/GetFolderByServerRelativePath(decodedUrl='"
                        + Uri.EscapeDataString(operation.Value.AncestorPath!.Replace("'", "''"))
                        + "')?$select=UniqueId,ServerRelativeUrl",
                };
                break;
            case "FinalParent":
            case "Parent":
                http = new HttpIntent
                {
                    RelativeUri =
                        "_api/web/GetFolderByServerRelativePath(decodedUrl='"
                        + Uri.EscapeDataString(operation.Value.ParentPath!.Replace("'", "''"))
                        + "')?$select=UniqueId,ServerRelativeUrl",
                };
                break;
            case "ConflictFile":
                http = new HttpIntent
                {
                    RelativeUri =
                        "_api/web/GetFileByServerRelativePath(decodedUrl='"
                        + Uri.EscapeDataString(
                            (operation.Value.ParentPath + "/" + binding.Candidate).Replace(
                                "'",
                                "''"
                            )
                        )
                        + "')?$select=UniqueId,ServerRelativeUrl",
                };
                break;
            case "Folder":
                http = SharePointRequests.FindFolder(
                    library.Target,
                    operation.Value.ParentPath!,
                    binding.Candidate
                );
                break;
            default:
                throw new EvaluationBlockedException("Unknown probe type.");
        }
        operation.Value.ProbeId = Guid.NewGuid();
        operation.Value.ProbeKind = kind;
        operation.Value.AbsenceVerified = false;
        if (!operation.Value.ExternalSubmitted)
            operation.Value.Status = "Inspecting";
        store.Save(operation);
        return new WorkerResult
        {
            Status = "Read",
            Key = operation.Value.Key,
            Token = dispatcher.Token,
            ProbeId = operation.Value.ProbeId,
            ProbeKind = kind,
            SiteUrl = library.Target.Web.AbsoluteUri.TrimEnd('/'),
            Http = http,
        };
    }

    private WorkerResult Block(
        StoredRow<OperationDocument> operation,
        StoredRow<DispatcherDocument> dispatcher,
        string code
    )
    {
        if (operation.Value.ExternalSubmitted && !operation.Value.ExternalResponseKnown)
        {
            operation.Value.Status = "ExternalUnknown";
            operation.Value.ErrorCode = code;
            store.Save(operation);
            return new WorkerResult { Status = "Quarantined", Key = operation.Value.Key };
        }
        operation.Value.Status = "Blocked";
        operation.Value.ErrorCode = code;
        store.Save(operation);
        Release(dispatcher);
        return new WorkerResult { Status = "Blocked", Key = operation.Value.Key };
    }

    private void Release(StoredRow<DispatcherDocument> dispatcher)
    {
        dispatcher.Value.RunId = null;
        dispatcher.Value.OperationKey = null;
        dispatcher.Value.Token = Guid.Empty;
        dispatcher.Value.RecoveryPermitted = false;
        dispatcher.Value.Status = "Idle";
        store.Save(dispatcher);
    }

    private WorkerResult Done(StoredRow<OperationDocument> operation)
    {
        var binding = operation.Value.Folders.FirstOrDefault();
        return new WorkerResult
        {
            Status = "Applied",
            Key = operation.Value.Key,
            PhysicalId = binding?.PhysicalId ?? operation.Value.ResultPhysicalId,
            LocationId = binding?.LocationId ?? operation.Value.ResultLocationId,
        };
    }

    private void Allowed(string table)
    {
        if (allowedTables != null && !allowedTables.Contains(table, StringComparer.Ordinal))
            throw new EvaluationBlockedException(
                "Record source is outside the approved runtime scope."
            );
    }

    private static void ValidateRun(string run)
    {
        if (string.IsNullOrWhiteSpace(run) || run.Length > 150 || run.Any(char.IsControl))
            throw new EvaluationBlockedException("Bounded actual flow-run identity required.");
    }

    private void Audit(string operation, string run, string kind) =>
        store.Create(
            "asx_attempt",
            new AttemptDocument
            {
                Key = "attempt:" + Guid.NewGuid().ToString("N"),
                Status = "Recorded",
                OperationKey = operation,
                RunId = run,
                Event = kind,
                AtUtc = clock(),
            }
        );
}
