using System;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class LibrarySetup : OperationDocument
{
    [DataMember]
    public Guid SiteId { get; set; }

    [DataMember]
    public Guid NativeSiteId { get; set; }

    [DataMember]
    public Guid WebId { get; set; }

    [DataMember]
    public string WebUrl { get; set; } = "";

    [DataMember]
    public string Name { get; set; } = "";

    [DataMember]
    public Guid ListId { get; set; }

    [DataMember]
    public Guid CatalogId { get; set; }

    [DataMember]
    public string? PolicyOperation { get; set; }

    [DataMember]
    public int OwnerGroup { get; set; }

    [DataMember]
    public int OwnerRole { get; set; }

    [DataMember]
    public PolicyRole ReadRole { get; set; } = new PolicyRole();

    [DataMember]
    public PolicyRole ContributeRole { get; set; } = new PolicyRole();

    [DataMember]
    public PolicyEntry[] Entries { get; set; } = Array.Empty<PolicyEntry>();

    /// <summary>The admin's consent to the initial team's broader access, applied with it later.</summary>
    [DataMember]
    public bool AcknowledgeBroaderAccess { get; set; }

    [DataMember]
    public CatalogLibraryObservation? Library { get; set; }

    [DataMember]
    public string Mutation { get; set; } = "";

    [DataMember]
    public HttpIntent? Intent { get; set; }

    /// <summary>
    /// For a prepared write: false until the shared connection permits sending it, then true.
    /// A write never permitted was never sent, so it has no unknown outcome in SharePoint.
    /// Null in setups stored before this was recorded (see LibraryProvisioning.NeverSent).
    /// </summary>
    [DataMember]
    public bool? WritePermitted { get; set; }

    /// <summary>
    /// True while Documents looks SharePoint up for a create whose answer was lost: the run only
    /// reads (ReconcileTitle, ReconcileUrl) and never prepares a write (spec 6.8).
    /// </summary>
    [DataMember]
    public bool Reconcile { get; set; }

    /// <summary>When SharePoint last answered 404 to the read of this title, before the create.</summary>
    [DataMember]
    public DateTime? AbsentUtc { get; set; }

    /// <summary>The lookup's finding; null until a lookup starts.</summary>
    [DataMember]
    public LibraryRecovery? Recovery { get; set; }

    /// <summary>The list the lookup found by title, kept between its two reads.</summary>
    [DataMember]
    public ReconcileObservation? TitleHit { get; set; }
}

[DataContract]
public sealed class NewLibraryBody
{
    [DataMember(Name = "__metadata")]
    public LibraryMetadata Metadata { get; set; } = new LibraryMetadata();

    [DataMember]
    public string Title { get; set; } = "";

    [DataMember]
    public int BaseTemplate { get; set; } = 101;

    [DataMember]
    public bool AllowContentTypes { get; set; } = true;
}

[DataContract]
public sealed class LibraryMetadata
{
    [DataMember(Name = "type")]
    public string Type { get; set; } = "SP.List";
}

[DataContract]
public sealed class CreatedLibrary
{
    [DataMember]
    public Guid Id { get; set; }

    [DataMember]
    public string Title { get; set; } = "";
}

public sealed class LibraryProvisioning
{
    private readonly IOrganizationService service;
    private readonly DocumentStore store;
    private readonly Func<DateTime> clock;

    public LibraryProvisioning(IOrganizationService service, Func<DateTime>? clock = null)
    {
        this.service = service;
        store = new DocumentStore(service);
        this.clock = clock ?? (() => DateTime.UtcNow);
    }

    /// <summary>
    /// Validates the site and requested access, then persists an idempotent library setup operation.
    /// </summary>
    /// <param name="request">The site, library name, initial team access, and stable request identity.</param>
    /// <returns>The existing or newly queued library setup result.</returns>
    // The fields of the library read; before the create the library is read by its title.
    private const string LibrarySelect =
        "$select=Id,HasUniqueRoleAssignments,RootFolder/UniqueId,RootFolder/ServerRelativeUrl&$expand=RootFolder";

    /// <summary>The refusal of a recovery choice the lookup's finding does not offer.</summary>
    public const string ChoiceNotOffered =
        "Documents doesn't offer that choice for what it found. Choose one of the choices shown on the setup, check again, or cancel the setup.";

    /// <summary>Why a library the catalog gives to another site, or removed, is not used.</summary>
    internal const string CatalogConflictNotice =
        "The library was created in SharePoint, but the Documents catalog already has a different entry for it (another site, or removed). Nothing was changed. Remove that entry if it is wrong, add the library as an existing library, then cancel this setup.";

    /// <summary>The query string of the read by title, which carries the title as an alias.</summary>
    private static string ByTitle(string name) =>
        Domain.SharePointAddress.Alias(name) + "&" + LibrarySelect;

    /// <summary>The query string of the lookup's read by title (spec 6.8).</summary>
    private static string LookupByTitle(string name) =>
        Domain.SharePointAddress.Alias(name) + "&" + LibraryReconcile.Select;

    /// <summary>The query string of the lookup's read by the address SharePoint derives from the title.</summary>
    private static string LookupByAddress(string webUrl, string name) =>
        Domain.SharePointAddress.Alias(LibraryReconcile.ExpectedUrl(webUrl, name))
        + "&"
        + LibraryReconcile.Select;

    /// <summary>
    /// The length of the longest query string a setup of this name on this site can send: the
    /// read by title before the create, and the two reads of a lookup after a lost create.
    /// </summary>
    private static int LongestRead(string webUrl, string name) =>
        Math.Max(
            ByTitle(name).Length,
            Math.Max(LookupByTitle(name).Length, LookupByAddress(webUrl, name).Length)
        );

    public CatalogResult Queue(CatalogRequest request)
    {
        Domain.FolderNames.Validate(request.Name);
        if (request.SiteId == Guid.Empty || request.RequestId == Guid.Empty)
            throw new EvaluationBlockedException("Select a site and provide a request identity.");
        if (
            request.Entries == null
            || request.Entries.Length > Domain.Bounds.TeamEntries
            || request.Entries.Any(e =>
                e.TeamId == Guid.Empty || !new[] { "Read", "Contribute" }.Contains(e.Access)
            )
            || request.Entries.Select(e => e.TeamId).Distinct().Count() != request.Entries.Length
        )
            throw new EvaluationBlockedException(
                "Select distinct teams with Read or Contribute access."
            );
        var teams = TeamDirectory.Read(service, request.Entries.Select(e => e.TeamId));
        foreach (var entry in request.Entries)
            TeamPrincipal.Validate(
                teams.TryGetValue(entry.TeamId, out var team)
                    ? team
                    : throw new EvaluationBlockedException(TeamDirectory.DeletedRefusal),
                request.AcknowledgeBroaderAccess
            );
        var site = service.Retrieve(
            "asx_site",
            request.SiteId,
            new ColumnSet("asx_nativeid", "asx_webid", "asx_url", "asx_approved")
        );
        if (!site.GetAttributeValue<bool>("asx_approved"))
            throw new EvaluationBlockedException(
                "Site validation must complete before creating a library."
            );
        // The setup reads the library by its title before creating it, and after a lost create
        // its lookup reads it by title and by address. A name too long for any of those reads,
        // once escaped into the address, is refused now with its reason rather than failing
        // later (Domain.SharePointAddress).
        int address = LongestRead(TemplateStore.Text(site, "asx_url"), request.Name);
        if (address > Domain.SharePointAddress.MaxQueryString)
            throw new EvaluationBlockedException(
                "This library name is too long for the HTTP connector: written into a SharePoint address it would be "
                    + address.ToString("N0", System.Globalization.CultureInfo.InvariantCulture)
                    + " characters, and the HTTP connector accepts "
                    + Domain.SharePointAddress.MaxQueryString.ToString(
                        "N0",
                        System.Globalization.CultureInfo.InvariantCulture
                    )
                    + ". Use a shorter name."
            );
        string key =
            "librarycreate:"
            + DocumentStore.Hash(
                request.SiteId.ToString("N") + ":" + request.Name.ToUpperInvariant()
            );
        var old = store.Find<LibrarySetup>("asx_operation", key);
        if (old != null && old.Value.Status == "Cancelled")
        {
            // A cancelled setup is started again: fresh reads decide everything, so a library
            // an earlier attempt may have made is found by them, never created twice.
            var again = old.Value;
            Unsend(again);
            again.Entries = request.Entries;
            again.AcknowledgeBroaderAccess = request.AcknowledgeBroaderAccess;
            again.ErrorCode = null;
            again.RetryCount = 0;
            store.Save(old);
            return Result(store.Require<LibrarySetup>("asx_operation", key));
        }
        if (old != null)
        {
            if (JsonWire.Write(old.Value.Entries) != JsonWire.Write(request.Entries))
                throw new EvaluationBlockedException(
                    "Library creation is already recorded with different access. Open its current status."
                );
            return Result(old);
        }
        store.Create(
            "asx_operation",
            new LibrarySetup
            {
                Key = key,
                SiteId = site.Id,
                NativeSiteId = site.GetAttributeValue<EntityReference>("asx_nativeid").Id,
                WebId = Guid.Parse(TemplateStore.Text(site, "asx_webid")),
                WebUrl = TemplateStore.Text(site, "asx_url"),
                Name = request.Name,
                Entries = request.Entries,
                AcknowledgeBroaderAccess = request.AcknowledgeBroaderAccess,
            }
        );
        return Result(store.Require<LibrarySetup>("asx_operation", key));
    }

    public CatalogResult Inspect(string key) =>
        Result(PickUp(store.Require<LibrarySetup>("asx_operation", key)));

    /// <summary>
    /// Retry or Cancel of a library setup from its card in Sites, through the catalog API, so a
    /// Documents Security Administrator needs no Operator role. The operator rules apply
    /// unchanged (see Manage): Retry of a create that may have reached SharePoint stays
    /// refused, and Cancel deletes nothing in SharePoint. Once the library is created, the
    /// setup waits on its first access run, so Retry and Cancel act on that run.
    /// </summary>
    /// <param name="key">The library setup key.</param>
    /// <param name="retry">True for Retry, false for Cancel.</param>
    /// <returns>The setup's current status, with what the action left in place.</returns>
    public CatalogResult ManageSetup(string key, bool retry)
    {
        if (!key.StartsWith("librarycreate:", StringComparison.Ordinal))
            throw new EvaluationBlockedException("Select a library setup.");
        var op = store.Require<LibrarySetup>("asx_operation", key);
        var command = new WorkerRequest { Command = retry ? "Retry" : "Cancel", Key = key };
        WorkerResult done;
        if (op.Value.Status == "AccessPending" && op.Value.PolicyOperation != null)
        {
            command.Key = op.Value.PolicyOperation;
            done = new SecurityWorker(service, clock).Execute(command, true);
        }
        else
            done = Manage(command, op);
        var result = Result(store.Require<LibrarySetup>("asx_operation", key));
        result.Notices = done.Notices;
        return result;
    }

    private CatalogResult Result(StoredRow<LibrarySetup> op)
    {
        string status = op.Value.Status == "Applied" ? "Ready" : op.Value.Status;
        string? issue = op.Value.ErrorCode,
            recovery = null;
        if (status == "AccessPending" && op.Value.PolicyOperation != null)
        {
            var policy = store
                .Require<SecurityOperation>("asx_operation", op.Value.PolicyOperation)
                .Value;
            status =
                policy.Status == "Applied" ? "Ready"
                : new[] { "Blocked", "Quarantined", "ExternalUnknown" }.Contains(policy.Status)
                    ? policy.Status
                : "Applying access";
            issue = policy.ErrorCode;
            recovery = op.Value.PolicyOperation;
        }
        return new CatalogResult
        {
            Key = op.Value.Key,
            Status = status,
            RowVersion = op.Row.RowVersion,
            CatalogId = op.Value.CatalogId,
            Issue = issue,
            RecoveryKey = recovery,
            Recovery = op.Value.Recovery,
        };
    }

    private string List(LibrarySetup op) => "_api/web/lists(guid'" + op.ListId.ToString("D") + "')";

    /// <summary>
    /// True when the setup's prepared write was never permitted, so SharePoint never received
    /// it: a pause, a lost Prepare response or a run that ended before its permit. A setup
    /// stored before the permit was recorded (null) counts as possibly sent: nothing stored by
    /// an earlier release tells the two apart, so it is looked up in SharePoint, and Cancel.
    /// </summary>
    /// <param name="op">The library setup.</param>
    internal static bool NeverSent(LibrarySetup op) =>
        op.ExternalSubmitted && !op.ExternalResponseKnown && op.WritePermitted == false;

    /// <summary>
    /// A create that may have reached SharePoint and whose answer was lost starts the SharePoint
    /// lookup, due when the lease of the run that sent it ends; any other unknown write goes to
    /// RecoveryRequired for Retry. Either way the site's writer is released (spec 6.8). Nothing
    /// on the site depends on that write: folder work and access runs of a new library start
    /// only after it is registered, and the setup reads SharePoint before any further write.
    /// </summary>
    internal static void AwaitRecovery(
        DocumentStore store,
        StoredRow<LibrarySetup> op,
        StoredRow<DispatcherDocument>? claim
    )
    {
        DateTime? leaseEnd = null;
        if (HeldBy(op, claim) is { } held)
        {
            leaseEnd = held.Value.LeaseUntilUtc;
            ReleaseHeld(store, held);
        }
        if (LostCreate(op.Value))
            StartLookup(op.Value, leaseEnd);
        else
            op.Value.Status = "RecoveryRequired";
        store.Save(op);
    }

    /// <summary>A create that may have reached SharePoint, whose answer was lost, for a library not yet known.</summary>
    internal static bool LostCreate(LibrarySetup op) =>
        op.ExternalSubmitted
        && !op.ExternalResponseKnown
        && !NeverSent(op)
        && op.Mutation == "CreateLibrary"
        && op.ListId == Guid.Empty;

    private static void StartLookup(LibrarySetup op, DateTime? notBefore)
    {
        op.Reconcile = true;
        op.Status = "Reconciling";
        op.Recovery = LibraryReconcile.Checking();
        op.TitleHit = null;
        op.ProbeId = Guid.Empty;
        op.RetryCount = 0;
        op.NextAttemptUtc = notBefore;
        op.ErrorCode =
            "SharePoint didn't confirm the library creation. Documents is checking SharePoint.";
    }

    private static void ReleaseHeld(DocumentStore store, StoredRow<DispatcherDocument> claim)
    {
        claim.Value.HttpOutstanding = false;
        claim.Value.OperationKey = null;
        claim.Value.RunId = null;
        claim.Value.Token = Guid.Empty;
        claim.Value.Status = "Idle";
        store.Save(claim);
    }

    /// <summary>
    /// The site's writer claim when this setup's own run still holds it (the claim names the setup
    /// and has a run); otherwise null. The only place that rule is written: AwaitRecovery,
    /// Settled and LookupClaimFailed all ask it.
    /// </summary>
    private static StoredRow<DispatcherDocument>? HeldBy(
        StoredRow<LibrarySetup> op,
        StoredRow<DispatcherDocument>? claim
    ) =>
        claim != null && claim.Value.OperationKey == op.Value.Key && claim.Value.RunId != null
            ? claim
            : null;

    /// <summary>Releases the site's writer if this setup's own run still holds it.</summary>
    private static void ReleaseIfHeld(
        DocumentStore store,
        StoredRow<LibrarySetup> op,
        StoredRow<DispatcherDocument>? claim
    )
    {
        if (HeldBy(op, claim) is { } held)
            ReleaseHeld(store, held);
    }

    /// <summary>The finding for a lookup read that failed for good: the failure's cause, as of now.</summary>
    private static LibraryRecovery LookupFailed(WorkerRequest failure, DateTime now) =>
        LibraryReconcile.ReadFailed(
            TransientFailure.Cause(failure.StatusCode, failure.ErrorCode, failure.Error),
            now
        );

    /// <summary>
    /// Holds a setup while its site is suspended: it waits in RetryWait with a notice, its next
    /// check backs off (at most 15 minutes apart), so it takes no dispatch slot meanwhile, and it
    /// resumes by itself once the site is approved again. A create that may have reached
    /// SharePoint is looked up instead. Null when the site is not suspended.
    /// </summary>
    private WorkerResult? Suspended(WorkerRequest request, StoredRow<LibrarySetup> op)
    {
        if (WorkCoordination.StopsSite(service, op.Value.SiteId) != "suspended")
            return null;
        var claim = store.Find<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        // Another run holding the site's writer is no reason to stay listed: this setup is not
        // claimed, so it waits with the same backoff and takes no dispatch slot meanwhile.
        bool held = claim?.Value.RunId != null && claim.Value.OperationKey == request.Key;
        if (held && claim!.Value.LeaseUntilUtc > clock())
            return new WorkerResult { Status = "Quarantined", Key = request.Key };
        if (NeverSent(op.Value))
            Unsend(op.Value);
        if (op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown)
            return null;
        if (held)
        {
            claim!.Value.HttpOutstanding = false;
            Release(claim);
        }
        op.Value.RetryCount++;
        op.Value.Status = "RetryWait";
        op.Value.NextAttemptUtc = WorkerCoordinator.RetryAt(clock(), op.Value.RetryCount, null);
        op.Value.ProbeId = Guid.Empty;
        op.Value.ErrorCode =
            "Waiting: the site is suspended. This library setup resumes by itself once the site is approved again; check "
            + op.Value.RetryCount.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + ".";
        store.Save(op);
        return new WorkerResult
        {
            Status = "RetryWait",
            Key = request.Key,
            Notices = new[] { op.Value.ErrorCode },
        };
    }

    /// <summary>Puts a write that was never sent back to Pending, to be read and prepared again.</summary>
    internal static void Unsend(LibrarySetup op)
    {
        op.Intent = null;
        op.Mutation = "";
        op.WritePermitted = null;
        WorkCoordination.Unsend(op);
    }

    /// <summary>
    /// The operator's Retry or Cancel, once no run holds a live claim. A write that was never
    /// sent is simply prepared again. A library create that may have been sent is looked up in
    /// SharePoint and resolved by the admin's choice, never by Retry; Cancel always works and
    /// deletes nothing in SharePoint.
    /// </summary>
    private WorkerResult Manage(WorkerRequest request, StoredRow<LibrarySetup> op)
    {
        bool retry = request.Command == "Retry";
        if (op.Value.Status == "Cancelled" && !retry)
            return new WorkerResult { Key = request.Key, Status = op.Value.Status };
        if (
            op.Value.Status == "Applied"
            || op.Value.Status == "AccessPending"
            || op.Value.Status == "Cancelled"
        )
            throw new EvaluationBlockedException(
                op.Value.Status == "Cancelled"
                    ? "Cancelled setup cannot retry. Create the library again to start over."
                    : "The library is created; its access run is managed with the library."
            );
        bool unknown = op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown;
        bool neverSent = NeverSent(op.Value);
        bool sentCreate =
            unknown
            && !neverSent
            && op.Value.Mutation == "CreateLibrary"
            && op.Value.ListId == Guid.Empty;
        if (retry && sentCreate)
            throw new EvaluationBlockedException(
                "Documents is checking SharePoint for this library. Use the choice it offers on the setup, or cancel the setup. Cancel deletes nothing in SharePoint."
            );
        bool released = store.ReleaseExpired(request.Key, clock(), request.Command);
        if (
            retry
            && !released
            && op.Value.Status != "Blocked"
            && op.Value.Status != "RetryWait"
            && op.Value.Status != "RecoveryRequired"
        )
            throw new EvaluationBlockedException("Only interrupted setup can retry.");
        op = store.Require<LibrarySetup>("asx_operation", request.Key);
        if (neverSent)
            Unsend(op.Value);
        else if (unknown && retry)
            // A boundary or owner-access write: the next run reads the library before any write.
            op.Value.ExternalResponseKnown = true;
        var notices = Array.Empty<string>();
        if (retry)
        {
            op.Value.Status = "Pending";
            op.Value.ErrorCode = null;
        }
        else
        {
            op.Value.Status = "Cancelled";
            // A lookup in progress stops with the setup; nothing it found is kept or changed, and
            // its "checking" text or finding sentence is no longer the setup's issue.
            if (op.Value.Reconcile || op.Value.Recovery != null)
                op.Value.ErrorCode = null;
            op.Value.Reconcile = false;
            op.Value.TitleHit = null;
            op.Value.Recovery = null;
            notices = new[]
            {
                sentCreate
                    ? "Cancelled. Nothing in SharePoint was deleted. If SharePoint created the library, add it as an existing library."
                    : "Cancelled. Nothing in SharePoint was deleted.",
            };
        }
        op.Value.NextAttemptUtc = null;
        op.Value.RetryCount = 0;
        op.Value.ProbeId = Guid.Empty;
        store.Save(op);
        return new WorkerResult
        {
            Key = request.Key,
            Status = op.Value.Status,
            Notices = notices,
        };
    }

    /// <summary>
    /// Advances library creation, permission setup, and catalog registration from persisted observations.
    /// </summary>
    /// <param name="request">The worker command, operation identity, and command-specific inputs.</param>
    /// <param name="transaction">Whether execution is inside a Dataverse transaction; must be true.</param>
    /// <returns>The step status and any HTTP intent or continuation information.</returns>
    public WorkerResult Execute(WorkerRequest request, bool transaction)
    {
        if (!transaction)
            throw new EvaluationBlockedException("Library setup requires a transaction.");
        var op = store.Require<LibrarySetup>("asx_operation", request.Key);
        if (request.Command == "FailUnclaimed")
            return store.FailUnclaimed<LibrarySetup>(request.Key, request, clock());
        if (request.Command == "Retry" || request.Command == "Cancel")
            return Manage(request, op);
        if (request.Command == "Claim")
        {
            if (
                !WorkCoordination.HasCapacity(
                    service,
                    WorkCoordination.Operation(service, request.Key),
                    clock()
                )
            )
                return new WorkerResult { Status = "Busy", Key = request.Key };
            if (
                new[] { "AccessPending", "Applied", "Cancelled", "Blocked" }.Contains(
                    op.Value.Status
                )
            )
                return new WorkerResult { Status = op.Value.Status, Key = request.Key };
            // A lost write waits for the admin's choice on the lookup's finding, Retry or Cancel;
            // it never starts again by itself, so the create is never sent twice.
            if (op.Value.Status == "RecoveryRequired")
                return new WorkerResult { Status = "Quarantined", Key = request.Key };
            if (op.Value.NextAttemptUtc > clock())
                return new WorkerResult { Status = "RetryWait", Key = request.Key };
            if (
                string.IsNullOrWhiteSpace(request.RunId)
                || request.RunId.Length > 150
                || request.RunId.Any(char.IsControl)
            )
                throw new EvaluationBlockedException("Actual worker run identity required.");
            var held = Suspended(request, op);
            if (held != null)
                return held;
            // Legacy shim: 0.1.0.3 evidence recovery set the list ID from the original create
            // response and left the answer marked unknown. The list ID settles it; the run reads
            // that library by ID. Remove after 0.1.0.5.
            if (
                op.Value.ExternalSubmitted
                && !op.Value.ExternalResponseKnown
                && op.Value.Mutation == "CreateLibrary"
                && op.Value.ListId != Guid.Empty
            )
                op.Value.ExternalResponseKnown = true;
            var site = service.Retrieve(
                "asx_site",
                op.Value.SiteId,
                new ColumnSet("asx_approved", "asx_webid", "asx_url")
            );
            // A lookup only reads, so it also runs while the site is suspended (decision D5).
            if (
                (!op.Value.Reconcile && !site.GetAttributeValue<bool>("asx_approved"))
                || TemplateStore.Text(site, "asx_webid") != op.Value.WebId.ToString("D")
            )
                throw new EvaluationBlockedException("Site identity or readiness changed.");
            // A re-pointed site keeps its web ID; setup follows its new address. Every claim
            // reads the library again before registering it.
            op.Value.WebUrl = TemplateStore.Text(site, "asx_url");
            var claim = store.Find<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(service, request.Key)
            );
            if (claim == null)
            {
                store.Create(
                    "asx_claim",
                    new DispatcherDocument
                    {
                        Key = WorkCoordination.Operation(service, request.Key),
                        Status = "Idle",
                    }
                );
                claim = store.Require<DispatcherDocument>(
                    "asx_claim",
                    WorkCoordination.Operation(service, request.Key)
                );
            }
            claim = WorkCoordination.Unstall(service, claim);
            if (claim.Value.RunId != null)
            {
                if (claim.Value.OperationKey != request.Key)
                    return new WorkerResult { Status = "Busy", Key = request.Key };
                bool live = claim.Value.LeaseUntilUtc > clock();
                if (op.Value.Reconcile)
                {
                    // A lookup only reads: once the lease of the run that held it ends, the next run takes it over.
                    if (live)
                        return new WorkerResult { Status = "Quarantined", Key = request.Key };
                    claim.Value.HttpOutstanding = false;
                }
                else
                {
                    // An expired lease is taken over only when no library write is outstanding; an
                    // unknown write is looked up in SharePoint, since a second create could duplicate it.
                    bool unknownWrite =
                        op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown;
                    // A prepared write the connection never permitted was never sent: the run that
                    // held it stopped first. It is read and prepared again, not recovered.
                    if (unknownWrite && claim.Value.LeaseUntilUtc <= clock() && NeverSent(op.Value))
                    {
                        Unsend(op.Value);
                        unknownWrite = false;
                    }
                    if (claim.Value.LeaseUntilUtc > clock() || unknownWrite)
                    {
                        if (unknownWrite && claim.Value.LeaseUntilUtc <= clock())
                        {
                            op.Value.ErrorCode =
                                op.Value.ErrorCode
                                ?? "Library request outcome is unknown. Reconcile the original run before retry.";
                            AwaitRecovery(store, op, claim);
                        }
                        return new WorkerResult { Status = "Quarantined", Key = request.Key };
                    }
                    claim.Value.HttpOutstanding = false;
                    op.Value.ExternalResponseKnown = true;
                }
            }
            claim.Value.OperationKey = request.Key;
            claim.Value.RunId = request.RunId;
            claim.Value.Token = Guid.NewGuid();
            claim.Value.LeaseUntilUtc = clock().AddMinutes(5);
            claim.Value.Status = "Claimed";
            store.Save(claim);
            if (op.Value.Reconcile)
                return Read(
                    op,
                    claim.Value,
                    "ReconcileTitle",
                    "_api/web/lists/GetByTitle(@p)?" + LookupByTitle(op.Value.Name)
                );
            return Read(
                op,
                claim.Value,
                "Owner",
                "_api/web/associatedownergroup?$select=Id,PrincipalType"
            );
        }
        var lease = store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        if (
            lease.Value.OperationKey != request.Key
            || lease.Value.RunId != request.RunId
            || request.Token == Guid.Empty
            || lease.Value.Token != request.Token
            || lease.Value.LeaseUntilUtc <= clock()
        )
            throw new EvaluationBlockedException("Stale library setup claim.");
        if (request.Command == "Renew")
        {
            lease.Value.LeaseUntilUtc = clock().AddMinutes(5);
            store.Save(lease);
            return new WorkerResult
            {
                Status = "Renewed",
                Key = request.Key,
                Token = request.Token,
            };
        }
        if (request.Command == "Fail")
        {
            // A lookup read that failed waits after a temporary failure, else reports why.
            if (op.Value.Reconcile)
                return TransientFailure.Is(request)
                    ? WaitLookup(op, lease, null, request.StatusCode, request.ErrorCode)
                    : Settled(store, op, lease, LookupFailed(request, clock()));
            // The run failed before its prepared write was permitted, so nothing was sent.
            if (NeverSent(op.Value))
                Unsend(op.Value);
            if (
                TransientFailure.Is(request)
                && !(op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown)
            )
                return Wait(op, lease, null, request.StatusCode, request.ErrorCode);
            return Block(op, lease, "Library setup interrupted. Existing content was preserved.");
        }
        if (request.Command == "PrepareCreate")
        {
            if (
                op.Value.Status != "ReadyToCreate"
                || op.Value.Intent == null
                || op.Value.ExternalSubmitted
            )
                throw new EvaluationBlockedException("No library mutation is ready.");
            op.Value.ExternalSubmitted = true;
            op.Value.ExternalResponseKnown = false;
            // Not sent until the shared connection permits it (WorkCoordination.BeginHttp).
            op.Value.WritePermitted = false;
            op.Value.Status = "ExternalUnknown";
            store.Save(op);
            return Work(op.Value, lease.Value, "Create", op.Value.Intent);
        }
        if (request.Command == "CreateResponse")
        {
            if (!op.Value.ExternalSubmitted || op.Value.ExternalResponseKnown)
                throw new EvaluationBlockedException("No library mutation is outstanding.");
            // An answer, even a lost one, means the write was sent.
            op.Value.WritePermitted = true;
            if (request.HttpStatus == 429)
            {
                // SharePoint does not execute a throttled request, so nothing was created or
                // changed. After the wait, setup re-reads and prepares the same change again.
                op.Value.ExternalSubmitted = false;
                op.Value.ExternalResponseKnown = false;
                return Wait(op, lease, request.RetryAfter, 429, null);
            }
            if (request.HttpStatus == 0 || request.HttpStatus == 408 || request.HttpStatus >= 500)
                return Block(
                    op,
                    lease,
                    "Library request outcome is unknown. Reconcile the original run before retry."
                );
            op.Value.ExternalResponseKnown = true;
            if (request.HttpStatus < 200 || request.HttpStatus >= 300)
                return Block(
                    op,
                    lease,
                    "SharePoint rejected "
                        + op.Value.Mutation
                        + ". Check application permissions and retry setup."
                );
            if (op.Value.Mutation == "CreateLibrary")
            {
                var created = JsonWire
                    .Read<ODataEnvelope<CreatedLibrary>>(request.ResponseBody ?? "")
                    .Data;
                if (created == null || created.Id == Guid.Empty || created.Title != op.Value.Name)
                    return Block(
                        op,
                        lease,
                        "Creation response did not establish the library identity."
                    );
                op.Value.ListId = created.Id;
            }
            return Read(
                op,
                lease.Value,
                "Library",
                List(op.Value)
                    + "?$select=Id,HasUniqueRoleAssignments,RootFolder/UniqueId,RootFolder/ServerRelativeUrl&$expand=RootFolder"
            );
        }
        if (
            request.Command != "Observe"
            || request.ProbeId != op.Value.ProbeId
            || request.ProbeKind != op.Value.ProbeKind
        )
            throw new EvaluationBlockedException("Library observation does not match its request.");
        if (
            request.HttpStatus == 429
            || request.HttpStatus >= 500
            || request.HttpStatus == 0
            || request.HttpStatus == 408
        )
        {
            if (op.Value.Reconcile)
                return WaitLookup(op, lease, request.RetryAfter, request.HttpStatus, null);
            if (op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown)
                return Block(op, lease, "Library write is unresolved.");
            return Wait(op, lease, request.RetryAfter, request.HttpStatus, null);
        }
        try
        {
            switch (op.Value.ProbeKind)
            {
                case "Owner":
                    var owner = SharePointObservations.Body<SiteGroup>(request);
                    if (owner.Id <= 0 || owner.Type != 8)
                        throw new EvaluationBlockedException(
                            "This site needs an associated owners group before library creation."
                        );
                    op.Value.OwnerGroup = owner.Id;
                    return Read(
                        op,
                        lease.Value,
                        "Roles",
                        "_api/web/roledefinitions?$select=Id,RoleTypeKind,BasePermissions"
                    );
                case "Roles":
                    var roles = SharePointObservations.Body<ODataRows<SecurityRoleObservation>>(
                        request
                    );
                    if (roles.Next != null || roles.Rows.Length > 100)
                        throw new EvaluationBlockedException(
                            "Permission definitions could not be read completely."
                        );
                    PolicyRole Pick(int type)
                    {
                        var x = roles.Rows.Single(v => v.Type == type);
                        return new PolicyRole
                        {
                            Id = x.Id,
                            High = x.Permissions.High,
                            Low = x.Permissions.Low,
                        };
                    }
                    op.Value.ReadRole = Pick(2);
                    op.Value.ContributeRole = Pick(3);
                    op.Value.OwnerRole = Pick(5).Id;
                    SecurityAdministration.ValidateRole(op.Value.ReadRole, true);
                    SecurityAdministration.ValidateRole(op.Value.ContributeRole, false);
                    return Read(
                        op,
                        lease.Value,
                        "Library",
                        // The title goes in the query string as a parameter alias, like folder
                        // paths (SharePointRequests.ByPath), so a long title never lengthens the
                        // URL path the connector limits.
                        op.Value.ListId == Guid.Empty
                            ? "_api/web/lists/GetByTitle(@p)?" + ByTitle(op.Value.Name)
                            : List(op.Value) + "?" + LibrarySelect
                    );
                case "Library":
                    if (request.HttpStatus == 404 && op.Value.ListId == Guid.Empty)
                    {
                        // A library created after this read is one the lookup may find (spec 6.8).
                        op.Value.AbsentUtc = clock();
                        return Prepare(
                            op,
                            lease.Value,
                            "CreateLibrary",
                            new HttpIntent
                            {
                                Method = "POST",
                                RelativeUri = "_api/web/lists",
                                Body = JsonWire.Write(new NewLibraryBody { Title = op.Value.Name }),
                            }
                        );
                    }
                    var library = SharePointObservations.Body<CatalogLibraryObservation>(request);
                    if (op.Value.ListId == Guid.Empty)
                        throw new EvaluationBlockedException(
                            "A library with this name already exists. Add the existing library instead."
                        );
                    if (
                        library.Id != op.Value.ListId
                        || library.Root?.Id == Guid.Empty
                        || library.Root == null
                    )
                        throw new EvaluationBlockedException(
                            "Library identity changed during setup."
                        );
                    new SharePointTarget(
                        op.Value.WebUrl,
                        op.Value.WebId,
                        op.Value.ListId,
                        library.Root.Path
                    );
                    op.Value.Library = library;
                    if (library.Unique != true)
                        return Prepare(
                            op,
                            lease.Value,
                            "LibraryBoundary",
                            new HttpIntent
                            {
                                Method = "POST",
                                RelativeUri =
                                    List(op.Value)
                                    + "/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)",
                            }
                        );
                    return OwnerGrant(op, lease.Value);
                case "Acl":
                    // Setups saved by 0.1.0.3 mid-read asked for the library's whole role
                    // assignment list. Its answer is not used: the owners group's own assignment
                    // is read instead. Remove after 0.1.0.5.
                    return OwnerGrant(op, lease.Value);
                case "OwnerGrant":
                    // Only the owners group's own assignment is read, so the library's size or
                    // the people shared with it never matter. SharePoint answers 404 when the
                    // group has no assignment on the list.
                    AclAssignment? owners = null;
                    if (request.HttpStatus != 404)
                    {
                        owners = SharePointObservations.Body<AclAssignment>(request);
                        if (owners.Member?.Id != op.Value.OwnerGroup)
                            throw new EvaluationBlockedException(
                                "SharePoint returned the role assignment of another principal."
                            );
                    }
                    if (owners?.Roles?.Rows?.Any(role => role.Id == op.Value.OwnerRole) != true)
                        return Prepare(
                            op,
                            lease.Value,
                            "OwnerAccess",
                            new HttpIntent
                            {
                                Method = "POST",
                                RelativeUri =
                                    List(op.Value)
                                    + "/roleassignments/addroleassignment(principalid="
                                    + op.Value.OwnerGroup
                                    + ",roledefid="
                                    + op.Value.OwnerRole
                                    + ")",
                            }
                        );
                    return Register(op, lease);
                case "ReconcileTitle":
                case "ReconcileUrl":
                    return Lookup(op, lease, request);
                default:
                    throw new EvaluationBlockedException("Unknown library setup phase.");
            }
        }
        catch (EvaluationBlockedException error)
        {
            return Block(op, lease, error.Message);
        }
        catch (InvalidOperationException)
        {
            return Block(op, lease, "SharePoint permission definitions are ambiguous.");
        }
    }

    private WorkerResult Register(StoredRow<LibrarySetup> op, StoredRow<DispatcherDocument> lease)
    {
        var v = op.Value;
        var folder = v.Library!.Root;
        v.CatalogId = SiteIdentity.LibraryId(v.SiteId, v.ListId);
        // While the setup awaited recovery, an admin may have added the created library as an
        // existing one. Approval gives the same row for the same site and list, so that row is
        // adopted; any other entry for the list is left alone and the setup stops with why.
        var listed = new QueryExpression("asx_library")
        {
            ColumnSet = new ColumnSet("asx_siteid", "statecode"),
            TopCount = 2,
        };
        listed.Criteria.AddCondition("asx_listid", ConditionOperator.Equal, v.ListId.ToString("D"));
        var entries = service.RetrieveMultiple(listed).Entities;
        if (entries.Count > 0)
        {
            var entry = entries.Count == 1 ? entries[0] : null;
            if (
                entry == null
                || entry.Id != v.CatalogId
                || entry.GetAttributeValue<EntityReference>("asx_siteid")?.Id != v.SiteId
                || entry.GetAttributeValue<OptionSetValue>("statecode")?.Value == 1
            )
                return Block(op, lease, CatalogConflictNotice);
            return Adopt(op, lease);
        }
        var nativeId = new NativeLocations(service).EnsureLibrary(
            v.SiteId,
            v.WebId,
            v.ListId,
            v.NativeSiteId,
            v.WebUrl,
            folder.Path,
            v.Name
        );
        var catalog = new Entity("asx_library", v.CatalogId)
        {
            ["asx_name"] = v.Name,
            ["asx_siteid"] = new EntityReference("asx_site", v.SiteId),
            ["asx_listid"] = v.ListId.ToString("D"),
            ["asx_entryid"] = folder.Id.ToString("D"),
            ["asx_entryurl"] = new Uri(v.WebUrl).GetLeftPart(UriPartial.Authority) + folder.Path,
            ["asx_nativeparentid"] = new EntityReference("sharepointdocumentlocation", nativeId),
            ["asx_approved"] = true,
            ["asx_policyapplied"] = false,
            ["asx_readrole"] = JsonWire.Write(v.ReadRole),
            ["asx_contributerole"] = JsonWire.Write(v.ContributeRole),
        };
        service.Create(catalog);
        Release(lease);
        return ApplyAccess(op);
    }

    /// <summary>
    /// Finishes with the catalog entry an admin added for this library while the setup awaited
    /// recovery. Access the admin set there is kept; with none, the setup's own access is
    /// applied, as it would have been.
    /// </summary>
    private WorkerResult Adopt(StoredRow<LibrarySetup> op, StoredRow<DispatcherDocument> lease)
    {
        var v = op.Value;
        Release(lease);
        if (store.Find<PolicyDocument>("asx_policy", "policy:" + v.CatalogId.ToString("N")) == null)
            return ApplyAccess(op);
        v.Status = "Applied";
        v.CompletedUtc = clock();
        v.ExternalSubmitted = false;
        v.ExternalResponseKnown = true;
        v.ErrorCode = null;
        store.Save(op);
        return new WorkerResult
        {
            Status = "Applied",
            Key = v.Key,
            Notices = new[]
            {
                "Library created. It was added to the catalog while this setup awaited recovery, so the access settings there are kept.",
            },
        };
    }

    /// <summary>Applies the setup's team access to its catalog entry and waits on that run.</summary>
    private WorkerResult ApplyAccess(StoredRow<LibrarySetup> op)
    {
        var v = op.Value;
        var policy = new SecurityAdministration(service).Execute(
            new SecurityRequest
            {
                Command = "ApplyPolicy",
                LibraryId = v.CatalogId,
                Entries = v.Entries,
                AcknowledgeBroaderAccess = v.AcknowledgeBroaderAccess,
                ReadRole = v.ReadRole,
                ContributeRole = v.ContributeRole,
            },
            true
        );
        v.PolicyOperation = policy.Policy!.OperationKey;
        var access = store.Require<SecurityOperation>("asx_operation", v.PolicyOperation!);
        access.Value.LibrarySetupKey = v.Key;
        store.Save(access);
        v.Status = "AccessPending";
        v.ExternalSubmitted = false;
        v.ExternalResponseKnown = true;
        v.ErrorCode = null;
        store.Save(op);
        return new WorkerResult
        {
            Status = "AccessPending",
            Key = v.Key,
            Notices = new[] { "Library created. Team access is being applied." },
        };
    }

    /// <summary>Reads the owners group's role assignment on the new library.</summary>
    private WorkerResult OwnerGrant(StoredRow<LibrarySetup> op, DispatcherDocument claim) =>
        Read(
            op,
            claim,
            "OwnerGrant",
            List(op.Value)
                + "/roleassignments/getbyprincipalid("
                + op.Value.OwnerGroup
                + ")?$select=Member/Id,RoleDefinitionBindings/Id&$expand=Member,RoleDefinitionBindings"
        );

    private WorkerResult Read(
        StoredRow<LibrarySetup> op,
        DispatcherDocument claim,
        string kind,
        string endpoint
    )
    {
        op.Value.ProbeKind = kind;
        op.Value.ProbeId = Guid.NewGuid();
        op.Value.Status = "Inspecting";
        store.Save(op);
        return Work(op.Value, claim, "Read", new HttpIntent { RelativeUri = endpoint });
    }

    private WorkerResult Prepare(
        StoredRow<LibrarySetup> op,
        DispatcherDocument claim,
        string kind,
        HttpIntent intent
    )
    {
        op.Value.Mutation = kind;
        op.Value.Intent = intent;
        op.Value.ExternalSubmitted = false;
        op.Value.ExternalResponseKnown = false;
        op.Value.Status = "ReadyToCreate";
        store.Save(op);
        return Work(op.Value, claim, "ReadyToCreate", null);
    }

    private static WorkerResult Work(
        LibrarySetup op,
        DispatcherDocument claim,
        string status,
        HttpIntent? http
    ) =>
        new WorkerResult
        {
            Key = op.Key,
            Status = status,
            Token = claim.Token,
            ProbeId = op.ProbeId,
            ProbeKind = op.ProbeKind,
            SiteUrl = op.WebUrl,
            Http = http,
        };

    private WorkerResult Block(
        StoredRow<LibrarySetup> op,
        StoredRow<DispatcherDocument> claim,
        string issue
    )
    {
        // A lookup that cannot go on reports why; its create stays unknown until the admin chooses.
        if (op.Value.Reconcile)
            return Settled(store, op, claim, LibraryReconcile.ReadFailed(issue, clock()));
        op.Value.ErrorCode = issue;
        bool unknown = op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown;
        // An unknown library write is never simply sent again: a second create could make a
        // duplicate library. A lost create is looked up in SharePoint (Reconciling, listed once
        // due); any other unknown write waits in RecoveryRequired, off the dispatch page.
        if (unknown)
            AwaitRecovery(store, op, claim);
        else
        {
            op.Value.Status = "Blocked";
            store.Save(op);
            Release(claim);
        }
        return new WorkerResult
        {
            Key = op.Value.Key,
            Status = unknown ? "Quarantined" : "Blocked",
            Notices = new[] { issue },
        };
    }

    /// <summary>
    /// Waits after a temporary failure with no attempt cap and releases the writer slot.
    /// </summary>
    private WorkerResult Wait(
        StoredRow<LibrarySetup> op,
        StoredRow<DispatcherDocument> lease,
        string? retryAfter,
        int? statusCode,
        string? errorCode
    )
    {
        DocumentStore.Wait(op.Value, statusCode, errorCode, retryAfter, clock());
        store.Save(op);
        Release(lease);
        return new WorkerResult
        {
            Status = "RetryWait",
            Key = op.Value.Key,
            Notices = new[] { op.Value.ErrorCode! },
        };
    }

    private void Release(StoredRow<DispatcherDocument> claim)
    {
        claim.Value.OperationKey = null;
        claim.Value.RunId = null;
        claim.Value.Token = Guid.Empty;
        claim.Value.Status = "Idle";
        store.Save(claim);
    }

    /// <summary>One lookup read's answer: the title read leads to the address read, which decides.</summary>
    private WorkerResult Lookup(
        StoredRow<LibrarySetup> op,
        StoredRow<DispatcherDocument> lease,
        WorkerRequest request
    )
    {
        if (request.HttpStatus != 200 && request.HttpStatus != 404)
            return Settled(
                store,
                op,
                lease,
                LibraryReconcile.ReadFailed(SharePointObservations.ErrorMessage(request), clock())
            );
        var hit =
            request.HttpStatus == 404
                ? null
                : SharePointObservations.Body<ReconcileObservation>(request);
        string expected = LibraryReconcile.ExpectedUrl(op.Value.WebUrl, op.Value.Name);
        if (op.Value.ProbeKind == "ReconcileTitle")
        {
            op.Value.TitleHit = hit;
            return Read(
                op,
                lease.Value,
                "ReconcileUrl",
                "_api/web/GetList(@p)?" + LookupByAddress(op.Value.WebUrl, op.Value.Name)
            );
        }
        var finding = LibraryReconcile.Decide(
            op.Value.SiteId,
            op.Value.Name,
            expected,
            op.Value.TitleHit,
            hit,
            Requested(op),
            CatalogRows(op.Value, op.Value.TitleHit, hit),
            clock()
        );
        return Settled(store, op, lease, finding);
    }

    /// <summary>Stores the finding, puts the setup in RecoveryRequired and releases the site's writer.</summary>
    internal static WorkerResult Settled(
        DocumentStore store,
        StoredRow<LibrarySetup> op,
        StoredRow<DispatcherDocument>? claim,
        LibraryRecovery finding
    )
    {
        ReleaseIfHeld(store, op, claim);
        string sentence = LibraryReconcile.Sentence(op.Value.Name, finding);
        op.Value.Recovery = finding;
        op.Value.Reconcile = false;
        op.Value.TitleHit = null;
        op.Value.Status = "RecoveryRequired";
        op.Value.ErrorCode = sentence;
        op.Value.ProbeId = Guid.Empty;
        op.Value.NextAttemptUtc = null;
        op.Value.RetryCount = 0;
        store.Save(op);
        return new WorkerResult
        {
            Key = op.Value.Key,
            Status = "RecoveryRequired",
            Notices = new[] { sentence },
            Recovery = finding,
        };
    }

    /// <summary>A temporary failure of a lookup read waits (A6) and stays a lookup.</summary>
    private WorkerResult WaitLookup(
        StoredRow<LibrarySetup> op,
        StoredRow<DispatcherDocument> lease,
        string? retryAfter,
        int? statusCode,
        string? errorCode
    )
    {
        DocumentStore.Wait(op.Value, statusCode, errorCode, retryAfter, clock());
        op.Value.Status = "Reconciling";
        store.Save(op);
        Release(lease);
        return new WorkerResult
        {
            Status = "RetryWait",
            Key = op.Value.Key,
            Notices = new[] { op.Value.ErrorCode! },
        };
    }

    /// <summary>The claim of a lookup failed before any read: wait after a temporary failure, else report why.</summary>
    internal static WorkerResult LookupClaimFailed(
        DocumentStore store,
        StoredRow<LibrarySetup> op,
        StoredRow<DispatcherDocument>? claim,
        WorkerRequest failure,
        DateTime now
    )
    {
        if (!TransientFailure.Is(failure))
            return Settled(store, op, claim, LookupFailed(failure, now));
        ReleaseIfHeld(store, op, claim);
        DocumentStore.Wait(op.Value, failure.StatusCode, failure.ErrorCode, null, now);
        op.Value.Status = "Reconciling";
        store.Save(op);
        return new WorkerResult
        {
            Status = "Reconciling",
            Key = op.Value.Key,
            Notices = new[] { op.Value.ErrorCode! },
        };
    }

    /// <summary>The admin's choice on a finding (spec 6.8): use a library it found, or create the library again.</summary>
    /// <param name="request">Key, Choice (UseLibrary with ListId, or CreateAgain) and the RowVersion the admin saw.</param>
    /// <returns>The setup, Pending, for the next dispatch.</returns>
    public CatalogResult ResolveSetup(CatalogRequest request)
    {
        var op = store.Require<LibrarySetup>("asx_operation", SetupKey(request.Key));
        var finding = op.Value.Recovery;
        if (
            op.Value.Status != "RecoveryRequired"
            || finding == null
            || finding.State == "Checking"
            || !LostCreate(op.Value)
        )
            throw new EvaluationBlockedException(
                "This setup has no SharePoint finding to act on. Check again first."
            );
        if (op.Row.RowVersion != request.RowVersion)
            throw new EvaluationBlockedException(
                "This setup changed. Look at its latest finding and choose again."
            );
        string receipt;
        if (request.Choice == "UseLibrary")
        {
            var candidate =
                finding.Candidates.FirstOrDefault(c => c.ListId == request.ListId)
                ?? throw new EvaluationBlockedException(
                    "Choose one of the libraries Documents found."
                );
            if (!candidate.IsLibrary)
                throw new EvaluationBlockedException("Only a document library can be used.");
            if (candidate.CatalogEntry == "Conflict")
                throw new EvaluationBlockedException(CatalogConflictNotice);
            if (!Offered(finding, "UseLibrary") && !Offered(finding, "UseCandidate"))
                throw new EvaluationBlockedException(ChoiceNotOffered);
            // The list is known now: the next run reads it by ID, then boundary, owner access and
            // Register, which adopts this site's catalog entry for it (A10). No create is prepared.
            op.Value.ListId = candidate.ListId;
            op.Value.ExternalResponseKnown = true;
            op.Value.Status = "Pending";
            receipt = "RecoveryUseLibrary:" + candidate.ListId.ToString("D");
        }
        else if (request.Choice == "CreateAgain")
        {
            if (finding.Candidates.Any(c => c.TitleMatches))
                throw new EvaluationBlockedException(
                    "SharePoint has a library with this name, so Documents won't create another. Use it, or cancel the setup."
                );
            // Only a finding that offers it: a read that failed shows nothing about the library.
            if (!Offered(finding, "CreateAgain"))
                throw new EvaluationBlockedException(ChoiceNotOffered);
            // The next run reads the title before preparing the create; a library that appeared
            // since the lookup stops it with the name-collision rule.
            Unsend(op.Value);
            receipt = "RecoveryCreateAgain";
        }
        else
            throw new EvaluationBlockedException("Choose UseLibrary or CreateAgain.");
        op.Value.Recovery = null;
        op.Value.ErrorCode = null;
        op.Value.RetryCount = 0;
        op.Value.NextAttemptUtc = null;
        store.Save(op);
        Receipt(op.Value.Key, receipt);
        return Result(store.Require<LibrarySetup>("asx_operation", op.Value.Key));
    }

    /// <summary>Check again: looks SharePoint up once more for a setup whose create is unknown.</summary>
    /// <param name="key">The library setup key.</param>
    /// <returns>The setup, Reconciling, with its finding Checking.</returns>
    public CatalogResult RecheckSetup(string key)
    {
        var op = store.Require<LibrarySetup>("asx_operation", SetupKey(key));
        if (op.Value.Status == "Reconciling")
            return Result(op);
        if (op.Value.Status != "RecoveryRequired" || !LostCreate(op.Value))
            throw new EvaluationBlockedException(
                "Only a library setup whose result is unknown can be checked again."
            );
        StartLookup(op.Value, null);
        store.Save(op);
        Receipt(key, "RecoveryCheckAgain");
        return Result(store.Require<LibrarySetup>("asx_operation", key));
    }

    /// <summary>
    /// Starts the lookup for a setup 0.1.0.3 left in RecoveryRequired before lookups existed, when
    /// Inspect or ListProblems first shows it (decision D6). Any other row is returned unchanged.
    /// </summary>
    /// <param name="op">The library setup as read.</param>
    /// <returns>The setup as it now stands.</returns>
    internal StoredRow<LibrarySetup> PickUp(StoredRow<LibrarySetup> op)
    {
        if (
            op.Value.Status != "RecoveryRequired"
            || op.Value.Recovery != null
            || !LostCreate(op.Value)
        )
            return op;
        StartLookup(op.Value, null);
        store.Save(op);
        return store.Require<LibrarySetup>("asx_operation", op.Value.Key);
    }

    // Whether the stored finding offers this choice (LibraryReconcile.Choices).
    private static bool Offered(LibraryRecovery finding, string choice) =>
        Array.IndexOf(finding.Choices ?? Array.Empty<string>(), choice) >= 0;

    private static string SetupKey(string key) =>
        key.StartsWith("librarycreate:", StringComparison.Ordinal)
            ? key
            : throw new EvaluationBlockedException("Select a library setup.");

    // When the title was last seen absent; setups stored before AbsentUtc use the row's createdon (D4).
    private DateTime Requested(StoredRow<LibrarySetup> op) =>
        op.Value.AbsentUtc
        ?? service
            .Retrieve("asx_operation", op.Row.Id, new ColumnSet("createdon"))
            .GetAttributeValue<DateTime>("createdon");

    // Catalog rows on this site with the requested name, and rows for either list found.
    private CatalogEntryRow[] CatalogRows(LibrarySetup setup, params ReconcileObservation?[] hits)
    {
        var query = new QueryExpression("asx_library")
        {
            ColumnSet = new ColumnSet("asx_siteid", "asx_name", "asx_listid", "statecode"),
        };
        query.Criteria.FilterOperator = LogicalOperator.Or;
        var named = new FilterExpression(LogicalOperator.And);
        named.AddCondition("asx_siteid", ConditionOperator.Equal, setup.SiteId);
        named.AddCondition("asx_name", ConditionOperator.Equal, setup.Name);
        query.Criteria.AddFilter(named);
        // List IDs are stored as text, in either case (sites-access.js reads them case-blind too).
        var lists = hits.Where(hit => hit != null)
            .SelectMany(hit =>
                new object[] { hit!.Id.ToString("D"), hit.Id.ToString("D").ToUpperInvariant() }
            )
            .ToArray();
        if (lists.Length > 0)
            query.Criteria.AddCondition("asx_listid", ConditionOperator.In, lists);
        return service
            .RetrieveMultiple(query)
            .Entities.Select(row => new CatalogEntryRow
            {
                Id = row.Id,
                SiteId = row.GetAttributeValue<EntityReference>("asx_siteid")?.Id ?? Guid.Empty,
                Name = row.GetAttributeValue<string>("asx_name") ?? "",
                ListId = Guid.TryParse(row.GetAttributeValue<string>("asx_listid"), out var list)
                    ? list
                    : Guid.Empty,
                Removed = row.GetAttributeValue<OptionSetValue>("statecode")?.Value == 1,
            })
            .ToArray();
    }

    // An audit row for the admin's recovery choice, beside the run attempts of the setup.
    private void Receipt(string key, string what) =>
        store.Create(
            "asx_attempt",
            new AttemptDocument
            {
                Key = "attempt:" + Guid.NewGuid().ToString("N"),
                Status = "Recorded",
                OperationKey = key,
                RunId = "admin",
                Event = what,
                AtUtc = clock(),
            }
        );
}
