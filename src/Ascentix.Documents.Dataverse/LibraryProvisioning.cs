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

    [DataMember]
    public CatalogLibraryObservation? Library { get; set; }

    [DataMember]
    public string Mutation { get; set; } = "";

    [DataMember]
    public HttpIntent? Intent { get; set; }
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
    public CatalogResult Queue(CatalogRequest request)
    {
        Domain.FolderNames.Validate(request.Name);
        if (request.SiteId == Guid.Empty || request.RequestId == Guid.Empty)
            throw new EvaluationBlockedException("Select a site and provide a request identity.");
        if (
            request.Entries == null
            || request.Entries.Length > 10
            || request.Entries.Any(e =>
                e.TeamId == Guid.Empty || !new[] { "Read", "Contribute" }.Contains(e.Access)
            )
            || request.Entries.Select(e => e.TeamId).Distinct().Count() != request.Entries.Length
        )
            throw new EvaluationBlockedException(
                "Select distinct teams with Read or Contribute access."
            );
        foreach (var entry in request.Entries)
        {
            var team = service.Retrieve(
                "team",
                entry.TeamId,
                new ColumnSet("teamtype", "isdefault")
            );
            if (
                team.GetAttributeValue<OptionSetValue>("teamtype")?.Value != 0
                || team.GetAttributeValue<bool>("isdefault")
            )
                throw new EvaluationBlockedException(
                    "Only non-default manual owner teams are supported."
                );
        }
        var site = service.Retrieve(
            "asx_site",
            request.SiteId,
            new ColumnSet("asx_nativeid", "asx_webid", "asx_url", "asx_approved")
        );
        if (!site.GetAttributeValue<bool>("asx_approved"))
            throw new EvaluationBlockedException(
                "Site validation must complete before creating a library."
            );
        string key =
            "librarycreate:"
            + DocumentStore.Hash(
                request.SiteId.ToString("N") + ":" + request.Name.ToUpperInvariant()
            );
        var old = store.Find<LibrarySetup>("asx_operation", key);
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
            }
        );
        return Result(store.Require<LibrarySetup>("asx_operation", key));
    }

    public CatalogResult Inspect(string key) =>
        Result(store.Require<LibrarySetup>("asx_operation", key));

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
        };
    }

    private string List(LibrarySetup op) => "_api/web/lists(guid'" + op.ListId.ToString("D") + "')";

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
        {
            var active = store.Find<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(service, request.Key)
            );
            if (
                active?.Value.OperationKey == request.Key
                || op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown
            )
                throw new EvaluationBlockedException(
                    "An outstanding library request must be reconciled before retry."
                );
            if (op.Value.Status != "Blocked" && op.Value.Status != "RetryWait")
                throw new EvaluationBlockedException("Only interrupted setup can retry.");
            op.Value.Status = request.Command == "Retry" ? "Pending" : "Cancelled";
            op.Value.NextAttemptUtc = null;
            op.Value.RetryCount = 0;
            store.Save(op);
            return new WorkerResult { Key = request.Key, Status = op.Value.Status };
        }
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
            if (op.Value.NextAttemptUtc > clock())
                return new WorkerResult { Status = "RetryWait", Key = request.Key };
            if (
                string.IsNullOrWhiteSpace(request.RunId)
                || request.RunId.Length > 150
                || request.RunId.Any(char.IsControl)
            )
                throw new EvaluationBlockedException("Actual worker run identity required.");
            var site = service.Retrieve(
                "asx_site",
                op.Value.SiteId,
                new ColumnSet("asx_approved", "asx_webid", "asx_url")
            );
            if (
                !site.GetAttributeValue<bool>("asx_approved")
                || TemplateStore.Text(site, "asx_webid") != op.Value.WebId.ToString("D")
                || TemplateStore.Text(site, "asx_url") != op.Value.WebUrl
            )
                throw new EvaluationBlockedException("Site identity or readiness changed.");
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
            if (claim.Value.RunId != null)
            {
                if (claim.Value.OperationKey != request.Key)
                    return new WorkerResult { Status = "Busy", Key = request.Key };
                // An expired lease is taken over only when no library write is outstanding; an
                // unknown write keeps operator recovery, since a second create could duplicate it.
                bool unknownWrite = op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown;
                if (
                    !claim.Value.RecoveryPermitted
                    && (claim.Value.LeaseUntilUtc > clock() || unknownWrite)
                )
                    return new WorkerResult { Status = "Quarantined", Key = request.Key };
                claim.Value.HttpOutstanding = false;
                if (
                    unknownWrite
                    && op.Value.Mutation == "CreateLibrary"
                    && op.Value.ListId == Guid.Empty
                )
                    throw new EvaluationBlockedException(
                        "Unknown library creation needs physical identity reconciliation; a same-name library must not be adopted automatically."
                    );
                op.Value.ExternalResponseKnown = true;
            }
            claim.Value.OperationKey = request.Key;
            claim.Value.RunId = request.RunId;
            claim.Value.Token = Guid.NewGuid();
            claim.Value.LeaseUntilUtc = clock().AddMinutes(5);
            claim.Value.RecoveryPermitted = false;
            claim.Value.Status = "Claimed";
            store.Save(claim);
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
            || lease.Value.RecoveryPermitted
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
            op.Value.Status = "ExternalUnknown";
            store.Save(op);
            return Work(op.Value, lease.Value, "Create", op.Value.Intent);
        }
        if (request.Command == "CreateResponse")
        {
            if (!op.Value.ExternalSubmitted || op.Value.ExternalResponseKnown)
                throw new EvaluationBlockedException("No library mutation is outstanding.");
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
                        (
                            op.Value.ListId == Guid.Empty
                                ? "_api/web/lists/GetByTitle('"
                                    + Uri.EscapeDataString(op.Value.Name.Replace("'", "''"))
                                    + "')"
                                : List(op.Value)
                        )
                            + "?$select=Id,HasUniqueRoleAssignments,RootFolder/UniqueId,RootFolder/ServerRelativeUrl&$expand=RootFolder"
                    );
                case "Library":
                    if (request.HttpStatus == 404 && op.Value.ListId == Guid.Empty)
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
                    return Read(
                        op,
                        lease.Value,
                        "Acl",
                        List(op.Value)
                            + "/roleassignments?$select=Member/Id,Member/PrincipalType,RoleDefinitionBindings/Id,RoleDefinitionBindings/BasePermissions&$expand=Member,RoleDefinitionBindings"
                    );
                case "Acl":
                    var acl = SharePointObservations.Body<ODataRows<AclAssignment>>(request);
                    var hash = SharePointObservations.AclHash(acl);
                    if (
                        !acl.Rows.Any(a =>
                            a.Member.Id == op.Value.OwnerGroup
                            && a.Roles.Rows.Any(role => role.Id == op.Value.OwnerRole)
                        )
                    )
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
                    return Register(op, lease, hash);
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

    private WorkerResult Register(
        StoredRow<LibrarySetup> op,
        StoredRow<DispatcherDocument> lease,
        string hash
    )
    {
        var v = op.Value;
        var folder = v.Library!.Root;
        var nativeId = new NativeLocations(service).EnsureLibrary(
            v.SiteId,
            v.WebId,
            v.ListId,
            v.NativeSiteId,
            v.WebUrl,
            folder.Path,
            v.Name
        );
        v.CatalogId = SiteIdentity.LibraryId(v.SiteId, v.ListId);
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
            ["asx_aclhash"] = hash,
            ["asx_readrole"] = JsonWire.Write(v.ReadRole),
            ["asx_contributerole"] = JsonWire.Write(v.ContributeRole),
        };
        service.Create(catalog);
        Release(lease);
        var policy = new SecurityAdministration(service).Execute(
            new SecurityRequest
            {
                Command = "ApplyPolicy",
                LibraryId = v.CatalogId,
                Entries = v.Entries,
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
        op.Value.ErrorCode = issue;
        bool unknown = op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown;
        op.Value.Status = unknown ? "ExternalUnknown" : "Blocked";
        store.Save(op);
        if (!unknown)
            Release(claim);
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
        claim.Value.RecoveryPermitted = false;
        claim.Value.Status = "Idle";
        store.Save(claim);
    }
}
