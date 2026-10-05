using System;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class CatalogRequest
{
    [DataMember]
    public string Command { get; set; } = "";

    [DataMember]
    public PolicyEntry[] Entries { get; set; } = Array.Empty<PolicyEntry>();

    [DataMember]
    public Guid CatalogId { get; set; }

    [DataMember]
    public Guid RequestId { get; set; }

    [DataMember]
    public Guid NativeSiteId { get; set; }

    [DataMember]
    public Guid SiteId { get; set; }

    [DataMember]
    public Guid ListId { get; set; }

    [DataMember]
    public Guid NativeParentId { get; set; }

    [DataMember]
    public string Key { get; set; } = "";

    [DataMember]
    public string RowVersion { get; set; } = "";

    [DataMember]
    public string? CatalogRowVersion { get; set; }

    [DataMember]
    public string Name { get; set; } = "";
}

[DataContract]
public sealed class LibraryChoice
{
    [DataMember]
    public Guid Id { get; set; }

    [DataMember]
    public string Title { get; set; } = "";

    [DataMember]
    public int BaseTemplate { get; set; }

    [DataMember]
    public bool Hidden { get; set; }
}

[DataContract]
public sealed class CatalogProbe : OperationDocument
{
    [DataMember]
    public bool DiscoverLibraries { get; set; }

    [DataMember]
    public string? DiscoveryEndpoint { get; set; }

    [DataMember]
    public string? NextLibraries { get; set; }

    [DataMember]
    public LibraryChoice[] Libraries { get; set; } = Array.Empty<LibraryChoice>();

    [DataMember]
    public bool AutoApprove { get; set; }

    [DataMember]
    public string DisplayName { get; set; } = "";

    [DataMember]
    public Guid NativeSiteId { get; set; }

    [DataMember]
    public Guid SiteId { get; set; }

    [DataMember]
    public Guid WebId { get; set; }

    [DataMember]
    public Guid CollectionId { get; set; }

    [DataMember]
    public Guid CatalogId { get; set; }

    [DataMember]
    public string WebUrl { get; set; } = "";

    [DataMember]
    public Guid ListId { get; set; }

    [DataMember]
    public Guid NativeParentId { get; set; }

    [DataMember]
    public Guid EntryId { get; set; }

    [DataMember]
    public string EntryUrl { get; set; } = "";

    [DataMember]
    public AclAssignment[] Acl { get; set; } = Array.Empty<AclAssignment>();

    [DataMember]
    public SecurityRoleObservation[] Roles { get; set; } = Array.Empty<SecurityRoleObservation>();

    [DataMember]
    public string Endpoint { get; set; } = "";

    [DataMember]
    public string[] Pages { get; set; } = Array.Empty<string>();

    [DataMember]
    public DateTime ObservedUtc { get; set; }
}

[DataContract]
public sealed class CatalogResult
{
    [DataMember]
    public string? RecoveryKey { get; set; }

    [DataMember]
    public string Status { get; set; } = "";

    [DataMember]
    public string Key { get; set; } = "";

    [DataMember]
    public string RowVersion { get; set; } = "";

    [DataMember]
    public Guid CatalogId { get; set; }

    [DataMember]
    public string? Issue { get; set; }

    [DataMember]
    public CatalogProbe? Observation { get; set; }
}

[DataContract]
public sealed class WebObservation
{
    [DataMember(Name = "Id")]
    public Guid Id { get; set; }

    [DataMember(Name = "Url")]
    public string Url { get; set; } = "";
}

[DataContract]
public sealed class CatalogLibraryObservation
{
    [DataMember(Name = "Id")]
    public Guid Id { get; set; }

    [DataMember(Name = "HasUniqueRoleAssignments")]
    public bool? Unique { get; set; }

    [DataMember(Name = "RootFolder")]
    public FolderObservation Root { get; set; } = null!;
}

public sealed class CatalogAdministration
{
    private readonly IOrganizationService service;
    private readonly DocumentStore store;
    private readonly Func<DateTime> clock;

    public CatalogAdministration(IOrganizationService service, Func<DateTime>? clock = null)
    {
        this.service = service;
        store = new DocumentStore(service);
        this.clock = clock ?? (() => DateTime.UtcNow);
    }

    public CatalogResult Execute(CatalogRequest request, bool transaction)
    {
        if (!transaction)
            throw new EvaluationBlockedException("Catalog approval requires a transaction.");
        request.Key = request.Key ?? string.Empty;
        if (request.Command == "CreateLibrary")
            return new LibraryProvisioning(service, clock).Queue(request);
        if (
            request.Command == "Inspect"
            && request.Key.StartsWith("librarycreate:", StringComparison.Ordinal)
        )
            return new LibraryProvisioning(service, clock).Inspect(request.Key);
        if (request.Command == "NextLibraries")
        {
            var prior = store.Require<CatalogProbe>("asx_operation", request.Key);
            if (
                !prior.Value.DiscoverLibraries
                || prior.Value.Status != "Discovered"
                || prior.Row.RowVersion != request.RowVersion
                || prior.Value.NextLibraries == null
            )
                throw new EvaluationBlockedException(
                    "Refresh the library discovery page before continuing."
                );
            prior.Value.DiscoveryEndpoint = prior.Value.NextLibraries;
            prior.Value.NextLibraries = null;
            prior.Value.Libraries = Array.Empty<LibraryChoice>();
            prior.Value.Status = "Pending";
            store.Save(prior);
            return Result(store.Require<CatalogProbe>("asx_operation", request.Key));
        }
        if (request.Command == "DiscoverLibraries")
        {
            if (request.RequestId == Guid.Empty)
                throw new EvaluationBlockedException(
                    "Library discovery request identity required."
                );
            var site = service.Retrieve(
                "asx_site",
                request.SiteId,
                new ColumnSet("asx_nativeid", "asx_approved", "asx_collectionid", "asx_webid")
            );
            if (!site.GetAttributeValue<bool>("asx_approved"))
                throw new EvaluationBlockedException("Validate the site first.");
            var native = site.GetAttributeValue<EntityReference>("asx_nativeid").Id;
            string key = "catalogprobe:" + request.RequestId.ToString("N");
            var existing = store.Find<CatalogProbe>("asx_operation", key);
            if (
                existing != null
                && (!existing.Value.DiscoverLibraries || existing.Value.SiteId != site.Id)
            )
                throw new EvaluationBlockedException(
                    "Discovery request ID was reused for different intent."
                );
            if (existing == null)
                store.Create(
                    "asx_operation",
                    new CatalogProbe
                    {
                        Key = key,
                        DiscoverLibraries = true,
                        DisplayName = "Existing libraries",
                        NativeSiteId = native,
                        SiteId = site.Id,
                        WebId = Guid.Parse(TemplateStore.Text(site, "asx_webid")),
                        CollectionId = Guid.Parse(TemplateStore.Text(site, "asx_collectionid")),
                        WebUrl = NativeSite(service, native),
                    }
                );
            return Result(store.Require<CatalogProbe>("asx_operation", key));
        }
        bool automatic = request.Command == "AddSite" || request.Command == "AddLibrary";
        if (request.Command == "AddSite")
            request.Command = "ProbeSite";
        if (request.Command == "AddLibrary")
            request.Command = "ProbeLibrary";
        if (request.Command == "ProbeSite" || request.Command == "ProbeLibrary")
        {
            if (request.RequestId == Guid.Empty)
                throw new EvaluationBlockedException("Stable probe request ID required.");
            bool library = request.Command == "ProbeLibrary";
            Guid nativeId = request.NativeSiteId;
            Guid webId = Guid.Empty;
            Guid collectionId = Guid.Empty;
            if (library)
            {
                var site = service.Retrieve(
                    "asx_site",
                    request.SiteId,
                    new ColumnSet("asx_nativeid", "asx_webid", "asx_collectionid", "asx_approved")
                );
                if (!site.GetAttributeValue<bool>("asx_approved"))
                    throw new EvaluationBlockedException(
                        "Site approval required before library discovery."
                    );
                nativeId =
                    site.GetAttributeValue<EntityReference>("asx_nativeid")?.Id ?? Guid.Empty;
                webId = Guid.Parse(TemplateStore.Text(site, "asx_webid"));
                collectionId = Guid.Parse(TemplateStore.Text(site, "asx_collectionid"));
                if (
                    request.ListId == Guid.Empty
                    || !automatic && request.NativeParentId == Guid.Empty
                )
                    throw new EvaluationBlockedException(
                        "Existing library GUID and native parent are required."
                    );
            }
            string url = NativeSite(service, nativeId);
            string key = "catalogprobe:" + request.RequestId.ToString("N");
            var existing = store.Find<CatalogProbe>("asx_operation", key);
            if (existing != null)
            {
                if (
                    existing.Value.DiscoverLibraries
                    || existing.Value.AutoApprove != automatic
                    || existing.Value.NativeSiteId != nativeId
                    || existing.Value.ListId != (library ? request.ListId : Guid.Empty)
                    || existing.Value.NativeParentId
                        != (library ? request.NativeParentId : Guid.Empty)
                )
                    throw new EvaluationBlockedException(
                        "Probe request ID was reused for different intent."
                    );
                return Result(existing);
            }
            store.Create(
                "asx_operation",
                new CatalogProbe
                {
                    Key = key,
                    AutoApprove = automatic,
                    DisplayName = string.IsNullOrWhiteSpace(request.Name)
                        ? "SharePoint site"
                        : request.Name,
                    NativeSiteId = nativeId,
                    SiteId = library ? request.SiteId : Guid.Empty,
                    WebId = webId,
                    CollectionId = collectionId,
                    WebUrl = url,
                    ListId = library ? request.ListId : Guid.Empty,
                    NativeParentId = library ? request.NativeParentId : Guid.Empty,
                    EntryUrl =
                        library && request.NativeParentId != Guid.Empty
                            ? new NativeLocations(service).ResolveParent(
                                request.NativeParentId,
                                nativeId
                            )
                            : "",
                }
            );
            return Result(store.Require<CatalogProbe>("asx_operation", key));
        }
        if (request.Command == "SuspendSite" || request.Command == "SuspendLibrary")
        {
            bool library = request.Command == "SuspendLibrary";
            string suspendTable = library ? "asx_library" : "asx_site";
            var row = service.Retrieve(
                suspendTable,
                request.CatalogId,
                new ColumnSet("asx_approved")
            );
            if (
                string.IsNullOrEmpty(request.CatalogRowVersion)
                || request.CatalogRowVersion != row.RowVersion
            )
                throw new EvaluationBlockedException(
                    "Refresh the catalog record before suspension."
                );
            var active = store.Find<DispatcherDocument>(
                "asx_claim",
                (
                    library
                        ? WorkCoordination.Library(service, row.Id)
                        : WorkCoordination.Site(service, row.Id)
                )
            );
            // Suspension always succeeds. A write already sent to SharePoint may finish and record
            // its result; every new claim or create re-reads approval and stops. Saving the writer
            // row serializes this change with a concurrent claim.
            if (active != null)
                store.Save(active);
            else
                store.Create(
                    "asx_claim",
                    new DispatcherDocument
                    {
                        Key = (
                            library
                                ? WorkCoordination.Library(service, row.Id)
                                : WorkCoordination.Site(service, row.Id)
                        ),
                        Status = "Idle",
                    }
                );
            var suspension = new Entity(suspendTable, row.Id)
            {
                RowVersion = row.RowVersion,
                ["asx_approved"] = false,
            };
            if (library)
                suspension["asx_policyapplied"] = false;
            service.Execute(
                new UpdateRequest
                {
                    Target = suspension,
                    ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                }
            );
            return new CatalogResult
            {
                Status = "Suspended; existing access is not revoked",
                CatalogId = row.Id,
            };
        }
        var probe = store.Require<CatalogProbe>("asx_operation", request.Key);
        if (request.Command == "Inspect")
            return Result(probe);
        if (request.Command == "CompleteSiteIdentity")
        {
            if (
                request.CatalogId == Guid.Empty
                || probe.Value.Status != "Captured"
                || request.RowVersion != probe.Row.RowVersion
                || probe.Value.ObservedUtc < clock().AddMinutes(-15)
                || probe.Value.ObservedUtc > clock()
                || probe.Value.ListId != Guid.Empty
            )
                throw new EvaluationBlockedException(
                    "Fresh site observation required for development identity upgrade."
                );
            var current = service.Retrieve(
                "asx_site",
                request.CatalogId,
                new ColumnSet(
                    "asx_nativeid",
                    "asx_webid",
                    "asx_url",
                    "asx_collectionid",
                    "asx_identity"
                )
            );
            var upgradeIdentity = SiteIdentity.Key(
                probe.Value.WebUrl,
                probe.Value.CollectionId,
                probe.Value.WebId
            );
            if (
                current.GetAttributeValue<EntityReference>("asx_nativeid")?.Id
                    != probe.Value.NativeSiteId
                || TemplateStore.Text(current, "asx_url") != probe.Value.WebUrl
                || TemplateStore.Text(current, "asx_webid") != probe.Value.WebId.ToString("D")
                || NativeSite(service, probe.Value.NativeSiteId) != probe.Value.WebUrl
            )
                throw new EvaluationBlockedException(
                    "Identity upgrade cannot reassign an existing site."
                );
            if (
                current.GetAttributeValue<string>("asx_identity") != null
                    && current.GetAttributeValue<string>("asx_identity") != upgradeIdentity
                || current.GetAttributeValue<string>("asx_collectionid") != null
                    && current.GetAttributeValue<string>("asx_collectionid")
                        != probe.Value.CollectionId.ToString("D")
            )
                throw new EvaluationBlockedException("Stored site collection identity differs.");
            service.Execute(
                new UpdateRequest
                {
                    Target = new Entity("asx_site", current.Id)
                    {
                        RowVersion = current.RowVersion,
                        ["asx_collectionid"] = probe.Value.CollectionId.ToString("D"),
                        ["asx_identity"] = upgradeIdentity,
                    },
                    ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                }
            );
            probe.Value.CatalogId = current.Id;
            probe.Value.Status = "Approved";
            store.Save(probe);
            return Result(probe);
        }
        if (
            request.Command != "Approve"
            || request.RowVersion != probe.Row.RowVersion
            || probe.Value.Status != "Captured"
            || probe.Value.ObservedUtc < clock().AddMinutes(-15)
            || probe.Value.ObservedUtc > clock()
            || string.IsNullOrWhiteSpace(request.Name)
            || request.Name.Length > 200
        )
            throw new EvaluationBlockedException(
                "Fresh captured observation, unchanged version and display name are required."
            );
        var claim = store.Find<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        // Approval proceeds while another run writes on the site; saving the writer row
        // serializes it with a concurrent claim.
        if (claim != null)
            store.Save(claim);
        else
            store.Create(
                "asx_claim",
                new DispatcherDocument
                {
                    Key = WorkCoordination.Operation(service, request.Key),
                    Status = "Idle",
                }
            );
        var value = probe.Value;
        if (NativeSite(service, value.NativeSiteId) != value.WebUrl)
            throw new EvaluationBlockedException("Native site changed after observation.");
        bool isLibrary = value.ListId != Guid.Empty;
        string table = isLibrary ? "asx_library" : "asx_site";
        string identity = SiteIdentity.Key(value.WebUrl, value.CollectionId, value.WebId);
        Guid id = isLibrary
            ? SiteIdentity.LibraryId(value.SiteId, value.ListId)
            : DocumentStore.StableId("site:" + identity);
        var query = new QueryExpression(table) { ColumnSet = new ColumnSet(true), TopCount = 2 };
        query.Criteria.AddCondition(
            isLibrary ? "asx_listid" : "asx_identity",
            ConditionOperator.Equal,
            isLibrary ? value.ListId.ToString("D") : identity
        );
        if (isLibrary)
            query.Criteria.AddCondition("asx_siteid", ConditionOperator.Equal, value.SiteId);
        var rows = service.RetrieveMultiple(query);
        if (rows.MoreRecords || rows.Entities.Count > 1)
            throw new EvaluationBlockedException("Catalog identity collision.");
        var old = rows.Entities.SingleOrDefault();
        if (old != null)
            id = old.Id;
        Entity target;
        if (isLibrary)
        {
            var site = service.Retrieve(
                "asx_site",
                value.SiteId,
                new ColumnSet("asx_webid", "asx_collectionid", "asx_approved")
            );
            if (
                !site.GetAttributeValue<bool>("asx_approved")
                || site.GetAttributeValue<string>("asx_webid") != value.WebId.ToString("D")
                || site.GetAttributeValue<string>("asx_collectionid")
                    != value.CollectionId.ToString("D")
            )
                throw new EvaluationBlockedException("Site approval changed.");
            if (value.NativeParentId == Guid.Empty)
                value.NativeParentId = new NativeLocations(service).EnsureLibrary(
                    value.SiteId,
                    value.WebId,
                    value.ListId,
                    value.NativeSiteId,
                    value.WebUrl,
                    value.LibraryRootPath!,
                    request.Name
                );
            new NativeLocations(service).ValidateParent(
                value.NativeParentId,
                value.EntryUrl,
                value.NativeSiteId
            );
            target = new Entity(table, id)
            {
                ["asx_name"] = request.Name,
                ["asx_siteid"] = new EntityReference("asx_site", value.SiteId),
                ["asx_listid"] = value.ListId.ToString("D"),
                ["asx_entryid"] = value.EntryId.ToString("D"),
                ["asx_entryurl"] = value.EntryUrl,
                ["asx_nativeparentid"] = new EntityReference(
                    "sharepointdocumentlocation",
                    value.NativeParentId
                ),
                ["asx_approved"] = true,
                ["asx_policyapplied"] = false,
                ["asx_aclhash"] = SharePointObservations.AclHash(
                    new ODataRows<AclAssignment> { Rows = value.Acl }
                ),
                ["asx_readrole"] = RoleJson(value, 2),
                ["asx_contributerole"] = RoleJson(value, 3),
            };
        }
        else
            target = new Entity(table, id)
            {
                ["asx_name"] = request.Name,
                ["asx_nativeid"] = new EntityReference("sharepointsite", value.NativeSiteId),
                ["asx_webid"] = value.WebId.ToString("D"),
                ["asx_collectionid"] = value.CollectionId.ToString("D"),
                ["asx_identity"] = identity,
                ["asx_url"] = value.WebUrl,
                ["asx_approved"] = true,
            };
        if (old == null)
            service.Create(target);
        else
        {
            if (old.Id != id || (!value.AutoApprove && request.CatalogRowVersion != old.RowVersion))
                throw new EvaluationBlockedException(
                    "Existing catalog requires its current row version for explicit reapproval."
                );
            foreach (
                var pair in target.Attributes.Where(p =>
                    p.Key != "asx_name"
                    && p.Key != "asx_approved"
                    && p.Key != "asx_policyapplied"
                    && p.Key != "asx_aclhash"
                )
            )
                if (!Equals(old.Contains(pair.Key) ? old[pair.Key] : null, pair.Value))
                    throw new EvaluationBlockedException(
                        "Catalog physical identity changed; no reassignment."
                    );
            if (
                isLibrary
                && store
                    .Find<PolicyDocument>("asx_policy", "policy:" + id.ToString("N"))
                    ?.Value.OperationKey != null
            )
                throw new EvaluationBlockedException(
                    "Finish or cancel the queued policy before reapproval."
                );
            var update = new Entity(table, id)
            {
                RowVersion = old.RowVersion,
                ["asx_name"] = request.Name,
                ["asx_approved"] = true,
            };
            if (isLibrary)
            {
                update["asx_policyapplied"] = false;
                update["asx_aclhash"] = target["asx_aclhash"];
            }
            service.Execute(
                new UpdateRequest
                {
                    Target = update,
                    ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                }
            );
        }
        probe.Value.CatalogId = id;
        probe.Value.Status = "Approved";
        store.Save(probe);
        return new CatalogResult
        {
            Status = "Approved",
            Key = value.Key,
            CatalogId = id,
        };
    }

    private static string RoleJson(CatalogProbe value, int type)
    {
        var role = value.Roles.Single(r => r.Type == type);
        return JsonWire.Write(
            new PolicyRole
            {
                Id = role.Id,
                High = role.Permissions.High,
                Low = role.Permissions.Low,
            }
        );
    }

    public static string NativeSite(IOrganizationService service, Guid id)
    {
        if (id == Guid.Empty)
            throw new EvaluationBlockedException("Existing native site required.");
        var site = service.Retrieve(
            "sharepointsite",
            id,
            new ColumnSet("absoluteurl", "statecode")
        );
        if (
            site.GetAttributeValue<OptionSetValue>("statecode")?.Value != 0
            || !Uri.TryCreate(
                site.GetAttributeValue<string>("absoluteurl"),
                UriKind.Absolute,
                out var url
            )
            || url.Scheme != "https"
            || url.UserInfo.Length > 0
            || url.Query.Length > 0
            || url.Fragment.Length > 0
        )
            throw new EvaluationBlockedException(
                "Active absolute native SharePoint site required."
            );
        return url.AbsoluteUri.TrimEnd('/');
    }

    private static CatalogResult Result(StoredRow<CatalogProbe> row) =>
        new CatalogResult
        {
            Status = row.Value.Status,
            Key = row.Value.Key,
            RowVersion = row.Row.RowVersion,
            Observation = row.Value,
            Issue = row.Value.ErrorCode,
            CatalogId = row.Value.Status == "Approved" ? row.Value.CatalogId : Guid.Empty,
        };
}

public sealed class CatalogWorker
{
    private readonly IOrganizationService service;
    private readonly DocumentStore store;
    private readonly Func<DateTime> clock;

    public CatalogWorker(IOrganizationService service, Func<DateTime>? clock = null)
    {
        this.service = service;
        store = new DocumentStore(service);
        this.clock = clock ?? (() => DateTime.UtcNow);
    }

    public WorkerResult Execute(WorkerRequest request, bool transaction)
    {
        if (!transaction)
            throw new EvaluationBlockedException("Catalog probe requires a transaction.");
        var op = store.Require<CatalogProbe>("asx_operation", request.Key);
        if (request.Command == "FailUnclaimed")
            return store.FailUnclaimed<CatalogProbe>(request.Key, request, clock());
        if (request.Command == "Retry" || request.Command == "Cancel")
        {
            // Probes only read, so Retry and Cancel are always safe once the claim expires.
            bool released = store.ReleaseExpired(request.Key, clock(), request.Command);
            if (
                !released
                && op.Value.Status != "Blocked"
                && op.Value.Status != "Pending"
                && op.Value.Status != "RetryWait"
            )
                throw new EvaluationBlockedException(
                    "Only blocked/pending probes can retry or cancel."
                );
            op.Value.Status = request.Command == "Retry" ? "Pending" : "Cancelled";
            op.Value.NextAttemptUtc = null;
            op.Value.RetryCount = 0;
            op.Value.ErrorCode = null;
            store.Save(op);
            return new WorkerResult { Status = op.Value.Status, Key = request.Key };
        }
        if (
            request.Command == "Claim"
            && (
                op.Value.Status == "Discovered"
                || op.Value.Status == "Captured"
                || op.Value.Status == "Cancelled"
                || op.Value.Status == "Approved"
                || op.Value.Status == "Blocked"
            )
        )
            return new WorkerResult { Status = op.Value.Status, Key = request.Key };
        if (CatalogAdministration.NativeSite(service, op.Value.NativeSiteId) != op.Value.WebUrl)
            throw new EvaluationBlockedException("Native site changed during probe.");
        if (
            request.Command == "Claim"
            && !WorkCoordination.HasCapacity(
                service,
                WorkCoordination.Operation(service, request.Key),
                clock()
            )
        )
            return new WorkerResult { Status = "Busy", Key = request.Key };
        var claim = store.Find<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        if (request.Command == "Claim")
        {
            if (op.Value.NextAttemptUtc > clock())
                return new WorkerResult { Status = "RetryWait", Key = request.Key };
            if (string.IsNullOrWhiteSpace(request.RunId) || request.RunId.Length > 150)
                throw new EvaluationBlockedException("Actual run identity required.");
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
                // Probes only read: once the lease expires the next claim takes over.
                if (
                    !claim.Value.RecoveryPermitted
                    && claim.Value.LeaseUntilUtc > clock()
                    && (claim.Value.Token != request.Token || claim.Value.RunId != request.RunId)
                )
                    return new WorkerResult { Status = "Quarantined", Key = request.Key };
                claim.Value.HttpOutstanding = false;
            }
            claim.Value.OperationKey = request.Key;
            claim.Value.RunId = request.RunId;
            claim.Value.Token = Guid.NewGuid();
            claim.Value.RecoveryPermitted = false;
            claim.Value.LeaseUntilUtc = clock().AddMinutes(5);
            claim.Value.Status = "Claimed";
            store.Save(claim);
            return Probe(op, claim.Value, "CatalogWeb");
        }
        if (
            claim == null
            || request.Token == Guid.Empty
            || claim.Value.OperationKey != request.Key
            || claim.Value.RunId != request.RunId
            || claim.Value.Token != request.Token
            || claim.Value.LeaseUntilUtc <= clock()
            || claim.Value.RecoveryPermitted
        )
            throw new EvaluationBlockedException("Stale catalog probe claim.");
        if (request.Command == "Renew")
        {
            claim.Value.LeaseUntilUtc = clock().AddMinutes(5);
            store.Save(claim);
            return new WorkerResult
            {
                Status = "Renewed",
                Key = request.Key,
                Token = claim.Value.Token,
            };
        }
        if (request.Command == "Fail")
        {
            if (TransientFailure.Is(request))
                return Wait(op, claim, null, request.StatusCode, request.ErrorCode);
            // A notice left by an earlier temporary wait is not the cause of this failure.
            if (
                op.Value.ErrorCode?.StartsWith("Waiting to retry", StringComparison.Ordinal) == true
            )
                op.Value.ErrorCode = null;
            op.Value.ErrorCode =
                op.Value.ErrorCode
                ?? "Catalog worker failed before approval. Inspect the failed flow action, then retry after repair.";
            op.Value.Status = "Blocked";
            store.Save(op);
            Release(claim);
            return new WorkerResult { Status = "Blocked", Key = request.Key };
        }
        if (request.Command == "Complete")
        {
            try
            {
                if (op.Value.Status != "Verified")
                    throw new EvaluationBlockedException("Complete catalog readback required.");
                op.Value.Status = op.Value.DiscoverLibraries ? "Discovered" : "Captured";
                op.Value.ObservedUtc = clock();
                store.Save(op);
                Release(claim);
                if (op.Value.AutoApprove)
                {
                    var fresh = store.Require<CatalogProbe>("asx_operation", request.Key);
                    var approved = new CatalogAdministration(service, clock).Execute(
                        new CatalogRequest
                        {
                            Command = "Approve",
                            Key = request.Key,
                            RowVersion = fresh.Row.RowVersion,
                            Name = fresh.Value.DisplayName,
                        },
                        true
                    );
                    return new WorkerResult { Status = "Approved", Key = request.Key };
                }
                return new WorkerResult { Status = op.Value.Status, Key = request.Key };
            }
            catch (EvaluationBlockedException error)
            {
                op = store.Require<CatalogProbe>("asx_operation", request.Key);
                op.Value.Status = "Blocked";
                op.Value.ErrorCode = error.Message;
                store.Save(op);
                var released = store.Require<DispatcherDocument>(
                    "asx_claim",
                    WorkCoordination.Operation(service, request.Key)
                );
                if (released.Value.RunId != null)
                    Release(released);
                return new WorkerResult
                {
                    Status = "Blocked",
                    Key = request.Key,
                    Notices = new[] { error.Message },
                };
            }
        }
        if (
            request.Command != "Observe"
            || request.ProbeId != op.Value.ProbeId
            || request.ProbeId == Guid.Empty
            || request.ProbeKind != op.Value.ProbeKind
        )
            throw new EvaluationBlockedException("Stale or unsupported catalog observation.");
        // Catalog probes only read, so a missing response (0 or 408) waits like a throttle.
        if (
            request.HttpStatus == 429
            || request.HttpStatus >= 500
            || request.HttpStatus == 0
            || request.HttpStatus == 408
        )
            return Wait(op, claim, request.RetryAfter, request.HttpStatus, null);
        try
        {
            switch (op.Value.ProbeKind)
            {
                case "CatalogWeb":
                    var web = SharePointObservations.Body<WebObservation>(request);
                    if (
                        web.Id == Guid.Empty
                        || web.Url.TrimEnd('/') != op.Value.WebUrl
                        || (op.Value.WebId != Guid.Empty && op.Value.WebId != web.Id)
                    )
                        throw new EvaluationBlockedException(
                            "Native and physical site identities differ."
                        );
                    op.Value.WebId = web.Id;
                    return Probe(op, claim.Value, "CatalogCollection");
                case "CatalogCollection":
                    var collection = SharePointObservations.Body<WebObservation>(request);
                    if (
                        collection.Id == Guid.Empty
                        || (
                            op.Value.CollectionId != Guid.Empty
                            && op.Value.CollectionId != collection.Id
                        )
                    )
                        throw new EvaluationBlockedException(
                            "Site collection identity changed; revalidate the selected site."
                        );
                    op.Value.CollectionId = collection.Id;
                    if (op.Value.DiscoverLibraries)
                        return Probe(op, claim.Value, "CatalogLibraries");
                    if (op.Value.ListId != Guid.Empty)
                        return Probe(op, claim.Value, "CatalogLibrary");
                    break;
                case "CatalogLibraries":
                    var choices = SharePointObservations.Body<ODataRows<LibraryChoice>>(request);
                    if (
                        choices.Rows == null
                        || choices.Rows.Length > 50
                        || choices.Rows.Any(c =>
                            c.Id == Guid.Empty || c.BaseTemplate != 101 || c.Hidden
                        )
                        || choices.Rows.Select(c => c.Id).Distinct().Count() != choices.Rows.Length
                    )
                        throw new EvaluationBlockedException(
                            "Library discovery page is incomplete or invalid."
                        );
                    op.Value.Libraries = choices.Rows;
                    op.Value.NextLibraries =
                        choices.Next == null
                            ? null
                            : SecurityPaging.Next(
                                new Uri(op.Value.WebUrl),
                                "_api/web/lists?$select=Id,Title,BaseTemplate,Hidden&$filter=BaseTemplate eq 101 and Hidden eq false&$orderby=Title&$top=50",
                                choices.Next,
                                op.Value.Pages
                            );
                    break;
                case "CatalogLibrary":
                case "CatalogFinal":
                    var library = SharePointObservations.Body<CatalogLibraryObservation>(request);
                    if (
                        library.Id != op.Value.ListId
                        || library.Unique != true
                        || library.Root == null
                        || library.Root.Id == Guid.Empty
                    )
                        throw new EvaluationBlockedException(
                            "Existing library needs its reviewed unique permission boundary."
                        );
                    string entry =
                        new Uri(op.Value.WebUrl).GetLeftPart(UriPartial.Authority)
                        + library.Root.Path;
                    new SharePointTarget(
                        op.Value.WebUrl,
                        op.Value.WebId,
                        op.Value.ListId,
                        library.Root.Path
                    );
                    if (op.Value.NativeParentId != Guid.Empty)
                        new NativeLocations(service).ValidateParent(
                            op.Value.NativeParentId,
                            op.Value.EntryUrl,
                            op.Value.NativeSiteId
                        );
                    else if (string.IsNullOrEmpty(op.Value.EntryUrl))
                        op.Value.EntryUrl = entry;
                    string desired = Uri.UnescapeDataString(new Uri(op.Value.EntryUrl).AbsolutePath)
                        .TrimEnd('/');
                    if (
                        desired != library.Root.Path
                        && !desired.StartsWith(library.Root.Path + "/", StringComparison.Ordinal)
                    )
                        throw new EvaluationBlockedException(
                            "Native entry is outside the selected library."
                        );
                    if (desired.Substring(library.Root.Path.Length).Count(c => c == '/') > 20)
                        throw new EvaluationBlockedException("Entry ancestor depth exceeds 20.");
                    bool final = op.Value.ProbeKind == "CatalogFinal";
                    if (
                        final
                        && (
                            op.Value.LibraryRootId != library.Root.Id
                            || op.Value.LibraryRootPath != library.Root.Path
                        )
                    )
                        throw new EvaluationBlockedException(
                            "Library root changed during observation."
                        );
                    op.Value.LibraryRootId = library.Root.Id;
                    op.Value.LibraryRootPath = library.Root.Path;
                    if (desired == library.Root.Path)
                    {
                        if (final && op.Value.EntryId != library.Root.Id)
                            throw new EvaluationBlockedException("Entry identity changed.");
                        op.Value.EntryId = library.Root.Id;
                        if (final)
                            break;
                        return Probe(op, claim.Value, "CatalogRoles");
                    }
                    op.Value.AncestorPath = desired;
                    return Probe(op, claim.Value, final ? "CatalogEntryFinal" : "CatalogEntry");
                case "CatalogEntry":
                case "CatalogEntryFinal":
                case "CatalogAncestor":
                case "CatalogAncestorFinal":
                    var folder = SharePointObservations.Body<FolderObservation>(request);
                    bool last = op.Value.ProbeKind.EndsWith("Final", StringComparison.Ordinal);
                    bool first = op.Value.ProbeKind.StartsWith(
                        "CatalogEntry",
                        StringComparison.Ordinal
                    );
                    if (folder.Id == Guid.Empty || folder.Path != op.Value.AncestorPath)
                        throw new EvaluationBlockedException(
                            "Approved entry or ancestor folder identity differs."
                        );
                    if (first)
                    {
                        if (last && folder.Id != op.Value.EntryId)
                            throw new EvaluationBlockedException(
                                "Entry identity changed during observation."
                            );
                        op.Value.EntryId = folder.Id;
                    }
                    string ancestor = folder.Path.Substring(0, folder.Path.LastIndexOf('/'));
                    if (ancestor != op.Value.LibraryRootPath)
                    {
                        if (
                            !ancestor.StartsWith(
                                op.Value.LibraryRootPath + "/",
                                StringComparison.Ordinal
                            )
                        )
                            throw new EvaluationBlockedException(
                                "Ancestor escaped the selected library."
                            );
                        op.Value.AncestorPath = ancestor;
                        return Probe(
                            op,
                            claim.Value,
                            last ? "CatalogAncestorFinal" : "CatalogAncestor"
                        );
                    }
                    if (last)
                        break;
                    return Probe(op, claim.Value, "CatalogRoles");
                case "CatalogRoles":
                    var roles = SharePointObservations.Body<ODataRows<SecurityRoleObservation>>(
                        request
                    );
                    if (roles.Rows == null)
                        throw new EvaluationBlockedException("Role page missing.");
                    op.Value.Roles = op.Value.Roles.Concat(roles.Rows).ToArray();
                    if (
                        op.Value.Roles.Length > 100
                        || op.Value.Roles.Select(r => r.Id).Distinct().Count()
                            != op.Value.Roles.Length
                    )
                        throw new EvaluationBlockedException(
                            "Role inventory incomplete or duplicate."
                        );
                    if (roles.Next != null)
                        return Continue(op, claim.Value, roles.Next);
                    if (
                        op.Value.Roles.Count(r => r.Type == 2) != 1
                        || op.Value.Roles.Count(r => r.Type == 3) != 1
                    )
                        throw new EvaluationBlockedException(
                            "Unambiguous Reader and Contributor roles required."
                        );
                    var reader = op.Value.Roles.Single(r => r.Type == 2);
                    var contributor = op.Value.Roles.Single(r => r.Type == 3);
                    SecurityAdministration.ValidateRole(
                        new PolicyRole
                        {
                            Id = reader.Id,
                            High = reader.Permissions?.High ?? "",
                            Low = reader.Permissions?.Low ?? "",
                        },
                        true
                    );
                    SecurityAdministration.ValidateRole(
                        new PolicyRole
                        {
                            Id = contributor.Id,
                            High = contributor.Permissions?.High ?? "",
                            Low = contributor.Permissions?.Low ?? "",
                        },
                        false
                    );
                    return Probe(op, claim.Value, "CatalogAcl");
                case "CatalogAcl":
                    var acl = SharePointObservations.Body<ODataRows<AclAssignment>>(request);
                    if (acl.Rows == null)
                        throw new EvaluationBlockedException("ACL page missing.");
                    op.Value.Acl = op.Value.Acl.Concat(acl.Rows).ToArray();
                    if (
                        op.Value.Acl.Length > 1000
                        || op.Value.Acl.Select(a => a.Member?.Id).Distinct().Count()
                            != op.Value.Acl.Length
                    )
                        throw new EvaluationBlockedException(
                            "ACL inventory incomplete or duplicate."
                        );
                    if (acl.Next != null)
                        return Continue(op, claim.Value, acl.Next);
                    SharePointObservations.AclHash(
                        new ODataRows<AclAssignment> { Rows = op.Value.Acl }
                    );
                    return Probe(op, claim.Value, "CatalogFinal");
                default:
                    throw new EvaluationBlockedException("Unknown catalog probe.");
            }
            op.Value.Status = "Verified";
            store.Save(op);
            return new WorkerResult
            {
                Status = "Verified",
                Key = op.Value.Key,
                Token = claim.Value.Token,
            };
        }
        catch (EvaluationBlockedException error)
        {
            op.Value.Status = "Blocked";
            op.Value.ErrorCode = error.Message;
            store.Save(op);
            Release(claim);
            return new WorkerResult
            {
                Status = "Blocked",
                Key = op.Value.Key,
                Notices = new[] { error.Message },
            };
        }
    }

    private WorkerResult Probe(StoredRow<CatalogProbe> op, DispatcherDocument claim, string kind)
    {
        string endpoint =
            kind == "CatalogWeb" ? "_api/web?$select=Id,Url"
            : kind == "CatalogCollection" ? "_api/site?$select=Id"
            : kind == "CatalogRoles"
                ? "_api/web/roledefinitions?$select=Id,RoleTypeKind,BasePermissions"
            : "_api/web/lists(guid'"
                + op.Value.ListId
                + "')"
                + (
                    kind == "CatalogAcl"
                        ? "/roleassignments?$select=Member/Id,Member/PrincipalType,RoleDefinitionBindings/Id,RoleDefinitionBindings/BasePermissions&$expand=Member,RoleDefinitionBindings"
                        : "?$select=Id,HasUniqueRoleAssignments,RootFolder/UniqueId,RootFolder/ServerRelativeUrl&$expand=RootFolder"
                );
        if (kind == "CatalogLibraries")
            endpoint =
                op.Value.DiscoveryEndpoint
                ?? "_api/web/lists?$select=Id,Title,BaseTemplate,Hidden&$filter=BaseTemplate eq 101 and Hidden eq false&$orderby=Title&$top=50";
        if (
            kind == "CatalogEntry"
            || kind == "CatalogEntryFinal"
            || kind == "CatalogAncestor"
            || kind == "CatalogAncestorFinal"
        )
            endpoint =
                "_api/web/GetFolderByServerRelativePath(decodedUrl='"
                + Uri.EscapeDataString(op.Value.AncestorPath!.Replace("'", "''"))
                + "')?$select=UniqueId,ServerRelativeUrl";
        if (kind == "CatalogAcl")
            op.Value.Acl = Array.Empty<AclAssignment>();
        if (kind == "CatalogRoles")
            op.Value.Roles = Array.Empty<SecurityRoleObservation>();
        op.Value.ProbeKind = kind;
        op.Value.ProbeId = Guid.NewGuid();
        op.Value.Endpoint = endpoint;
        op.Value.Pages = new[] { endpoint };
        op.Value.Status = "Inspecting";
        store.Save(op);
        return Result(op.Value, claim, endpoint);
    }

    private WorkerResult Continue(StoredRow<CatalogProbe> op, DispatcherDocument claim, string next)
    {
        string path = SecurityPaging.Next(
            new Uri(op.Value.WebUrl),
            op.Value.Endpoint,
            next,
            op.Value.Pages
        );
        op.Value.Pages = op.Value.Pages.Concat(new[] { path }).ToArray();
        op.Value.ProbeId = Guid.NewGuid();
        store.Save(op);
        return Result(op.Value, claim, path);
    }

    private static WorkerResult Result(
        CatalogProbe op,
        DispatcherDocument claim,
        string endpoint
    ) =>
        new WorkerResult
        {
            Status = "Read",
            Key = op.Key,
            Token = claim.Token,
            ProbeId = op.ProbeId,
            ProbeKind = op.ProbeKind,
            SiteUrl = op.WebUrl,
            Http = new HttpIntent { RelativeUri = endpoint },
        };

    /// <summary>
    /// Waits after a temporary failure with no attempt cap and releases the writer slot.
    /// </summary>
    private WorkerResult Wait(
        StoredRow<CatalogProbe> op,
        StoredRow<DispatcherDocument> claim,
        string? retryAfter,
        int? statusCode,
        string? errorCode
    )
    {
        DocumentStore.Wait(op.Value, statusCode, errorCode, retryAfter, clock());
        store.Save(op);
        Release(claim);
        return new WorkerResult
        {
            Status = "RetryWait",
            Key = op.Value.Key,
            Notices = new[] { op.Value.ErrorCode! },
        };
    }

    private void Release(StoredRow<DispatcherDocument> claim)
    {
        claim.Value.RunId = null;
        claim.Value.OperationKey = null;
        claim.Value.Token = Guid.Empty;
        claim.Value.RecoveryPermitted = false;
        claim.Value.Status = "Idle";
        store.Save(claim);
    }
}
