using System;
using System.Collections.Generic;
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

    /// <summary>CreateLibrary: the admin accepted the initial team's broader access.</summary>
    [DataMember]
    public bool AcknowledgeBroaderAccess { get; set; }

    /// <summary>
    /// AddLibrary, ProbeLibrary and Approve: the admin saw and accepted that Documents stops the
    /// library's permission inheritance (see CatalogAdministration.InheritanceWarning).
    /// </summary>
    [DataMember]
    public bool BreakInheritance { get; set; }
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

    /// <summary>False when the library inherits its site's permissions; null when unknown.</summary>
    [DataMember(Name = "HasUniqueRoleAssignments")]
    public bool? Unique { get; set; }
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

    /// <summary>The library inherits its site's permissions, as last read.</summary>
    [DataMember]
    public bool Inherits { get; set; }

    /// <summary>The admin accepted that Documents stops the inheritance (AddLibrary consent).</summary>
    [DataMember]
    public bool BreakInheritance { get; set; }

    /// <summary>
    /// Re-point: re-reads an approved site by its web ID, or library by its list GUID and entry
    /// folder ID, and follows a rename or move of the same site or library.
    /// </summary>
    [DataMember]
    public bool Repoint { get; set; }

    /// <summary>Re-point: the site URL or library entry URL before re-pointing.</summary>
    [DataMember]
    public string PreviousUrl { get; set; } = "";

    /// <summary>Re-point: the library's entry folder no longer exists; it is not recreated.</summary>
    [DataMember]
    public bool EntryMissing { get; set; }

    /// <summary>Re-point: what changed, such as "old URL → new URL", for the admin.</summary>
    [DataMember]
    public string[] Changes { get; set; } = Array.Empty<string>();

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

    /// <summary>What a command left in place or changed, for the admin.</summary>
    [DataMember]
    public string[] Notices { get; set; } = Array.Empty<string>();
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
    /// <summary>Shown to the admin before approving a library that inherits its site's permissions.</summary>
    public const string InheritanceWarning =
        "This library inherits permissions from the site. When you approve it, Documents stops the inheritance, keeps a copy of the current site permissions, and then manages team access on it.";

    public const string LibraryGone =
        "This library no longer exists on the site. Remove it, or register the new library.";

    /// <summary>The key prefix of re-point probes; the catalog guard lets only them change addresses.</summary>
    public const string RepointPrefix = "catalogprobe:repoint:";

    /// <summary>The library discovery query, which also reads whether each library inherits.</summary>
    public const string DiscoveryEndpoint =
        "_api/web/lists?$select=Id,Title,BaseTemplate,Hidden,HasUniqueRoleAssignments&$filter=BaseTemplate eq 101 and Hidden eq false&$orderby=Title&$top=50";

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
        if (request.Command == "RetrySetup" || request.Command == "CancelSetup")
            return new LibraryProvisioning(service, clock).ManageSetup(
                request.Key,
                request.Command == "RetrySetup"
            );
        if (request.Command == "RepointLibrary" || request.Command == "RepointSite")
            return QueueRepoint(request);
        if (request.Command == "RemoveLibrary")
            return RemoveLibrary(request.CatalogId);
        if (request.Command == "RemoveSite")
            return RemoveSite(request.CatalogId);
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
                    BreakInheritance = library && request.BreakInheritance,
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
        // Consent, not a guard: the admin chooses whether Documents stops the inheritance.
        if (isLibrary && value.Inherits && !request.BreakInheritance && !value.BreakInheritance)
            throw new EvaluationBlockedException(
                InheritanceWarning + " Approve it with that acknowledgement to continue."
            );
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
                    p.Key != "asx_name" && p.Key != "asx_approved" && p.Key != "asx_policyapplied"
                )
            )
                if (!Equals(old.Contains(pair.Key) ? old[pair.Key] : null, pair.Value))
                    throw new EvaluationBlockedException(
                        "Catalog physical identity changed; no reassignment."
                    );
            // A queued access run (held while the library was suspended) resumes by itself after
            // reapproval. Only a run with a write sent to SharePoint whose result is not known yet
            // must be settled first, since reapproval re-reads the library under it.
            var queued = isLibrary
                ? store
                    .Find<PolicyDocument>("asx_policy", "policy:" + id.ToString("N"))
                    ?.Value.OperationKey
                : null;
            var access =
                queued == null ? null : store.Find<SecurityOperation>("asx_operation", queued);
            if (
                access != null
                && access.Value.ExternalSubmitted
                && !access.Value.ExternalResponseKnown
            )
                throw new EvaluationBlockedException(
                    "The library's access run has a write to SharePoint whose result is not known yet. Wait for it to be read back, or Cancel the access run, then approve the library again."
                );
            var update = new Entity(table, id)
            {
                RowVersion = old.RowVersion,
                ["asx_name"] = request.Name,
                ["asx_approved"] = true,
            };
            // Adding a removed site or library again makes it active.
            if (IsRemoved(old))
            {
                update["statecode"] = new OptionSetValue(0);
                update["statuscode"] = new OptionSetValue(1);
            }
            if (isLibrary)
                update["asx_policyapplied"] = false;
            service.Execute(
                new UpdateRequest
                {
                    Target = update,
                    ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                }
            );
        }
        if (isLibrary && value.Inherits)
            RecordInheritanceConsent(id);
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

    /// <summary>
    /// Queues the reads that re-point a site or library. Identity never changes: a library is
    /// re-read by its list GUID and entry folder ID on its site, a site by its web ID at the
    /// address its SharePoint site record in Dataverse now has.
    /// </summary>
    private CatalogResult QueueRepoint(CatalogRequest request)
    {
        if (request.RequestId == Guid.Empty || request.CatalogId == Guid.Empty)
            throw new EvaluationBlockedException(
                "Select a site or library and a request identity."
            );
        bool library = request.Command == "RepointLibrary";
        string key = RepointPrefix + request.RequestId.ToString("N");
        var existing = store.Find<CatalogProbe>("asx_operation", key);
        if (existing != null)
        {
            if (
                !existing.Value.Repoint
                || existing.Value.CatalogId != request.CatalogId
                || (existing.Value.ListId != Guid.Empty) != library
            )
                throw new EvaluationBlockedException(
                    "Re-point request ID was reused for different intent."
                );
            return Result(existing);
        }
        Entity? row = null;
        Guid siteId = request.CatalogId;
        if (library)
        {
            row = service.Retrieve(
                "asx_library",
                request.CatalogId,
                new ColumnSet(
                    "asx_name",
                    "asx_siteid",
                    "asx_listid",
                    "asx_entryid",
                    "asx_entryurl",
                    "statecode"
                )
            );
            if (IsRemoved(row))
                throw new EvaluationBlockedException(
                    "This library was removed. Add it again to use it."
                );
            siteId =
                row.GetAttributeValue<EntityReference>("asx_siteid")?.Id
                ?? throw new EvaluationBlockedException("Library site missing.");
        }
        var site = service.Retrieve(
            "asx_site",
            siteId,
            new ColumnSet("asx_name", "asx_nativeid", "asx_webid", "asx_collectionid", "asx_url")
        );
        var native =
            site.GetAttributeValue<EntityReference>("asx_nativeid")?.Id
            ?? throw new EvaluationBlockedException("Site has no SharePoint site record.");
        string url = NativeSite(service, native);
        string stored = TemplateStore.Text(site, "asx_url").TrimEnd('/');
        if (library && url != stored)
            throw new EvaluationBlockedException(
                "The site's address changed. Re-point the site first."
            );
        // Documents can call only allowed hosts and registered site records (the transport
        // gate), so a site moved outside them is named here, before anything is queued.
        string host = new Uri(url).Host;
        if (!RuntimeProfile.ReadHosts(service).Contains(host, StringComparer.OrdinalIgnoreCase))
            throw new EvaluationBlockedException(
                "The host "
                    + host
                    + " is not one of the SharePoint hosts allowed in the Runtime panel. Add it there, then re-point again."
            );
        if (!RuntimeProfile.Registered(service, url))
            throw new EvaluationBlockedException(
                "No active SharePoint site record in Dataverse has the address "
                    + url
                    + ". Set the site's SharePoint site record in Dataverse to exactly that address, then re-point again."
            );
        Guid.TryParse(site.GetAttributeValue<string>("asx_collectionid"), out var collection);
        store.Create(
            "asx_operation",
            new CatalogProbe
            {
                Key = key,
                Repoint = true,
                DisplayName = TemplateStore.Text(row ?? site, "asx_name"),
                NativeSiteId = native,
                SiteId = siteId,
                CatalogId = request.CatalogId,
                WebId = Guid.Parse(TemplateStore.Text(site, "asx_webid")),
                CollectionId = collection,
                WebUrl = url,
                ListId = library ? Guid.Parse(TemplateStore.Text(row!, "asx_listid")) : Guid.Empty,
                EntryId = library
                    ? Guid.Parse(TemplateStore.Text(row!, "asx_entryid"))
                    : Guid.Empty,
                EntryUrl = library ? TemplateStore.Text(row!, "asx_entryurl") : "",
                PreviousUrl = library ? TemplateStore.Text(row!, "asx_entryurl") : stored,
            }
        );
        return Result(store.Require<CatalogProbe>("asx_operation", key));
    }

    /// <summary>
    /// Applies a verified re-point in one transaction: the catalog row's addresses and the
    /// library's own Dataverse document location. Folder work already stored follows the new
    /// address at its next step (see WorkerCoordinator.Follow); child record locations are
    /// relative to the library's location, so they keep working.
    /// </summary>
    public CatalogResult ApplyRepoint(string key)
    {
        var probe = store.Require<CatalogProbe>("asx_operation", key);
        var value = probe.Value;
        if (!value.Repoint || value.Status != "Captured")
            throw new EvaluationBlockedException("A verified re-point observation is required.");
        var changes = new List<string>();
        if (value.ListId == Guid.Empty)
            RepointSite(value, changes);
        else
            RepointLibrary(value, changes);
        value.Changes = changes.ToArray();
        value.Status = "Approved";
        store.Save(probe);
        return Result(store.Require<CatalogProbe>("asx_operation", key));
    }

    private void RepointLibrary(CatalogProbe value, List<string> changes)
    {
        var library = service.Retrieve(
            "asx_library",
            value.CatalogId,
            new ColumnSet(
                "asx_siteid",
                "asx_listid",
                "asx_entryurl",
                "asx_nativeparentid",
                "statecode"
            )
        );
        if (IsRemoved(library))
            throw new EvaluationBlockedException(
                "This library was removed. Add it again to use it."
            );
        var site = service.Retrieve("asx_site", value.SiteId, new ColumnSet("asx_url"));
        if (
            library.GetAttributeValue<EntityReference>("asx_siteid")?.Id != value.SiteId
            || TemplateStore.Text(library, "asx_listid") != value.ListId.ToString("D")
        )
            throw new EvaluationBlockedException("Library identity changed during re-point.");
        if (
            TemplateStore.Text(site, "asx_url").TrimEnd('/') != value.WebUrl
            || NativeSite(service, value.NativeSiteId) != value.WebUrl
        )
            throw new EvaluationBlockedException(
                "The site's address changed. Re-point the site first."
            );
        string old = TemplateStore.Text(library, "asx_entryurl");
        if (value.EntryMissing)
        {
            changes.Add(
                "The library's entry folder (ID "
                    + value.EntryId.ToString("D")
                    + ") no longer exists. Documents did not recreate it; folder work for this library waits until the folder is restored."
            );
            return;
        }
        string entry = value.EntryUrl;
        string path = Uri.UnescapeDataString(new Uri(entry).AbsolutePath);
        new SharePointTarget(value.WebUrl, value.WebId, value.ListId, path);
        // Every read and check comes first; nothing is written until all of them pass.
        var native = new NativeLocations(service);
        var location = library.GetAttributeValue<EntityReference>("asx_nativeparentid");
        var locations =
            location == null
                ? Array.Empty<Entity>()
                : native.Follow(location.Id, value.NativeSiteId, value.WebUrl, path);
        if (entry != old)
        {
            service.Execute(
                new UpdateRequest
                {
                    Target = new Entity("asx_library", library.Id)
                    {
                        RowVersion = library.RowVersion,
                        ["asx_entryurl"] = entry,
                    },
                    ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                }
            );
            changes.Add(old + " → " + entry);
        }
        native.Write(locations);
        if (locations.Length > 0)
            changes.Add("The library's Dataverse document location now points to " + entry + ".");
        if (changes.Count == 0)
            changes.Add("The library's address has not changed.");
    }

    private void RepointSite(CatalogProbe value, List<string> changes)
    {
        var site = service.Retrieve(
            "asx_site",
            value.CatalogId,
            new ColumnSet("asx_url", "asx_webid", "asx_collectionid", "asx_identity")
        );
        if (TemplateStore.Text(site, "asx_webid") != value.WebId.ToString("D"))
            throw new EvaluationBlockedException("Site identity changed during re-point.");
        if (NativeSite(service, value.NativeSiteId) != value.WebUrl)
            throw new EvaluationBlockedException(
                "The site's address changed again; re-point again."
            );
        string old = TemplateStore.Text(site, "asx_url").TrimEnd('/');
        string url = value.WebUrl;
        if (old == url)
        {
            changes.Add("The site's address has not changed.");
            return;
        }
        var update = new Entity("asx_site", site.Id)
        {
            RowVersion = site.RowVersion,
            ["asx_url"] = url,
        };
        // The stored identity includes the host, so a move to another host keeps it current.
        if (
            site.GetAttributeValue<string>("asx_identity") != null
            && value.CollectionId != Guid.Empty
        )
            update["asx_identity"] = SiteIdentity.Key(url, value.CollectionId, value.WebId);
        var updates = new List<Entity> { update };
        changes.Add(old + " → " + url);
        string oldPath = Uri.UnescapeDataString(new Uri(old).AbsolutePath).TrimEnd('/');
        string newPath = Uri.UnescapeDataString(new Uri(url).AbsolutePath).TrimEnd('/');
        string origin = new Uri(url).GetLeftPart(UriPartial.Authority);
        var query = new QueryExpression("asx_library")
        {
            ColumnSet = new ColumnSet("asx_entryurl", "asx_listid"),
        };
        query.Criteria.AddCondition("asx_siteid", ConditionOperator.Equal, site.Id);
        // Every library of the site moves with it, however many there are.
        foreach (var library in CompleteQuery.Read(service, query))
        {
            string entry = TemplateStore.Text(library, "asx_entryurl");
            string path = Uri.UnescapeDataString(new Uri(entry).AbsolutePath);
            if (!path.StartsWith(oldPath + "/", StringComparison.Ordinal))
                continue;
            string moved = origin + newPath + path.Substring(oldPath.Length);
            new SharePointTarget(
                url,
                value.WebId,
                Guid.Parse(TemplateStore.Text(library, "asx_listid")),
                Uri.UnescapeDataString(new Uri(moved).AbsolutePath)
            );
            updates.Add(
                new Entity("asx_library", library.Id)
                {
                    RowVersion = library.RowVersion,
                    ["asx_entryurl"] = moved,
                }
            );
            changes.Add(entry + " → " + moved);
        }
        // Every read and check came first; the writes follow together.
        foreach (var target in updates)
            service.Execute(
                new UpdateRequest
                {
                    Target = target,
                    ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                }
            );
        WorkCoordination.MoveWriter(service, old, url);
    }

    /// <summary>Shown after a removal: Documents never deletes or changes anything in SharePoint.</summary>
    public const string RemovalNotice =
        "Nothing was deleted or changed in SharePoint: the library, its folders and permissions, and the Documents groups and their access stay as they are. Remove them in SharePoint if they are no longer needed.";

    /// <summary>A removed catalog row: inactive, unapproved and kept only for history.</summary>
    public static bool IsRemoved(Entity row) =>
        row.GetAttributeValue<OptionSetValue>("statecode")?.Value == 1;

    /// <summary>
    /// Removes a library no Draft or published template revision uses. Unfinished access work is
    /// cancelled with the existing Cancel semantics; folder work stops at its next step (folder
    /// creates already sent may finish). A library nothing refers to any more is deleted;
    /// otherwise it is kept, Removed, for history. Nothing in SharePoint is changed.
    /// </summary>
    private CatalogResult RemoveLibrary(Guid id)
    {
        if (id == Guid.Empty)
            throw new EvaluationBlockedException("Select a library.");
        var library = service.Retrieve(
            "asx_library",
            id,
            new ColumnSet("asx_name", "asx_nativeparentid", "statecode")
        );
        TouchWriter(WorkCoordination.Library(service, id));
        var destinations = new QueryExpression("asx_destination")
        {
            ColumnSet = new ColumnSet("asx_revisionid"),
        };
        destinations.Criteria.AddCondition("asx_libraryid", ConditionOperator.Equal, id);
        var revisions = CompleteQuery
            .Read(service, destinations)
            .Select(d => d.GetAttributeValue<EntityReference>("asx_revisionid")?.Id ?? Guid.Empty)
            .Where(r => r != Guid.Empty)
            .Distinct()
            .ToArray();
        var users = new List<string>();
        foreach (var revisionId in revisions)
        {
            var revision = service.Retrieve(
                "asx_revision",
                revisionId,
                new ColumnSet("asx_status", "asx_templateid")
            );
            var templateRef = revision.GetAttributeValue<EntityReference>("asx_templateid");
            var template =
                templateRef == null
                    ? null
                    : service.Retrieve(
                        "asx_template",
                        templateRef.Id,
                        new ColumnSet("asx_name", "asx_publishedrevisionid")
                    );
            string? state =
                revision.GetAttributeValue<string>("asx_status") == "Draft" ? "draft"
                : template?.GetAttributeValue<EntityReference>("asx_publishedrevisionid")?.Id
                == revisionId
                    ? "published"
                : null;
            // Superseded and retired revisions keep their history and never block removal.
            if (state != null)
                users.Add(
                    "Used by template '"
                        + (template?.GetAttributeValue<string>("asx_name") ?? "unnamed")
                        + "' ("
                        + state
                        + ")."
                );
        }
        if (users.Count > 0)
            throw new EvaluationBlockedException(
                string.Join(" ", users.Distinct())
                    + (
                        users.Count == 1
                            ? " Change the template first."
                            : " Change the templates first."
                    )
            );
        var notices = new List<string>();
        string policyKey = "policy:" + id.ToString("N");
        var policy = store.Find<PolicyDocument>("asx_policy", policyKey);
        if (policy?.Value.OperationKey != null)
        {
            var queued = store.Find<SecurityOperation>("asx_operation", policy.Value.OperationKey);
            if (
                queued != null
                && queued.Value.Status != "Applied"
                && queued.Value.Status != "Cancelled"
            )
            {
                var claim = store.Find<DispatcherDocument>(
                    "asx_claim",
                    WorkCoordination.Operation(service, queued.Value.Key)
                );
                // Remove always succeeds: a run working now cancels itself at its next step
                // (SecurityWorker.StopRemoved); idle work is cancelled here.
                if (
                    claim?.Value.OperationKey == queued.Value.Key
                    && claim.Value.RunId != null
                    && (claim.Value.LeaseUntilUtc > clock() || claim.Value.RecoveryPermitted)
                )
                    notices.Add(
                        "The access run working on this library stops at its next step. Nothing it already changed in SharePoint is undone."
                    );
                else
                {
                    new SecurityWorker(service, clock).Execute(
                        new WorkerRequest { Command = "Cancel", Key = queued.Value.Key },
                        true
                    );
                    notices.Add("Unfinished access work for this library was cancelled.");
                }
            }
            policy = store.Require<PolicyDocument>("asx_policy", policyKey);
        }
        var nativeParent = library.GetAttributeValue<EntityReference>("asx_nativeparentid");
        bool referenced =
            revisions.Length > 0
            || policy != null
            || nativeParent != null && Records(nativeParent.Id);
        library = service.Retrieve("asx_library", id, new ColumnSet("asx_name", "statecode"));
        notices.Add(RemovalNotice);
        if (!referenced)
        {
            service.Execute(
                new DeleteRequest
                {
                    Target = new EntityReference("asx_library", id)
                    {
                        RowVersion = library.RowVersion,
                    },
                    ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                }
            );
            return new CatalogResult
            {
                Status = "Deleted",
                CatalogId = id,
                Notices = notices.ToArray(),
            };
        }
        if (policy != null)
        {
            // Access sync for a removed library stops; its teams no longer schedule it.
            policy.Value.Status = "Removed";
            store.Save(policy);
            foreach (var team in policy.Value.ManagedTeams)
            {
                var reference = store.Find<PolicyTeamReference>(
                    "asx_policyentry",
                    "policyteam:" + id.ToString("N") + ":" + team.ToString("N")
                );
                if (reference != null && reference.Value.Status != "Inactive")
                {
                    reference.Value.Status = "Inactive";
                    store.Save(reference);
                }
            }
        }
        if (!IsRemoved(library))
            Retire("asx_library", library);
        return new CatalogResult
        {
            Status = "Removed",
            CatalogId = id,
            Notices = notices.ToArray(),
        };
    }

    /// <summary>
    /// Removes a site that has no libraries left. A site nothing refers to any more is deleted;
    /// one still referred to by removed libraries or library setup history is kept, Removed.
    /// </summary>
    private CatalogResult RemoveSite(Guid id)
    {
        if (id == Guid.Empty)
            throw new EvaluationBlockedException("Select a site.");
        var site = service.Retrieve("asx_site", id, new ColumnSet("asx_url", "statecode"));
        TouchWriter(WorkCoordination.Site(service, id));
        var query = new QueryExpression("asx_library")
        {
            ColumnSet = new ColumnSet("asx_name", "statecode"),
        };
        query.Criteria.AddCondition("asx_siteid", ConditionOperator.Equal, id);
        var libraries = CompleteQuery.Read(service, query);
        var active = libraries.Where(l => !IsRemoved(l)).ToArray();
        if (active.Length > 0)
            throw new EvaluationBlockedException(
                "Remove the site's libraries first: "
                    + string.Join(", ", active.Select(l => l.GetAttributeValue<string>("asx_name")))
                    + "."
            );
        var setups = new QueryExpression("asx_operation")
        {
            ColumnSet = new ColumnSet(false),
            TopCount = 1,
        };
        setups.Criteria.AddCondition("asx_workkind", ConditionOperator.Equal, "LibrarySetup");
        setups.Criteria.AddCondition(
            "asx_siteurl",
            ConditionOperator.Equal,
            TemplateStore.Text(site, "asx_url")
        );
        bool referenced =
            libraries.Length > 0 || service.RetrieveMultiple(setups).Entities.Count > 0;
        var notices = new[]
        {
            "Nothing was deleted or changed in SharePoint: the site, its libraries and the Documents groups stay as they are.",
        };
        if (!referenced)
        {
            service.Execute(
                new DeleteRequest
                {
                    Target = new EntityReference("asx_site", id) { RowVersion = site.RowVersion },
                    ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                }
            );
            return new CatalogResult
            {
                Status = "Deleted",
                CatalogId = id,
                Notices = notices,
            };
        }
        if (!IsRemoved(site))
            Retire("asx_site", site);
        return new CatalogResult
        {
            Status = "Removed",
            CatalogId = id,
            Notices = notices,
        };
    }

    /// <summary>
    /// Saves the site's writer row, as suspension does, so a removal and a concurrent claim or
    /// write permit on the same site serialize on its version: whichever commits second
    /// re-reads and sees the other.
    /// </summary>
    private void TouchWriter(string key)
    {
        var writer = store.Find<DispatcherDocument>("asx_claim", key);
        if (writer != null)
            store.Save(writer);
        else
            store.Create("asx_claim", new DispatcherDocument { Key = key, Status = "Idle" });
    }

    // Folder locations Documents made for records under the library's own location.
    private bool Records(Guid nativeParent)
    {
        var query = new QueryExpression("sharepointdocumentlocation")
        {
            ColumnSet = new ColumnSet(false),
            TopCount = 1,
        };
        query.Criteria.AddCondition("parentsiteorlocation", ConditionOperator.Equal, nativeParent);
        query.Criteria.AddCondition(
            "description",
            ConditionOperator.BeginsWith,
            "AscentixDocuments:"
        );
        return service.RetrieveMultiple(query).Entities.Count > 0;
    }

    // The Removed state: unapproved and inactive, hidden from pickers, planning and access sync.
    private void Retire(string table, Entity row)
    {
        var target = new Entity(table, row.Id)
        {
            RowVersion = row.RowVersion,
            ["asx_approved"] = false,
            ["statecode"] = new OptionSetValue(1),
            ["statuscode"] = new OptionSetValue(2),
        };
        if (table == "asx_library")
            target["asx_policyapplied"] = false;
        service.Execute(
            new UpdateRequest
            {
                Target = target,
                ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
            }
        );
    }

    /// <summary>
    /// Keeps the admin's consent on the library's access policy until the first access run
    /// stops the inheritance. A library with no policy yet gets one that, like a missing policy,
    /// has no access applied.
    /// </summary>
    private void RecordInheritanceConsent(Guid library)
    {
        string key = "policy:" + library.ToString("N");
        var policy = store.Find<PolicyDocument>("asx_policy", key);
        if (policy == null)
        {
            store.Create(
                "asx_policy",
                new PolicyDocument
                {
                    Key = key,
                    Status = "Missing",
                    LibraryId = library,
                    BreakInheritance = true,
                }
            );
            return;
        }
        policy.Value.BreakInheritance = true;
        store.Save(policy);
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
            // A re-point that failed while applying was rolled back; its message is the cause.
            if (op.Value.Repoint && !string.IsNullOrWhiteSpace(request.Error))
                op.Value.ErrorCode = Bounded(request.Error!);
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
                if (op.Value.Repoint)
                    return ApplyRepoint(request.Key);
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
                // A re-point is all or nothing: its failure rolls back, then Fail records it.
                if (op.Value.Repoint)
                    throw;
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
                    if (op.Value.Repoint && web.Id != Guid.Empty && web.Id != op.Value.WebId)
                        throw new EvaluationBlockedException(
                            "A different site now answers at "
                                + op.Value.WebUrl
                                + ". Re-point follows the same site only; register the other site separately."
                        );
                    if (op.Value.Repoint && web.Url.TrimEnd('/') != op.Value.WebUrl)
                        throw new EvaluationBlockedException(
                            "SharePoint reports this site at "
                                + web.Url.TrimEnd('/')
                                + ". Update the site's SharePoint site record in Dataverse to that address, then re-point again."
                        );
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
                    if (op.Value.Repoint && op.Value.ListId != Guid.Empty)
                        return Probe(op, claim.Value, "RepointLibrary");
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
                                CatalogAdministration.DiscoveryEndpoint,
                                choices.Next,
                                op.Value.Pages
                            );
                    break;
                case "CatalogLibrary":
                case "CatalogFinal":
                    var library = SharePointObservations.Body<CatalogLibraryObservation>(request);
                    if (
                        library.Id != op.Value.ListId
                        || library.Root == null
                        || library.Root.Id == Guid.Empty
                    )
                        throw new EvaluationBlockedException(
                            "The library or its root folder could not be read."
                        );
                    // An inheriting library is registered with the admin's consent; approval
                    // asks for it (InheritanceWarning).
                    op.Value.Inherits = library.Unique != true;
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
                case "RepointLibrary":
                    // The list GUID survives a rename; a 404 means the library is gone.
                    if (request.HttpStatus == 404)
                        throw new EvaluationBlockedException(CatalogAdministration.LibraryGone);
                    var moved = SharePointObservations.Body<CatalogLibraryObservation>(request);
                    if (
                        moved.Id != op.Value.ListId
                        || moved.Root == null
                        || moved.Root.Id == Guid.Empty
                        || string.IsNullOrWhiteSpace(moved.Root.Path)
                    )
                        throw new EvaluationBlockedException(CatalogAdministration.LibraryGone);
                    op.Value.LibraryRootId = moved.Root.Id;
                    op.Value.LibraryRootPath = moved.Root.Path;
                    if (op.Value.EntryId == moved.Root.Id)
                    {
                        op.Value.EntryUrl = Origin(op.Value) + moved.Root.Path;
                        break;
                    }
                    return Probe(op, claim.Value, "RepointEntry");
                case "RepointEntry":
                    // The entry folder is re-read by its unique ID and followed; a missing one
                    // is reported, never recreated.
                    if (request.HttpStatus == 404)
                    {
                        op.Value.EntryMissing = true;
                        break;
                    }
                    var entryFolder = SharePointObservations.Body<FolderObservation>(request);
                    if (
                        entryFolder.Id != op.Value.EntryId
                        || !entryFolder.Path.StartsWith(
                            op.Value.LibraryRootPath + "/",
                            StringComparison.Ordinal
                        )
                    )
                        throw new EvaluationBlockedException(
                            "The library's entry folder moved outside the library. Remove the library, or register it again."
                        );
                    op.Value.EntryUrl = Origin(op.Value) + entryFolder.Path;
                    break;
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
                    // The library's role assignments are not read: approval never gates on
                    // entries Documents does not own.
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

    /// <summary>
    /// Applies a verified re-point in this transaction. Any failure propagates, so Dataverse
    /// rolls the whole step back with no partial write; the flow's failure branch then reports
    /// it with Fail, which records the Blocked probe and its message in a separate step.
    /// </summary>
    private WorkerResult ApplyRepoint(string key)
    {
        new CatalogAdministration(service, clock).ApplyRepoint(key);
        return new WorkerResult { Status = "Approved", Key = key };
    }

    private static string Origin(CatalogProbe op) =>
        new Uri(op.WebUrl).GetLeftPart(UriPartial.Authority);

    private static string Bounded(string text)
    {
        text = new string(text.Where(c => !char.IsControl(c)).ToArray()).Trim();
        return text.Length > 600 ? text.Substring(0, 600) + "..." : text;
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
                + "?$select=Id,HasUniqueRoleAssignments,RootFolder/UniqueId,RootFolder/ServerRelativeUrl&$expand=RootFolder";
        if (kind == "CatalogLibraries")
            endpoint = op.Value.DiscoveryEndpoint ?? CatalogAdministration.DiscoveryEndpoint;
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
        if (kind == "RepointEntry")
            endpoint =
                "_api/web/GetFolderById('"
                + op.Value.EntryId.ToString("D")
                + "')?$select=UniqueId,ServerRelativeUrl";
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
