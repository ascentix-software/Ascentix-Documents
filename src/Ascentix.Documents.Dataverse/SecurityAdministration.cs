using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public sealed class SecurityCatalog
{
    public Guid LibraryId,
        SiteId;
    public SharePointTarget Target { get; }
    public Entity Library { get; }

    public SecurityCatalog(IOrganizationService service, Guid libraryId)
    {
        Library = service.Retrieve(
            "asx_library",
            libraryId,
            new ColumnSet(
                "asx_siteid",
                "asx_listid",
                "asx_entryurl",
                "asx_approved",
                "asx_policyrevision",
                "asx_policyapplied",
                "asx_readrole",
                "asx_contributerole"
            )
        );
        SiteId =
            Library.GetAttributeValue<EntityReference>("asx_siteid")?.Id
            ?? throw new EvaluationBlockedException("Library site missing.");
        LibraryId = libraryId;
        var site = service.Retrieve(
            "asx_site",
            SiteId,
            new ColumnSet("asx_url", "asx_webid", "asx_approved")
        );
        if (
            !Library.GetAttributeValue<bool>("asx_approved")
            || !site.GetAttributeValue<bool>("asx_approved")
        )
            throw new EvaluationBlockedException("Security destination is not approved.");
        var url = new Uri(TemplateStore.Text(Library, "asx_entryurl"));
        if (
            url.GetLeftPart(UriPartial.Authority)
            != new Uri(TemplateStore.Text(site, "asx_url")).GetLeftPart(UriPartial.Authority)
        )
            throw new EvaluationBlockedException("Library origin mismatch.");
        Target = new SharePointTarget(
            TemplateStore.Text(site, "asx_url"),
            Guid.Parse(TemplateStore.Text(site, "asx_webid")),
            Guid.Parse(TemplateStore.Text(Library, "asx_listid")),
            Uri.UnescapeDataString(url.AbsolutePath)
        );
    }

    public static void UpdateLibrary(
        IOrganizationService service,
        Entity old,
        Guid generation,
        bool applied
    )
    {
        if (string.IsNullOrEmpty(old.RowVersion))
            throw new EvaluationBlockedException("Library approval version missing.");
        var target = new Entity(old.LogicalName, old.Id)
        {
            RowVersion = old.RowVersion,
            ["asx_policyrevision"] = generation.ToString("D"),
            ["asx_policyapplied"] = applied,
        };
        service.Execute(
            new UpdateRequest
            {
                Target = target,
                ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
            }
        );
    }
}

public sealed class SecurityAdministration
{
    private readonly IOrganizationService service;
    private readonly DocumentStore store;

    public SecurityAdministration(IOrganizationService service)
    {
        this.service = service;
        store = new DocumentStore(service);
    }

    public SecurityResult Execute(
        SecurityRequest request,
        bool transaction,
        bool workerRefresh = false
    )
    {
        if (!transaction)
            throw new EvaluationBlockedException("Security administration requires a transaction.");
        if (request.Command == "RegisterTeam")
        {
            if (request.TeamId == Guid.Empty)
                throw new EvaluationBlockedException("Team ID required.");
            Entity? team = null;
            if (request.Enabled)
            {
                team =
                    TeamDirectory.Find(service, request.TeamId)
                    ?? throw new EvaluationBlockedException(TeamDirectory.DeletedRefusal);
                TeamPrincipal.Validate(team, request.AcknowledgeBroaderAccess);
            }
            string key = "team:" + request.TeamId.ToString("N");
            var old = store.Find<TeamRegistration>("asx_teamregistration", key);
            if (old == null)
            {
                if (!string.IsNullOrEmpty(request.RowVersion))
                    throw new EvaluationBlockedException("Registration no longer exists.");
                store.Create(
                    "asx_teamregistration",
                    new TeamRegistration
                    {
                        Key = key,
                        TeamId = request.TeamId,
                        Enabled = request.Enabled,
                        Group = team == null ? (bool?)null : TeamPrincipal.IsGroup(team),
                        Name = team == null ? null : TeamPrincipal.Name(team),
                        Status = request.Enabled ? "Enabled" : "Revoking",
                    }
                );
            }
            else
            {
                Version(old.Row, request.RowVersion);
                old.Value.Enabled = request.Enabled;
                if (team != null)
                {
                    old.Value.Group = TeamPrincipal.IsGroup(team);
                    old.Value.Name = TeamPrincipal.Name(team);
                }
                old.Value.Status = request.Enabled ? "Enabled" : "Revoking";
                store.Save(old);
            }
            var current = store.Require<TeamRegistration>("asx_teamregistration", key);
            // Opt-out is durable, never a deleted mapping. The security dispatcher schedules affected policies below.
            return new SecurityResult
            {
                Status = current.Value.Status,
                RowVersion = current.Row.RowVersion,
            };
        }
        if (request.LibraryId == Guid.Empty)
            throw new EvaluationBlockedException("Library ID required.");
        string policyKey = "policy:" + request.LibraryId.ToString("N");
        var existing = store.Find<PolicyDocument>("asx_policy", policyKey);
        if (request.Command == "GetPolicy")
            return Result(existing);
        // Before the catalog check: a stuck run of a suspended library can be cancelled too.
        if (request.Command == "RetryAccessRun" || request.Command == "CancelAccessRun")
            return ManageRun(existing, request);
        if (
            (request.Command == "ApplyPolicy" || workerRefresh && request.Command == "ApplyPending")
            && existing != null
        )
            existing = ReplaceQueuedRun(existing, request);
        if (request.Command == "ApplyPending")
        {
            // The next access review applies what the admin applied while the run could not
            // be replaced; until it can be, the change keeps waiting.
            if (existing == null || !existing.Value.ApplyPending)
                return Result(existing);
            if (existing.Value.OperationKey != null)
                return Result(existing);
            request.RowVersion = existing.Row.RowVersion;
        }
        var catalog = new SecurityCatalog(service, request.LibraryId);
        string? expectedVersion = request.RowVersion;
        if (request.Command == "SavePolicy" || request.Command == "ApplyPolicy")
        {
            if (request.ReadRole == null || request.ReadRole.Id == 0)
                request.ReadRole = JsonWire.Read<PolicyRole>(
                    TemplateStore.Text(catalog.Library, "asx_readrole")
                );
            if (request.ContributeRole == null || request.ContributeRole.Id == 0)
                request.ContributeRole = JsonWire.Read<PolicyRole>(
                    TemplateStore.Text(catalog.Library, "asx_contributerole")
                );
            Validate(request.Entries, request.ReadRole, request.ContributeRole);
            // A team deleted in Dataverse leaves the library's teams, whatever access its row
            // still shows: the other teams' changes go on, and the access run queued below
            // removes Documents' grant for the deleted team's group (see the queued entries).
            var teams = TeamDirectory.Read(service, request.Entries.Select(e => e.TeamId));
            foreach (var deleted in request.Entries.Where(e => !teams.ContainsKey(e.TeamId)))
                TeamDirectory.Retire(store, deleted.TeamId);
            request.Entries = request.Entries.Where(e => teams.ContainsKey(e.TeamId)).ToArray();
            foreach (var entry in request.Entries.Where(e => e.Access != "None"))
            {
                // Saving a draft validates eligibility; only applying explicitly onboards syncing.
                var team = teams[entry.TeamId];
                TeamPrincipal.Validate(team, request.AcknowledgeBroaderAccess);
                if (request.Command == "ApplyPolicy")
                {
                    // teamtype is fixed when a team is created, so the kind recorded here stays
                    // true; membership events for group teams are then skipped at the source.
                    bool group = TeamPrincipal.IsGroup(team);
                    string name = TeamPrincipal.Name(team);
                    string teamKey = "team:" + entry.TeamId.ToString("N");
                    var registration = store.Find<TeamRegistration>(
                        "asx_teamregistration",
                        teamKey
                    );
                    if (registration == null)
                        store.Create(
                            "asx_teamregistration",
                            new TeamRegistration
                            {
                                Key = teamKey,
                                TeamId = entry.TeamId,
                                Enabled = true,
                                Group = group,
                                Name = name,
                                Status = "Enabled",
                            }
                        );
                    else if (
                        !registration.Value.Enabled
                        || registration.Value.Group != group
                        || registration.Value.Name != name
                    )
                    {
                        registration.Value.Enabled = true;
                        registration.Value.Group = group;
                        registration.Value.Name = name;
                        registration.Value.Status = "Enabled";
                        store.Save(registration);
                    }
                }
            }
            if (existing == null)
            {
                if (!string.IsNullOrEmpty(request.RowVersion))
                    throw new EvaluationBlockedException("Policy no longer exists.");
                store.Create(
                    "asx_policy",
                    new PolicyDocument
                    {
                        Key = policyKey,
                        Status = "Draft",
                        LibraryId = request.LibraryId,
                        Desired = request.Entries,
                        ReadRole = request.ReadRole,
                        ContributeRole = request.ContributeRole,
                    }
                );
            }
            else
            {
                Version(existing.Row, request.RowVersion);
                // A queued or running access run keeps its own copy of the teams, so the team
                // list may change meanwhile; the run's status stays until it ends.
                existing.Value.Desired = request.Entries;
                existing.Value.ReadRole = request.ReadRole;
                existing.Value.ContributeRole = request.ContributeRole;
                if (existing.Value.OperationKey == null)
                    existing.Value.Status = "Draft";
                store.Save(existing);
            }
            existing = store.Require<PolicyDocument>("asx_policy", policyKey);
            if (request.Command == "SavePolicy")
                return Result(existing);
            if (existing.Value.OperationKey != null)
            {
                // The run in progress could not be replaced yet: a flow holds it, or SharePoint
                // has not answered its write. The change waits and the next access review, a
                // minute later, applies it as soon as it can.
                existing.Value.ApplyPending = true;
                if (request.BreakInheritance)
                    existing.Value.BreakInheritance = true;
                existing.Value.NextReviewUtc = DateTime.UtcNow;
                store.Save(existing);
                return Result(store.Require<PolicyDocument>("asx_policy", policyKey));
            }
            expectedVersion = existing.Row.RowVersion; // Apply continues in this same transaction; any failure rolls back the save.
        }
        if (
            (
                request.Command != "QueuePolicy"
                && request.Command != "ApplyPolicy"
                && !(workerRefresh && request.Command == "RefreshPolicy")
                && !(workerRefresh && request.Command == "ApplyPending")
            )
            || existing == null
        )
            throw new EvaluationBlockedException("Unsupported security command or missing policy.");
        Version(existing.Row, expectedVersion);
        if (existing.Value.OperationKey != null)
            throw new EvaluationBlockedException("A policy generation is already queued.");
        // Serialize queuing with a concurrent claim. Queuing writes nothing to SharePoint; the
        // queued operation takes the writer claim itself later, so an active writer is no reason
        // to refuse.
        var dispatcher = store.Find<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Library(service, request.LibraryId)
        );
        if (dispatcher == null)
        {
            store.Create(
                "asx_claim",
                new DispatcherDocument
                {
                    Key = WorkCoordination.Library(service, request.LibraryId),
                    Status = "Idle",
                }
            );
            dispatcher = store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Library(service, request.LibraryId)
            );
        }
        if (dispatcher != null)
            store.Save(dispatcher);
        // A scheduled refresh repeats what was applied; an admin's apply, also one that waited
        // for the run before it, applies the teams the admin chose.
        bool admin = !workerRefresh || request.Command == "ApplyPending";
        var desired = admin ? existing.Value.Desired : existing.Value.Approved;
        var queued = desired
            .Concat(
                existing
                    .Value.ManagedTeams.Where(team => !desired.Any(d => d.TeamId == team))
                    .Select(team => new PolicyEntry { TeamId = team, Access = "None" })
            )
            .ToArray();
        // A team deleted in Dataverse keeps no access, also when the scheduled refresh repeats
        // what was applied: its entry asks for none, so the run removes Documents' grant for its
        // Documents group, and its registration is retired. The group itself stays.
        var live = TeamDirectory.Read(service, queued.Select(e => e.TeamId));
        var entries = queued
            .Select(e =>
                live.ContainsKey(e.TeamId)
                    ? e
                    : new PolicyEntry { TeamId = e.TeamId, Access = "None" }
            )
            .ToArray();
        foreach (var entry in entries)
        {
            if (live.TryGetValue(entry.TeamId, out var team))
                TeamDirectory.Remember(store, team);
            else
                TeamDirectory.Retire(store, entry.TeamId);
            new TeamSnapshotReader(service).Read(entry.TeamId);
        }
        var generation = Guid.NewGuid();
        string operationKey = "policywork:" + generation.ToString("N");
        // The admin's consent is used by this run only; a later reset to inheritance asks again.
        bool breakInheritance = request.BreakInheritance || existing.Value.BreakInheritance;
        existing.Value.BreakInheritance = false;
        if (breakInheritance)
            existing.Value.Inherits = false;
        if (admin)
        {
            existing.Value.ApprovedReadRole = existing.Value.ReadRole;
            existing.Value.ApprovedContributeRole = existing.Value.ContributeRole;
        }
        existing.Value.Approved = entries;
        existing.Value.ManagedTeams = existing
            .Value.ManagedTeams.Concat(entries.Select(e => e.TeamId))
            .Distinct()
            .ToArray();
        existing.Value.Generation = generation;
        existing.Value.Queued = entries;
        existing.Value.OperationKey = operationKey;
        existing.Value.ApplyPending = false;
        existing.Value.Status = "Queued";
        store.Save(existing);
        store.Create(
            "asx_operation",
            new SecurityOperation
            {
                Key = operationKey,
                LibraryId = request.LibraryId,
                SiteId = catalog.SiteId,
                WebId = catalog.Target.WebId,
                ListId = catalog.Target.ListId,
                WebUrl = catalog.Target.Web.AbsoluteUri,
                PolicyKey = policyKey,
                PolicyRevision = generation,
                Entries = entries,
                ReadRole = existing.Value.ApprovedReadRole,
                ContributeRole = existing.Value.ApprovedContributeRole,
                BreakInheritance = breakInheritance,
            }
        );
        foreach (var entry in entries)
        {
            string referenceKey =
                "policyteam:" + request.LibraryId.ToString("N") + ":" + entry.TeamId.ToString("N");
            var reference = store.Find<PolicyTeamReference>("asx_policyentry", referenceKey);
            if (reference == null)
                store.Create(
                    "asx_policyentry",
                    new PolicyTeamReference
                    {
                        Key = referenceKey,
                        TeamId = entry.TeamId,
                        PolicyKey = policyKey,
                        Status = "Active",
                    }
                );
            else
            {
                reference.Value.Status = "Active";
                store.Save(reference);
            }
        }
        SecurityCatalog.UpdateLibrary(service, catalog.Library, generation, false);
        return Result(store.Require<PolicyDocument>("asx_policy", policyKey));
    }

    /// <summary>
    /// The admin's Apply access replaces the library's queued access run, the same way a newer
    /// team snapshot replaces idle work in SecurityRefresh: the run is cancelled with the
    /// existing Cancel semantics (nothing in SharePoint is undone) and the apply continues in
    /// this transaction, so the admin's newer intent wins. A run with a write SharePoint has not
    /// answered, or one a flow holds right now, is never replaced: the change then waits
    /// (ApplyPending) and the next access review applies it once the run can be replaced.
    /// </summary>
    private StoredRow<PolicyDocument> ReplaceQueuedRun(
        StoredRow<PolicyDocument> policy,
        SecurityRequest request
    )
    {
        if (policy.Value.OperationKey == null)
            return policy;
        var queued = store.Require<SecurityOperation>("asx_operation", policy.Value.OperationKey);
        if (queued.Value.ExternalSubmitted && !queued.Value.ExternalResponseKnown)
            return policy;
        var claim = store.Find<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, queued.Value.Key)
        );
        if (
            claim?.Value.OperationKey == queued.Value.Key
            && claim.Value.RunId != null
            && claim.Value.LeaseUntilUtc > DateTime.UtcNow
        )
            return policy;
        Version(policy.Row, request.RowVersion);
        new SecurityWorker(service).Execute(
            new WorkerRequest { Command = "Cancel", Key = queued.Value.Key },
            true
        );
        var replaced = store.Require<SecurityOperation>("asx_operation", queued.Value.Key);
        replaced.Value.ErrorCode = "Replaced by a newer access change.";
        store.Save(replaced);
        var current = store.Require<PolicyDocument>("asx_policy", policy.Value.Key);
        request.RowVersion = current.Row.RowVersion;
        return current;
    }

    /// <summary>
    /// Retry or Cancel of the library's queued access run, from the library in Sites. It acts
    /// only on the run the admin saw, and keeps the operator rules (SecurityWorker.Manage): a
    /// run a flow still holds waits for its claim to expire, Retry reads back any write whose
    /// outcome is unknown before writing again, and Cancel never undoes anything in SharePoint.
    /// </summary>
    private SecurityResult ManageRun(StoredRow<PolicyDocument>? policy, SecurityRequest request)
    {
        if (
            policy == null
            || string.IsNullOrEmpty(request.OperationKey)
            || policy.Value.OperationKey != request.OperationKey
        )
            throw new EvaluationBlockedException(
                "This library's access run changed. Refresh the library and try again."
            );
        new SecurityWorker(service).Execute(
            new WorkerRequest
            {
                Command = request.Command == "RetryAccessRun" ? "Retry" : "Cancel",
                Key = request.OperationKey!,
            },
            true
        );
        return Result(store.Require<PolicyDocument>("asx_policy", policy.Value.Key));
    }

    public static void Validate(PolicyEntry[] entries, PolicyRole read, PolicyRole contribute)
    {
        if (
            entries == null
            || entries.Length > Domain.Bounds.TeamEntries
            || entries.Any(e =>
                e.TeamId == Guid.Empty || !new[] { "None", "Read", "Contribute" }.Contains(e.Access)
            )
            || entries.Select(e => e.TeamId).Distinct().Count() != entries.Length
        )
            throw new EvaluationBlockedException(
                "At most ten unique team policy entries with named access levels required."
            );
        ValidateRole(read, true);
        ValidateRole(contribute, false);
        if (read.Id == contribute.Id)
            throw new EvaluationBlockedException("Read and Contribute roles must be distinct.");
    }

    // Rights that administer the site rather than its documents, as SharePoint's PermissionKind
    // names them, with their one-based bit position in the 64-bit permission mask
    // (https://learn.microsoft.com/dotnet/api/microsoft.sharepoint.client.permissionkind).
    // Documents grants teams only Read and Contribute, so a level carrying any of these is
    // refused. Every other right is the site's own choice: a customized level is accepted.
    // ManageAlerts (bit 39) stays accepted: it only lets a member manage other users' alerts,
    // that is who is e-mailed about changes they can already see, and grants no access.
    private static readonly (string Name, int Bit)[] AdministrativeRights =
    {
        ("ManageLists", 12),
        // Adds and edits pages and web parts, which can carry script.
        ("AddAndCustomizePages", 19),
        ("ManageSubwebs", 24),
        ("CreateGroups", 25),
        ("ManagePermissions", 26),
        ("ManageWeb", 31),
        ("EnumeratePermissions", 63),
    };

    /// <summary>
    /// Accepts the site's Read or Contribute permission level, customized or not, unless it
    /// carries administrative rights; the refusal names them. Read again at every access run.
    /// </summary>
    public static void ValidateRole(PolicyRole role, bool read)
    {
        if (
            role == null
            || role.Id <= 0
            || !uint.TryParse(
                role.High,
                System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture,
                out var high
            )
            || !uint.TryParse(
                role.Low,
                System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture,
                out var low
            )
            || role.High != high.ToString(System.Globalization.CultureInfo.InvariantCulture)
            || role.Low != low.ToString(System.Globalization.CultureInfo.InvariantCulture)
        )
            throw new EvaluationBlockedException(
                "SharePoint returned an unreadable permission mask for the site's "
                    + (read ? "Read" : "Contribute")
                    + " permission level."
            );
        ulong mask = ((ulong)high << 32) | low;
        // FullMask (Full Control) sets every bit SharePoint defines.
        var held =
            high == int.MaxValue && low == uint.MaxValue
                ? new[] { "FullMask" }
                : AdministrativeRights
                    .Where(right => (mask & (1UL << (right.Bit - 1))) != 0)
                    .Select(right => right.Name)
                    .ToArray();
        if (held.Length > 0)
            throw new EvaluationBlockedException(
                "The site's "
                    + (read ? "Read" : "Contribute")
                    + " permission level includes administrative rights ("
                    + string.Join(", ", held)
                    + "). Documents grants teams only Read and Contribute, so remove these rights from the permission level in SharePoint, then try again."
            );
    }

    private static void Version(Entity row, string? expected)
    {
        if (string.IsNullOrEmpty(expected) || row.RowVersion != expected)
            throw new EvaluationBlockedException(
                "Security configuration changed; refresh the diff before saving or queuing."
            );
    }

    /// <summary>
    /// The policy with its diff, and the queued access run's status and first notice, so the
    /// library shows a run that stopped or waits instead of "Applying access" forever.
    /// </summary>
    private SecurityResult Result(StoredRow<PolicyDocument>? row)
    {
        var result = Describe(row);
        var run =
            row?.Value.OperationKey == null
                ? null
                : store.Find<SecurityOperation>("asx_operation", row.Value.OperationKey)?.Value;
        if (run != null)
        {
            result.RunStatus = run.Status;
            result.RunNotice = run.ErrorCode ?? run.Notices.FirstOrDefault();
            result.RunNextAttemptUtc = run.Status == "RetryWait" ? run.NextAttemptUtc : null;
        }
        if (row != null)
            result.Teams = Teams(row.Value);
        return result;
    }

    /// <summary>
    /// The policy's teams with their names, read with one query, so a team deleted in Dataverse
    /// is shown as deleted by its last known name instead of failing the read.
    /// </summary>
    private PolicyTeam[] Teams(PolicyDocument policy)
    {
        var ids = policy
            .Desired.Concat(policy.Applied)
            .Select(e => e.TeamId)
            .Where(id => id != Guid.Empty)
            .Distinct()
            .ToArray();
        var live = TeamDirectory.Read(service, ids);
        return ids.Select(id =>
                live.TryGetValue(id, out var team)
                    ? new PolicyTeam { TeamId = id, Name = TeamPrincipal.Name(team) }
                    : new PolicyTeam
                    {
                        TeamId = id,
                        Deleted = true,
                        Name = store
                            .Find<TeamRegistration>(
                                "asx_teamregistration",
                                "team:" + id.ToString("N")
                            )
                            ?.Value.Name,
                    }
            )
            .ToArray();
    }

    private static SecurityResult Describe(StoredRow<PolicyDocument>? row) =>
        new SecurityResult
        {
            Status = row?.Value.Status ?? "Missing",
            RowVersion = row?.Row.RowVersion ?? "",
            Policy = row?.Value,
            Diff =
                row == null
                    ? Array.Empty<string>()
                    : row
                        .Value.Desired.Concat(
                            row.Value.Applied.Where(a =>
                                    !row.Value.Desired.Any(d => d.TeamId == a.TeamId)
                                )
                                .Select(a => new PolicyEntry { TeamId = a.TeamId, Access = "None" })
                        )
                        .Select(e =>
                            e.TeamId
                            + ": "
                            + (
                                row.Value.Applied.SingleOrDefault(a => a.TeamId == e.TeamId)?.Access
                                ?? "None"
                            )
                            + " → "
                            + e.Access
                        )
                        .ToArray(),
        };
}
