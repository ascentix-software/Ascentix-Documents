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
            if (request.Enabled)
                TeamPrincipal.Validate(
                    service.Retrieve("team", request.TeamId, TeamPrincipal.Columns())
                );
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
                        Status = request.Enabled ? "Enabled" : "Revoking",
                    }
                );
            }
            else
            {
                Version(old.Row, request.RowVersion);
                old.Value.Enabled = request.Enabled;
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
            foreach (var entry in request.Entries.Where(e => e.Access != "None"))
            {
                // Saving a draft validates eligibility; only applying explicitly onboards syncing.
                TeamPrincipal.Validate(
                    service.Retrieve("team", entry.TeamId, TeamPrincipal.Columns())
                );
                if (request.Command == "ApplyPolicy")
                {
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
                                Status = "Enabled",
                            }
                        );
                    else if (!registration.Value.Enabled)
                    {
                        registration.Value.Enabled = true;
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
                if (existing.Value.OperationKey != null)
                    throw new EvaluationBlockedException(
                        "Queued policy must finish or be reviewed before editing."
                    );
                existing.Value.Desired = request.Entries;
                existing.Value.ReadRole = request.ReadRole;
                existing.Value.ContributeRole = request.ContributeRole;
                existing.Value.Status = "Draft";
                store.Save(existing);
            }
            existing = store.Require<PolicyDocument>("asx_policy", policyKey);
            if (request.Command == "SavePolicy")
                return Result(existing);
            expectedVersion = existing.Row.RowVersion; // Apply continues in this same transaction; any failure rolls back the save.
        }
        if (
            (
                request.Command != "QueuePolicy"
                && request.Command != "ApplyPolicy"
                && !(workerRefresh && request.Command == "RefreshPolicy")
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
        var desired = workerRefresh ? existing.Value.Approved : existing.Value.Desired;
        var entries = desired
            .Concat(
                existing
                    .Value.ManagedTeams.Where(team => !desired.Any(d => d.TeamId == team))
                    .Select(team => new PolicyEntry { TeamId = team, Access = "None" })
            )
            .ToArray();
        foreach (var entry in entries)
            new TeamSnapshotReader(service).Read(entry.TeamId);
        var generation = Guid.NewGuid();
        string operationKey = "policywork:" + generation.ToString("N");
        if (!workerRefresh)
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

    public static void Validate(PolicyEntry[] entries, PolicyRole read, PolicyRole contribute)
    {
        if (
            entries == null
            || entries.Length > 10
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

    public static void ValidateRole(PolicyRole role, bool read)
    {
        // PermissionKind uses one-based bit positions. Allow the documented default role permissions,
        // require document read/write fundamentals, and reject all unknown or administrative bits.
        const uint readHigh = 176,
            readLow = 138612833,
            contributeHigh = 432,
            contributeLow = 1011028719;
        const uint readRequired = 200737,
            contributeRequired = 200751;
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
            || (high & ~(read ? readHigh : contributeHigh)) != 0
            || (low & ~(read ? readLow : contributeLow)) != 0
            || (low & (read ? readRequired : contributeRequired))
                != (read ? readRequired : contributeRequired)
        )
            throw new EvaluationBlockedException(
                "Reviewed role permissions do not match bounded Read/Contribute semantics."
            );
    }

    private static void Version(Entity row, string? expected)
    {
        if (string.IsNullOrEmpty(expected) || row.RowVersion != expected)
            throw new EvaluationBlockedException(
                "Security configuration changed; refresh the diff before saving or queuing."
            );
    }

    private static SecurityResult Result(StoredRow<PolicyDocument>? row) =>
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
