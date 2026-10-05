using System;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class SecurityLibraryObservation
{
    [DataMember(Name = "Id")]
    public Guid Id { get; set; }

    [DataMember(Name = "HasUniqueRoleAssignments")]
    public bool? UniquePermissions { get; set; }
}

[DataContract]
public sealed class SecurityRoleObservation
{
    [DataMember(Name = "Id")]
    public int Id { get; set; }

    [DataMember(Name = "RoleTypeKind")]
    public int Type { get; set; }

    [DataMember(Name = "BasePermissions")]
    public PermissionMask Permissions { get; set; } = null!;
}

public sealed class SecurityWorker
{
    private readonly IOrganizationService service;
    private readonly DocumentStore store;
    private readonly Func<DateTime> clock;

    public SecurityWorker(IOrganizationService service, Func<DateTime>? clock = null)
    {
        this.service = service;
        store = new DocumentStore(service);
        this.clock = clock ?? (() => DateTime.UtcNow);
    }

    /// <summary>
    /// Advances a claimed policy operation through membership and library-permission readback.
    /// </summary>
    /// <param name="request">The worker command, policy-operation identity, and command-specific inputs.</param>
    /// <param name="transaction">Whether execution is inside a Dataverse transaction; must be true.</param>
    /// <returns>The step status and any HTTP intent or continuation information.</returns>
    public WorkerResult Execute(WorkerRequest request, bool transaction)
    {
        if (!transaction)
            throw new EvaluationBlockedException("Security coordination requires a transaction.");
        var op = store.Require<SecurityOperation>("asx_operation", request.Key);
        if (request.Command == "FailUnclaimed")
            return store.FailUnclaimed<SecurityOperation>(request.Key, request, clock());
        if (request.Command == "Claim")
            return Claim(request, op);
        if (request.Command == "Retry" || request.Command == "Cancel")
            return Manage(request, op);
        var claim = Assert(request);
        var catalog = Current(op.Value);
        switch (request.Command)
        {
            case "Renew":
                claim.Value.LeaseUntilUtc = clock().AddMinutes(5);
                store.Save(claim);
                return new WorkerResult
                {
                    Status = "Renewed",
                    Key = request.Key,
                    Token = claim.Value.Token,
                };
            case "Observe":
                return Observe(request, op, claim, catalog);
            case "PrepareCreate":
                if (
                    op.Value.Mutation == null
                    || op.Value.Status != "ReadyToCreate"
                    || op.Value.ExternalSubmitted
                )
                    throw new EvaluationBlockedException(
                        "Fresh reviewed security mutation required."
                    );
                VerifyTeam(op.Value);
                op.Value.ExternalSubmitted = true;
                op.Value.ExternalResponseKnown = false;
                op.Value.Status = "ExternalUnknown";
                store.Save(op);
                Audit(op.Value.Key, request.RunId, "Submit:" + op.Value.MutationKind);
                return Result(op.Value, claim.Value, "Create", catalog, op.Value.Mutation);
            case "CreateResponse":
                if (!op.Value.ExternalSubmitted || op.Value.ExternalResponseKnown)
                    throw new EvaluationBlockedException(
                        "No matching outstanding security mutation."
                    );
                if (request.HttpStatus == 429)
                {
                    // SharePoint does not execute a throttled request, so nothing changed. The
                    // mutation is dropped and re-derived from fresh reads after the wait.
                    Audit(op.Value.Key, request.RunId, "Throttled:" + op.Value.MutationKind);
                    ClearMutation(op.Value);
                    return Wait(op, claim, request.RetryAfter, 429, null);
                }
                if (
                    request.HttpStatus == 0
                    || request.HttpStatus == 408
                    || request.HttpStatus >= 500
                )
                    return Block(op, claim, "AmbiguousSecurityWrite");
                if (request.HttpStatus < 200 || request.HttpStatus >= 300)
                {
                    op.Value.ExternalResponseKnown = true;
                    if (
                        op.Value.MutationKind != "MemberAdd"
                        && op.Value.MutationKind != "MemberRemove"
                    )
                        return Block(op, claim, "SecurityWriteRejected");
                    // One member SharePoint will not take is skipped; everyone else still syncs.
                    SkipMember(op.Value, request);
                    ClearMutation(op.Value);
                    return Probe(op, claim.Value, catalog, "SecurityMembers");
                }
                op.Value.ExternalResponseKnown = true;
                return Probe(
                    op,
                    claim.Value,
                    catalog,
                    op.Value.MutationKind.StartsWith("Grant", StringComparison.Ordinal)
                        ? "SecurityAcl"
                        : "SecurityGroup"
                );
            case "Complete":
                return Complete(op, claim, catalog);
            case "Fail":
                if (
                    TransientFailure.Is(request)
                    && !(op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown)
                )
                    return Wait(op, claim, null, request.StatusCode, request.ErrorCode);
                return Block(op, claim, "SecurityWorkerFailed");
            default:
                throw new EvaluationBlockedException("Unsupported security worker command.");
        }
    }

    private WorkerResult Manage(WorkerRequest request, StoredRow<SecurityOperation> op)
    {
        // A live claim belongs to a running flow. Once it expires the operator may Retry or
        // Cancel: Cancel never deletes anything in SharePoint, and the next run reads back any
        // write whose outcome is unknown before writing again.
        bool released = store.ReleaseExpired(op.Value.Key, clock(), request.Command);
        bool unknown = op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown;
        if (op.Value.Status == "Applied")
            return new WorkerResult { Status = "Applied", Key = op.Value.Key };
        if (request.Command == "Retry")
        {
            if (
                op.Value.Status != "Blocked"
                && op.Value.Status != "RetryWait"
                && !released
                && !unknown
            )
                throw new EvaluationBlockedException(
                    "Only blocked, waiting or interrupted work can retry."
                );
            if (unknown)
            {
                op.Value.Reprobe = true;
                op.Value.ExternalResponseKnown = true;
            }
            op.Value.Status = "Pending";
            op.Value.NextAttemptUtc = null;
            op.Value.RetryCount = 0;
            store.Save(op);
            return new WorkerResult { Status = "Pending", Key = op.Value.Key };
        }
        op.Value.Status = "Cancelled";
        store.Save(op);
        var policy = store.Require<PolicyDocument>("asx_policy", op.Value.PolicyKey);
        if (policy.Value.OperationKey != op.Value.Key)
            throw new EvaluationBlockedException("Policy generation changed.");
        policy.Value.OperationKey = null;
        policy.Value.Queued = Array.Empty<PolicyEntry>();
        policy.Value.Status = "NeedsReview";
        store.Save(policy);
        var catalog = new SecurityCatalog(service, op.Value.LibraryId);
        SecurityCatalog.UpdateLibrary(service, catalog.Library, op.Value.PolicyRevision, false);
        return new WorkerResult { Status = "Cancelled", Key = op.Value.Key };
    }

    private WorkerResult Claim(WorkerRequest request, StoredRow<SecurityOperation> op)
    {
        if (
            op.Value.Status == "Applied"
            || op.Value.Status == "Cancelled"
            || op.Value.Status == "Blocked"
        )
            return new WorkerResult { Status = op.Value.Status, Key = request.Key };
        if (op.Value.NextAttemptUtc > clock())
            return new WorkerResult { Status = "RetryWait", Key = request.Key };
        if (
            string.IsNullOrWhiteSpace(request.RunId)
            || request.RunId.Length > 150
            || request.RunId.Any(char.IsControl)
        )
            throw new EvaluationBlockedException("Actual bounded run identity required.");
        if (
            !WorkCoordination.HasCapacity(
                service,
                WorkCoordination.Operation(service, request.Key),
                clock()
            )
        )
            return new WorkerResult { Status = "Busy", Key = request.Key };
        var catalog = Current(op.Value);
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
            bool same =
                claim.Value.RunId == request.RunId
                && claim.Value.Token == request.Token
                && claim.Value.LeaseUntilUtc > clock();
            // Once the lease expires the next claim takes over and reads back any unknown write.
            if (!same && !claim.Value.RecoveryPermitted && claim.Value.LeaseUntilUtc > clock())
                return new WorkerResult { Status = "Quarantined", Key = request.Key };
            if (!same)
            {
                claim.Value.HttpOutstanding = false;
                if (op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown)
                    op.Value.Reprobe = true;
                op.Value.ExternalResponseKnown = true;
            }
        }
        op.Value.Steps = 0;
        claim.Value.OperationKey = request.Key;
        claim.Value.RunId = request.RunId;
        claim.Value.Token = Guid.NewGuid();
        claim.Value.LeaseUntilUtc = clock().AddMinutes(5);
        claim.Value.RecoveryPermitted = false;
        claim.Value.Status = "Claimed";
        store.Save(claim);
        // Recovery reads the outstanding target before any new write; never repeats an ambiguous POST.
        return Probe(
            op,
            claim.Value,
            catalog,
            op.Value.ExternalSubmitted
                ? (
                    op.Value.MutationKind.StartsWith("Grant", StringComparison.Ordinal)
                        ? "SecurityAcl"
                        : "SecurityGroup"
                )
                : "SecurityLibrary"
        );
    }

    private WorkerResult Observe(
        WorkerRequest request,
        StoredRow<SecurityOperation> op,
        StoredRow<DispatcherDocument> claim,
        SecurityCatalog catalog
    )
    {
        if (
            request.ProbeId == Guid.Empty
            || request.ProbeId != op.Value.ProbeId
            || request.ProbeKind != op.Value.ProbeKind
        )
            throw new EvaluationBlockedException("Stale security observation.");
        try
        {
            if (
                request.HttpStatus == 429
                || request.HttpStatus >= 500
                || request.HttpStatus == 0
                || request.HttpStatus == 408
            )
            {
                if (op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown)
                    return Block(op, claim, "UnknownSecurityWrite");
                return Wait(op, claim, request.RetryAfter, request.HttpStatus, null);
            }
            switch (op.Value.ProbeKind)
            {
                case "SecurityLibrary":
                    var library = SharePointObservations.Body<SecurityLibraryObservation>(request);
                    if (library.Id != catalog.Target.ListId || library.UniquePermissions != true)
                        throw new EvaluationBlockedException(
                            "Library must have its explicitly reviewed unique permission boundary before policy application."
                        );
                    return Probe(op, claim.Value, catalog, "SecurityReadRole");
                case "SecurityReadRole":
                case "SecurityContributeRole":
                    var role = SharePointObservations.Body<SecurityRoleObservation>(request);
                    bool read = op.Value.ProbeKind == "SecurityReadRole";
                    var expected = read ? op.Value.ReadRole : op.Value.ContributeRole;
                    if (
                        role.Id != expected.Id
                        || role.Type != (read ? 2 : 3)
                        || role.Permissions == null
                        || role.Permissions.High != expected.High
                        || role.Permissions.Low != expected.Low
                    )
                        throw new EvaluationBlockedException(
                            "Reviewed role ID/type/permission mask changed."
                        );
                    SecurityAdministration.ValidateRole(expected, read);
                    if (read)
                        return Probe(op, claim.Value, catalog, "SecurityContributeRole");
                    return NextTeam(op, claim.Value, catalog);
                case "SecurityAcl":
                    // Only the Documents group's own entry matters. Sharing links, Limited Access
                    // entries and people added by hand belong to admins and are never compared.
                    var acl = SharePointObservations.Body<ODataRows<AclAssignment>>(request);
                    if (acl.Rows == null)
                        throw new EvaluationBlockedException("Missing ACL page.");
                    int owned = store
                        .Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey)
                        .Value.GroupId;
                    op.Value.Acl = op
                        .Value.Acl.Concat(acl.Rows.Where(a => a?.Member?.Id == owned))
                        .ToArray();
                    if (acl.Next != null)
                        return Continue(op, claim.Value, catalog, acl.Next);
                    return ReconcileGrant(op, claim.Value, catalog);
                case "SecurityGroup":
                    var groups = SharePointObservations.Body<ODataRows<SiteGroup>>(request);
                    if (groups.Rows == null || groups.Next != null || groups.Rows.Length > 1)
                        throw new EvaluationBlockedException(
                            "Group lookup incomplete or ambiguous."
                        );
                    var group = store.Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey);
                    var found = groups.Rows.SingleOrDefault();
                    if (
                        found == null
                        && op.Value.Reprobe
                        && op.Value.MutationKind == "GroupCreate"
                        && group.Value.GroupId == 0
                    )
                    {
                        // Read back after an unknown create: no group, so it was not created.
                        op.Value.Reprobe = false;
                        ClearMutation(op.Value);
                    }
                    if (found == null)
                    {
                        if (group.Value.GroupId != 0 || op.Value.ExternalSubmitted)
                            throw new EvaluationBlockedException(
                                "Owned group missing or create outcome unresolved."
                            );
                        return Prepare(
                            op,
                            claim.Value,
                            catalog,
                            "GroupCreate",
                            new HttpIntent
                            {
                                Method = "POST",
                                RelativeUri = "_api/web/sitegroups",
                                Body = JsonWire.Write(
                                    new GroupCreateBody
                                    {
                                        Title = group.Value.Title,
                                        Description = group.Value.Marker,
                                    }
                                ),
                            }
                        );
                    }
                    if (group.Value.GroupId == 0 && found.Description != group.Value.Marker)
                        throw new EvaluationBlockedException(
                            "A SharePoint group already uses this Dataverse team name and is not managed for this team. Rename the conflicting group or Dataverse team, then retry."
                        );
                    if (
                        found.Id <= 0
                        || found.Type != 8
                        || (group.Value.GroupId == 0 && found.Title != group.Value.Title)
                        || found.Description != group.Value.Marker
                        || (group.Value.GroupId != 0 && group.Value.GroupId != found.Id)
                    )
                        throw new EvaluationBlockedException(
                            "Managed group ownership or identity changed."
                        );
                    if (op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown)
                        return Block(op, claim, "UnknownGroupWrite");
                    if (group.Value.GroupId == 0)
                    {
                        group.Value.GroupId = found.Id;
                        group.Value.Status = "Owned";
                        store.Save(group);
                    }
                    if (op.Value.MutationKind == "GroupCreate")
                        ClearMutation(op.Value);
                    return Probe(op, claim.Value, catalog, "SecurityMembers");
                case "SecurityMembers":
                    var members = SharePointObservations.Body<ODataRows<SitePerson>>(request);
                    if (members.Rows == null)
                        throw new EvaluationBlockedException("Membership page missing.");
                    string groupTitle = store
                        .Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey)
                        .Value.Title;
                    foreach (var member in members.Rows.Where(m => m != null))
                    {
                        // People are synced from the team. Anything else an admin put in the
                        // Documents group, such as an Entra group, is left in place.
                        if (
                            member.Type != 1
                            || member.Id <= 0
                            || string.IsNullOrWhiteSpace(member.Login)
                        )
                        {
                            Notice(
                                op.Value,
                                "Team '"
                                    + groupTitle
                                    + "': "
                                    + (
                                        string.IsNullOrWhiteSpace(member.Login)
                                            ? "SharePoint principal " + member.Id
                                            : member.Login
                                    )
                                    + " in the Documents group is not a person and was left in place."
                            );
                            continue;
                        }
                        if (
                            op.Value.Members.Any(m =>
                                m.Id == member.Id
                                || string.Equals(
                                    m.Login,
                                    member.Login,
                                    StringComparison.OrdinalIgnoreCase
                                )
                            )
                        )
                            continue;
                        op.Value.Members = op.Value.Members.Concat(new[] { member }).ToArray();
                    }
                    // A team holds at most 2000 people; the bound keeps the run's stored
                    // operation inside the Dataverse payload limit.
                    if (op.Value.Members.Length > 2000)
                        throw new EvaluationBlockedException(
                            "The Documents group has more people than one access run can hold."
                        );
                    if (members.Next != null)
                        return Continue(op, claim.Value, catalog, members.Next);
                    return ReconcileMembers(op, claim.Value, catalog);
                default:
                    throw new EvaluationBlockedException("Unknown security probe.");
            }
        }
        catch (EvaluationBlockedException error)
        {
            return Block(op, claim, error.Message);
        }
    }

    private WorkerResult NextTeam(
        StoredRow<SecurityOperation> op,
        DispatcherDocument claim,
        SecurityCatalog catalog
    )
    {
        if (op.Value.TeamIndex >= op.Value.Entries.Length)
        {
            // Each team's group and grant were read back after its last write.
            foreach (var done in op.Value.Entries)
                VerifyMembership(op.Value, done.TeamId);
            op.Value.Status = "Verified";
            store.Save(op);
            return Result(op.Value, claim, "Verified", catalog);
        }
        var entry = op.Value.Entries[op.Value.TeamIndex];
        string groupKey =
            "group:" + catalog.SiteId.ToString("N") + ":" + entry.TeamId.ToString("N");
        var group = store.Find<ManagedGroup>("asx_managedgroup", groupKey);
        if (group == null && entry.Access == "None")
        {
            op.Value.TeamIndex++;
            return NextTeam(op, claim, catalog);
        }
        if (group == null)
        {
            var title = service
                .Retrieve("team", entry.TeamId, new Microsoft.Xrm.Sdk.Query.ColumnSet("name"))
                .GetAttributeValue<string>("name");
            if (string.IsNullOrWhiteSpace(title))
                throw new EvaluationBlockedException(
                    "The Dataverse team must have a name before its SharePoint group can be created."
                );
            store.Create(
                "asx_managedgroup",
                new ManagedGroup
                {
                    Key = groupKey,
                    SiteId = catalog.SiteId,
                    TeamId = entry.TeamId,
                    Nonce = Guid.NewGuid(),
                    Title = title,
                }
            );
        }
        op.Value.GroupKey = groupKey;
        op.Value.MembershipKey =
            "membership:"
            + op.Value.PolicyRevision.ToString("N")
            + ":"
            + entry.TeamId.ToString("N");
        var membership = store.Find<MembershipDocument>("asx_membership", op.Value.MembershipKey);
        if (membership == null)
        {
            var snapshot = new TeamSnapshotReader(service).Snapshot(entry.TeamId);
            store.Create(
                "asx_membership",
                new MembershipDocument
                {
                    Key = op.Value.MembershipKey,
                    GroupKey = groupKey,
                    Generation = op.Value.PolicyRevision,
                    Desired = snapshot.People,
                    Skipped = snapshot.Skipped,
                }
            );
            membership = store.Require<MembershipDocument>(
                "asx_membership",
                op.Value.MembershipKey
            );
        }
        foreach (var skipped in membership.Value.Skipped)
            Notice(op.Value, skipped);
        return Probe(op, claim, catalog, "SecurityGroup");
    }

    private WorkerResult ReconcileMembers(
        StoredRow<SecurityOperation> op,
        DispatcherDocument claim,
        SecurityCatalog catalog
    )
    {
        var snapshot = store.Require<MembershipDocument>("asx_membership", op.Value.MembershipKey);
        VerifyTeam(op.Value);
        if (op.Value.MutationKind == "MemberAdd" || op.Value.MutationKind == "MemberRemove")
        {
            bool adding = op.Value.MutationKind == "MemberAdd";
            bool present = adding
                ? op.Value.Members.Any(m =>
                    string.Equals(
                        m.Login,
                        op.Value.MutationLogin,
                        StringComparison.OrdinalIgnoreCase
                    )
                )
                : op.Value.Members.Any(m => m.Id == op.Value.MutationMemberId);
            // After an unknown write the readback decides: a missing change was not applied and
            // is prepared again below from the same comparison.
            if (!op.Value.ExternalResponseKnown || (present != adding && !op.Value.Reprobe))
                throw new EvaluationBlockedException(
                    "Membership independent readback differs from the submitted change."
                );
            op.Value.Reprobe = false;
            ClearMutation(op.Value);
        }
        snapshot.Value.Observed = op.Value.Members;
        snapshot.Value.Complete = true;
        snapshot.Value.Status = "Observed";
        store.Save(snapshot);
        var group = store.Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey);
        var desired = snapshot.Value.Desired.Select(p => p.Login).ToArray();
        var remove = op.Value.Members.FirstOrDefault(p =>
            !desired.Contains(p.Login, StringComparer.OrdinalIgnoreCase)
            && !op.Value.SkippedMembers.Contains(op.Value.GroupKey + "|#" + p.Id)
        );
        if (remove != null)
        {
            op.Value.MutationMemberId = remove.Id;
            return Prepare(
                op,
                claim,
                catalog,
                "MemberRemove",
                SharePointRequests.RemoveMember(group.Value.GroupId, remove.Id)
            );
        }
        var add = desired.FirstOrDefault(login =>
            !op.Value.Members.Any(p =>
                string.Equals(p.Login, login, StringComparison.OrdinalIgnoreCase)
            )
            && !op.Value.SkippedMembers.Contains(
                op.Value.GroupKey + "|" + login,
                StringComparer.OrdinalIgnoreCase
            )
        );
        if (add != null)
        {
            op.Value.MutationLogin = add;
            return Prepare(
                op,
                claim,
                catalog,
                "MemberAdd",
                SharePointRequests.AddMember(group.Value.GroupId, add)
            );
        }
        snapshot = store.Require<MembershipDocument>("asx_membership", op.Value.MembershipKey);
        snapshot.Value.Status = "Applied";
        store.Save(snapshot);
        group.Value.MembershipHash = TeamSnapshotReader.Hash(snapshot.Value.Desired);
        group.Value.Status = "Applied";
        store.Save(group);
        return Probe(op, claim, catalog, "SecurityAcl");
    }

    /// <summary>
    /// Brings the Documents group's grant on the library to the configured level. The group is
    /// Documents' own and the Dataverse team is the source of truth, so a level changed, removed
    /// or added by hand is put back and reported in a notice rather than blocking.
    /// </summary>
    private WorkerResult ReconcileGrant(
        StoredRow<SecurityOperation> op,
        DispatcherDocument claim,
        SecurityCatalog catalog
    )
    {
        var entry = op.Value.Entries[op.Value.TeamIndex];
        var group = store.Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey).Value;
        string key = "grant:" + catalog.LibraryId.ToString("N") + ":" + entry.TeamId.ToString("N");
        var grant = store.Find<ManagedGrant>("asx_managedgrant", key);
        if (op.Value.Acl.Any(a => a.Member.Type != 8))
            throw new EvaluationBlockedException("Managed group principal type changed.");
        int[] roles = op
            .Value.Acl.SelectMany(a => a.Roles?.Rows ?? Array.Empty<AclRole>())
            .Select(r => r.Id)
            .Distinct()
            .OrderBy(r => r)
            .ToArray();
        bool grantWrite = op.Value.MutationKind.StartsWith("Grant", StringComparison.Ordinal);
        if (
            op.Value.Reprobe
            && grantWrite
            && roles.Contains(op.Value.MutationRole) != (op.Value.MutationKind == "GrantAdd")
        )
        {
            // Read back after an unknown write: the grant change was not applied. The comparison
            // below prepares it again.
            op.Value.Reprobe = false;
            ClearMutation(op.Value);
            grantWrite = false;
        }
        if (grantWrite)
        {
            if (!op.Value.ExternalResponseKnown)
                throw new EvaluationBlockedException(
                    "Unknown grant write requires controlled recovery."
                );
            bool add = op.Value.MutationKind == "GrantAdd";
            if (roles.Contains(op.Value.MutationRole) != add)
                throw new EvaluationBlockedException(
                    "Grant independent readback did not match the submitted change."
                );
            if (grant == null)
                throw new EvaluationBlockedException("Prepared managed grant receipt missing.");
            if (add)
                grant.Value.RoleId = op.Value.MutationRole;
            else if (grant.Value.RoleId == op.Value.MutationRole)
                grant.Value.RoleId = 0;
            grant.Value.GroupId = group.GroupId;
            grant.Value.Status = "Applied";
            grant.Value.Generation = op.Value.PolicyRevision;
            store.Save(grant);
            op.Value.Reprobe = false;
            ClearMutation(op.Value);
            grant = store.Require<ManagedGrant>("asx_managedgrant", key);
        }
        int desired =
            entry.Access == "Read" ? op.Value.ReadRole.Id
            : entry.Access == "Contribute" ? op.Value.ContributeRole.Id
            : 0;
        if (
            !store
                .Require<TeamRegistration>(
                    "asx_teamregistration",
                    "team:" + entry.TeamId.ToString("N")
                )
                .Value.Enabled
        )
            desired = 0;
        int[] target = desired == 0 ? Array.Empty<int>() : new[] { desired };
        // What Documents last applied. A receipt still Pending means Documents' own write may
        // have landed without a confirmed outcome, so a difference is not reported as a hand edit.
        bool known = grant == null || grant.Value.Status == "Applied";
        int recorded =
            grant != null && grant.Value.GroupId == group.GroupId ? grant.Value.RoleId : 0;
        int[] expected = recorded == 0 ? Array.Empty<int>() : new[] { recorded };
        if (roles.SequenceEqual(target))
        {
            // Already at the configured level, including a Documents write whose outcome was
            // unknown but landed. Record it and move on.
            if (grant != null && (recorded != desired || grant.Value.Status != "Applied"))
            {
                grant.Value.RoleId = desired;
                grant.Value.GroupId = group.GroupId;
                grant.Value.Status = "Applied";
                grant.Value.Generation = op.Value.PolicyRevision;
                store.Save(grant);
            }
            op.Value.TeamIndex++;
            return NextTeam(op, claim, catalog);
        }
        string changed =
            "Team '"
            + group.Title
            + "': the Documents group's permission on this library was changed outside Documents";
        // Reported once per team: the first reading is what the admin changed.
        if (
            known
            && !roles.SequenceEqual(expected)
            && !op.Value.Notices.Any(n => n.StartsWith(changed, StringComparison.Ordinal))
        )
            Notice(
                op.Value,
                changed
                    + " (found "
                    + Levels(op.Value, roles)
                    + ", configured "
                    + Levels(op.Value, target)
                    + "). Documents restored the configured level. To keep a different level, change this team's access in Documents."
            );
        if (grant == null)
        {
            store.Create(
                "asx_managedgrant",
                new ManagedGrant
                {
                    Key = key,
                    LibraryId = catalog.LibraryId,
                    TeamId = entry.TeamId,
                    GroupId = group.GroupId,
                    Generation = op.Value.PolicyRevision,
                }
            );
        }
        else
        {
            // Pending until the write is read back, so an unknown outcome is not later taken
            // for a hand edit.
            grant.Value.Status = "Pending";
            store.Save(grant);
        }
        int extra = roles.FirstOrDefault(r => r != desired);
        op.Value.MutationRole = extra != 0 ? extra : desired;
        return Prepare(
            op,
            claim,
            catalog,
            extra != 0 ? "GrantRemove" : "GrantAdd",
            SharePointRequests.ChangeOwnedGrant(
                catalog.Target,
                group.GroupId,
                op.Value.MutationRole,
                extra == 0
            )
        );
    }

    private static string Levels(SecurityOperation op, int[] roles) =>
        roles.Length == 0
            ? "no access"
            : string.Join(
                " and ",
                roles.Select(r =>
                    r == op.ReadRole.Id ? "Read"
                    : r == op.ContributeRole.Id ? "Contribute"
                    : "permission level " + r
                )
            );

    private WorkerResult Complete(
        StoredRow<SecurityOperation> op,
        StoredRow<DispatcherDocument> claim,
        SecurityCatalog catalog
    )
    {
        if (op.Value.Status != "Verified" || op.Value.ExternalSubmitted)
            throw new EvaluationBlockedException("Final security readback required.");
        foreach (var entry in op.Value.Entries)
            VerifyMembership(op.Value, entry.TeamId);
        var policy = store.Require<PolicyDocument>("asx_policy", op.Value.PolicyKey);
        policy.Value.Applied = op
            .Value.Entries.Select(e => new PolicyEntry
            {
                TeamId = e.TeamId,
                Access = store
                    .Require<TeamRegistration>(
                        "asx_teamregistration",
                        "team:" + e.TeamId.ToString("N")
                    )
                    .Value.Enabled
                    ? e.Access
                    : "None",
            })
            .ToArray();
        policy.Value.Queued = Array.Empty<PolicyEntry>();
        policy.Value.OperationKey = null;
        policy.Value.Status = "Applied";
        policy.Value.Notices = op.Value.Notices;
        policy.Value.ResidualAccess = new[]
        {
            "Documents manages only its own team groups and their grants on this library. People may still have access through other groups, direct shares, links, item permissions or site administration.",
        };
        store.Save(policy);
        foreach (var entry in policy.Value.Applied)
        {
            var reference = store.Require<PolicyTeamReference>(
                "asx_policyentry",
                "policyteam:" + op.Value.LibraryId.ToString("N") + ":" + entry.TeamId.ToString("N")
            );
            reference.Value.Status = entry.Access == "None" ? "Inactive" : "Active";
            store.Save(reference);
        }
        foreach (var entry in op.Value.Entries)
        {
            var receipt = store.Find<MembershipDocument>(
                "asx_membership",
                "membership:"
                    + op.Value.PolicyRevision.ToString("N")
                    + ":"
                    + entry.TeamId.ToString("N")
            );
            if (receipt?.Value.Status == "Applied")
            {
                receipt.Value.RetireAfterUtc = clock().AddDays(30);
                store.Save(receipt);
            }
        }
        if (op.Value.LibrarySetupKey != null)
        {
            var setup = store.Require<LibrarySetup>("asx_operation", op.Value.LibrarySetupKey);
            setup.Value.Status = "Applied";
            setup.Value.CompletedUtc = clock();
            store.Save(setup);
        }
        SecurityCatalog.UpdateLibrary(service, catalog.Library, op.Value.PolicyRevision, true);
        op.Value.Status = "Applied";
        store.Save(op);
        Audit(op.Value.Key, claim.Value.RunId!, "PolicyApplied");
        Release(claim);
        return new WorkerResult
        {
            Status = "Applied",
            Key = op.Value.Key,
            Notices = op.Value.Notices.Concat(policy.Value.ResidualAccess).ToArray(),
        };
    }

    private SecurityCatalog Current(SecurityOperation op)
    {
        var catalog = new SecurityCatalog(service, op.LibraryId);
        var policy = store.Require<PolicyDocument>("asx_policy", op.PolicyKey).Value;
        if (
            catalog.Target.WebId != op.WebId
            || catalog.Target.ListId != op.ListId
            || catalog.Target.Web.AbsoluteUri != op.WebUrl
            || catalog.SiteId != op.SiteId
            || policy.Generation != op.PolicyRevision
            || policy.OperationKey != op.Key
            || catalog.Library.GetAttributeValue<string>("asx_policyrevision")
                != op.PolicyRevision.ToString("D")
        )
            throw new EvaluationBlockedException("Security approval generation changed.");
        return catalog;
    }

    private void VerifyTeam(SecurityOperation op)
    {
        if (string.IsNullOrEmpty(op.MembershipKey))
            return;
        var snapshot = store.Require<MembershipDocument>("asx_membership", op.MembershipKey).Value;
        if (
            TeamSnapshotReader.Hash(snapshot.Desired)
            != TeamSnapshotReader.Hash(
                new TeamSnapshotReader(service).Read(op.Entries[op.TeamIndex].TeamId)
            )
        )
            throw new EvaluationBlockedException(
                "Team changed during reconciliation; a fresh generation is required."
            );
    }

    private void VerifyMembership(SecurityOperation op, Guid teamId)
    {
        var snapshot = store.Find<MembershipDocument>(
            "asx_membership",
            "membership:" + op.PolicyRevision.ToString("N") + ":" + teamId.ToString("N")
        );
        if (snapshot == null)
        {
            if (op.Entries.Single(e => e.TeamId == teamId).Access != "None")
                throw new EvaluationBlockedException("Membership receipt missing.");
            return;
        }
        if (
            !snapshot.Value.Complete
            || snapshot.Value.Status != "Applied"
            || TeamSnapshotReader.Hash(snapshot.Value.Desired)
                != TeamSnapshotReader.Hash(new TeamSnapshotReader(service).Read(teamId))
        )
            throw new EvaluationBlockedException(
                "Final membership generation changed or is incomplete."
            );
    }

    private WorkerResult Prepare(
        StoredRow<SecurityOperation> op,
        DispatcherDocument claim,
        SecurityCatalog catalog,
        string kind,
        HttpIntent http
    )
    {
        op.Value.Mutation = http;
        op.Value.MutationKind = kind;
        op.Value.Status = "ReadyToCreate";
        op.Value.ExternalSubmitted = false;
        op.Value.ExternalResponseKnown = false;
        store.Save(op);
        return Result(op.Value, claim, "ReadyToCreate", catalog);
    }

    private WorkerResult Probe(
        StoredRow<SecurityOperation> op,
        DispatcherDocument claim,
        SecurityCatalog catalog,
        string kind
    )
    {
        if (++op.Value.Steps > 40 && !op.Value.ExternalSubmitted)
        {
            op.Value.Status = "Pending";
            store.Save(op);
            Release(
                store.Require<DispatcherDocument>(
                    "asx_claim",
                    WorkCoordination.Operation(service, op.Value.Key)
                )
            );
            return new WorkerResult { Status = "Pending", Key = op.Value.Key };
        }
        string list = "_api/web/lists(guid'" + catalog.Target.ListId + "')";
        string endpoint;
        switch (kind)
        {
            case "SecurityLibrary":
                endpoint = list + "?$select=Id,HasUniqueRoleAssignments";
                break;
            case "SecurityReadRole":
            case "SecurityContributeRole":
                endpoint =
                    "_api/web/roledefinitions("
                    + (
                        kind == "SecurityReadRole"
                            ? op.Value.ReadRole.Id
                            : op.Value.ContributeRole.Id
                    )
                    + ")?$select=Id,RoleTypeKind,BasePermissions";
                break;
            case "SecurityAcl":
                op.Value.Acl = Array.Empty<AclAssignment>();
                endpoint =
                    list
                    + "/roleassignments?$select=Member/Id,Member/PrincipalType,RoleDefinitionBindings/Id,RoleDefinitionBindings/BasePermissions&$expand=Member,RoleDefinitionBindings";
                break;
            case "SecurityGroup":
                var group = store
                    .Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey)
                    .Value;
                endpoint =
                    "_api/web/sitegroups?$select=Id,Title,Description,PrincipalType&$top=2&$filter="
                    + Uri.EscapeDataString(
                        group.GroupId > 0
                            ? "Id eq " + group.GroupId
                            : "Title eq '" + group.Title.Replace("'", "''") + "'"
                    );
                break;
            case "SecurityMembers":
                op.Value.Members = Array.Empty<SitePerson>();
                endpoint =
                    "_api/web/sitegroups("
                    + store
                        .Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey)
                        .Value.GroupId
                    + ")/users?$select=Id,LoginName,PrincipalType&$top=500";
                break;
            default:
                throw new EvaluationBlockedException("Unsupported security read.");
        }
        op.Value.ProbeKind = kind;
        op.Value.ProbeId = Guid.NewGuid();
        op.Value.PageEndpoint = endpoint;
        op.Value.PageLinks = new[] { endpoint };
        op.Value.Status = op.Value.ExternalSubmitted ? "ExternalUnknown" : "Inspecting";
        store.Save(op);
        return Result(op.Value, claim, "Read", catalog, new HttpIntent { RelativeUri = endpoint });
    }

    private WorkerResult Continue(
        StoredRow<SecurityOperation> op,
        DispatcherDocument claim,
        SecurityCatalog catalog,
        string next
    )
    {
        string path = SecurityPaging.Next(
            catalog.Target.Web,
            op.Value.PageEndpoint,
            next,
            op.Value.PageLinks
        );
        op.Value.PageLinks = op.Value.PageLinks.Concat(new[] { path }).ToArray();
        op.Value.ProbeId = Guid.NewGuid();
        store.Save(op);
        return Result(op.Value, claim, "Read", catalog, new HttpIntent { RelativeUri = path });
    }

    private StoredRow<DispatcherDocument> Assert(WorkerRequest request)
    {
        var claim = store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        if (
            request.Token == Guid.Empty
            || claim.Value.Token != request.Token
            || claim.Value.RunId != request.RunId
            || claim.Value.OperationKey != request.Key
            || claim.Value.LeaseUntilUtc <= clock()
            || claim.Value.RecoveryPermitted
        )
            throw new EvaluationBlockedException("Stale security claim.");
        return claim;
    }

    private WorkerResult Block(
        StoredRow<SecurityOperation> op,
        StoredRow<DispatcherDocument> claim,
        string error
    )
    {
        op.Value.ErrorCode = error;
        bool unknown = op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown;
        op.Value.Status = unknown ? "ExternalUnknown" : "Blocked";
        store.Save(op);
        if (!unknown)
            Release(claim);
        return new WorkerResult
        {
            Status = unknown ? "Quarantined" : "Blocked",
            Key = op.Value.Key,
            Notices = new[] { error },
        };
    }

    /// <summary>
    /// Waits after a temporary failure with no attempt cap and releases the writer slot.
    /// </summary>
    private WorkerResult Wait(
        StoredRow<SecurityOperation> op,
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

    // Bounded so the operation stays well inside the 500,000-character JSON payload limit.
    private const int NoticeLimit = 200;

    /// <summary>Records something an admin should see after the run, once.</summary>
    private static void Notice(SecurityOperation op, string text)
    {
        text = new string(text.Where(c => !char.IsControl(c)).ToArray());
        if (text.Length > 600)
            text = text.Substring(0, 600) + "...";
        if (op.Notices.Length >= NoticeLimit)
            text = "More notices were not recorded.";
        if (op.Notices.Length > NoticeLimit || op.Notices.Contains(text, StringComparer.Ordinal))
            return;
        op.Notices = op.Notices.Concat(new[] { text }).ToArray();
    }

    /// <summary>Records a member add or remove SharePoint rejected, with SharePoint's message.</summary>
    private void SkipMember(SecurityOperation op, WorkerRequest request)
    {
        bool adding = op.MutationKind == "MemberAdd";
        string login = adding
            ? op.MutationLogin
            : op.Members.FirstOrDefault(m => m.Id == op.MutationMemberId)?.Login
                ?? "SharePoint user " + op.MutationMemberId;
        op.SkippedMembers = op
            .SkippedMembers.Concat(
                new[]
                {
                    op.GroupKey + (adding ? "|" + op.MutationLogin : "|#" + op.MutationMemberId),
                }
            )
            .ToArray();
        string title = store.Require<ManagedGroup>("asx_managedgroup", op.GroupKey).Value.Title;
        Notice(
            op,
            "SharePoint did not "
                + (adding ? "add " : "remove ")
                + login.Substring(login.LastIndexOf('|') + 1)
                + (adding ? " to" : " from")
                + " the group for team '"
                + title
                + "': "
                + SharePointObservations.ErrorMessage(request)
        );
    }

    private static void ClearMutation(SecurityOperation op)
    {
        op.ExternalSubmitted = false;
        op.ExternalResponseKnown = false;
        op.Mutation = null;
        op.MutationKind = "";
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

    private void Audit(string key, string run, string kind) =>
        store.Create(
            "asx_attempt",
            new AttemptDocument
            {
                Key = "attempt:" + Guid.NewGuid().ToString("N"),
                Status = "Recorded",
                OperationKey = key,
                RunId = run,
                Event = kind,
                AtUtc = clock(),
            }
        );

    private static WorkerResult Result(
        SecurityOperation op,
        DispatcherDocument claim,
        string status,
        SecurityCatalog catalog,
        HttpIntent? http = null
    ) =>
        new WorkerResult
        {
            Status = status,
            Key = op.Key,
            Token = claim.Token,
            ProbeId = op.ProbeId,
            ProbeKind = op.ProbeKind,
            SiteUrl = catalog.Target.Web.AbsoluteUri.TrimEnd('/'),
            Http = http,
        };
}
