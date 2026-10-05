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
                    return Block(op, claim, "SecurityWriteRejected");
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
        SecurityCatalog.UpdateLibrary(
            service,
            catalog.Library,
            op.Value.PolicyRevision,
            false,
            op.Value.BaselineHash
        );
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
                    return Probe(
                        op,
                        claim.Value,
                        catalog,
                        read ? "SecurityContributeRole" : "SecurityBaseline"
                    );
                case "SecurityBaseline":
                case "SecurityAcl":
                case "SecurityFinal":
                    var acl = SharePointObservations.Body<ODataRows<AclAssignment>>(request);
                    if (acl.Rows == null)
                        throw new EvaluationBlockedException("Missing ACL page.");
                    op.Value.Acl = op.Value.Acl.Concat(acl.Rows).ToArray();
                    if (
                        op.Value.Acl.Length > 1000
                        || op.Value.Acl.Select(a => a.Member?.Id).Distinct().Count()
                            != op.Value.Acl.Length
                    )
                        throw new EvaluationBlockedException("Incomplete or duplicate ACL pages.");
                    if (acl.Next != null)
                        return Continue(op, claim.Value, catalog, acl.Next);
                    string hash = SharePointObservations.AclHash(
                        new ODataRows<AclAssignment> { Rows = op.Value.Acl }
                    );
                    if (op.Value.ProbeKind == "SecurityBaseline")
                    {
                        if (hash != op.Value.BaselineHash)
                            throw new EvaluationBlockedException(
                                "Library ACL differs from the reviewed preservation baseline."
                            );
                        return NextTeam(op, claim.Value, catalog);
                    }
                    if (op.Value.ProbeKind == "SecurityFinal")
                    {
                        if (hash != op.Value.BaselineHash)
                            throw new EvaluationBlockedException(
                                "Library ACL drifted before completion."
                            );
                        foreach (var entry in op.Value.Entries)
                            VerifyMembership(op.Value, entry.TeamId);
                        op.Value.Status = "Verified";
                        store.Save(op);
                        return Result(op.Value, claim.Value, "Verified", catalog);
                    }
                    return ReconcileGrant(op, claim.Value, catalog, hash);
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
                    op.Value.Members = op.Value.Members.Concat(members.Rows).ToArray();
                    if (
                        op.Value.Members.Length > 2000
                        || op.Value.Members.Any(m =>
                            m.Id <= 0 || m.Type != 1 || string.IsNullOrWhiteSpace(m.Login)
                        )
                        || op.Value.Members.Select(m => m.Id).Distinct().Count()
                            != op.Value.Members.Length
                        || op.Value.Members.Select(m => m.Login)
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .Count() != op.Value.Members.Length
                    )
                        throw new EvaluationBlockedException(
                            "Ambiguous or unsupported group membership."
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
            return Probe(op, claim, catalog, "SecurityFinal");
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
        if (store.Find<MembershipDocument>("asx_membership", op.Value.MembershipKey) == null)
            store.Create(
                "asx_membership",
                new MembershipDocument
                {
                    Key = op.Value.MembershipKey,
                    GroupKey = groupKey,
                    Generation = op.Value.PolicyRevision,
                    Desired = new TeamSnapshotReader(service).Read(entry.TeamId),
                }
            );
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

    private WorkerResult ReconcileGrant(
        StoredRow<SecurityOperation> op,
        DispatcherDocument claim,
        SecurityCatalog catalog,
        string hash
    )
    {
        var entry = op.Value.Entries[op.Value.TeamIndex];
        var group = store.Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey).Value;
        string key = "grant:" + catalog.LibraryId.ToString("N") + ":" + entry.TeamId.ToString("N");
        var grant = store.Find<ManagedGrant>("asx_managedgrant", key);
        var actual = op.Value.Acl.SingleOrDefault(a => a.Member.Id == group.GroupId);
        if (actual != null && actual.Member.Type != 8)
            throw new EvaluationBlockedException("Managed group principal type changed.");
        int[] roles = actual?.Roles.Rows.Select(r => r.Id).ToArray() ?? Array.Empty<int>();
        if (
            op.Value.Reprobe
            && op.Value.MutationKind.StartsWith("Grant", StringComparison.Ordinal)
            && hash == op.Value.BaselineHash
            && hash != op.Value.ExpectedAclHash
        )
        {
            // Read back after an unknown write: the ACL is unchanged, so the grant change was
            // not applied. The comparison below prepares it again.
            op.Value.Reprobe = false;
            ClearMutation(op.Value);
        }
        if (op.Value.MutationKind.StartsWith("Grant", StringComparison.Ordinal))
        {
            if (!op.Value.ExternalResponseKnown)
                throw new EvaluationBlockedException(
                    "Unknown grant write requires controlled recovery."
                );
            if (hash != op.Value.ExpectedAclHash)
                throw new EvaluationBlockedException(
                    "ACL changed beyond the one prepared owned grant mutation."
                );
            bool add = op.Value.MutationKind == "GrantAdd";
            if (roles.Contains(op.Value.MutationRole) != add)
                throw new EvaluationBlockedException(
                    "Grant independent readback did not match the submitted change."
                );
            if (grant == null)
                throw new EvaluationBlockedException("Prepared managed grant receipt missing.");
            grant.Value.RoleId = add ? op.Value.MutationRole : 0;
            grant.Value.Status = "Applied";
            grant.Value.Generation = op.Value.PolicyRevision;
            store.Save(grant);
            op.Value.Reprobe = false;
            ClearMutation(op.Value);
            grant = store.Require<ManagedGrant>("asx_managedgrant", key);
        }
        else if (hash != op.Value.BaselineHash)
            throw new EvaluationBlockedException("ACL changed outside the current managed write.");
        if (
            roles.Length > 1
            || (
                roles.Length == 1
                && (
                    grant == null
                    || grant.Value.GroupId != group.GroupId
                    || grant.Value.RoleId != roles[0]
                )
            )
        )
            throw new EvaluationBlockedException(
                "Unowned or unexpected assignment on the managed principal; no adoption/removal allowed."
            );
        if (grant != null && grant.Value.RoleId != 0 && !roles.Contains(grant.Value.RoleId))
            throw new EvaluationBlockedException(
                "Managed grant disappeared outside reconciliation."
            );
        // Preserve the entire independently observed ACL after each owned transition. Final read must match it.
        op.Value.BaselineHash = hash;
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
        int current = roles.SingleOrDefault();
        if (current != desired)
        {
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
            op.Value.MutationRole = current != 0 ? current : desired;
            var after = op.Value.Acl.Where(a => a.Member.Id != group.GroupId).ToList();
            if (current == 0)
            {
                var permission =
                    desired == op.Value.ReadRole.Id ? op.Value.ReadRole : op.Value.ContributeRole;
                after.Add(
                    new AclAssignment
                    {
                        Member = new AclMember { Id = group.GroupId, Type = 8 },
                        Roles = new ODataRows<AclRole>
                        {
                            Rows = new[]
                            {
                                new AclRole
                                {
                                    Id = desired,
                                    Permissions = new PermissionMask
                                    {
                                        High = permission.High,
                                        Low = permission.Low,
                                    },
                                },
                            },
                        },
                    }
                );
            }
            op.Value.ExpectedAclHash = SharePointObservations.AclHash(
                new ODataRows<AclAssignment> { Rows = after.ToArray() }
            );
            return Prepare(
                op,
                claim,
                catalog,
                current != 0 ? "GrantRemove" : "GrantAdd",
                SharePointRequests.ChangeOwnedGrant(
                    catalog.Target,
                    group.GroupId,
                    op.Value.MutationRole,
                    current == 0
                )
            );
        }
        op.Value.Residual = op
            .Value.Acl.Where(a => a.Member.Id != group.GroupId)
            .Select(a =>
                "Other library principal "
                + a.Member.Id
                + " may retain access; item shares, links and site administrators are not exhaustively evaluated."
            )
            .ToArray();
        op.Value.TeamIndex++;
        return NextTeam(op, claim, catalog);
    }

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
        policy.Value.ResidualAccess = op
            .Value.Residual.Concat(
                new[]
                {
                    "Managed access is verified at library scope; effective user access may remain through other groups, direct shares, links, item scopes or site administration.",
                }
            )
            .ToArray();
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
        SecurityCatalog.UpdateLibrary(
            service,
            catalog.Library,
            op.Value.PolicyRevision,
            true,
            op.Value.BaselineHash
        );
        op.Value.Status = "Applied";
        store.Save(op);
        Audit(op.Value.Key, claim.Value.RunId!, "PolicyApplied");
        Release(claim);
        return new WorkerResult
        {
            Status = "Applied",
            Key = op.Value.Key,
            Notices = policy.Value.ResidualAccess,
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
            case "SecurityBaseline":
            case "SecurityAcl":
            case "SecurityFinal":
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
