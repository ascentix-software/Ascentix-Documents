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
    /// <summary>A library reset to inherit its site's permissions after Documents stopped it.</summary>
    public const string InheritsAgainNotice =
        "This library inherits permissions again. Use Apply access to let Documents stop the inheritance again.";

    /// <summary>The setup of a new library whose first access run was cancelled or replaced.</summary>
    public const string SetupAccessCancelled =
        "Library created. Its first access run was cancelled; apply access on the library.";

    public const string InheritanceStoppedNotice =
        "Documents stopped this library's permission inheritance and kept a copy of the site's permissions.";

    private const string BreakInheritanceKind = "BreakInheritance";

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
        if (request.Command != "Retry" && request.Command != "Cancel")
        {
            var stopped = StopRemoved(request, op) ?? StopSuspended(request, op);
            if (stopped != null)
                return stopped;
        }
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
                {
                    // Resolving a group claim changes no access and can be repeated, so a lost
                    // response waits and resolves again.
                    if (op.Value.MutationKind == "PrincipalEnsure")
                    {
                        ClearMutation(op.Value);
                        return Wait(op, claim, request.RetryAfter, request.HttpStatus, null);
                    }
                    // Whether inheritance stopped is read back after the wait: the library read
                    // decides, so the break is never sent twice.
                    if (op.Value.MutationKind == BreakInheritanceKind)
                    {
                        op.Value.Reprobe = true;
                        op.Value.ExternalResponseKnown = true;
                        return Wait(op, claim, request.RetryAfter, request.HttpStatus, null);
                    }
                    return Block(op, claim, "AmbiguousSecurityWrite");
                }
                if (request.HttpStatus < 200 || request.HttpStatus >= 300)
                {
                    op.Value.ExternalResponseKnown = true;
                    // 401/403 means Documents lost its permission on the site, which no other
                    // member would get past either, so that still stops the run.
                    if (
                        (
                            op.Value.MutationKind != "MemberAdd"
                            && op.Value.MutationKind != "MemberRemove"
                            && op.Value.MutationKind != "PrincipalEnsure"
                        )
                        || request.HttpStatus == 401
                        || request.HttpStatus == 403
                    )
                        return Block(op, claim, "SecurityWriteRejected");
                    // One member SharePoint will not take is skipped; everyone else still syncs.
                    SkipMember(op.Value, request);
                    ClearMutation(op.Value);
                    return Probe(op, claim.Value, catalog, "SecurityMembers");
                }
                op.Value.ExternalResponseKnown = true;
                return Probe(op, claim.Value, catalog, ReadBack(op.Value));
            case "Complete":
                return Complete(op, claim, catalog);
            case "Fail":
                // Resolving a group claim changes no access and can be repeated, so its unknown
                // outcome is dropped and the claim resolved again; the failure itself decides.
                if (
                    op.Value.MutationKind == "PrincipalEnsure"
                    && op.Value.ExternalSubmitted
                    && !op.Value.ExternalResponseKnown
                )
                    ClearMutation(op.Value);
                // An inheritance break of unknown outcome is read back by the next run.
                if (
                    op.Value.MutationKind == BreakInheritanceKind
                    && op.Value.ExternalSubmitted
                    && !op.Value.ExternalResponseKnown
                )
                {
                    op.Value.Reprobe = true;
                    op.Value.ExternalResponseKnown = true;
                }
                if (
                    TransientFailure.Is(request)
                    && !(op.Value.ExternalSubmitted && !op.Value.ExternalResponseKnown)
                )
                    return Wait(op, claim, null, request.StatusCode, request.ErrorCode);
                // The failed call's own message, such as why a check stopped the run, is the
                // reason the library shows.
                return Block(
                    op,
                    claim,
                    string.IsNullOrWhiteSpace(request.Error)
                        ? "SecurityWorkerFailed"
                        : Bounded(request.Error!)
                );
            default:
                throw new EvaluationBlockedException("Unsupported security worker command.");
        }
    }

    /// <summary>
    /// Cancels an access run whose library was removed from Documents, at its next step, so
    /// Remove never waits for a running flow. Nothing in SharePoint is undone or deleted.
    /// </summary>
    private WorkerResult? StopRemoved(WorkerRequest request, StoredRow<SecurityOperation> op)
    {
        if (
            op.Value.Status == "Applied"
            || op.Value.Status == "Cancelled"
            || !new WorkerCatalog(service).Removed(op.Value.LibraryId)
        )
            return null;
        var claim = store.Find<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        if (claim?.Value.OperationKey == request.Key && claim.Value.RunId != null)
        {
            if (request.Command != "Claim")
                claim = Assert(request);
            else if (claim.Value.LeaseUntilUtc > clock())
                return new WorkerResult { Status = "Quarantined", Key = request.Key };
            claim!.Value.HttpOutstanding = false;
            Release(claim);
        }
        op = store.Require<SecurityOperation>("asx_operation", request.Key);
        op.Value.Status = "Cancelled";
        op.Value.ErrorCode = "DestinationRemoved";
        op.Value.NextAttemptUtc = null;
        store.Save(op);
        var policy = store.Find<PolicyDocument>("asx_policy", op.Value.PolicyKey);
        if (policy?.Value.OperationKey == op.Value.Key)
        {
            policy.Value.OperationKey = null;
            policy.Value.Queued = Array.Empty<PolicyEntry>();
            store.Save(policy);
        }
        Audit(op.Value.Key, request.RunId ?? "worker", "Cancel");
        return new WorkerResult
        {
            Status = "Cancelled",
            Key = request.Key,
            Notices = new[]
            {
                "The library was removed from Documents, so this access run stopped. Nothing in SharePoint was undone or deleted.",
            },
        };
    }

    /// <summary>
    /// Holds an access run while its library or site is suspended, at its next step: it waits
    /// in RetryWait with a notice, its next check backs off (at most 15 minutes apart) so it
    /// takes no dispatch slot meanwhile, and it resumes by itself once both are approved again.
    /// Nothing is written while suspended. A write already sent is read back after the wait
    /// before anything is written again, as after an operator Retry.
    /// </summary>
    private WorkerResult? StopSuspended(WorkerRequest request, StoredRow<SecurityOperation> op)
    {
        if (
            op.Value.Status == "Applied"
            || op.Value.Status == "Cancelled"
            || op.Value.Status == "Blocked"
            || request.Command == "Claim" && op.Value.NextAttemptUtc > clock()
            || WorkCoordination.Stops(service, op.Value.LibraryId) != "suspended"
        )
            return null;
        // A step must hold the run's claim, as every step does: a stale step of a run that
        // already waits must not push its next check further out.
        if (request.Command != "Claim")
            Assert(request);
        var claim = store.Find<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, request.Key)
        );
        if (claim?.Value.OperationKey == request.Key && claim.Value.RunId != null)
        {
            if (request.Command != "Claim")
                claim = Assert(request);
            else if (claim.Value.LeaseUntilUtc > clock())
                return new WorkerResult { Status = "Quarantined", Key = request.Key };
            claim!.Value.HttpOutstanding = false;
            Release(claim);
        }
        op = store.Require<SecurityOperation>("asx_operation", request.Key);
        if (op.Value.ExternalSubmitted)
        {
            op.Value.Reprobe = true;
            op.Value.ExternalResponseKnown = true;
            op.Value.ReadbackMisses = 0;
        }
        else
            ClearMutation(op.Value);
        op.Value.RetryCount++;
        op.Value.Status = "RetryWait";
        op.Value.NextAttemptUtc = WorkerCoordinator.RetryAt(clock(), op.Value.RetryCount, null);
        op.Value.ProbeId = Guid.Empty;
        op.Value.ErrorCode =
            "Waiting: the library or its site is suspended. This access run resumes by itself once both are approved again; check "
            + op.Value.RetryCount.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + ".";
        store.Save(op);
        Audit(op.Value.Key, request.RunId ?? "worker", "DestinationSuspended");
        return new WorkerResult
        {
            Status = "RetryWait",
            Key = request.Key,
            Notices = new[] { op.Value.ErrorCode },
        };
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
            // A write still pending after read-back misses is read again from scratch: if it is
            // still not there it is prepared and written again.
            if (op.Value.ExternalSubmitted)
                op.Value.Reprobe = true;
            op.Value.ReadbackMisses = 0;
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
        // The library a setup created stays; its setup ends here instead of showing its first
        // access run as applying for ever. Access is applied again from the library.
        if (op.Value.LibrarySetupKey != null)
        {
            var setup = store.Find<LibrarySetup>("asx_operation", op.Value.LibrarySetupKey);
            if (setup?.Value.Status == "AccessPending")
            {
                setup.Value.Status = "Applied";
                setup.Value.CompletedUtc = clock();
                setup.Value.ErrorCode = SetupAccessCancelled;
                store.Save(setup);
            }
        }
        // Cancel works whatever the library's approval: a suspended or removed library too.
        var library = service.Retrieve(
            "asx_library",
            op.Value.LibraryId,
            new Microsoft.Xrm.Sdk.Query.ColumnSet("asx_policyrevision")
        );
        SecurityCatalog.UpdateLibrary(service, library, op.Value.PolicyRevision, false);
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
        claim = WorkCoordination.Unstall(service, claim);
        if (claim.Value.RunId != null)
        {
            if (claim.Value.OperationKey != request.Key)
                return new WorkerResult { Status = "Busy", Key = request.Key };
            bool same =
                claim.Value.RunId == request.RunId
                && claim.Value.Token == request.Token
                && claim.Value.LeaseUntilUtc > clock();
            // Once the lease expires the next claim takes over and reads back any unknown write.
            if (!same && claim.Value.LeaseUntilUtc > clock())
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
        claim.Value.Status = "Claimed";
        store.Save(claim);
        // Recovery reads the outstanding target before any new write; never repeats an ambiguous POST.
        return Probe(
            op,
            claim.Value,
            catalog,
            op.Value.ExternalSubmitted ? ReadBack(op.Value) : "SecurityLibrary"
        );
    }

    /// <summary>The read that shows whether the run's last write took effect.</summary>
    private static string ReadBack(SecurityOperation op) =>
        op.MutationKind == BreakInheritanceKind ? "SecurityLibrary"
        : op.MutationKind.StartsWith("Grant", StringComparison.Ordinal) ? "SecurityAcl"
        : "SecurityGroup";

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
                    if (library.Id != catalog.Target.ListId)
                        throw new EvaluationBlockedException(
                            "SharePoint returned a different library than the approved one."
                        );
                    return ObserveInheritance(
                        op,
                        claim,
                        catalog,
                        library.UniquePermissions == true
                    );
                case "SecurityReadRole":
                case "SecurityContributeRole":
                    var role = SharePointObservations.Body<SecurityRoleObservation>(request);
                    bool read = op.Value.ProbeKind == "SecurityReadRole";
                    var expected = read ? op.Value.ReadRole : op.Value.ContributeRole;
                    if (role.Id != expected.Id || role.Type != (read ? 2 : 3))
                        throw new EvaluationBlockedException(
                            "SharePoint returned a different permission level than the site's "
                                + (read ? "Read" : "Contribute")
                                + " level."
                        );
                    // The level is read again at every run: a customization since approval is
                    // accepted unless it adds administrative rights.
                    var current = new PolicyRole
                    {
                        Id = role.Id,
                        High = role.Permissions?.High ?? "",
                        Low = role.Permissions?.Low ?? "",
                    };
                    SecurityAdministration.ValidateRole(current, read);
                    if (read)
                        op.Value.ReadRole = current;
                    else
                        op.Value.ContributeRole = current;
                    if (read)
                        return Probe(op, claim.Value, catalog, "SecurityContributeRole");
                    return NextTeam(op, claim.Value, catalog);
                case "SecurityAcl":
                    // Only the Documents group's own assignment is read. Sharing links, Limited
                    // Access entries and people added by hand belong to admins and are never
                    // read or compared, so library size never matters.
                    int owned = store
                        .Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey)
                        .Value.GroupId;
                    AclAssignment? assignment = null;
                    // getbyprincipalid answers 404 when the group has no assignment on the list.
                    if (request.HttpStatus != 404)
                    {
                        assignment = SharePointObservations.Body<AclAssignment>(request);
                        if (assignment.Member?.Id != owned)
                            throw new EvaluationBlockedException(
                                "SharePoint returned the role assignment of another principal."
                            );
                    }
                    op.Value.Acl =
                        assignment == null ? Array.Empty<AclAssignment>() : new[] { assignment };
                    return ReconcileGrant(op, claim, catalog);
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
                        if (op.Value.ExternalSubmitted)
                            throw new EvaluationBlockedException(
                                "Owned group missing or create outcome unresolved."
                            );
                        if (group.Value.GroupId != 0)
                        {
                            // The Documents group was deleted in SharePoint. It is Documents' own,
                            // so it is created again; members and the grant follow the normal sync.
                            Notice(
                                op.Value,
                                "Team '"
                                    + group.Value.Title
                                    + "': the Documents group (SharePoint ID "
                                    + group.Value.GroupId
                                    + ") was deleted in SharePoint. Documents created it again and restored its members and access."
                            );
                            group.Value.GroupId = 0;
                            group.Value.MembershipHash = "";
                            group.Value.Status = "Pending";
                            store.Save(group);
                            return Probe(op, claim.Value, catalog, "SecurityGroup");
                        }
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
                        || (group.Value.GroupId != 0 && group.Value.GroupId != found.Id)
                    )
                        throw new EvaluationBlockedException(
                            "Managed group ownership or identity changed."
                        );
                    // Once Documents holds the group's SharePoint ID, that ID is the proof of
                    // ownership; the description marker only matters when first adopting it.
                    if (group.Value.GroupId != 0 && found.Description != group.Value.Marker)
                        Notice(
                            op.Value,
                            "Team '"
                                + group.Value.Title
                                + "': the Documents group's description was changed in SharePoint. Documents keeps managing the group by its SharePoint ID."
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
                    // A team too large for one run leaves its group's members as they are; the
                    // group still gets its access.
                    if (op.Value.SkippedMembers.Contains(op.Value.GroupKey + "|*"))
                    {
                        op.Value.Members = Array.Empty<SitePerson>();
                        return ReconcileMembers(op, claim.Value, catalog);
                    }
                    return Probe(op, claim.Value, catalog, "SecurityMembers");
                case "SecurityMembers":
                    var members = SharePointObservations.Body<ODataRows<SitePerson>>(request);
                    if (members.Rows == null)
                        throw new EvaluationBlockedException("Membership page missing.");
                    var owner = store
                        .Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey)
                        .Value;
                    string groupTitle = owner.Title;
                    var wanted = store
                        .Require<MembershipDocument>("asx_membership", op.Value.MembershipKey)
                        .Value;
                    // The group claims of a group team, wanted now or put there by Documents
                    // before, are Documents' to manage like people.
                    var principals = wanted
                        .Desired.Where(p => p.Group)
                        .Select(p => p.Login)
                        .Concat(owner.Principals ?? Array.Empty<string>())
                        .ToArray();
                    // The members read so far are kept on the run and in the team's membership
                    // row; reading stops where either would pass its column size.
                    int room = Math.Min(
                        SkipDetailBudget - JsonWire.Write(op.Value).Length,
                        PayloadMaxLength
                            - (
                                JsonWire.Write(wanted).Length
                                - JsonWire.Write(wanted.Observed).Length
                                + JsonWire.Write(op.Value.Members).Length
                            )
                    );
                    bool full = false;
                    foreach (var member in members.Rows.Where(m => m != null))
                    {
                        // People are synced from the team. Anything else an admin put in the
                        // Documents group, such as an Entra group, is left in place.
                        bool principal =
                            member.Id > 0
                            && !string.IsNullOrWhiteSpace(member.Login)
                            && principals.Contains(member.Login, StringComparer.OrdinalIgnoreCase);
                        if (
                            !principal
                            && (
                                member.Type != 1
                                || member.Id <= 0
                                || string.IsNullOrWhiteSpace(member.Login)
                            )
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
                        room -= JsonWire.Write(member).Length + 1;
                        if (room < 0)
                        {
                            full = true;
                            break;
                        }
                        op.Value.Members = op.Value.Members.Concat(new[] { member }).ToArray();
                    }
                    if (full)
                    {
                        // Without every member read, a removal could be wrong: the group's
                        // members are left as they are this run, and its access still applies.
                        if (!op.Value.SkippedMembers.Contains(op.Value.GroupKey + "|*"))
                            op.Value.SkippedMembers = op
                                .Value.SkippedMembers.Concat(new[] { op.Value.GroupKey + "|*" })
                                .ToArray();
                        Notice(
                            op.Value,
                            "Team '"
                                + groupTitle
                                + "': the Documents group has more people than one access run can store (a 500,000-character Dataverse row). Team membership was not synced: people removed from the team keep access, and people added get none, until this is resolved. The team's access to the library is still applied, and every scheduled refresh tries again."
                        );
                        return ReconcileMembers(op, claim.Value, catalog);
                    }
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

    /// <summary>
    /// The first step of every run, before any group or grant write. A library with its own
    /// permissions continues. An inheriting library has its inheritance stopped when the admin
    /// consented (keeping a copy of the site's permissions); otherwise the run stops with a
    /// notice, honoring an admin who reset the library to inherit.
    /// </summary>
    private WorkerResult ObserveInheritance(
        StoredRow<SecurityOperation> op,
        StoredRow<DispatcherDocument> claim,
        SecurityCatalog catalog,
        bool unique
    )
    {
        bool breaking = op.Value.MutationKind == BreakInheritanceKind && op.Value.ExternalSubmitted;
        if (unique)
        {
            if (breaking)
            {
                // Read back after the break, including one whose response was lost.
                op.Value.Reprobe = false;
                ClearMutation(op.Value);
                Notice(op.Value, InheritanceStoppedNotice);
            }
            return Probe(op, claim.Value, catalog, "SecurityReadRole");
        }
        if (breaking)
        {
            if (!op.Value.Reprobe)
            {
                // SharePoint confirmed the break but the read does not show it yet; reads can
                // lag a write, so read again after a short wait within the read-back window.
                if (++op.Value.ReadbackMisses >= ReadbackWindow)
                    throw new EvaluationBlockedException(
                        "The library still inherits permissions after SharePoint confirmed stopping it."
                    );
                return Wait(op, claim, null, null, "InheritanceReadbackPending");
            }
            // Read back after an unknown outcome: the break did not happen. It is sent once now.
            op.Value.Reprobe = false;
            ClearMutation(op.Value);
        }
        if (!op.Value.BreakInheritance)
        {
            var policy = store.Require<PolicyDocument>("asx_policy", op.Value.PolicyKey);
            policy.Value.Inherits = true;
            store.Save(policy);
            throw new EvaluationBlockedException(InheritsAgainNotice);
        }
        return Prepare(
            op,
            claim.Value,
            catalog,
            BreakInheritanceKind,
            SharePointRequests.BreakInheritance(catalog.Target)
        );
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
            // A team deleted in Dataverse since this run was queued never had a group here, so
            // there is no access to remove: the run moves on, and its end records that.
            var team = TeamDirectory.Find(service, entry.TeamId);
            if (team == null)
            {
                op.Value.TeamIndex++;
                return NextTeam(op, claim, catalog);
            }
            var title = team.GetAttributeValue<string>("name");
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
                    Incomplete = snapshot.Incomplete,
                }
            );
            membership = store.Require<MembershipDocument>(
                "asx_membership",
                op.Value.MembershipKey
            );
        }
        foreach (var skipped in membership.Value.Skipped)
            Notice(op.Value, skipped);
        // A team with more people than one run can store keeps its group's members as they are;
        // its skipped notice says so, and the group still gets its access.
        string stop = groupKey + "|*";
        if (membership.Value.Incomplete && !op.Value.SkippedMembers.Contains(stop))
            op.Value.SkippedMembers = op.Value.SkippedMembers.Concat(new[] { stop }).ToArray();
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
        if (op.Value.MutationKind == "PrincipalEnsure")
        {
            // SharePoint resolved the group claim; it is added next. After a takeover the
            // outcome is unknown, and resolving again is harmless.
            if (!op.Value.Reprobe)
                op.Value.EnsuredLogin = op.Value.MutationLogin;
            op.Value.Reprobe = false;
            ClearMutation(op.Value);
        }
        else if (op.Value.MutationKind == "MemberAdd" || op.Value.MutationKind == "MemberRemove")
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
            if (!op.Value.ExternalResponseKnown)
                throw new EvaluationBlockedException(
                    "Membership independent readback differs from the submitted change."
                );
            // After an unknown write the readback decides: a missing change was not applied and
            // is prepared again below from the same comparison.
            if (present != adding && !op.Value.Reprobe)
                SkipUnconfirmedMember(op.Value, snapshot.Value.Observed, adding);
            op.Value.Reprobe = false;
            ClearMutation(op.Value);
        }
        snapshot.Value.Observed = op.Value.Members;
        snapshot.Value.Complete = true;
        snapshot.Value.Status = "Observed";
        store.Save(snapshot);
        var group = store.Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey);
        var desired = snapshot.Value.Desired.Select(p => p.Login).ToArray();
        // Past the recordable number of skips, the rest of this group's member changes wait for
        // the next scheduled refresh.
        bool stopped = op.Value.SkippedMembers.Contains(op.Value.GroupKey + "|*");
        var remove = stopped
            ? null
            : op.Value.Members.FirstOrDefault(p =>
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
        var add = stopped
            ? null
            : desired.FirstOrDefault(login =>
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
            if (
                snapshot
                    .Value.Desired.First(p =>
                        string.Equals(p.Login, add, StringComparison.OrdinalIgnoreCase)
                    )
                    .Group
            )
            {
                // From here the claim is Documents' own in this group, even if the add's outcome
                // is never known, so a later change of the team's group removes it.
                var owned = group.Value.Principals ?? Array.Empty<string>();
                if (!owned.Contains(add, StringComparer.OrdinalIgnoreCase))
                {
                    group.Value.Principals = owned.Concat(new[] { add }).ToArray();
                    store.Save(group);
                }
                if (!string.Equals(op.Value.EnsuredLogin, add, StringComparison.OrdinalIgnoreCase))
                    return Prepare(
                        op,
                        claim,
                        catalog,
                        "PrincipalEnsure",
                        SharePointRequests.EnsurePrincipal(add)
                    );
            }
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
        // A group with skipped members does not record the team's hash, so every scheduled
        // refresh sees a difference and retries them.
        string applied = TeamSnapshotReader.Hash(snapshot.Value.Desired);
        group.Value.MembershipHash = op.Value.SkippedMembers.Any(k =>
            k.StartsWith(op.Value.GroupKey + "|", StringComparison.Ordinal)
        )
            ? RetryHashPrefix + applied
            : applied;
        // A claim Documents removed is no longer its own; one SharePoint kept stays managed.
        // Without a complete read of the group nothing was removed, so all stay managed.
        if (!stopped)
            group.Value.Principals = (group.Value.Principals ?? Array.Empty<string>())
                .Where(p =>
                    desired.Contains(p, StringComparer.OrdinalIgnoreCase)
                    || op.Value.Members.Any(m =>
                        string.Equals(m.Login, p, StringComparison.OrdinalIgnoreCase)
                    )
                )
                .ToArray();
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
        StoredRow<DispatcherDocument> claimRow,
        SecurityCatalog catalog
    )
    {
        var claim = claimRow.Value;
        var entry = op.Value.Entries[op.Value.TeamIndex];
        var group = store.Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey).Value;
        string key = "grant:" + catalog.LibraryId.ToString("N") + ":" + entry.TeamId.ToString("N");
        var grant = store.Find<ManagedGrant>("asx_managedgrant", key);
        if (op.Value.Acl.Any(a => a.Member.Type != 8))
            throw new EvaluationBlockedException("Managed group principal type changed.");
        // Limited Access (RoleTypeKind 1) is added and removed by SharePoint itself when
        // something is shared with the group. It is not Documents' grant and is never touched.
        int[] roles = op
            .Value.Acl.SelectMany(a => a.Roles?.Rows ?? Array.Empty<AclRole>())
            .Where(r => r.Type != LimitedAccessKind)
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
            {
                // SharePoint confirmed the write but the read does not show it yet; reads can
                // lag a write. Read again after a short wait. Three consecutive read-backs are
                // the consistency window for one write, not a cap on work.
                if (++op.Value.ReadbackMisses >= ReadbackWindow)
                    throw new EvaluationBlockedException(
                        "Grant independent readback did not match the submitted change."
                    );
                return Wait(op, claimRow, null, null, "GrantReadbackPending");
            }
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
        if (!Granted(entry.TeamId))
            desired = 0;
        int[] target = desired == 0 ? Array.Empty<int>() : new[] { desired };
        // What Documents last applied. A receipt still Pending means Documents' own write may
        // have landed without a confirmed outcome, so a difference is not reported as a hand edit.
        bool known = grant == null || grant.Value.Status == "Applied";
        int recorded = RecordedRole(grant?.Value, group);
        int[] expected = recorded == 0 ? Array.Empty<int>() : new[] { recorded };
        if (roles.SequenceEqual(target))
        {
            // Already at the configured level, including a Documents write whose outcome was
            // unknown but landed. Record it and move on.
            // A receipt of an earlier group (deleted in SharePoint and created again) is stale:
            // that group's grant went with it, so the receipt is brought to the current group.
            if (
                grant != null
                && (
                    recorded != desired
                    || grant.Value.Status != "Applied"
                    || grant.Value.GroupId != group.GroupId
                )
            )
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

    // SharePoint RoleTypeKind for Limited Access.
    private const int LimitedAccessKind = 1;

    // Consecutive read-backs of one confirmed grant write before it is treated as not applied.
    private const int ReadbackWindow = 3;

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

    /// <summary>
    /// Whether the team keeps its configured access: it is still registered and still exists in
    /// Dataverse. A team deleted since its run was queued gets none, also before its
    /// registration is retired.
    /// </summary>
    private bool Granted(Guid teamId) =>
        store
            .Require<TeamRegistration>("asx_teamregistration", "team:" + teamId.ToString("N"))
            .Value.Enabled
        && TeamDirectory.Find(service, teamId) != null;

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
        var before = policy.Value.Applied;
        var live = TeamDirectory.Read(service, op.Value.Entries.Select(e => e.TeamId));
        policy.Value.Applied = op
            .Value.Entries.Select(e => new PolicyEntry
            {
                TeamId = e.TeamId,
                Access =
                    live.ContainsKey(e.TeamId)
                    && store
                        .Require<TeamRegistration>(
                            "asx_teamregistration",
                            "team:" + e.TeamId.ToString("N")
                        )
                        .Value.Enabled
                        ? e.Access
                        : "None",
            })
            .ToArray();
        foreach (var entry in policy.Value.Applied)
        {
            var reference = store.Require<PolicyTeamReference>(
                "asx_policyentry",
                "policyteam:" + op.Value.LibraryId.ToString("N") + ":" + entry.TeamId.ToString("N")
            );
            reference.Value.Status = entry.Access == "None" ? "Inactive" : "Active";
            store.Save(reference);
        }
        // A team deleted in Dataverse no longer has access here. Its registration is finished
        // once no library's access refers to it; its Documents group stays in SharePoint.
        foreach (var entry in op.Value.Entries.Where(e => !live.ContainsKey(e.TeamId)))
        {
            TeamDirectory.Retire(store, entry.TeamId);
            bool finished = TeamDirectory.Finish(service, store, entry.TeamId);
            if (!before.Any(a => a.TeamId == entry.TeamId && a.Access != "None"))
                continue;
            var group = store.Find<ManagedGroup>(
                "asx_managedgroup",
                "group:" + catalog.SiteId.ToString("N") + ":" + entry.TeamId.ToString("N")
            );
            string name =
                store
                    .Find<TeamRegistration>(
                        "asx_teamregistration",
                        "team:" + entry.TeamId.ToString("N")
                    )
                    ?.Value.Name
                ?? group?.Value.Title
                ?? entry.TeamId.ToString("D");
            Notice(op.Value, TeamDirectory.DeletedNotice(name, finished, group?.Value.GroupId > 0));
        }
        policy.Value.Queued = Array.Empty<PolicyEntry>();
        policy.Value.OperationKey = null;
        policy.Value.Status = "Applied";
        policy.Value.Inherits = false;
        policy.Value.Notices = op.Value.Notices;
        // Members SharePoint refused, or a team too large for one run, leave membership unsynced
        // even though the grants are applied; the library shows that it needs attention.
        policy.Value.MembershipIncomplete = op.Value.SkippedMembers.Length > 0;
        policy.Value.ResidualAccess = new[]
        {
            "Documents manages only its own team groups and their grants on this library. People may still have access through other groups, direct shares, links, item permissions or site administration.",
        };
        store.Save(policy);
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
        // A re-pointed site keeps its web ID; the run follows its new address.
        if (catalog.Target.WebId == op.WebId && catalog.Target.Web.AbsoluteUri != op.WebUrl)
            op.WebUrl = catalog.Target.Web.AbsoluteUri;
        if (
            catalog.Target.WebId != op.WebId
            || catalog.Target.ListId != op.ListId
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
            // A team deleted before the run reached it has no group here and no receipt.
            if (op.Entries.Single(e => e.TeamId == teamId).Access != "None" && Granted(teamId))
                throw new EvaluationBlockedException("Membership receipt missing.");
            return;
        }
        bool granted = Granted(teamId);
        string revoked =
            TeamDirectory.Find(service, teamId) == null
                ? TeamDirectory.DeletedDuringRun
                : TeamDirectory.TurnedOffDuringRun;
        if (
            !snapshot.Value.Complete
            || snapshot.Value.Status != "Applied"
            || TeamSnapshotReader.Hash(snapshot.Value.Desired)
                != TeamSnapshotReader.Hash(new TeamSnapshotReader(service).Read(teamId))
        )
            throw new EvaluationBlockedException(
                granted ? "Final membership generation changed or is incomplete." : revoked
            );
        // A team with no members deleted or turned off after its grant was confirmed leaves no
        // membership difference, only its grant: the run stops rather than record no access
        // while the grant stays, and the next run removes it.
        if (!granted && RecordedRole(store, op.LibraryId, op.SiteId, teamId) > 0)
            throw new EvaluationBlockedException(revoked);
    }

    /// <summary>
    /// The role Documents' grant receipt records for the team's current Documents group, or 0.
    /// A receipt naming another group, one deleted in SharePoint and created again since, is
    /// stale: that group's grant went with it. The grant reconcile, the end-of-run check and the
    /// refresh's replace rule all use this.
    /// </summary>
    internal static int RecordedRole(ManagedGrant? grant, ManagedGroup? group) =>
        grant != null && group != null && grant.GroupId == group.GroupId ? grant.RoleId : 0;

    internal static int RecordedRole(
        DocumentStore store,
        Guid libraryId,
        Guid siteId,
        Guid teamId
    ) =>
        RecordedRole(
            store
                .Find<ManagedGrant>(
                    "asx_managedgrant",
                    "grant:" + libraryId.ToString("N") + ":" + teamId.ToString("N")
                )
                ?.Value,
            store
                .Find<ManagedGroup>(
                    "asx_managedgroup",
                    "group:" + siteId.ToString("N") + ":" + teamId.ToString("N")
                )
                ?.Value
        );

    private static string Bounded(string text)
    {
        text = new string(text.Where(c => !char.IsControl(c)).ToArray()).Trim();
        return text.Length > NoticeLength ? text.Substring(0, NoticeLength) + "..." : text;
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
                    + "/roleassignments/getbyprincipalid("
                    + store
                        .Require<ManagedGroup>("asx_managedgroup", op.Value.GroupKey)
                        .Value.GroupId
                    + ")?$select=Member/Id,Member/PrincipalType,RoleDefinitionBindings/Id,RoleDefinitionBindings/RoleTypeKind&$expand=Member,RoleDefinitionBindings";
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
    private const int NoticeLength = 600;

    // Marks a group hash that must not match the team, so the next refresh retries skips.
    private const string RetryHashPrefix = "retry:";

    /// <summary>Records something an admin should see after the run, once.</summary>
    private static void Notice(SecurityOperation op, string text)
    {
        text = new string(text.Where(c => !char.IsControl(c)).ToArray());
        if (text.Length > NoticeLength)
            text = text.Substring(0, NoticeLength) + "...";
        if (op.Notices.Length >= NoticeLimit)
            text = "More notices were not recorded.";
        if (op.Notices.Length > NoticeLimit || op.Notices.Contains(text, StringComparer.Ordinal))
            return;
        op.Notices = op.Notices.Concat(new[] { text }).ToArray();
    }

    /// <summary>Records a member add or remove SharePoint rejected, with SharePoint's message.</summary>
    private void SkipMember(SecurityOperation op, WorkerRequest request)
    {
        // A group claim SharePoint cannot resolve is skipped like a person it will not add.
        bool adding = op.MutationKind != "MemberRemove";
        string login = adding
            ? op.MutationLogin
            : op.Members.FirstOrDefault(m => m.Id == op.MutationMemberId)?.Login
                ?? "SharePoint user " + op.MutationMemberId;
        string title = store.Require<ManagedGroup>("asx_managedgroup", op.GroupKey).Value.Title;
        op.SkippedCount++;
        if (!Remember(op, adding ? "|" + op.MutationLogin : "|#" + op.MutationMemberId, title))
            return;
        Notice(
            op,
            "SharePoint did not "
                + (adding ? "add " : "remove ")
                + Upn(login)
                + (adding ? " to" : " from")
                + " the group for team '"
                + title
                + "': "
                + SharePointObservations.ErrorMessage(request)
        );
    }

    /// <summary>
    /// SharePoint accepted a member change but the group does not show it, for example because
    /// it lists the person under a renamed sign-in name. The member is skipped with a notice;
    /// people who appeared with the add are kept so the next pass does not remove them.
    /// </summary>
    private void SkipUnconfirmedMember(SecurityOperation op, SitePerson[] before, bool adding)
    {
        string title = store.Require<ManagedGroup>("asx_managedgroup", op.GroupKey).Value.Title;
        string login = adding
            ? op.MutationLogin
            : op.Members.FirstOrDefault(m => m.Id == op.MutationMemberId)?.Login
                ?? "SharePoint user " + op.MutationMemberId;
        var keys = adding
            ? new[] { "|" + op.MutationLogin }
                .Concat(
                    op.Members.Where(m => !before.Any(b => b.Id == m.Id)).Select(m => "|#" + m.Id)
                )
                .ToArray()
            : new[] { "|#" + op.MutationMemberId };
        op.SkippedCount++;
        foreach (var key in keys)
            if (!Remember(op, key, title))
                return;
        Notice(
            op,
            "Team '"
                + title
                + "': SharePoint accepted "
                + (adding ? "adding " : "removing ")
                + Upn(login)
                + " but the group "
                + (adding ? "does not list that sign-in name" : "still lists it")
                + ", for example because the sign-in name changed. Documents left the group as SharePoint shows it."
        );
    }

    // A person is shown by sign-in name; a group claim (c:0...) is shown whole, since its last
    // part is only an object ID.
    private static string Upn(string login) =>
        login.StartsWith("c:", StringComparison.Ordinal)
            ? login
            : login.Substring(login.LastIndexOf('|') + 1);

    // asx_operation.asx_payload holds at most 500,000 characters (MaxLength in
    // asx_operation/Entity.xml, also enforced by JsonWire.Read). Skip details stop where the
    // stored operation, plus room for the full set of notices this run can still record, would
    // pass that column size.
    private const int PayloadMaxLength = 500000;
    private const int SkipDetailBudget = PayloadMaxLength - NoticeLimit * (NoticeLength + 16);

    /// <summary>
    /// Records a skipped member change for this run's group. Returns false once the details
    /// would no longer fit the operation payload; the group then stops member changes for this
    /// run and the next scheduled refresh retries them.
    /// </summary>
    private static bool Remember(SecurityOperation op, string member, string title)
    {
        string key = op.GroupKey + member;
        string stop = op.GroupKey + "|*";
        if (op.SkippedMembers.Contains(stop, StringComparer.Ordinal))
            return false;
        if (JsonWire.Write(op).Length + key.Length + 4 <= SkipDetailBudget)
        {
            op.SkippedMembers = op.SkippedMembers.Concat(new[] { key }).ToArray();
            return true;
        }
        op.SkippedMembers = op.SkippedMembers.Concat(new[] { stop }).ToArray();
        Notice(
            op,
            "Team '"
                + title
                + "': "
                + op.SkippedCount
                + " member changes were skipped, more than one run can record. Team membership was not synced: people removed from the team keep access, and people added get none, until this is resolved. The remaining member changes for this team are retried on the next scheduled refresh."
        );
        return false;
    }

    private static void ClearMutation(SecurityOperation op)
    {
        op.ReadbackMisses = 0;
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
