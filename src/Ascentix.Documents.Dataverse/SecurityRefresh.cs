using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public static class CompleteQuery
{
    // Dataverse returns at most 5,000 rows per page ("Page results using QueryExpression",
    // https://learn.microsoft.com/power-apps/developer/data-platform/org-service/queryexpression/page-results).
    private const int PageSize = 5000;

    /// <summary>
    /// Reads every page, however many rows there are, and rejects duplicate rows or invalid
    /// continuations.
    /// </summary>
    /// <param name="service">The Dataverse service used to retrieve query pages.</param>
    /// <param name="query">The query to execute; its paging information is updated during traversal.</param>
    /// <returns>All rows returned by the completed query.</returns>
    public static Entity[] Read(IOrganizationService service, QueryExpression query)
    {
        if (query.TopCount.HasValue)
            throw new EvaluationBlockedException("Complete queries cannot use TopCount.");
        query.PageInfo = new PagingInfo { Count = PageSize, PageNumber = 1 };
        // Pages follow each other only in a unique order (same article): the primary key.
        if (query.Orders.Count == 0)
            query.AddOrder(query.EntityName + "id", OrderType.Ascending);
        var rows = new List<Entity>();
        var ids = new HashSet<Guid>();
        var cookies = new HashSet<string>();
        while (true)
        {
            var page = service.RetrieveMultiple(query);
            foreach (var row in page.Entities)
                if (!ids.Add(row.Id))
                    throw new EvaluationBlockedException("Duplicate/changing complete query page.");
                else
                    rows.Add(row);
            if (!page.MoreRecords)
                return rows.ToArray();
            if (
                page.Entities.Count == 0
                || string.IsNullOrEmpty(page.PagingCookie)
                || !cookies.Add(page.PagingCookie)
            )
                throw new EvaluationBlockedException("Incomplete query pagination.");
            query.PageInfo.PageNumber++;
            query.PageInfo.PagingCookie = page.PagingCookie;
        }
    }
}

public sealed class SecurityRefresh
{
    private readonly IOrganizationService service;
    private readonly DocumentStore store;
    private readonly Func<DateTime> clock;
    private readonly Func<TimeSpan>? elapsed;
    private bool waiting;

    /// <param name="service">The Dataverse service.</param>
    /// <param name="clock">The current time; review times are set from it.</param>
    /// <param name="elapsed">Time spent in this scan so far; by default a stopwatch.</param>
    public SecurityRefresh(
        IOrganizationService service,
        Func<DateTime>? clock = null,
        Func<TimeSpan>? elapsed = null
    )
    {
        this.service = service;
        store = new DocumentStore(service);
        this.clock = clock ?? (() => DateTime.UtcNow);
        this.elapsed = elapsed;
    }

    // The flow's Refresh_approved_security step is one call of the worker custom API, which
    // Dataverse stops after 2 minutes ("Analyze plug-in performance",
    // https://learn.microsoft.com/power-apps/developer/data-platform/analyze-performance; the
    // same limit sizes EventRegistrations.MaxNewTablesPerSave). The scan uses half of it and
    // leaves the rest of the due policies to the next run, a minute later.
    public static readonly TimeSpan ScanBudget = TimeSpan.FromMinutes(1);

    // Due policies are read a page at a time, like the dispatch page (DocumentStore.Pending);
    // the scan goes on page after page until none is due or the time budget is used.
    private const int ScanPage = 20;

    /// <summary>
    /// Reviews every policy whose daily review is due, oldest first, until none is left or the
    /// time budget is used; the next run continues with the rest.
    /// </summary>
    public WorkerResult Scan()
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        Func<TimeSpan> spent = elapsed ?? (() => watch.Elapsed);
        var keys = new List<string>();
        var reviewed = new HashSet<Guid>();
        while (true)
        {
            var query = new QueryExpression("asx_policy")
            {
                ColumnSet = new ColumnSet("asx_payload"),
                PageInfo = new PagingInfo { Count = ScanPage, PageNumber = 1 },
            };
            query.Criteria.AddCondition("asx_reviewafter", ConditionOperator.LessEqual, clock());
            query.AddOrder("asx_reviewafter", OrderType.Ascending);
            // A reviewed policy is no longer due, so each read starts with the next ones.
            var rows = service
                .RetrieveMultiple(query)
                .Entities.Where(row => !reviewed.Contains(row.Id))
                .ToArray();
            if (rows.Length == 0)
                break;
            foreach (var row in rows)
            {
                if (spent() >= ScanBudget)
                    return new WorkerResult
                    {
                        Status = "ReviewedDuePolicies",
                        Keys = keys.ToArray(),
                    };
                reviewed.Add(row.Id);
                keys.AddRange(Refresh(new[] { row }, Guid.Empty));
                var pointer = JsonWire.Read<PolicyDocument>(TemplateStore.Text(row, "asx_payload"));
                var current = store.Require<PolicyDocument>("asx_policy", pointer.Key);
                // A change still waiting for the run before it is looked at again by the next
                // dispatch run, a minute later (the flow's recurrence).
                current.Value.NextReviewUtc = current.Value.ApplyPending
                    ? clock().AddMinutes(1)
                    : clock().AddDays(1);
                store.Save(current);
            }
        }
        return new WorkerResult { Status = "ReviewedDuePolicies", Keys = keys.ToArray() };
    }

    /// <summary>
    /// Refreshes one team-policy page and advances its cursor after affected work is resolved.
    /// </summary>
    /// <param name="job">The versioned team invalidation, including its page cursor and work receipt.</param>
    /// <returns>The invalidation status and affected policy-operation keys.</returns>
    public WorkerResult Consume(StoredRow<OutboxDocument> job)
    {
        if (job.Value.Status != "Pending")
            return new WorkerResult
            {
                Status = job.Value.Status,
                Key = job.Value.Key,
                Keys = job.Value.Operations,
            };
        var page = Page(job.Value.SecurityTeamId, job.Value.SecurityPage, job.Value.SecurityCookie);
        var policies = page.Entities.Select(row =>
            store
                .Require<PolicyDocument>(
                    "asx_policy",
                    JsonWire
                        .Read<PolicyTeamReference>(TemplateStore.Text(row, "asx_payload"))
                        .PolicyKey
                )
                .Row
        );
        waiting = false;
        var keys = Refresh(policies, job.Value.SecurityTeamId);
        if (waiting)
            return new WorkerResult
            {
                Status = "Pending",
                Key = job.Value.Key,
                Keys = keys,
            };
        // A team deleted in Dataverse that no library's access refers to any more, such as one
        // whose libraries were all removed from Documents, has nothing left to remove: its
        // registration is finished now. Otherwise each library's next access run finishes it.
        if (!page.MoreRecords && TeamDirectory.Find(service, job.Value.SecurityTeamId) == null)
        {
            TeamDirectory.Retire(store, job.Value.SecurityTeamId);
            TeamDirectory.Finish(service, store, job.Value.SecurityTeamId);
        }
        job.Value.Operations = keys;
        job.Value.SecurityPage = page.MoreRecords ? job.Value.SecurityPage + 1 : 1;
        job.Value.SecurityCookie = page.MoreRecords ? page.PagingCookie : null;
        job.Value.Status = page.MoreRecords ? "Pending" : "Planned";
        store.Save(job);
        return new WorkerResult
        {
            Status = job.Value.Status,
            Key = job.Value.Key,
            Keys = keys,
        };
    }

    private EntityCollection Page(Guid team, int number, string? cookie)
    {
        var query = new QueryExpression("asx_policyentry")
        {
            ColumnSet = new ColumnSet("asx_payload"),
            PageInfo = new PagingInfo
            {
                Count = 5,
                PageNumber = number,
                PagingCookie = cookie,
            },
        };
        query.Criteria.AddCondition("asx_teamkey", ConditionOperator.Equal, team.ToString("N"));
        query.Criteria.AddCondition("asx_status", ConditionOperator.Equal, "Active");
        query.AddOrder("asx_policyentryid", OrderType.Ascending);
        var page = service.RetrieveMultiple(query);
        if (
            page.Entities.Count > 5
            || page.Entities.Select(e => e.Id).Distinct().Count() != page.Entities.Count
            || page.MoreRecords
                && (
                    page.Entities.Count == 0
                    || string.IsNullOrEmpty(page.PagingCookie)
                    || page.PagingCookie == cookie
                )
        )
            throw new EvaluationBlockedException("Team policy continuation is incomplete.");
        return page;
    }

    private string[] Refresh(IEnumerable<Entity> rows, Guid onlyTeam)
    {
        var queued = new List<string>();
        foreach (var row in rows)
        {
            var pointer = JsonWire.Read<PolicyDocument>(
                row.GetAttributeValue<string>("asx_payload")
            );
            var policy = store.Require<PolicyDocument>("asx_policy", pointer.Key);
            if (
                policy.Value.Generation == Guid.Empty
                || onlyTeam != Guid.Empty && !policy.Value.Approved.Any(e => e.TeamId == onlyTeam)
            )
                continue;
            // A suspended or removed library (or site) is the admin's stop for that library
            // only: it is skipped with a notice, and every other library still refreshes. This
            // reads the library as it is now, not the policy's status: a library removed and
            // then added again is refreshed again, though its policy still says Removed.
            if (Paused(policy))
                continue;
            // A library added again before re-adding resumed its access refresh still has its
            // team references Inactive; they are reactivated so team events reach it again.
            if (policy.Value.Status == "Removed")
                ResumeTeamReferences(store, policy.Value);
            policy = store.Require<PolicyDocument>("asx_policy", pointer.Key);
            if (policy.Value.ApplyPending)
            {
                // An admin's change that waited for the run before it: applied once that run
                // can be replaced, otherwise it waits for the next review.
                var applied = new SecurityAdministration(service).Execute(
                    new SecurityRequest
                    {
                        Command = "ApplyPending",
                        LibraryId = policy.Value.LibraryId,
                        RowVersion = policy.Row.RowVersion,
                    },
                    true,
                    true
                );
                if (applied.Policy?.ApplyPending == false && applied.Policy.OperationKey != null)
                    queued.Add(applied.Policy.OperationKey);
                continue;
            }
            if (
                store
                    .Find<DispatcherDocument>(
                        "asx_claim",
                        WorkCoordination.Library(service, policy.Value.LibraryId)
                    )
                    ?.Value.RunId != null
            )
            {
                waiting = true;
                continue;
            }
            var catalog = new SecurityCatalog(service, policy.Value.LibraryId);
            bool changed = false,
                snapshotChanged = false,
                deletedGrant = false;
            var operation =
                policy.Value.OperationKey == null
                    ? null
                    : store.Require<SecurityOperation>("asx_operation", policy.Value.OperationKey);
            var approved = policy
                .Value.Approved.Where(e => onlyTeam == Guid.Empty || e.TeamId == onlyTeam)
                .ToArray();
            var live = TeamDirectory.Read(service, approved.Select(e => e.TeamId));
            foreach (var entry in approved)
            {
                // A team deleted in Dataverse is retired here too, so its access is removed even
                // when its delete event was missed or came while the library was out of
                // Documents (removed or suspended): the check below then sees it revoked.
                if (!live.ContainsKey(entry.TeamId))
                {
                    TeamDirectory.Retire(store, entry.TeamId);
                    // Documents' grant for it is still recorded on this library, for example
                    // when the team was deleted after its run confirmed the grant.
                    deletedGrant |=
                        store
                            .Find<ManagedGrant>(
                                "asx_managedgrant",
                                "grant:"
                                    + catalog.LibraryId.ToString("N")
                                    + ":"
                                    + entry.TeamId.ToString("N")
                            )
                            ?.Value.RoleId > 0;
                }
                string hash = TeamSnapshotReader.Hash(
                    new TeamSnapshotReader(service).Read(entry.TeamId)
                );
                var group = store.Find<ManagedGroup>(
                    "asx_managedgroup",
                    "group:" + catalog.SiteId.ToString("N") + ":" + entry.TeamId.ToString("N")
                );
                if (group != null && group.Value.MembershipHash != hash)
                    changed = true;
                if (
                    !store
                        .Require<TeamRegistration>(
                            "asx_teamregistration",
                            "team:" + entry.TeamId.ToString("N")
                        )
                        .Value.Enabled
                    && policy.Value.Applied.Any(a => a.TeamId == entry.TeamId && a.Access != "None")
                )
                    changed = true;
                if (operation != null)
                {
                    var snapshot = store.Find<MembershipDocument>(
                        "asx_membership",
                        "membership:"
                            + operation.Value.PolicyRevision.ToString("N")
                            + ":"
                            + entry.TeamId.ToString("N")
                    );
                    if (snapshot != null && TeamSnapshotReader.Hash(snapshot.Value.Desired) != hash)
                        snapshotChanged = true;
                }
            }
            if (operation != null)
            {
                // Only a proved newer team snapshot may replace idle, unsubmitted work. Unknown/active grants never overlap.
                // A run stopped because its team was deleted mid-run (TeamDirectory.DeletedDuringRun)
                // is replaced too, so the new run removes the deleted team's grant.
                bool replace =
                    snapshotChanged || deletedGrant && operation.Value.Status == "Blocked";
                if (
                    !replace
                    || operation.Value.ExternalSubmitted
                    || (operation.Value.Status != "Pending" && operation.Value.Status != "Blocked")
                )
                {
                    if (replace)
                        waiting = true;
                    continue;
                }
                new SecurityWorker(service).Execute(
                    new WorkerRequest { Command = "Cancel", Key = operation.Value.Key },
                    true
                );
                policy = store.Require<PolicyDocument>("asx_policy", pointer.Key);
                changed = true;
            }
            if (!changed)
                continue;
            var result = new SecurityAdministration(service).Execute(
                new SecurityRequest
                {
                    Command = "RefreshPolicy",
                    LibraryId = policy.Value.LibraryId,
                    RowVersion = policy.Row.RowVersion,
                },
                true,
                true
            );
            if (result.Policy?.OperationKey != null)
                queued.Add(result.Policy.OperationKey);
        }
        return queued.ToArray();
    }

    /// <summary>
    /// Reactivates the team references of a library's applied access, as the access run that
    /// applied it left them, so team events schedule the library again after it was removed
    /// from Documents and added again.
    /// </summary>
    public static void ResumeTeamReferences(DocumentStore store, PolicyDocument policy)
    {
        foreach (var entry in policy.Applied.Where(e => e.Access != "None"))
        {
            var reference = store.Find<PolicyTeamReference>(
                "asx_policyentry",
                "policyteam:" + policy.LibraryId.ToString("N") + ":" + entry.TeamId.ToString("N")
            );
            if (reference == null || reference.Value.Status == "Active")
                continue;
            reference.Value.Status = "Active";
            store.Save(reference);
        }
    }

    public const string SuspendedNotice =
        "Access refresh is paused because this library or its site is suspended. Existing access is unchanged; the refresh resumes by itself once both are approved again.";

    public const string RemovedNotice =
        "Access refresh stopped because this library or its site was removed from Documents. Existing access is unchanged.";

    /// <summary>
    /// True when the policy's library or site is suspended or removed, after recording that on
    /// the policy once. Otherwise clears an earlier pause notice.
    /// </summary>
    private bool Paused(StoredRow<PolicyDocument> policy)
    {
        string? stop = WorkCoordination.Stops(service, policy.Value.LibraryId);
        string? notice =
            stop == "suspended" ? SuspendedNotice
            : stop == null ? null
            : RemovedNotice;
        var kept = policy
            .Value.Notices.Where(n => n != SuspendedNotice && n != RemovedNotice)
            .ToArray();
        var notices = notice == null ? kept : kept.Concat(new[] { notice }).ToArray();
        if (!notices.SequenceEqual(policy.Value.Notices))
        {
            policy.Value.Notices = notices;
            store.Save(policy);
        }
        return stop != null;
    }
}
