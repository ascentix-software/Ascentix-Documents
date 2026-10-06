using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public static class CompleteQuery
{
    /// <summary>
    /// Reads every page within the row limit and rejects duplicate rows or invalid continuations.
    /// </summary>
    /// <param name="service">The Dataverse service used to retrieve query pages.</param>
    /// <param name="query">The query to execute; its paging information is updated during traversal.</param>
    /// <param name="maximum">The maximum total row count accepted for the complete result.</param>
    /// <returns>All rows returned by the completed query.</returns>
    public static Entity[] Read(
        IOrganizationService service,
        QueryExpression query,
        int maximum = 1000
    )
    {
        if (query.TopCount.HasValue)
            throw new EvaluationBlockedException("Complete queries cannot use TopCount.");
        query.PageInfo = new PagingInfo { Count = Math.Min(500, maximum), PageNumber = 1 };
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
            if (rows.Count > maximum)
                throw new EvaluationBlockedException("Complete query bound exceeded.");
            if (!page.MoreRecords)
                return rows.ToArray();
            if (
                page.Entities.Count == 0
                || string.IsNullOrEmpty(page.PagingCookie)
                || !cookies.Add(page.PagingCookie)
                || ++query.PageInfo.PageNumber > 20
            )
                throw new EvaluationBlockedException("Incomplete query pagination.");
            query.PageInfo.PagingCookie = page.PagingCookie;
        }
    }
}

public sealed class SecurityRefresh
{
    private readonly IOrganizationService service;
    private readonly DocumentStore store;
    private readonly Func<DateTime> clock;
    private bool waiting;

    public SecurityRefresh(IOrganizationService service, Func<DateTime>? clock = null)
    {
        this.service = service;
        store = new DocumentStore(service);
        this.clock = clock ?? (() => DateTime.UtcNow);
    }

    public WorkerResult Scan()
    {
        var query = new QueryExpression("asx_policy")
        {
            ColumnSet = new ColumnSet("asx_payload"),
            TopCount = 5,
        };
        query.Criteria.AddCondition("asx_reviewafter", ConditionOperator.LessEqual, clock());
        query.AddOrder("asx_reviewafter", OrderType.Ascending);
        var rows = service.RetrieveMultiple(query).Entities;
        var keys = Refresh(rows, Guid.Empty);
        foreach (var row in rows)
        {
            var pointer = JsonWire.Read<PolicyDocument>(TemplateStore.Text(row, "asx_payload"));
            var current = store.Require<PolicyDocument>("asx_policy", pointer.Key);
            current.Value.NextReviewUtc = clock().AddDays(1);
            store.Save(current);
        }
        return new WorkerResult { Status = "ReviewedDuePolicies", Keys = keys };
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
                || policy.Value.Status == "Removed"
                || onlyTeam != Guid.Empty && !policy.Value.Approved.Any(e => e.TeamId == onlyTeam)
            )
                continue;
            // A suspended or removed library (or site) is the admin's stop for that library
            // only: it is skipped with a notice, and every other library still refreshes.
            if (Paused(policy))
                continue;
            policy = store.Require<PolicyDocument>("asx_policy", pointer.Key);
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
                snapshotChanged = false;
            var operation =
                policy.Value.OperationKey == null
                    ? null
                    : store.Require<SecurityOperation>("asx_operation", policy.Value.OperationKey);
            foreach (
                var entry in policy.Value.Approved.Where(e =>
                    onlyTeam == Guid.Empty || e.TeamId == onlyTeam
                )
            )
            {
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
                if (
                    !snapshotChanged
                    || operation.Value.ExternalSubmitted
                    || (operation.Value.Status != "Pending" && operation.Value.Status != "Blocked")
                )
                {
                    if (snapshotChanged)
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
