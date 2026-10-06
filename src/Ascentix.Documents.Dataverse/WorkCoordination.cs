using System;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class ConnectionBudget : StoredDocument
{
    [DataMember]
    public string[] Writers { get; set; } = Array.Empty<string>();

    [DataMember]
    public DateTime NextStartUtc { get; set; }

    [DataMember]
    public DateTime PauseUntilUtc { get; set; }

    [DataMember]
    public int PaceSeconds { get; set; } = 1;

    [DataMember]
    public DateTime RecoverPaceUtc { get; set; }
}

public static class WorkCoordination
{
    public const string BudgetKey = "connection-budget:v1";

    public static string SiteUrl(string url) =>
        "site-writer:" + DocumentStore.Hash(url.TrimEnd('/').ToLowerInvariant());

    public static string Site(IOrganizationService service, Guid id) =>
        SiteUrl(
            TemplateStore.Text(
                service.Retrieve("asx_site", id, new ColumnSet("asx_url")),
                "asx_url"
            )
        );

    public static string Library(IOrganizationService service, Guid id) =>
        Site(
            service,
            service
                .Retrieve("asx_library", id, new ColumnSet("asx_siteid"))
                .GetAttributeValue<EntityReference>("asx_siteid")
                .Id
        );

    public static string Operation(IOrganizationService service, string key)
    {
        var store = new DocumentStore(service);
        if (key.StartsWith("catalogprobe:", StringComparison.Ordinal))
            return SiteUrl(store.Require<CatalogProbe>("asx_operation", key).Value.WebUrl);
        // Library setup and access runs use their site's current address, like folder jobs, so
        // a re-pointed site keeps one writer (see MoveWriter).
        if (key.StartsWith("librarycreate:", StringComparison.Ordinal))
        {
            var setup = store.Require<LibrarySetup>("asx_operation", key).Value;
            return SiteUrl(Current(service, "asx_site", setup.SiteId) ?? setup.WebUrl);
        }
        if (key.StartsWith("policywork:", StringComparison.Ordinal))
        {
            var access = store.Require<SecurityOperation>("asx_operation", key).Value;
            var library = Find(service, "asx_library", access.LibraryId, "asx_siteid");
            var site = library?.GetAttributeValue<EntityReference>("asx_siteid");
            return SiteUrl(
                (site == null ? null : Current(service, "asx_site", site.Id)) ?? access.WebUrl
            );
        }
        var op = store.Require<OperationDocument>("asx_operation", key).Value;
        return Library(
            service,
            op.HistoryCompacted || op.Folders.Length == 0 ? op.ResultLibraryId : op.Folder.LibraryId
        );
    }

    /// <summary>
    /// Moves a re-pointed site's writer row to its new address, including a run that holds it,
    /// so the run's next step finds its claim and the site still has one writer.
    /// </summary>
    public static void MoveWriter(IOrganizationService service, string oldUrl, string newUrl)
    {
        string from = SiteUrl(oldUrl),
            to = SiteUrl(newUrl);
        if (from == to)
            return;
        var store = new DocumentStore(service);
        var old = store.Find<DispatcherDocument>("asx_claim", from);
        if (old == null || old.Value.RunId == null)
            return;
        var moved = store.Find<DispatcherDocument>("asx_claim", to);
        if (moved == null)
        {
            store.Create("asx_claim", new DispatcherDocument { Key = to, Status = "Idle" });
            moved = store.Require<DispatcherDocument>("asx_claim", to);
        }
        if (moved.Value.RunId != null)
            return;
        moved.Value.OperationKey = old.Value.OperationKey;
        moved.Value.RunId = old.Value.RunId;
        moved.Value.Token = old.Value.Token;
        moved.Value.LeaseUntilUtc = old.Value.LeaseUntilUtc;
        moved.Value.HttpOutstanding = old.Value.HttpOutstanding;
        moved.Value.RecoveryPermitted = old.Value.RecoveryPermitted;
        moved.Value.TerminationEvidence = old.Value.TerminationEvidence;
        moved.Value.Status = old.Value.Status;
        store.Save(moved);
        old = store.Require<DispatcherDocument>("asx_claim", from);
        old.Value.HttpOutstanding = false;
        old.Value.RunId = null;
        old.Value.OperationKey = null;
        old.Value.Token = Guid.Empty;
        old.Value.RecoveryPermitted = false;
        old.Value.Status = "Idle";
        store.Save(old);
    }

    private static Entity? Find(IOrganizationService service, string table, Guid id, string column)
    {
        var query = new QueryExpression(table) { ColumnSet = new ColumnSet(column), TopCount = 1 };
        query.Criteria.AddCondition(table + "id", ConditionOperator.Equal, id);
        return service.RetrieveMultiple(query).Entities.FirstOrDefault();
    }

    // The site's current address; null when the row no longer exists.
    private static string? Current(IOrganizationService service, string table, Guid id) =>
        Find(service, table, id, "asx_url")?.GetAttributeValue<string>("asx_url");

    private static StoredRow<ConnectionBudget> Budget(DocumentStore store)
    {
        var budget = store.Find<ConnectionBudget>("asx_claim", BudgetKey);
        if (budget == null)
        {
            store.Create("asx_claim", new ConnectionBudget { Key = BudgetKey, Status = "Budget" });
            budget = store.Require<ConnectionBudget>("asx_claim", BudgetKey);
        }
        return budget;
    }

    /// <summary>
    /// Records the site writer in the shared connection row in the site claim transaction, so
    /// pacing, backoff and the runtime idle check see it; unresolved writers stay listed. Each
    /// site has one writer, which keeps its writes in order; sites do not limit one another.
    /// </summary>
    /// <param name="service">The Dataverse service participating in the site claim transaction.</param>
    /// <param name="claim">The site writer's current ownership and outstanding-request state.</param>
    public static void SaveWriter(IOrganizationService service, DispatcherDocument claim)
    {
        if (!claim.Key.StartsWith("site-writer:", StringComparison.Ordinal))
            throw new EvaluationBlockedException("Writer must have a site scope.");
        if (claim.RunId == null && claim.HttpOutstanding)
            throw new EvaluationBlockedException(
                "Cannot release a writer with an unresolved HTTP request."
            );
        var store = new DocumentStore(service);
        var budget = Budget(store);
        var other = budget.Value.Writers.Where(key => key != claim.Key).ToArray();
        budget.Value.Writers =
            claim.RunId == null ? other : other.Concat(new[] { claim.Key }).ToArray();
        store.Save(budget);
    }

    public static bool Busy(IOrganizationService service) =>
        new DocumentStore(service)
            .Find<ConnectionBudget>("asx_claim", BudgetKey)
            ?.Value.Writers.Length > 0;

    public static void RequireIdle(IOrganizationService service)
    {
        var store = new DocumentStore(service);
        var budget = Budget(store);
        if (budget.Value.Writers.Length != 0)
            throw new EvaluationBlockedException(
                "Finish or reconcile active writers before changing runtime configuration."
            );
        // Writes the shared row using its current version, including when its values are unchanged.
        store.Save(budget);
    }

    public static bool HasCapacity(IOrganizationService service, string scope, DateTime? now = null)
    {
        var budget = new DocumentStore(service)
            .Find<ConnectionBudget>("asx_claim", BudgetKey)
            ?.Value;
        // A site that already writes keeps going; a new writer waits only out a shared
        // Retry-After pause.
        return budget == null
            || budget.Writers.Contains(scope)
            || budget.PauseUntilUtc <= (now ?? DateTime.UtcNow);
    }

    /// <summary>
    /// Renews the site lease and grants one HTTP request when shared pacing and backoff allow it.
    /// </summary>
    /// <param name="service">The transactional Dataverse service used to update claims and the budget.</param>
    /// <param name="request">The operation key, run identity, and current claim token.</param>
    /// <param name="now">The UTC time used for lease renewal and request admission.</param>
    /// <returns>A Permit result or a Wait result containing the required delay in seconds.</returns>
    public static WorkerResult BeginHttp(
        IOrganizationService service,
        WorkerRequest request,
        DateTime now
    )
    {
        var store = new DocumentStore(service);
        var claim = store.Require<DispatcherDocument>("asx_claim", Operation(service, request.Key));
        Assert(claim.Value, request, now);
        if (claim.Value.HttpOutstanding)
            throw new EvaluationBlockedException(
                "The previous HTTP request must complete before another can start."
            );
        var stopped = StopWrite(service, store, claim, request.Key);
        if (stopped != null)
            return stopped;
        var budget = Budget(store);
        if (!budget.Value.Writers.Contains(claim.Value.Key))
            throw new EvaluationBlockedException(
                "Writer is not admitted to the shared connection."
            );
        var until =
            budget.Value.NextStartUtc > budget.Value.PauseUntilUtc
                ? budget.Value.NextStartUtc
                : budget.Value.PauseUntilUtc;
        claim.Value.LeaseUntilUtc = now.AddMinutes(5);
        store.Save(claim);
        if (until > now)
            return new WorkerResult
            {
                Status = "Wait",
                Key = request.Key,
                WaitSeconds = Math.Min(
                    60,
                    Math.Max(1, (int)Math.Ceiling((until - now).TotalSeconds))
                ),
            };
        // SaveWriter refreshed the same row; read its new row version before the admission CAS.
        budget = Budget(store);
        if (budget.Value.PaceSeconds > 1 && budget.Value.RecoverPaceUtc <= now)
        {
            budget.Value.PaceSeconds = Math.Max(1, budget.Value.PaceSeconds / 2);
            budget.Value.RecoverPaceUtc = now.AddMinutes(10);
        }
        budget.Value.NextStartUtc = now.AddSeconds(Math.Max(1, budget.Value.PaceSeconds));
        store.Save(budget);
        claim = store.Require<DispatcherDocument>("asx_claim", claim.Value.Key);
        claim.Value.HttpOutstanding = true;
        store.Save(claim);
        // A library setup records that its prepared write may now be sent. Until then a run
        // that stops leaves nothing unknown in SharePoint (LibraryProvisioning.NeverSent).
        if (request.Key.StartsWith("librarycreate:", StringComparison.Ordinal))
        {
            var setup = store.Require<LibrarySetup>("asx_operation", request.Key);
            if (Writing(setup.Value))
            {
                setup.Value.WritePermitted = true;
                store.Save(setup);
            }
        }
        return new WorkerResult { Status = "Permit", Key = request.Key };
    }

    /// <summary>
    /// Applies connection-wide backoff and releases the outstanding request flag for known responses.
    /// </summary>
    /// <param name="service">The transactional Dataverse service used to update claims and the budget.</param>
    /// <param name="request">The claim identity, HTTP status, and optional Retry-After response value.</param>
    /// <param name="now">The UTC time used to validate the claim and calculate backoff.</param>
    /// <returns>A quarantine result for an unknown outcome, or null for a known response.</returns>
    public static WorkerResult? Response(
        IOrganizationService service,
        WorkerRequest request,
        DateTime now
    )
    {
        var store = new DocumentStore(service);
        var claim = store.Require<DispatcherDocument>("asx_claim", Operation(service, request.Key));
        Assert(claim.Value, request, now);
        if (request.HttpStatus == 429 || request.HttpStatus >= 500)
        {
            var budget = Budget(store);
            // RetryAt never throws and caps an honored Retry-After at 15 minutes, so a long or
            // malformed hint pauses shared SharePoint work for at most one re-check interval.
            var pause = WorkerCoordinator.RetryAt(now, 1, request.RetryAfter);
            if (pause > budget.Value.PauseUntilUtc)
                budget.Value.PauseUntilUtc = pause;
            budget.Value.PaceSeconds = Math.Min(4, Math.Max(1, budget.Value.PaceSeconds) * 2);
            budget.Value.RecoverPaceUtc = pause.AddMinutes(10);
            store.Save(budget);
        }
        // A read (Observe) has no side effects, so a missing response is just a failed read: the
        // worker waits and releases the slot. Only an unknown write outcome keeps the request
        // outstanding until a later claim reads back what SharePoint did.
        if ((request.HttpStatus == 0 || request.HttpStatus == 408) && request.Command != "Observe")
            return new WorkerResult
            {
                Status = "Quarantined",
                Key = request.Key,
                Token = request.Token,
                Notices = new[]
                {
                    "Request completion is unknown. The site's writer stays reserved until the original run is reconciled.",
                },
            };
        claim.Value.HttpOutstanding = false;
        store.Save(claim);
        return null;
    }

    /// <summary>
    /// The last moment before a SharePoint write: a write for a destination removed or
    /// suspended since the step was prepared is not permitted. It was never sent, so it is put
    /// back as not submitted and Pending, and the run's claim is released; the job's next claim
    /// then runs the usual stop (cancelled for a removed destination, stopped at claim for a
    /// suspended one). Reads are still permitted: they change nothing.
    /// </summary>
    private static WorkerResult? StopWrite(
        IOrganizationService service,
        DocumentStore store,
        StoredRow<DispatcherDocument> claim,
        string key
    )
    {
        string? reason;
        if (key.StartsWith("policywork:", StringComparison.Ordinal))
        {
            var access = store.Require<SecurityOperation>("asx_operation", key);
            if (!Writing(access.Value) || (reason = Stops(service, access.Value.LibraryId)) == null)
                return null;
            access.Value.Mutation = null;
            access.Value.MutationKind = "";
            Unsend(access.Value);
            store.Save(access);
        }
        else if (key.StartsWith("librarycreate:", StringComparison.Ordinal))
        {
            var setup = store.Require<LibrarySetup>("asx_operation", key);
            if (!Writing(setup.Value) || (reason = StopsSite(service, setup.Value.SiteId)) == null)
                return null;
            setup.Value.Intent = null;
            setup.Value.Mutation = "";
            Unsend(setup.Value);
            store.Save(setup);
        }
        else if (key.StartsWith("catalogprobe:", StringComparison.Ordinal))
            return null;
        else
        {
            var job = store.Require<OperationDocument>("asx_operation", key);
            if (
                !Writing(job.Value)
                || job.Value.Folders.Length == 0
                || (reason = Stops(service, job.Value.Folder.LibraryId)) == null
            )
                return null;
            Unsend(job.Value);
            store.Save(job);
        }
        claim.Value.HttpOutstanding = false;
        claim.Value.RunId = null;
        claim.Value.OperationKey = null;
        claim.Value.Token = Guid.Empty;
        claim.Value.RecoveryPermitted = false;
        claim.Value.Status = "Idle";
        store.Save(claim);
        store.Create(
            "asx_attempt",
            new AttemptDocument
            {
                Key = "attempt:" + Guid.NewGuid().ToString("N"),
                Status = "Recorded",
                OperationKey = key,
                RunId = "worker",
                Event = "WriteNotPermitted",
                AtUtc = DateTime.UtcNow,
            }
        );
        return new WorkerResult
        {
            Status = "Stopped",
            Key = key,
            Notices = new[]
            {
                "The destination was "
                    + reason
                    + " before this write was sent. Nothing was written to SharePoint; the work stops at its next step.",
            },
        };
    }

    // A prepared write is about to be sent: submitted, with no response yet.
    private static bool Writing(OperationDocument op) =>
        op.ExternalSubmitted && !op.ExternalResponseKnown;

    internal static void Unsend(OperationDocument op)
    {
        op.ExternalSubmitted = false;
        op.ExternalResponseKnown = false;
        op.AbsenceVerified = false;
        op.ProbeId = Guid.Empty;
        op.NextAttemptUtc = null;
        op.Status = "Pending";
    }

    // "removed" or "suspended" when the library (or its site) no longer takes writes.
    internal static string? Stops(IOrganizationService service, Guid libraryId)
    {
        var library = Find(service, "asx_library", libraryId, "asx_siteid");
        var state = library == null ? null : Read(service, "asx_library", libraryId);
        if (library == null || state == null || state.Value.removed)
            return "removed";
        var site = library.GetAttributeValue<EntityReference>("asx_siteid");
        return !state.Value.approved ? "suspended"
            : site == null ? "removed"
            : StopsSite(service, site.Id);
    }

    internal static string? StopsSite(IOrganizationService service, Guid siteId)
    {
        var state = Read(service, "asx_site", siteId);
        return state == null || state.Value.removed ? "removed"
            : !state.Value.approved ? "suspended"
            : null;
    }

    private static (bool approved, bool removed)? Read(
        IOrganizationService service,
        string table,
        Guid id
    )
    {
        var query = new QueryExpression(table)
        {
            ColumnSet = new ColumnSet("asx_approved", "statecode"),
            TopCount = 1,
        };
        query.Criteria.AddCondition(table + "id", ConditionOperator.Equal, id);
        var row = service.RetrieveMultiple(query).Entities.FirstOrDefault();
        if (row == null)
            return null;
        return (
            row.GetAttributeValue<bool>("asx_approved"),
            row.GetAttributeValue<OptionSetValue>("statecode")?.Value == 1
        );
    }

    private static void Assert(DispatcherDocument claim, WorkerRequest request, DateTime now)
    {
        if (
            request.Token == Guid.Empty
            || claim.Token != request.Token
            || claim.OperationKey != request.Key
            || claim.RunId != request.RunId
            || claim.LeaseUntilUtc <= now
            || claim.RecoveryPermitted
        )
            throw new EvaluationBlockedException(
                "HTTP admission requires the current site writer."
            );
    }
}
