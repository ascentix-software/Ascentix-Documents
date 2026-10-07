using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

/// <summary>A template re-run as Monitor shows it (spec 6.5, 3.3 Template re-runs).</summary>
[DataContract]
public sealed class TemplateRunState
{
    [DataMember]
    public Guid TemplateId { get; set; }

    [DataMember]
    public string TemplateName { get; set; } = "";

    [DataMember]
    public string Table { get; set; } = "";

    [DataMember]
    public string TableLabel { get; set; } = "";

    [DataMember]
    public int Version { get; set; }

    /// <summary>Running, Waiting, Retrying, Paused, Blocked, Done or Cancelled.</summary>
    [DataMember]
    public string State { get; set; } = "";

    [DataMember]
    public int Planned { get; set; }

    [DataMember]
    public int Queued { get; set; }

    /// <summary>
    /// The records the run covers. Until the run is done it comes from the platform's daily
    /// row-count snapshot, raised to what the run has already met; done, it is the exact count.
    /// </summary>
    [DataMember]
    public int Total { get; set; }

    /// <summary>True until the run is done: the page shows the total as "about N".</summary>
    [DataMember]
    public bool TotalEstimated { get; set; }

    [DataMember]
    public DateTime? StartedUtc { get; set; }

    [DataMember]
    public string? StartedBy { get; set; }

    [DataMember]
    public DateTime? NextAttemptUtc { get; set; }

    [DataMember]
    public DateTime? EstimatedFinishUtc { get; set; }

    [DataMember]
    public DateTime? EndedUtc { get; set; }

    [DataMember]
    public string? Problem { get; set; }
}

/// <summary>
/// Re-runs a template's published version for every record of its table in the background. One
/// asx_outbox row per template is reused for each run (decision D7). The dispatcher plans it like
/// any outbox row; each Plan queues at most one page, and only once the previous page has left
/// Pending, so new record events never wait behind more than one page of re-run work.
/// </summary>
public sealed class TemplateRun
{
    public const string Prefix = "templaterun:";

    /// <summary>asx_table of a run row, so lists find runs without reading payloads; the real table is in the payload.</summary>
    public const string IndexMarker = "templaterun";

    /// <summary>
    /// Records queued per page: one fewer than the dispatcher's outbox page, so a page and its run
    /// row fit in one dispatch. ListOutbox lists the run after its page, so with no other work the
    /// run finds its page planned and queues the next in the same dispatch: a page per dispatch.
    /// A larger page (spec 6.9 had 25, TargetedReplan's size) cannot be planned in one dispatch,
    /// which halves the pace. It also bounds the re-run work that waits ahead of new events and
    /// keeps each Plan call short.
    /// </summary>
    public const int PageSize = DocumentStore.DispatchPage - 1;

    /// <summary>asx_priority of re-run folder jobs: after access and catalog work (0) and normal folder jobs (1).</summary>
    public const int Priority = 2;

    public const string TableGone = "The table no longer exists.";

    private static readonly string[] Active = { "Pending", "Paused", "Blocked" };
    private readonly IOrganizationService service;
    private readonly IOrganizationService reads;
    private readonly DocumentStore store;
    private readonly string[] allowed;
    private readonly Func<DateTime> clock;

    /// <param name="service">The caller's service: every write, and reads of Documents' own tables.</param>
    /// <param name="allowed">The tables enabled in Documents.</param>
    /// <param name="clock">The UTC clock.</param>
    /// <param name="worker">
    /// The worker's service, for the reads an operator may lack the privilege for: the table's
    /// row count and the starter's name. Null when the caller is the worker (the dispatcher).
    /// </param>
    public TemplateRun(
        IOrganizationService service,
        string[] allowed,
        Func<DateTime>? clock = null,
        IOrganizationService? worker = null
    )
    {
        this.service = service;
        reads = worker ?? service;
        store = new DocumentStore(service);
        this.allowed = allowed;
        this.clock = clock ?? (() => DateTime.UtcNow);
    }

    public static string Key(Guid templateId) =>
        Prefix + DocumentStore.Hash(templateId.ToString("N"));

    /// <summary>Starts a re-run, or returns the active one (one per template).</summary>
    /// <param name="request">TemplateId and a RequestId that makes the call idempotent.</param>
    /// <param name="caller">Who started it.</param>
    public WorkerResult Start(WorkerRequest request, Guid caller)
    {
        if (request.TemplateId == Guid.Empty || request.RequestId == Guid.Empty)
            throw new EvaluationBlockedException("Choose a template and send a request ID.");
        var template =
            TemplateLifecycle.Find(service, request.TemplateId)
            ?? throw new EvaluationBlockedException("The template was deleted.");
        var published =
            template.GetAttributeValue<EntityReference>("asx_publishedrevisionid")
            ?? throw new EvaluationBlockedException("Publish the template first.");
        string table = TemplateStore.Text(template, "asx_table");
        if (!allowed.Contains(table))
            throw new EvaluationBlockedException(WorkerCoordinator.TableNotEnabled(table));
        string key = Key(request.TemplateId);
        var existing = store.Find<OutboxDocument>("asx_outbox", key);
        if (
            existing != null
            && (
                existing.Value.RunRequestId == request.RequestId
                || Active.Contains(existing.Value.Status)
            )
        )
        {
            var result = Result(existing.Value);
            if (existing.Value.RunRequestId != request.RequestId)
                result.Notices = new[] { "A re-run of this template is already in progress." };
            return result;
        }
        if (TableInfo.Find(service, table) == null)
            throw new EvaluationBlockedException(TableGone);
        var run = existing?.Value ?? new OutboxDocument { Key = key };
        run.Status = "Pending";
        run.TemplateId = request.TemplateId;
        run.RevisionId = published.Id;
        run.Table = table;
        run.Priority = Priority;
        run.SourcePage = 1;
        run.SourceCookie = null;
        run.PageKeys = Array.Empty<string>();
        run.Planned = 0;
        run.PlannedAtResume = 0;
        run.Total = TableInfo.Records(reads, table);
        run.StartedBy = caller;
        run.StartedByName = Name(reads, "systemuser", caller, "fullname");
        run.StartedUtc = clock();
        run.ResumedUtc = run.StartedUtc;
        run.EndedUtc = null;
        run.RunRequestId = request.RequestId;
        run.Notices = Array.Empty<string>();
        run.NextAttemptUtc = null;
        run.Attempts = 0;
        if (existing == null)
            store.Create("asx_outbox", run);
        else
            store.Save(existing);
        return Result(store.Require<OutboxDocument>("asx_outbox", key).Value);
    }

    public WorkerResult Pause(string key) =>
        Change(
            key,
            "Pending",
            run => run.Status = "Paused",
            "Only a running re-run can be paused."
        );

    public WorkerResult Resume(string key) =>
        Change(
            key,
            "Paused",
            run =>
            {
                run.Status = "Pending";
                run.ResumedUtc = clock();
                run.PlannedAtResume = Progress(run).Planned;
            },
            "Only a paused re-run can be resumed."
        );

    /// <summary>Stops the run: records not yet planned are skipped; the current page's waiting rows are cancelled; nothing in SharePoint is undone.</summary>
    public WorkerResult Cancel(string key)
    {
        var row = Require(key);
        if (!Active.Contains(row.Value.Status))
            return Result(row.Value);
        CancelPage(row.Value, "Re-run cancelled.");
        return End(row, "Cancelled", "Re-run cancelled.");
    }

    /// <summary>
    /// About how many records a re-run of the template covers: the platform's daily row-count
    /// snapshot of its table, read as the worker. Writes nothing and runs no aggregate.
    /// </summary>
    public WorkerResult Count(Guid templateId)
    {
        var template =
            TemplateLifecycle.Find(service, templateId)
            ?? throw new EvaluationBlockedException("The template was deleted.");
        string table = TemplateStore.Text(template, "asx_table");
        if (TableInfo.Find(service, table) == null)
            throw new EvaluationBlockedException(TableGone);
        return new WorkerResult
        {
            Status = "Counted",
            Run = new TemplateRunState
            {
                TemplateId = templateId,
                Table = table,
                Total = TableInfo.Records(reads, table),
                TotalEstimated = true,
            },
        };
    }

    /// <summary>One dispatch of a run row (spec 6.5 steps 1-6).</summary>
    public WorkerResult Consume(StoredRow<OutboxDocument> job)
    {
        var run = job.Value;
        if (run.Status != "Pending")
            return Result(run);
        // Listed again, so a wait after a temporary failure is over; any save below ends it, and
        // a Consume that fails again rolls back and keeps it (as WorkerCoordinator.Plan does).
        bool waited = run.NextAttemptUtc != null || run.Attempts != 0;
        run.NextAttemptUtc = null;
        run.Attempts = 0;
        if (waited)
            run.Notices = Array.Empty<string>();
        if (!TemplateLifecycle.Active(TemplateLifecycle.Find(service, run.TemplateId), clock()))
            return End(job, "Cancelled", "The template is off, so the re-run stopped.");
        if (!allowed.Contains(run.Table))
            return End(job, "Cancelled", WorkerCoordinator.TableNotEnabled(run.Table));
        // Looked up without a fault, which would end the transaction: the page's rows could not
        // be planned either, so they stop with the run.
        var table = TableInfo.Find(service, run.Table);
        if (table == null)
        {
            CancelPage(run, "The table no longer exists, so the re-run stopped.");
            return End(job, "Cancelled", "The table no longer exists, so the re-run stopped.");
        }
        // One page in flight: the next page waits until every row of this one has left Pending.
        if (
            run.PageKeys.Any(k =>
                store.Find<OutboxDocument>("asx_outbox", k)?.Value.Status == "Pending"
            )
        )
        {
            if (waited)
                store.Save(job);
            return Result(run);
        }
        run.Planned += run.PageKeys.Length;
        run.PageKeys = Array.Empty<string>();
        var page = ReadPage(run, table.PrimaryIdAttribute);
        var coordinator = new WorkerCoordinator(service, clock, allowed);
        var queued = new List<string>();
        foreach (var id in page.Ids)
        {
            if (
                store.Find<OutboxDocument>(
                    "asx_outbox",
                    WorkerCoordinator.RetirementKey(run.Table, id)
                ) != null
            )
            {
                run.Planned++;
                continue;
            }
            // The request ID follows the run and the record, so a page read twice queues nothing twice.
            var result = coordinator.Queue(
                new WorkerRequest
                {
                    TemplateId = run.TemplateId,
                    RecordId = id,
                    RequestId = DocumentStore.StableId(
                        run.Key + ":" + run.RunRequestId.ToString("N") + ":" + id.ToString("N")
                    ),
                },
                Priority
            );
            if (result.Status == "Pending")
                queued.Add(result.Key);
            else
                run.Planned++;
        }
        run.PageKeys = queued.ToArray();
        run.SourcePage++;
        run.SourceCookie = page.Cookie;
        if (!page.More && run.PageKeys.Length == 0)
        {
            // Every record has been met: the count is exact now.
            run.Total = run.Planned;
            return End(job, "Planned", null);
        }
        // The snapshot can be a day old; the run never shows fewer records than it has met.
        run.Total = Math.Max(run.Total, run.Planned + run.PageKeys.Length);
        store.Save(job);
        return Result(run);
    }

    /// <summary>The run's state for Monitor, with an estimated finish.</summary>
    public TemplateRunState Describe(OutboxDocument run)
    {
        DateTime now = clock();
        var (planned, waiting) = Progress(run);
        string state = run.Status switch
        {
            "Paused" => "Paused",
            "Blocked" => "Blocked",
            "Planned" => "Done",
            "Cancelled" => "Cancelled",
            _ => run.NextAttemptUtc > now ? "Retrying"
            : waiting > 0 ? "Waiting"
            : "Running",
        };
        var described = new TemplateRunState
        {
            TemplateId = run.TemplateId,
            TemplateName =
                Name(service, "asx_template", run.TemplateId, "asx_name") ?? "Deleted template",
            Table = run.Table,
            TableLabel = TableInfo.Label(service, run.Table),
            Version = Version(run.RevisionId),
            State = state,
            Planned = planned,
            Queued = waiting,
            Total = run.Total,
            TotalEstimated = state != "Done",
            StartedUtc = run.StartedUtc,
            StartedBy = run.StartedByName,
            NextAttemptUtc = run.NextAttemptUtc > now ? run.NextAttemptUtc : null,
            EndedUtc = run.EndedUtc,
            Problem =
                state == "Blocked" || state == "Cancelled" ? run.Notices.FirstOrDefault() : null,
        };
        // Seconds per record since the last start or resume, once a full page has finished (D15).
        int done = planned - run.PlannedAtResume;
        bool moving = state == "Running" || state == "Waiting" || state == "Retrying";
        if (moving && run.ResumedUtc != null && done >= PageSize)
        {
            double perRecord = (now - run.ResumedUtc.Value).TotalSeconds / done;
            described.EstimatedFinishUtc = now.AddSeconds(
                perRecord * Math.Max(0, run.Total - planned)
            );
        }
        return described;
    }

    /// <summary>Records planned so far, counting the finished rows of the page in flight, and its rows still Pending.</summary>
    private (int Planned, int Waiting) Progress(OutboxDocument run)
    {
        int waiting = run.PageKeys.Count(k =>
            store.Find<OutboxDocument>("asx_outbox", k)?.Value.Status == "Pending"
        );
        return (run.Planned + run.PageKeys.Length - waiting, waiting);
    }

    /// <summary>Cancels the rows of the page in flight that are still waiting to be planned.</summary>
    private void CancelPage(OutboxDocument run, string notice)
    {
        foreach (var pageKey in run.PageKeys)
        {
            var queued = store.Find<OutboxDocument>("asx_outbox", pageKey);
            if (queued?.Value.Status != "Pending")
                continue;
            queued.Value.Status = "Cancelled";
            queued.Value.NextAttemptUtc = null;
            queued.Value.Notices = new[] { notice };
            store.Save(queued);
        }
    }

    private WorkerResult Change(
        string key,
        string from,
        Action<OutboxDocument> apply,
        string refusal
    )
    {
        var row = Require(key);
        if (row.Value.Status != from)
        {
            // Pausing a paused run, or resuming a running one, has nothing to do.
            if (
                row.Value.Status == "Paused" && from == "Pending"
                || row.Value.Status == "Pending" && from == "Paused"
            )
                return Result(row.Value);
            throw new EvaluationBlockedException(refusal);
        }
        apply(row.Value);
        store.Save(row);
        return Result(store.Require<OutboxDocument>("asx_outbox", key).Value);
    }

    private StoredRow<OutboxDocument> Require(string key) =>
        key.StartsWith(Prefix, StringComparison.Ordinal)
            ? store.Require<OutboxDocument>("asx_outbox", key)
            : throw new EvaluationBlockedException("Choose a template re-run.");

    private WorkerResult End(StoredRow<OutboxDocument> job, string status, string? notice)
    {
        job.Value.Status = status;
        job.Value.EndedUtc = clock();
        job.Value.NextAttemptUtc = null;
        job.Value.Attempts = 0;
        job.Value.Notices = notice == null ? Array.Empty<string>() : new[] { notice };
        store.Save(job);
        return Result(job.Value);
    }

    private WorkerResult Result(OutboxDocument run) =>
        new WorkerResult
        {
            Status = run.Status,
            Key = run.Key,
            Run = Describe(run),
            Notices = run.Notices,
        };

    private (Guid[] Ids, bool More, string? Cookie) ReadPage(OutboxDocument run, string primaryId)
    {
        var query = new QueryExpression(run.Table)
        {
            ColumnSet = new ColumnSet(false),
            PageInfo = new PagingInfo
            {
                Count = PageSize,
                PageNumber = run.SourcePage,
                PagingCookie = run.SourceCookie,
            },
        };
        query.AddOrder(primaryId, OrderType.Ascending);
        var page = service.RetrieveMultiple(query);
        // The same continuation check as TargetedReplan.Consume: a page that claims more but gives
        // no way on fails, and the dispatcher's FailOutbox retries it at the same cursor.
        if (
            page.Entities.Count > PageSize
            || page.MoreRecords
                && (
                    page.Entities.Count == 0
                    || string.IsNullOrEmpty(page.PagingCookie)
                    || page.PagingCookie == run.SourceCookie
                )
        )
            throw new EvaluationBlockedException(
                "The record page could not be read completely; the re-run tries this page again."
            );
        return (page.Entities.Select(e => e.Id).ToArray(), page.MoreRecords, page.PagingCookie);
    }

    private static string? Name(IOrganizationService service, string table, Guid id, string column)
    {
        if (id == Guid.Empty)
            return null;
        var query = new QueryExpression(table) { ColumnSet = new ColumnSet(column), TopCount = 1 };
        query.Criteria.AddCondition(table + "id", ConditionOperator.Equal, id);
        return service
            .RetrieveMultiple(query)
            .Entities.FirstOrDefault()
            ?.GetAttributeValue<string>(column);
    }

    private int Version(Guid revision)
    {
        if (revision == Guid.Empty)
            return 0;
        var query = new QueryExpression("asx_revision")
        {
            ColumnSet = new ColumnSet("asx_version"),
            TopCount = 1,
        };
        query.Criteria.AddCondition("asx_revisionid", ConditionOperator.Equal, revision);
        return service
                .RetrieveMultiple(query)
                .Entities.FirstOrDefault()
                ?.GetAttributeValue<int>("asx_version")
            ?? 0;
    }
}
