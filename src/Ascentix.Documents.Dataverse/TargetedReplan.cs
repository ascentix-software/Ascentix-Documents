using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public sealed class TargetedReplan
{
    private readonly IOrganizationService service;
    private readonly DocumentStore store;
    private readonly string[] allowed;

    public TargetedReplan(IOrganizationService service, string[] allowed)
    {
        this.service = service;
        store = new DocumentStore(service);
        this.allowed = allowed;
    }

    public static string SourceKey(string table, Guid id) =>
        DocumentStore.Hash(table + ":" + id.ToString("N"));

    /// <summary>
    /// Appends a related-record invalidation when a current record-selection receipt references it.
    /// </summary>
    /// <param name="service">The Dataverse service used to find referencing record selections.</param>
    /// <param name="store">The durable store used to find or create the outbox invalidation.</param>
    /// <param name="table">The logical name of the changed related-record table.</param>
    /// <param name="id">The changed related-record ID.</param>
    /// <param name="correlation">The event correlation ID used in the idempotency key.</param>
    public static void Append(
        IOrganizationService service,
        DocumentStore store,
        string table,
        Guid id,
        Guid correlation
    )
    {
        var matches = new QueryExpression("asx_outbox")
        {
            ColumnSet = new ColumnSet(false),
            TopCount = 1,
        };
        matches.Criteria.FilterOperator = LogicalOperator.Or;
        for (int i = 0; i < 6; i++)
            matches.Criteria.AddCondition(
                "asx_source" + i,
                ConditionOperator.Equal,
                SourceKey(table, id)
            );
        if (service.RetrieveMultiple(matches).Entities.Count == 0)
            return;
        string key =
            "source-event:" + correlation.ToString("N") + ":" + table + ":" + id.ToString("N");
        if (store.Find<OutboxDocument>("asx_outbox", key) == null)
            store.Create(
                "asx_outbox",
                new OutboxDocument
                {
                    Key = key,
                    RelatedTable = table,
                    RelatedRecordId = id,
                }
            );
    }

    public WorkerResult Consume(StoredRow<OutboxDocument> job)
    {
        if (job.Value.Status != "Pending")
            return new WorkerResult { Key = job.Value.Key, Status = job.Value.Status };
        if (!allowed.Contains(job.Value.RelatedTable) || job.Value.RelatedRecordId == Guid.Empty)
            throw new EvaluationBlockedException("Related event is outside runtime source scope.");
        var query = new QueryExpression("asx_outbox")
        {
            ColumnSet = new ColumnSet("asx_payload"),
            PageInfo = new PagingInfo
            {
                Count = 25,
                PageNumber = job.Value.SourcePage,
                PagingCookie = job.Value.SourceCookie,
            },
        };
        var source = new FilterExpression(LogicalOperator.Or);
        for (int i = 0; i < 6; i++)
            source.AddCondition(
                "asx_source" + i,
                ConditionOperator.Equal,
                SourceKey(job.Value.RelatedTable, job.Value.RelatedRecordId)
            );
        query.Criteria.AddFilter(source);
        query.AddOrder("asx_outboxid", OrderType.Ascending);
        var page = service.RetrieveMultiple(query);
        if (
            page.Entities.Count > 25
            || page.MoreRecords
                && (
                    page.Entities.Count == 0
                    || string.IsNullOrEmpty(page.PagingCookie)
                    || page.PagingCookie == job.Value.SourceCookie
                )
        )
            throw new EvaluationBlockedException("Related-record continuation is incomplete.");
        var coordinator = new WorkerCoordinator(service, allowedTables: allowed);
        foreach (var row in page.Entities)
        {
            var selection = JsonWire.Read<RecordPlanDocument>(
                TemplateStore.Text(row, "asx_payload")
            );
            if (
                selection.RecordId == Guid.Empty
                || selection.TemplateId == Guid.Empty
                || !selection.Sources.Any(s =>
                    s.Table == job.Value.RelatedTable && s.Id == job.Value.RelatedRecordId
                )
            )
                throw new EvaluationBlockedException(
                    "Related selection index differs from its receipt."
                );
            // Root updates have their own filtered event; this fan-out is for direct related sources only.
            if (
                store.Find<OutboxDocument>(
                    "asx_outbox",
                    WorkerCoordinator.RetirementKey(selection.Table, selection.RecordId)
                ) != null
            )
                continue;
            coordinator.Execute(
                new WorkerRequest
                {
                    Command = "Queue",
                    TemplateId = selection.TemplateId,
                    RecordId = selection.RecordId,
                    RequestId = DocumentStore.StableId(job.Value.Key + ":" + selection.Key),
                },
                true
            );
        }
        job.Value.SourcePage++;
        job.Value.SourceCookie = page.PagingCookie;
        job.Value.Status = page.MoreRecords ? "Pending" : "Planned";
        store.Save(job);
        return new WorkerResult { Key = job.Value.Key, Status = job.Value.Status };
    }
}
