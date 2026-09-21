using System;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Domain;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class BatchRecord
{
    [DataMember]
    public Guid Id { get; set; }

    [DataMember]
    public SourceVersion[] Sources { get; set; } = Array.Empty<SourceVersion>();

    [DataMember]
    public string[] Impact { get; set; } = Array.Empty<string>();
}

[DataContract]
public sealed class BatchDocument : StoredDocument
{
    [DataMember]
    public Guid TemplateId { get; set; }

    [DataMember]
    public Guid RevisionId { get; set; }

    [DataMember]
    public BatchRecord[] Records { get; set; } = Array.Empty<BatchRecord>();

    [DataMember]
    public string[] Requests { get; set; } = Array.Empty<string>();
}

public sealed class BatchReplan
{
    private readonly IOrganizationService service;
    private readonly DocumentStore store;
    private readonly string[] allowed;
    private readonly Func<DateTime> clock;

    public BatchReplan(IOrganizationService service, string[] allowed, Func<DateTime>? clock = null)
    {
        this.service = service;
        store = new DocumentStore(service);
        this.allowed = allowed;
        this.clock = clock ?? (() => DateTime.UtcNow);
    }

    public WorkerResult Execute(WorkerRequest request, bool transaction)
    {
        if (!transaction)
            throw new EvaluationBlockedException("Batch review requires a transaction.");
        if (request.Command == "PreviewBatch")
        {
            if (
                request.RequestId == Guid.Empty
                || request.TemplateId == Guid.Empty
                || request.RecordIds == null
                || request.RecordIds.Length < 1
                || request.RecordIds.Length > 5
                || request.RecordIds.Any(id => id == Guid.Empty)
                || request.RecordIds.Distinct().Count() != request.RecordIds.Length
            )
                throw new EvaluationBlockedException(
                    "Select one to five unique records for a bounded batch review."
                );
            string key = "batch:" + request.RequestId.ToString("N");
            var old = store.Find<BatchDocument>("asx_outbox", key);
            if (old != null)
            {
                if (
                    old.Value.TemplateId != request.TemplateId
                    || !old.Value.Records.Select(r => r.Id).SequenceEqual(request.RecordIds)
                )
                    throw new EvaluationBlockedException("Batch request identity was reused.");
                return Result(old);
            }
            var header = service.Retrieve(
                "asx_template",
                request.TemplateId,
                new ColumnSet("asx_publishedrevisionid")
            );
            Guid revision =
                header.GetAttributeValue<EntityReference>("asx_publishedrevisionid")?.Id
                ?? throw new EvaluationBlockedException("Publish before batch review.");
            var template = new TemplateStore(service).Read(revision);
            if (template.Sources.Any(s => !allowed.Contains(s.Table)))
                throw new EvaluationBlockedException("Batch source exceeds runtime scope.");
            int total = 0;
            var records = request
                .RecordIds.Select(id =>
                {
                    if (
                        store.Find<OutboxDocument>(
                            "asx_outbox",
                            WorkerCoordinator.RetirementKey(template.Table, id)
                        ) != null
                    )
                        throw new EvaluationBlockedException(
                            "Deleted records require decommission review."
                        );
                    var snapshot = new SnapshotReader(service).Read(template, id);
                    var intents = FolderPlanner.Plan(template, id, snapshot.Values);
                    total += intents.Count;
                    if (total > 1000)
                        throw new EvaluationBlockedException(
                            "Batch preview exceeds 1000 folder intents; select fewer records."
                        );
                    return new BatchRecord
                    {
                        Id = id,
                        Sources = snapshot
                            .Records.Select(
                                (record, i) =>
                                    new SourceVersion
                                    {
                                        Table = record.LogicalName,
                                        Id = record.Id,
                                        Version = snapshot.Versions[i],
                                    }
                            )
                            .ToArray(),
                        Impact = intents
                            .Select(intent =>
                            {
                                return "Resolve or create expected path: "
                                    + intent.Section
                                    + "/"
                                    + intent.RelativePath;
                            })
                            .Concat(
                                new[]
                                {
                                    "Folders excluded by conditions remain intact. Physical collisions and shared access are verified during execution.",
                                }
                            )
                            .ToArray(),
                    };
                })
                .ToArray();
            store.Create(
                "asx_outbox",
                new BatchDocument
                {
                    Key = key,
                    Status = "BatchReview",
                    TemplateId = template.Id,
                    RevisionId = revision,
                    Records = records,
                }
            );
            return Result(store.Require<BatchDocument>("asx_outbox", key));
        }
        if (request.Command != "QueueBatch")
            throw new EvaluationBlockedException("Unsupported batch command.");
        var batch = store.Require<BatchDocument>("asx_outbox", request.Key);
        if (batch.Value.Status == "BatchQueued")
            return Result(batch);
        if (batch.Value.Status != "BatchReview" || batch.Row.RowVersion != request.RowVersion)
            throw new EvaluationBlockedException("Refresh the exact batch review before queuing.");
        var current = TemplateLifecycle.Find(service, batch.Value.TemplateId);
        if (!TemplateLifecycle.Active(current, clock()))
            throw new EvaluationBlockedException(
                "Template is deleted, deactivated or outside its schedule."
            );
        if (
            current!.GetAttributeValue<EntityReference>("asx_publishedrevisionid")?.Id
            != batch.Value.RevisionId
        )
            throw new EvaluationBlockedException(
                "Publication changed; request a new batch review."
            );
        foreach (var record in batch.Value.Records)
        foreach (var source in record.Sources)
        {
            if (
                !allowed.Contains(source.Table)
                || string.IsNullOrEmpty(source.Version)
                || service.Retrieve(source.Table, source.Id, new ColumnSet(false)).RowVersion
                    != source.Version
            )
                throw new EvaluationBlockedException("Source changed; request a new batch review.");
        }
        var coordinator = new WorkerCoordinator(service, clock, allowed);
        batch.Value.Requests = batch
            .Value.Records.Select(record =>
                coordinator
                    .Execute(
                        new WorkerRequest
                        {
                            Command = "Replan",
                            TemplateId = batch.Value.TemplateId,
                            RecordId = record.Id,
                            RequestId = DocumentStore.StableId(
                                batch.Value.Key + ":" + record.Id.ToString("N")
                            ),
                        },
                        true
                    )
                    .Key
            )
            .ToArray();
        foreach (var record in batch.Value.Records)
        {
            var job = store.Require<OutboxDocument>(
                "asx_outbox",
                "request:"
                    + DocumentStore
                        .StableId(batch.Value.Key + ":" + record.Id.ToString("N"))
                        .ToString("N")
            );
            job.Value.PinnedRevision = true;
            job.Value.ReviewedSources = record.Sources;
            store.Save(job);
        }
        batch.Value.Status = "BatchQueued";
        store.Save(batch);
        return Result(store.Require<BatchDocument>("asx_outbox", request.Key));
    }

    private static WorkerResult Result(StoredRow<BatchDocument> row) =>
        new WorkerResult
        {
            Key = row.Value.Key,
            Status = row.Value.Status,
            RowVersion = row.Row.RowVersion,
            Batch = row.Value,
            Keys = row.Value.Requests,
        };
}
