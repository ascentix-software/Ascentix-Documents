using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class BatchReplanTests
{
    [Fact]
    public void ReviewIsNonDispatchableAndQueueIsExplicitIdempotentAndVersioned()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var api = new BatchReplan(f.Service, new[] { "account" });
        var review = api.Execute(
            new WorkerRequest
            {
                Command = "PreviewBatch",
                TemplateId = f.TemplateId,
                RecordIds = new[] { f.RecordId },
                RequestId = Guid.NewGuid(),
            },
            true
        );
        Assert.Equal("BatchReview", review.Status);
        Assert.NotEmpty(review.Batch!.Records.Single().Impact);
        Assert.Empty(f.Store.Pending("asx_outbox"));
        Assert.Throws<EvaluationBlockedException>(() =>
            api.Execute(
                new WorkerRequest
                {
                    Command = "QueueBatch",
                    Key = review.Key,
                    RowVersion = "stale",
                },
                true
            )
        );
        var queued = f.Service.Transaction(() =>
            api.Execute(
                new WorkerRequest
                {
                    Command = "QueueBatch",
                    Key = review.Key,
                    RowVersion = review.RowVersion,
                },
                true
            )
        );
        Assert.Single(queued.Keys);
        Assert.Equal(
            queued.Keys,
            api.Execute(
                new WorkerRequest
                {
                    Command = "QueueBatch",
                    Key = review.Key,
                    RowVersion = review.RowVersion,
                },
                true
            ).Keys
        );
        var job = f.Store.Require<OutboxDocument>("asx_outbox", queued.Keys.Single());
        Assert.True(job.Value.Replan);
        Assert.True(job.Value.PinnedRevision);
        Assert.Equal(
            "Planned",
            f.Coordinator.Execute(
                new WorkerRequest { Command = "Plan", Key = job.Value.Key },
                true
            ).Status
        );
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ChangedSourceRequiresNewReviewBeforeQueueOrPlanning(bool afterQueue)
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var api = new BatchReplan(f.Service, new[] { "account" });
        var review = api.Execute(
            new WorkerRequest
            {
                Command = "PreviewBatch",
                TemplateId = f.TemplateId,
                RecordIds = new[] { f.RecordId },
                RequestId = Guid.NewGuid(),
            },
            true
        );
        WorkerResult? queued = afterQueue
            ? api.Execute(
                new WorkerRequest
                {
                    Command = "QueueBatch",
                    Key = review.Key,
                    RowVersion = review.RowVersion,
                },
                true
            )
            : null;
        f.Service.Seed(f.Service.Rows[f.RecordId]);
        if (afterQueue)
            Assert.Equal(
                "NeedsReview",
                f.Coordinator.Execute(
                    new WorkerRequest { Command = "Plan", Key = queued!.Keys.Single() },
                    true
                ).Status
            );
        else
            Assert.Throws<EvaluationBlockedException>(() =>
                api.Execute(
                    new WorkerRequest
                    {
                        Command = "QueueBatch",
                        Key = review.Key,
                        RowVersion = review.RowVersion,
                    },
                    true
                )
            );
    }

    [Fact]
    public void BatchBoundAndScopeAreEnforcedBeforeWorkIsCreated()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var api = new BatchReplan(f.Service, Array.Empty<string>());
        Assert.Throws<EvaluationBlockedException>(() =>
            api.Execute(
                new WorkerRequest
                {
                    Command = "PreviewBatch",
                    TemplateId = f.TemplateId,
                    RecordIds = new[] { f.RecordId },
                    RequestId = Guid.NewGuid(),
                },
                true
            )
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            api.Execute(
                new WorkerRequest
                {
                    Command = "PreviewBatch",
                    TemplateId = f.TemplateId,
                    RecordIds = Enumerable.Range(0, 6).Select(_ => Guid.NewGuid()).ToArray(),
                    RequestId = Guid.NewGuid(),
                },
                true
            )
        );
        Assert.Empty(f.Store.Pending("asx_outbox"));
    }

    [Fact]
    public void PreviewCoversUpToTheTracedNumberOfRecords()
    {
        var error = Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            new BatchReplan(new DurableWorkerTests.MemoryService(), new[] { "account" }).Execute(
                new WorkerRequest
                {
                    Command = "PreviewBatch",
                    RequestId = Guid.NewGuid(),
                    TemplateId = Guid.NewGuid(),
                    RecordIds = Enumerable
                        .Range(0, Ascentix.Documents.Domain.Bounds.PreviewRecords + 1)
                        .Select(_ => Guid.NewGuid())
                        .ToArray(),
                },
                true
            )
        );
        Assert.Equal(
            "Select one to "
                + Ascentix.Documents.Domain.Bounds.PreviewRecords
                + " unique records for a bounded batch review.",
            error.Message
        );
    }
}
