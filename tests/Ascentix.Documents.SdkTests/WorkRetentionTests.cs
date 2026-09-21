using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class WorkRetentionTests
{
    [Fact]
    public void LatestDestinationKeepsOnlyResultWhileOlderHistoryExpires()
    {
        var f = new DurableWorkerTests.Fixture();
        var job = f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key);
        job.Value.Status = "Applied";
        job.Value.Folders[0].Status = "Applied";
        job.Value.Folders[0].LocationId = Guid.NewGuid();
        job.Value.RetireAfterUtc = DateTime.UtcNow.AddDays(-1);
        f.Store.Save(job);
        f.Store.Create(
            "asx_outbox",
            new RecordPlanDocument
            {
                Key = "recordplan:" + f.TemplateId.ToString("N") + ":" + f.RecordId.ToString("N"),
                Status = "Selection",
                Operations = new[] { job.Value.Key },
            }
        );
        var nativeBefore = f.Service.Rows.Values.Count(r =>
            r.LogicalName == "sharepointdocumentlocation"
        );
        WorkRetention.Purge(f.Service, DateTime.UtcNow);
        var current = f.Store.Require<OperationDocument>("asx_operation", job.Value.Key).Value;
        Assert.True(current.HistoryCompacted);
        Assert.Empty(current.Folders);
        Assert.Null(current.RetireAfterUtc);
        Assert.Equal(job.Value.Folders[0].LocationId, current.ResultLocationId);
        Assert.Equal(
            nativeBefore,
            f.Service.Rows.Values.Count(r => r.LogicalName == "sharepointdocumentlocation")
        );
        Assert.Equal(
            "Applied",
            f.Coordinator.Execute(
                new WorkerRequest
                {
                    Command = "Claim",
                    Key = job.Value.Key,
                    RunId = "later",
                },
                true
            ).Status
        );
    }

    [Fact]
    public void RetentionRemovesExpiredRequestsButPreservesPendingAndUncertainEvidence()
    {
        var f = new DurableWorkerTests.Fixture();
        f.Store.Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = "finished",
                Status = "Planned",
                RetireAfterUtc = DateTime.UtcNow.AddDays(-1),
            }
        );
        f.Store.Create("asx_outbox", new OutboxDocument { Key = "unfinished", Status = "Pending" });
        f.Store.Create(
            "asx_attempt",
            new AttemptDocument
            {
                Key = "uncertain-receipt",
                OperationKey = f.Operation.Key,
                RetireAfterUtc = DateTime.UtcNow.AddDays(-1),
            }
        );
        WorkRetention.Purge(f.Service, DateTime.UtcNow);
        Assert.Null(f.Store.Find<OutboxDocument>("asx_outbox", "finished"));
        Assert.NotNull(f.Store.Find<OutboxDocument>("asx_outbox", "unfinished"));
        Assert.NotNull(f.Store.Find<AttemptDocument>("asx_attempt", "uncertain-receipt"));
        Assert.Throws<EvaluationBlockedException>(() =>
            WorkRetention.ValidateDelete(
                f.Service,
                f.Service.Rows.Values.First(r => r.LogicalName == "sharepointdocumentlocation"),
                DateTime.UtcNow
            )
        );
    }
}
