using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class RecordInspectionTests
{
    [Fact]
    public void RecordSummaryReportsPartialDestinationWithoutPermanentFolderInventory()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        f.Service.Seed(
            new Entity("asx_folder", Guid.NewGuid())
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", f.RevisionId),
                ["asx_key"] = "child",
                ["asx_sectionkey"] = "general",
                ["asx_parentkey"] = "root",
                ["asx_expression"] = "Child",
                ["asx_order"] = 1,
            }
        );
        var queued = f.Coordinator.Execute(
            new WorkerRequest
            {
                Command = "Queue",
                TemplateId = f.TemplateId,
                RecordId = f.RecordId,
                RequestId = Guid.NewGuid(),
            },
            true
        );
        var plan = f.Coordinator.Execute(
            new WorkerRequest { Command = "Plan", Key = queued.Key },
            true
        );
        Assert.Single(plan.Keys);
        var applied = f.Store.Require<OperationDocument>("asx_operation", plan.Keys[0]);
        applied.Value.Folders[0].PhysicalId = Guid.NewGuid();
        applied.Value.Folders[0].Status = "Applied";
        f.Store.Save(applied);
        var result = RecordInspection.Read(
            f.Service,
            new WorkerRequest { TemplateId = f.TemplateId, RecordId = f.RecordId },
            new[] { "account" }
        );
        Assert.Equal("Partial", result.Status);
        Assert.Equal(2, result.Record!.Folders.Length);
        Assert.Contains("Pending", result.Record.OperationStates);
        Assert.DoesNotContain(
            f.Service.Rows.Values,
            r => r.LogicalName == "asx_binding" || r.LogicalName == "asx_physicalfolder"
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            RecordInspection.Read(
                f.Service,
                new WorkerRequest { TemplateId = f.TemplateId, RecordId = f.RecordId },
                Array.Empty<string>()
            )
        );
    }

    [Fact]
    public void ExcessiveDateRetryAfterRequiresReviewJustLikeExcessiveSeconds()
    {
        var now = new DateTime(2026, 9, 8, 12, 0, 0, DateTimeKind.Utc);
        Assert.Throws<EvaluationBlockedException>(() =>
            WorkerCoordinator.RetryAt(
                now,
                1,
                now.AddDays(8).ToString("r", System.Globalization.CultureInfo.InvariantCulture)
            )
        );
    }
}
