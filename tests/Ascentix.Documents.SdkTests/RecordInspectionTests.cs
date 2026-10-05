using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
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
    public void RecordWhoseOnlyOutboxRowIsBlockedReportsBlockedWithItsNotice()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var request = new WorkerRequest { TemplateId = f.TemplateId, RecordId = f.RecordId };
        Assert.Equal(
            "NotPlanned",
            RecordInspection.Read(f.Service, request, new[] { "account" }).Status
        );
        var key = f.BlockRecord();
        var result = RecordInspection.Read(f.Service, request, new[] { "account" });
        Assert.Equal("Blocked", result.Status);
        Assert.Equal(key, result.Key);
        Assert.Equal(new[] { DurableWorkerTests.PlanningFailedNotice }, result.Notices);
        Assert.Equal(
            "Pending",
            f.Execute(new WorkerRequest { Command = "RetryOutbox", Key = key }).Status
        );
        Assert.Equal(
            "NotPlanned",
            RecordInspection.Read(f.Service, request, new[] { "account" }).Status
        );
        Assert.Equal(
            "Planned",
            f.Execute(new WorkerRequest { Command = "Plan", Key = key }).Status
        );
        Assert.Equal(
            "Pending",
            RecordInspection.Read(f.Service, request, new[] { "account" }).Status
        );
    }

    [Fact]
    public void InspectionFindsBlockedRowsByTheIndexedRecordColumn()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var key = f.BlockRecord();
        var queries = new List<QueryExpression>();
        f.Service.QueryHook = q =>
        {
            if (q.EntityName == "asx_outbox")
                queries.Add(q);
            return null;
        };
        var request = new WorkerRequest { TemplateId = f.TemplateId, RecordId = f.RecordId };
        Assert.Equal(
            "Blocked",
            RecordInspection.Read(f.Service, request, new[] { "account" }).Status
        );
        var lookup = queries.Single(q =>
            q.ColumnSet.Columns.Contains("asx_payload") && q.TopCount == null
        );
        Assert.Contains(
            lookup.Criteria.Conditions,
            c =>
                c.AttributeName == "asx_recordid"
                && c.Operator == ConditionOperator.Equal
                && (string)c.Values[0] == f.RecordId.ToString("D")
        );
        // A row written before the column was populated is not found until it is retried.
        f.Service.QueryHook = null;
        var row = f.Store.Require<OutboxDocument>("asx_outbox", key).Row;
        f.Service.Rows[row.Id].Attributes.Remove("asx_recordid");
        Assert.Equal(
            "NotPlanned",
            RecordInspection.Read(f.Service, request, new[] { "account" }).Status
        );
        Assert.Equal(
            "Pending",
            f.Execute(new WorkerRequest { Command = "RetryOutbox", Key = key }).Status
        );
        f.Service.Rows[f.LibraryId]["asx_approved"] = false;
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Execute(new WorkerRequest { Command = "Plan", Key = key })
        );
        f.Execute(new WorkerRequest { Command = "FailOutbox", Key = key });
        Assert.Equal(
            "Blocked",
            RecordInspection.Read(f.Service, request, new[] { "account" }).Status
        );
    }

    [Fact]
    public void NewerPendingRequestWinsOverAnOlderBlockedRow()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        f.BlockRecord();
        var request = new WorkerRequest { TemplateId = f.TemplateId, RecordId = f.RecordId };
        var queued = f.Execute(
            new WorkerRequest
            {
                Command = "Queue",
                TemplateId = f.TemplateId,
                RecordId = f.RecordId,
                RequestId = Guid.NewGuid(),
            }
        );
        Assert.Equal("Pending", queued.Status);
        Assert.Equal(
            "NotPlanned",
            RecordInspection.Read(f.Service, request, new[] { "account" }).Status
        );
    }

    [Fact]
    public void BlockedRowForAnotherRecordDoesNotChangeThisRecordsInspection()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        f.BlockRecord();
        var unrelated = Guid.NewGuid();
        f.Service.Seed(new Entity("account", unrelated) { ["name"] = "Other" });
        Assert.Equal(
            "NotPlanned",
            RecordInspection
                .Read(
                    f.Service,
                    new WorkerRequest { TemplateId = f.TemplateId, RecordId = unrelated },
                    new[] { "account" }
                )
                .Status
        );
    }

    [Fact]
    public void ExcessiveDateRetryAfterIsCappedJustLikeExcessiveSeconds()
    {
        var now = new DateTime(2026, 9, 8, 12, 0, 0, DateTimeKind.Utc);
        Assert.Equal(
            now.AddMinutes(15),
            WorkerCoordinator.RetryAt(
                now,
                1,
                now.AddDays(8).ToString("r", System.Globalization.CultureInfo.InvariantCulture)
            )
        );
    }
}
