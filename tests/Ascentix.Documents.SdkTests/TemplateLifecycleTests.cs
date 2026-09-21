using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class TemplateLifecycleTests
{
    [Fact]
    public void ScheduleIncludesStartAndExcludesEndAndCanBeCleared()
    {
        var now = new DateTime(2026, 9, 8, 12, 0, 0, DateTimeKind.Utc);
        var row = new Entity("asx_template")
        {
            ["asx_startsutc"] = now,
            ["asx_endsutc"] = now.AddHours(1),
        };
        Assert.False(TemplateLifecycle.Active(row, now.AddTicks(-1)));
        Assert.True(TemplateLifecycle.Active(row, now));
        Assert.False(TemplateLifecycle.Active(row, now.AddHours(1)));
        row["asx_disabled"] = true;
        Assert.False(TemplateLifecycle.Active(row, now));
        row["asx_disabled"] = false;
        Assert.True(TemplateLifecycle.Active(row, now));
        Assert.Throws<EvaluationBlockedException>(() =>
            TemplateLifecycle.Validate(new Entity("asx_template") { ["asx_endsutc"] = now }, row)
        );
        TemplateLifecycle.Validate(
            new Entity("asx_template") { ["asx_startsutc"] = null, ["asx_endsutc"] = null },
            row
        );
        Assert.False(TemplateLifecycle.Active(null, now));
    }

    static void Stop(DurableWorkerTests.Fixture f, string reason)
    {
        if (reason == "deleted")
            f.Service.Rows.Remove(f.TemplateId);
        else if (reason == "disabled")
            f.Service.Rows[f.TemplateId]["asx_disabled"] = true;
        else if (reason == "future")
            f.Service.Rows[f.TemplateId]["asx_startsutc"] = f.Now.AddHours(1);
        else
            f.Service.Rows[f.TemplateId]["asx_endsutc"] = f.Now;
    }

    [Theory]
    [InlineData("deleted")]
    [InlineData("disabled")]
    [InlineData("future")]
    [InlineData("ended")]
    public void UnavailableTemplateCannotQueueReplanOrClaim(string reason)
    {
        var f = new DurableWorkerTests.Fixture();
        Stop(f, reason);
        foreach (var command in new[] { "Queue", "Replan" })
            Assert.Equal(
                "Inactive",
                f.Coordinator.Execute(
                    new WorkerRequest
                    {
                        Command = command,
                        TemplateId = f.TemplateId,
                        RecordId = f.RecordId,
                        RequestId = Guid.NewGuid(),
                    },
                    true
                ).Status
            );
        Assert.Equal("Cancelled", f.Claim().Status);
        Assert.False(
            f.Store.Require<OperationDocument>(
                "asx_operation",
                f.Operation.Key
            ).Value.ExternalSubmitted
        );
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "sharepointdocumentlocation");
    }

    [Theory]
    [InlineData("deleted")]
    [InlineData("disabled")]
    [InlineData("ended")]
    public void PendingPlanStopsCleanly(string reason)
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
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
        Stop(f, reason);
        Assert.Equal(
            "Cancelled",
            f.Coordinator.Execute(
                new WorkerRequest { Command = "Plan", Key = queued.Key },
                true
            ).Status
        );
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_operation");
    }

    [Theory]
    [InlineData("deleted")]
    [InlineData("disabled")]
    [InlineData("ended")]
    public void StopImmediatelyBeforeDispatchCancelsWithoutSendingCreate(string reason)
    {
        var f = new DurableWorkerTests.Fixture();
        var ready = f.Observe(f.Preflight(f.Claim()), "{}");
        Stop(f, reason);
        var result = f.Call("PrepareCreate", ready);
        Assert.Equal("Cancelled", result.Status);
        Assert.Null(result.Http);
        Assert.Null(
            f.Store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(f.Service, f.Operation.Key)
            ).Value.RunId
        );
        Assert.False(
            f.Store.Require<OperationDocument>(
                "asx_operation",
                f.Operation.Key
            ).Value.ExternalSubmitted
        );
    }

    [Fact]
    public void SubmittedRequestCanReconcileAfterTemplateDeletion()
    {
        var f = new DurableWorkerTests.Fixture();
        var create = f.Call("PrepareCreate", f.Observe(f.Preflight(f.Claim()), "{}"));
        Stop(f, "deleted");
        var read = f.Call(
            "CreateResponse",
            create,
            JsonWire.Write(
                new ODataEnvelope<CreateObservation>
                {
                    Data = new CreateObservation
                    {
                        Fields = new ODataRows<FieldValidation>
                        {
                            Rows = new[]
                            {
                                new FieldValidation { Field = "FileLeafRef", HasException = false },
                            },
                        },
                    },
                }
            ),
            200
        );
        var verified = f.ObserveAndFinalize(read, f.Item(Guid.NewGuid(), null));
        Assert.Equal("Applied", f.Call("Complete", verified).Status);
        Assert.NotEqual(
            Guid.Empty,
            f.Store.Require<OperationDocument>(
                "asx_operation",
                f.Operation.Key
            ).Value.Folder.PhysicalId
        );
    }

    [Fact]
    public void FreshEvaluationAfterReactivationRestartsOnlyLifecycleCancelledWork()
    {
        var f = new DurableWorkerTests.Fixture();
        f.SeedTemplate();
        Stop(f, "disabled");
        Assert.Equal("Cancelled", f.Claim().Status);
        f.Service.Rows[f.TemplateId]["asx_disabled"] = false;
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
        var planned = f.Coordinator.Execute(
            new WorkerRequest { Command = "Plan", Key = queued.Key },
            true
        );
        Assert.Equal(
            "Pending",
            f.Store.Require<OperationDocument>("asx_operation", planned.Keys.Single()).Value.Status
        );
        Assert.Equal(
            "Cancelled",
            f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value.Status
        );
    }

    [Fact]
    public void ReactivationAllowsNewRequestsAndDeletedHistoryIsReadable()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        Stop(f, "disabled");
        f.Service.Rows[f.TemplateId]["asx_disabled"] = false;
        Assert.Equal(
            "Pending",
            f.Coordinator.Execute(
                new WorkerRequest
                {
                    Command = "Queue",
                    TemplateId = f.TemplateId,
                    RecordId = f.RecordId,
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Status
        );
        Stop(f, "deleted");
        Assert.Equal(
            "TemplateDeleted",
            RecordInspection
                .Read(
                    f.Service,
                    new WorkerRequest { TemplateId = f.TemplateId, RecordId = f.RecordId },
                    new[] { "account" }
                )
                .Status
        );
    }
}
