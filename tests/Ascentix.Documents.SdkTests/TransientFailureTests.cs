using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class TransientFailureTests
{
    [Theory]
    [InlineData(429, null, null, true)]
    [InlineData(500, null, null, true)]
    [InlineData(502, null, null, true)]
    [InlineData(503, null, null, true)]
    [InlineData(504, null, null, true)]
    [InlineData(408, null, null, true)]
    [InlineData(0, null, null, true)]
    [InlineData(null, null, null, true)]
    [InlineData(400, "0x80072322", null, true)]
    [InlineData(400, "0x80072321", null, true)]
    [InlineData(400, "0x80072326", null, true)]
    [InlineData(412, "0x80060882", null, true)]
    [InlineData(400, "-2147088254", null, true)]
    [InlineData(500, "0x80044151", null, true)]
    [InlineData(500, "0x80040255", null, true)]
    [InlineData(500, "0x80044150", "Generic SQL error. Sql Number: 1205", true)]
    [InlineData(500, "0x80044150", "Generic SQL error. Sql Number: 2627", false)]
    [InlineData(400, null, null, false)]
    [InlineData(403, null, null, false)]
    [InlineData(404, null, null, false)]
    [InlineData(412, null, null, false)]
    [InlineData(400, "0x80040265", "ISV code aborted the operation.", false)]
    [InlineData(500, "0x80040224", "Unexpected exception from plug-in", false)]
    [InlineData(500, "0x8004418D", "Sandbox Worker process crashed", false)]
    [InlineData(400, "0x80060883", "ConcurrencyVersionNotProvided", false)]
    [InlineData(400, "not-a-code", null, false)]
    public void ClassifiesTemporaryAndGenuineFailures(
        int? status,
        string? code,
        string? error,
        bool transient
    ) => Assert.Equal(transient, TransientFailure.Is(status, code, error));

    [Fact]
    public void OnlyARequestThatReportsACauseIsClassified()
    {
        Assert.False(TransientFailure.Is(new WorkerRequest()));
        Assert.True(TransientFailure.Is(new WorkerRequest { StatusCode = 0 }));
        Assert.True(TransientFailure.Is(new WorkerRequest { StatusCode = 429 }));
        Assert.False(TransientFailure.Is(new WorkerRequest { StatusCode = 400 }));
        Assert.Equal(
            "Waiting to retry after a temporary error (HTTP 429, 0x80072322); attempt 3.",
            TransientFailure.Notice(429, "0x80072322", 3)
        );
        Assert.Equal(
            "Waiting to retry after a temporary error (no response); attempt 1.",
            TransientFailure.Notice(0, null, 1)
        );
    }

    private static string QueueFailingPlan(DurableWorkerTests.Fixture f)
    {
        f.Service.Rows[f.LibraryId]["asx_approved"] = false;
        var queued = f.Execute(
            new WorkerRequest
            {
                Command = "Queue",
                TemplateId = f.TemplateId,
                RecordId = f.RecordId,
                RequestId = Guid.NewGuid(),
            }
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Execute(new WorkerRequest { Command = "Plan", Key = queued.Key })
        );
        return queued.Key;
    }

    [Fact]
    public void TransientPlanFailureLeavesTheOutboxRowWaitingWithANotice()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var key = QueueFailingPlan(f);
        var result = f.Execute(
            new WorkerRequest
            {
                Command = "FailOutbox",
                Key = key,
                StatusCode = 429,
                ErrorCode = "0x80072322",
                Error = "Number of requests exceeded the limit",
            }
        );
        Assert.Equal("Pending", result.Status);
        var row = f.Store.Require<OutboxDocument>("asx_outbox", key);
        Assert.Equal("Pending", row.Value.Status);
        Assert.Equal(1, row.Value.Attempts);
        Assert.Equal(f.Now.AddSeconds(30), row.Value.NextAttemptUtc);
        Assert.Equal(
            new[] { "Waiting to retry after a temporary error (HTTP 429, 0x80072322); attempt 1." },
            row.Value.Notices
        );
        Assert.Equal(
            f.Now.AddSeconds(30),
            f.Service.Rows[row.Row.Id].GetAttributeValue<DateTime?>("asx_nextattempt")
        );
        f.Execute(
            new WorkerRequest
            {
                Command = "FailOutbox",
                Key = key,
                StatusCode = 0,
            }
        );
        row = f.Store.Require<OutboxDocument>("asx_outbox", key);
        Assert.Equal(2, row.Value.Attempts);
        Assert.Equal(f.Now.AddSeconds(60), row.Value.NextAttemptUtc);
        Assert.Contains("attempt 2", row.Value.Notices.Single());
    }

    [Fact]
    public void GenuinePlanFailureStillBlocks()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var key = QueueFailingPlan(f);
        var result = f.Execute(
            new WorkerRequest
            {
                Command = "FailOutbox",
                Key = key,
                StatusCode = 500,
                ErrorCode = "0x80040224",
                Error = "Unexpected exception from plug-in",
            }
        );
        Assert.Equal("Blocked", result.Status);
        var row = f.Store.Require<OutboxDocument>("asx_outbox", key).Value;
        Assert.Equal(new[] { WorkerCoordinator.PlanningFailedNotice }, row.Notices);
        Assert.Null(row.NextAttemptUtc);
    }

    [Fact]
    public void AnOldFlowThatSendsNoCauseStillBlocks()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var key = QueueFailingPlan(f);
        Assert.Equal(
            "Blocked",
            f.Execute(new WorkerRequest { Command = "FailOutbox", Key = key }).Status
        );
    }

    [Fact]
    public void ListOutboxSkipsRowsThatAreNotYetDueAndPlanClearsTheWait()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var key = QueueFailingPlan(f);
        f.Execute(
            new WorkerRequest
            {
                Command = "FailOutbox",
                Key = key,
                StatusCode = 503,
            }
        );
        Assert.Empty(f.Execute(new WorkerRequest { Command = "ListOutbox" }).Keys);
        f.Now = f.Now.AddSeconds(30);
        Assert.Equal(new[] { key }, f.Execute(new WorkerRequest { Command = "ListOutbox" }).Keys);
        f.Service.Rows[f.LibraryId]["asx_approved"] = true;
        Assert.Equal(
            "Planned",
            f.Execute(new WorkerRequest { Command = "Plan", Key = key }).Status
        );
        var row = f.Store.Require<OutboxDocument>("asx_outbox", key);
        Assert.Null(row.Value.NextAttemptUtc);
        Assert.Equal(0, row.Value.Attempts);
        Assert.Empty(row.Value.Notices);
        Assert.Null(f.Service.Rows[row.Row.Id].GetAttributeValue<DateTime?>("asx_nextattempt"));
    }

    [Fact]
    public void ListOutboxKeepsCreatedOrderForDueRows()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var first = QueueFailingPlan(f);
        var second = f.Execute(
            new WorkerRequest
            {
                Command = "Queue",
                TemplateId = f.TemplateId,
                RecordId = f.RecordId,
                RequestId = Guid.NewGuid(),
            }
        ).Key;
        f.Execute(
            new WorkerRequest
            {
                Command = "FailOutbox",
                Key = first,
                StatusCode = 503,
            }
        );
        Assert.Equal(
            new[] { second },
            f.Execute(new WorkerRequest { Command = "ListOutbox" }).Keys
        );
        f.Now = f.Now.AddMinutes(1);
        Assert.Equal(
            new[] { first, second },
            f.Execute(new WorkerRequest { Command = "ListOutbox" }).Keys
        );
    }

    [Fact]
    public void RetryOutboxClearsAWaitSoTheRowPlansNow()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var key = QueueFailingPlan(f);
        f.Execute(
            new WorkerRequest
            {
                Command = "FailOutbox",
                Key = key,
                StatusCode = 503,
            }
        );
        var retried = f.Execute(new WorkerRequest { Command = "RetryOutbox", Key = key });
        Assert.Equal("Pending", retried.Status);
        Assert.Null(f.Store.Require<OutboxDocument>("asx_outbox", key).Value.NextAttemptUtc);
        Assert.Equal(new[] { key }, f.Execute(new WorkerRequest { Command = "ListOutbox" }).Keys);
    }

    [Fact]
    public void WaitingRecordInspectionShowsTheNotice()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var key = QueueFailingPlan(f);
        f.Execute(
            new WorkerRequest
            {
                Command = "FailOutbox",
                Key = key,
                StatusCode = 503,
            }
        );
        var inspected = RecordInspection.Read(
            f.Service,
            new WorkerRequest { TemplateId = f.TemplateId, RecordId = f.RecordId },
            new[] { "account" }
        );
        Assert.Equal("WaitingToRetry", inspected.Status);
        Assert.Contains("attempt 1", inspected.Notices.Single());
    }

    [Fact]
    public void TransientClaimFailureWaitsInsteadOfBlocking()
    {
        var f = new DurableWorkerTests.Fixture();
        var result = f.Execute(
            new WorkerRequest
            {
                Command = "FailUnclaimed",
                Key = f.Operation.Key,
                StatusCode = 412,
                ErrorCode = "0x80060882",
            }
        );
        Assert.Equal("RetryWait", result.Status);
        var op = f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value;
        Assert.Equal("RetryWait", op.Status);
        Assert.Equal(1, op.RetryCount);
        Assert.Equal(f.Now.AddSeconds(30), op.NextAttemptUtc);
        Assert.Equal(
            "Waiting to retry after a temporary error (HTTP 412, 0x80060882); attempt 1.",
            op.ErrorCode
        );
        Assert.Equal("RetryWait", f.Claim("run-2").Status);
        f.Now = f.Now.AddSeconds(30);
        Assert.Equal("Read", f.Claim("run-2").Status);
    }

    [Fact]
    public void GenuineOrUnreportedClaimFailureStillBlocks()
    {
        var f = new DurableWorkerTests.Fixture();
        Assert.Equal(
            "Blocked",
            f.Execute(
                new WorkerRequest
                {
                    Command = "FailUnclaimed",
                    Key = f.Operation.Key,
                    StatusCode = 400,
                    ErrorCode = "0x80040265",
                }
            ).Status
        );
        var g = new DurableWorkerTests.Fixture();
        Assert.Equal(
            "Blocked",
            g.Execute(new WorkerRequest { Command = "FailUnclaimed", Key = g.Operation.Key }).Status
        );
    }

    [Fact]
    public void TransientDriveFailureWaitsAndReleasesTheWriter()
    {
        var f = new DurableWorkerTests.Fixture();
        var claim = f.Claim();
        var result = f.Execute(
            new WorkerRequest
            {
                Command = "Fail",
                Key = f.Operation.Key,
                RunId = "run-1",
                Token = claim.Token,
                StatusCode = 503,
            }
        );
        Assert.Equal("RetryWait", result.Status);
        var op = f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value;
        Assert.Equal("RetryWait", op.Status);
        Assert.Contains("HTTP 503", op.ErrorCode);
        Assert.Null(
            f.Store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(f.Service, f.Operation.Key)
            ).Value.RunId
        );
        f.Now = f.Now.AddMinutes(1);
        Assert.Equal("Read", f.Claim("run-2").Status);
    }

    [Fact]
    public void GenuineDriveFailureStillBlocks()
    {
        var f = new DurableWorkerTests.Fixture();
        var claim = f.Claim();
        var result = f.Execute(
            new WorkerRequest
            {
                Command = "Fail",
                Key = f.Operation.Key,
                RunId = "run-1",
                Token = claim.Token,
                StatusCode = 400,
                ErrorCode = "0x80040224",
            }
        );
        Assert.Equal("Blocked", result.Status);
        Assert.Equal(
            "WorkerFailed",
            f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value.ErrorCode
        );
    }

    [Fact]
    public void ThrottledReadsKeepBackingOffWithNoAttemptCap()
    {
        var f = new DurableWorkerTests.Fixture();
        for (int attempt = 1; attempt <= 8; attempt++)
        {
            var claim = f.Claim("run-" + attempt);
            Assert.Equal("Read", claim.Status);
            var waiting = f.Call("Observe", claim, "{}", 429);
            Assert.Equal("RetryWait", waiting.Status);
            var op = f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value;
            Assert.Equal(attempt, op.RetryCount);
            Assert.Equal(
                "Waiting to retry after a temporary error (HTTP 429); attempt " + attempt + ".",
                op.ErrorCode
            );
            Assert.Equal(
                f.Now.AddSeconds(Math.Min(900, 30 * Math.Pow(2, attempt - 1))),
                op.NextAttemptUtc
            );
            f.Now = op.NextAttemptUtc!.Value;
        }
    }

    [Theory]
    [InlineData("not a date")]
    [InlineData("-5")]
    [InlineData("1.5")]
    public void UnusableRetryAfterFallsBackToExponentialBackoff(string retryAfter)
    {
        var now = new DateTime(2026, 9, 8, 8, 0, 0, DateTimeKind.Utc);
        Assert.Equal(now.AddSeconds(60), WorkerCoordinator.RetryAt(now, 2, retryAfter));
    }

    [Fact]
    public void HonoredRetryAfterIsCappedAtTheFifteenMinuteInterval()
    {
        var now = new DateTime(2026, 9, 8, 8, 0, 0, DateTimeKind.Utc);
        Assert.Equal(now.AddMinutes(2), WorkerCoordinator.RetryAt(now, 1, "120"));
        Assert.Equal(now.AddMinutes(15), WorkerCoordinator.RetryAt(now, 1, "3600"));
        Assert.Equal(now.AddMinutes(15), WorkerCoordinator.RetryAt(now, 1, "999999999999"));
        Assert.Equal(
            now.AddMinutes(15),
            WorkerCoordinator.RetryAt(
                now,
                1,
                now.AddDays(8).ToString("r", System.Globalization.CultureInfo.InvariantCulture)
            )
        );
        Assert.Equal(now.AddMinutes(15), WorkerCoordinator.RetryAt(now, 40, null));
    }

    [Fact]
    public void ThrottledFolderCreateWasNotExecutedSoItWaitsAndCreatesLater()
    {
        var f = new DurableWorkerTests.Fixture();
        var absent = f.Observe(f.Preflight(f.Claim()), "{}");
        Assert.Equal("ReadyToCreate", absent.Status);
        var prepared = f.Call("PrepareCreate", absent);
        Assert.Equal("Create", prepared.Status);
        var request = new WorkerRequest
        {
            Command = "CreateResponse",
            Key = f.Operation.Key,
            RunId = "run-1",
            Token = prepared.Token,
            HttpStatus = 429,
            RetryAfter = "120",
        };
        var throttled = f.Service.Transaction(() =>
        {
            Assert.Null(WorkCoordination.Response(f.Service, request, f.Now));
            return f.Coordinator.Execute(request, true);
        });
        Assert.Equal("RetryWait", throttled.Status);
        var op = f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value;
        Assert.False(op.ExternalSubmitted);
        Assert.False(op.ExternalResponseKnown);
        Assert.Equal(f.Now.AddMinutes(2), op.NextAttemptUtc);
        Assert.Contains("HTTP 429", op.ErrorCode);
        var writer = f.Store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(f.Service, f.Operation.Key)
        );
        Assert.Null(writer.Value.RunId);
        Assert.False(writer.Value.HttpOutstanding);
        Assert.Empty(
            f.Store.Require<ConnectionBudget>("asx_claim", WorkCoordination.BudgetKey).Value.Writers
        );
        f.Now = f.Now.AddMinutes(2);
        var again = f.Observe(f.Preflight(f.Claim("run-2")), "{}");
        Assert.Equal("ReadyToCreate", again.Status);
        Assert.Equal("Create", f.Call("PrepareCreate", again).Status);
    }

    [Fact]
    public void HugeOrBadRetryAfterNeverPausesSharedWorkForDays()
    {
        foreach (var retryAfter in new[] { "999999999999", "garbage" })
        {
            var f = new DurableWorkerTests.Fixture();
            var claim = f.Claim();
            var request = new WorkerRequest
            {
                Command = "Observe",
                Key = f.Operation.Key,
                RunId = "run-1",
                Token = claim.Token,
                HttpStatus = 429,
                RetryAfter = retryAfter,
            };
            f.Service.Transaction(() => WorkCoordination.Response(f.Service, request, f.Now));
            var budget = f.Store.Require<ConnectionBudget>("asx_claim", WorkCoordination.BudgetKey);
            Assert.True(budget.Value.PauseUntilUtc <= f.Now.AddMinutes(15));
        }
    }
}
