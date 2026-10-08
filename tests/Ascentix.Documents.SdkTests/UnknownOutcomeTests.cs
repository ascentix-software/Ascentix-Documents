using System;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class UnknownOutcomeTests
{
    private static DispatcherDocument Writer(DurableWorkerTests.Fixture f) =>
        f
            .Store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(f.Service, f.Operation.Key)
            )
            .Value;

    private static OperationDocument Op(DurableWorkerTests.Fixture f) =>
        f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value;

    /// <summary>Sends a response through the shared admission step, then the worker, as the API does.</summary>
    private static WorkerResult Respond(
        DurableWorkerTests.Fixture f,
        string command,
        WorkerResult previous,
        int status
    )
    {
        var request = new WorkerRequest
        {
            Command = command,
            Key = f.Operation.Key,
            RunId = "run-1",
            Token = previous.Token,
            ProbeId = previous.ProbeId,
            ProbeKind = previous.ProbeKind,
            HttpStatus = status,
        };
        return f.Service.Transaction(() =>
            WorkCoordination.Response(f.Service, request, f.Now)
            ?? f.Coordinator.Execute(request, true)
        );
    }

    [Theory]
    [InlineData(0)]
    [InlineData(408)]
    public void ReadTimeoutWaitsAndReleasesTheWriterSlot(int status)
    {
        var f = new DurableWorkerTests.Fixture();
        var read = f.Claim();
        var result = Respond(f, "Observe", read, status);
        Assert.Equal("RetryWait", result.Status);
        Assert.Equal("RetryWait", Op(f).Status);
        Assert.Null(Writer(f).RunId);
        Assert.False(Writer(f).HttpOutstanding);
        Assert.Empty(
            f.Store.Require<ConnectionBudget>("asx_claim", WorkCoordination.BudgetKey).Value.Writers
        );
    }

    private static WorkerResult TimedOutCreate(DurableWorkerTests.Fixture f)
    {
        var prepared = f.Call("PrepareCreate", f.Observe(f.Preflight(f.Claim()), "{}"));
        Assert.Equal("Create", prepared.Status);
        var unknown = Respond(f, "CreateResponse", prepared, 0);
        Assert.Equal("Quarantined", unknown.Status);
        Assert.Equal("ExternalUnknown", Op(f).Status);
        return prepared;
    }

    [Fact]
    public void FolderCreateTimeoutIsTakenOverAfterTheLeaseAndAdoptsTheFolder()
    {
        var f = new DurableWorkerTests.Fixture();
        TimedOutCreate(f);
        Assert.Equal("Quarantined", f.Claim("run-2").Status);
        f.Now = f.Now.AddMinutes(6);
        Assert.Contains(
            f.Operation.Key,
            f.Execute(new WorkerRequest { Command = "ListOperations" }).Keys
        );
        var read = f.Claim("run-2");
        Assert.Equal("Read", read.Status);
        Assert.False(Writer(f).HttpOutstanding);
        var physical = Guid.NewGuid();
        var verified = f.ObserveAndFinalize(f.Preflight(read), f.Item(physical, null));
        Assert.Equal("Applied", f.Call("Complete", verified).Status);
        Assert.Single(f.Results, r => r.Status == "Create");
        Assert.Equal(physical, Op(f).Folders[0].PhysicalId);
    }

    [Fact]
    public void RecoveryRequiredLibraryWorkDoesNotStarveNewerWork()
    {
        var service = new DurableWorkerTests.MemoryService();
        var store = new DocumentStore(service);
        for (int i = 0; i < 25; i++)
            store.Create(
                "asx_operation",
                new LibrarySetup
                {
                    Key = "librarycreate:" + i,
                    Status = "RecoveryRequired",
                    Name = "Library " + i,
                }
            );
        store.Create("asx_operation", new OperationDocument { Key = "folderjob:newer" });
        Assert.Equal(new[] { "folderjob:newer" }, store.Pending("asx_operation"));
    }

    [Fact]
    public void AnExpiredRowThatCannotBeClaimedBacksOffAndReleasesItsSlot()
    {
        var f = new DurableWorkerTests.Fixture();
        Assert.Equal("Read", f.Claim().Status);
        // The library's entry folder changed, so claiming the job fails until it is re-pointed.
        f.Service.Rows[f.LibraryId]["asx_entryid"] = Guid.NewGuid().ToString();
        var fail = new WorkerRequest
        {
            Command = "FailUnclaimed",
            Key = f.Operation.Key,
            StatusCode = 400,
            ErrorCode = "0x80040265",
            Error = "Approved destination entry changed.",
        };
        // While the run's claim is live, nothing changes.
        Assert.Equal("Quarantined", f.Execute(fail).Status);
        f.Now = f.Now.AddMinutes(6);
        Assert.Contains(
            f.Operation.Key,
            f.Execute(new WorkerRequest { Command = "ListOperations" }).Keys
        );
        Assert.ThrowsAny<Exception>(() => f.Claim("run-2"));
        Assert.Equal("RetryWait", f.Execute(fail).Status);
        var op = Op(f);
        Assert.Equal(f.Now.AddSeconds(30), op.NextAttemptUtc);
        Assert.Contains("entry changed", op.ErrorCode);
        Assert.Contains("attempt 1", op.ErrorCode);
        Assert.Null(Writer(f).RunId);
        Assert.Empty(
            f.Store.Require<ConnectionBudget>("asx_claim", WorkCoordination.BudgetKey).Value.Writers
        );
        Assert.DoesNotContain(
            f.Operation.Key,
            f.Execute(new WorkerRequest { Command = "ListOperations" }).Keys
        );
        f.Now = f.Now.AddSeconds(30);
        Assert.Contains(
            f.Operation.Key,
            f.Execute(new WorkerRequest { Command = "ListOperations" }).Keys
        );
    }

    [Fact]
    public void AnExpiredUnknownCreateThatCannotBeClaimedRereadsLater()
    {
        var f = new DurableWorkerTests.Fixture();
        TimedOutCreate(f);
        f.Now = f.Now.AddMinutes(6);
        var result = f.Execute(
            new WorkerRequest
            {
                Command = "FailUnclaimed",
                Key = f.Operation.Key,
                StatusCode = 503,
            }
        );
        Assert.Equal("RetryWait", result.Status);
        Assert.True(Op(f).Reprobe);
        Assert.Null(Writer(f).RunId);
        f.Now = f.Now.AddMinutes(1);
        var verified = f.ObserveAndFinalize(
            f.Preflight(f.Claim("run-2")),
            f.Item(Guid.NewGuid(), null)
        );
        Assert.Equal("Applied", f.Call("Complete", verified).Status);
        Assert.Single(f.Results, r => r.Status == "Create");
    }

    [Fact]
    public void FolderCreateTimeoutWithNoFolderCreatesItOnce()
    {
        var f = new DurableWorkerTests.Fixture();
        TimedOutCreate(f);
        f.Now = f.Now.AddMinutes(6);
        var absent = f.Observe(f.Preflight(f.Claim("run-2")), "{}");
        Assert.Equal("ReadyToCreate", absent.Status);
        Assert.False(Op(f).ExternalSubmitted);
        var again = f.Call("PrepareCreate", absent);
        Assert.Equal("Create", again.Status);
        var created = f.Call("CreateResponse", again, DurableWorkerTests.CreateBody(), 201);
        var verified = f.ObserveAndFinalize(created, f.Item(Guid.NewGuid(), null));
        Assert.Equal("Applied", f.Call("Complete", verified).Status);
        Assert.Equal(2, f.Results.Count(r => r.Status == "Create"));
    }

    [Fact]
    public void OperatorCancelAndRetryWorkAfterTheLeaseExpires()
    {
        var f = new DurableWorkerTests.Fixture();
        TimedOutCreate(f);
        f.Now = f.Now.AddMinutes(6);
        Assert.Equal(
            "Pending",
            f.Execute(new WorkerRequest { Command = "Retry", Key = f.Operation.Key }).Status
        );
        Assert.Null(Writer(f).RunId);
        Assert.False(Writer(f).HttpOutstanding);
        var g = new DurableWorkerTests.Fixture();
        TimedOutCreate(g);
        Assert.ThrowsAny<Exception>(() =>
            g.Execute(new WorkerRequest { Command = "Cancel", Key = g.Operation.Key })
        );
        g.Now = g.Now.AddMinutes(6);
        Assert.Equal(
            "Cancelled",
            g.Execute(new WorkerRequest { Command = "Cancel", Key = g.Operation.Key }).Status
        );
        Assert.Null(Writer(g).RunId);
    }

    [Fact]
    public void RetriedUnknownCreateReprobesAndAdoptsWithoutAnotherPost()
    {
        var f = new DurableWorkerTests.Fixture();
        TimedOutCreate(f);
        f.Now = f.Now.AddMinutes(6);
        f.Execute(new WorkerRequest { Command = "Retry", Key = f.Operation.Key });
        var verified = f.ObserveAndFinalize(
            f.Preflight(f.Claim("run-2")),
            f.Item(Guid.NewGuid(), null)
        );
        Assert.Equal("Applied", f.Call("Complete", verified).Status);
        Assert.Single(f.Results, r => r.Status == "Create");
    }

    [Fact]
    public void AReconcilingSetupIsListedOnlyOnceDue()
    {
        var service = new DurableWorkerTests.MemoryService();
        var store = new DocumentStore(service);
        var now = new DateTime(2026, 10, 6, 12, 0, 0, DateTimeKind.Utc);
        store.Create(
            "asx_operation",
            new LibrarySetup { Key = "librarycreate:wait", Name = "Wait" }
        );
        var setup = store.Require<LibrarySetup>("asx_operation", "librarycreate:wait");
        setup.Value.Status = "Reconciling";
        setup.Value.Reconcile = true;
        setup.Value.NextAttemptUtc = now.AddMinutes(3);
        store.Save(setup);
        Assert.Empty(store.Pending("asx_operation", now: now));
        Assert.Equal(
            new[] { "librarycreate:wait" },
            store.Pending("asx_operation", now: now.AddMinutes(4))
        );
    }
}
