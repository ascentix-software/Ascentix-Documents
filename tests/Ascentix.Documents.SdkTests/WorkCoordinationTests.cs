using System;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class WorkCoordinationTests
{
    [Fact]
    public void EverySiteCanHoldItsOwnWriterAtOnceAndReleaseRemovesIt()
    {
        var f = new Fixture();
        var a = f.Claim("a");
        f.Claim("b");
        f.Claim("c");
        Assert.True(
            WorkCoordination.HasCapacity(
                f.Service,
                WorkCoordination.SiteUrl("https://example.sharepoint.com/sites/d")
            )
        );
        f.Claim("d");
        Assert.Equal(
            4,
            f.Store.Require<ConnectionBudget>(
                "asx_claim",
                WorkCoordination.BudgetKey
            ).Value.Writers.Length
        );
        var claim = f.Store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(f.Service, a.Key)
        );
        claim.Value.RunId = null;
        claim.Value.Status = "Idle";
        f.Store.Save(claim);
        Assert.Equal(
            3,
            f.Store.Require<ConnectionBudget>(
                "asx_claim",
                WorkCoordination.BudgetKey
            ).Value.Writers.Length
        );
    }

    [Fact]
    public void SharedPacingAndRetryAfterApplyAcrossSites()
    {
        var f = new Fixture();
        var a = f.Claim("a");
        var b = f.Claim("b");
        Assert.Equal("Permit", f.Begin(a).Status);
        Assert.Equal("Wait", f.Begin(b).Status);
        a.HttpStatus = 429;
        a.RetryAfter = "120";
        f.Response(a);
        f.Now = f.Now.AddSeconds(90);
        var wait = f.Begin(b);
        Assert.Equal("Wait", wait.Status);
        Assert.Equal(30, wait.WaitSeconds);
        f.Now = f.Now.AddSeconds(30);
        Assert.Equal("Permit", f.Begin(b).Status);
    }

    [Fact]
    public void UnknownRequestKeepsSlotAndRequiresExactRecoveryBeforeAnotherRequest()
    {
        var f = new Fixture();
        var a = f.Claim("a");
        Assert.Equal("Permit", f.Begin(a).Status);
        Assert.Throws<EvaluationBlockedException>(() => f.Begin(a));
        a.HttpStatus = 0;
        Assert.Equal("Quarantined", f.Response(a)!.Status);
        Assert.Throws<EvaluationBlockedException>(() => WorkCoordination.RequireIdle(f.Service));
        var claim = f.Store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(f.Service, a.Key)
        );
        claim.Value.RunId = null;
        Assert.Throws<EvaluationBlockedException>(() => f.Store.Save(claim));
        f.Now = f.Now.AddMinutes(6);
        Assert.Throws<EvaluationBlockedException>(() => f.Begin(a));
        a.Evidence = "Original flow terminated; outstanding HTTP completion verified.";
        new WorkerCoordinator(f.Service, () => f.Now).PermitRecovery(a, true);
        Assert.False(
            f.Store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(f.Service, a.Key)
            ).Value.HttpOutstanding
        );
        Assert.True(WorkCoordination.Busy(f.Service)); // Recovery permits takeover; it does not itself release the writer.
    }

    [Fact]
    public void ThrottlingReducesSharedPaceAndQuietRecoveryNeverExceedsOriginalBudget()
    {
        var f = new Fixture();
        var a = f.Claim("a");
        f.Begin(a);
        a.HttpStatus = 429;
        a.RetryAfter = "120";
        f.Response(a);
        Assert.Equal(
            2,
            f.Store.Require<ConnectionBudget>(
                "asx_claim",
                WorkCoordination.BudgetKey
            ).Value.PaceSeconds
        );
        Assert.False(
            WorkCoordination.HasCapacity(
                f.Service,
                WorkCoordination.SiteUrl("https://example.sharepoint.com/sites/b"),
                f.Now
            )
        );
        f.Now = f.Now.AddSeconds(120);
        f.Begin(a);
        Assert.Equal(
            f.Now.AddSeconds(2),
            f.Store.Require<ConnectionBudget>(
                "asx_claim",
                WorkCoordination.BudgetKey
            ).Value.NextStartUtc
        );
        a.HttpStatus = 200;
        f.Response(a);
        f.Now = f.Now.AddMinutes(11);
        var claim = f.Store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(f.Service, a.Key)
        );
        claim.Value.LeaseUntilUtc = f.Now.AddMinutes(5);
        f.Store.Save(claim);
        f.Begin(a);
        Assert.Equal(
            1,
            f.Store.Require<ConnectionBudget>(
                "asx_claim",
                WorkCoordination.BudgetKey
            ).Value.PaceSeconds
        );
    }

    private sealed class Fixture
    {
        public readonly DurableWorkerTests.MemoryService Service =
            new DurableWorkerTests.MemoryService();
        public DocumentStore Store => new DocumentStore(Service);
        public DateTime Now = new DateTime(2026, 9, 9, 12, 0, 0, DateTimeKind.Utc);

        public WorkerRequest Claim(string site) =>
            Service.Transaction(() =>
            {
                var key = "catalogprobe:" + site;
                var url = "https://example.sharepoint.com/sites/" + site;
                if (Store.Find<CatalogProbe>("asx_operation", key) == null)
                    Store.Create("asx_operation", new CatalogProbe { Key = key, WebUrl = url });
                var scope = WorkCoordination.SiteUrl(url);
                if (Store.Find<DispatcherDocument>("asx_claim", scope) == null)
                    Store.Create(
                        "asx_claim",
                        new DispatcherDocument { Key = scope, Status = "Idle" }
                    );
                var row = Store.Require<DispatcherDocument>("asx_claim", scope);
                row.Value.OperationKey = key;
                row.Value.RunId = site;
                row.Value.Token = Guid.NewGuid();
                row.Value.LeaseUntilUtc = Now.AddMinutes(5);
                row.Value.Status = "Claimed";
                Store.Save(row);
                return new WorkerRequest
                {
                    Key = key,
                    RunId = site,
                    Token = row.Value.Token,
                };
            });

        public WorkerResult Begin(WorkerRequest request) =>
            Service.Transaction(() => WorkCoordination.BeginHttp(Service, request, Now));

        public WorkerResult? Response(WorkerRequest request) =>
            Service.Transaction(() => WorkCoordination.Response(Service, request, Now));
    }
}
