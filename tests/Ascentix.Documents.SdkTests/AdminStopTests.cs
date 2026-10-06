using System;
using Ascentix.Documents.Dataverse;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class AdminStopTests
{
    /// <summary>Makes another run hold the site writer, as a busy site does.</summary>
    internal static void HoldWriter(DocumentStore store, string key, DateTime now)
    {
        if (store.Find<DispatcherDocument>("asx_claim", key) == null)
            store.Create("asx_claim", new DispatcherDocument { Key = key, Status = "Idle" });
        var claim = store.Require<DispatcherDocument>("asx_claim", key);
        claim.Value.RunId = "other/run";
        claim.Value.OperationKey = "folderjob:other";
        claim.Value.Token = Guid.NewGuid();
        claim.Value.LeaseUntilUtc = now.AddMinutes(5);
        claim.Value.Status = "Claimed";
        store.Save(claim);
    }

    private static CatalogResult Suspend(DurableWorkerTests.Fixture f) =>
        f.Service.Transaction(() =>
            new CatalogAdministration(f.Service, () => f.Now).Execute(
                new CatalogRequest
                {
                    Command = "SuspendLibrary",
                    CatalogId = f.LibraryId,
                    CatalogRowVersion = f.Service.Rows[f.LibraryId].RowVersion,
                },
                true
            )
        );

    [Fact]
    public void SuspendSucceedsWhileAnotherRunWritesOnTheSite()
    {
        var f = new DurableWorkerTests.Fixture();
        HoldWriter(f.Store, WorkCoordination.Library(f.Service, f.LibraryId), f.Now);
        Assert.StartsWith("Suspended", Suspend(f).Status);
        Assert.False(f.Service.Rows[f.LibraryId].GetAttributeValue<bool>("asx_approved"));
        Assert.Equal(
            "other/run",
            f.Store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Library(f.Service, f.LibraryId)
            ).Value.RunId
        );
    }

    [Fact]
    public void AfterSuspendNoNewFolderJobIsClaimedOrCreated()
    {
        var f = new DurableWorkerTests.Fixture();
        var ready = f.Observe(f.Preflight(f.Claim()), "{}");
        Assert.Equal("ReadyToCreate", ready.Status);
        Suspend(f);
        // The job is held until the library is approved again; it prepares no write.
        var stopped = f.Call("PrepareCreate", ready);
        Assert.Equal("RetryWait", stopped.Status);
        Assert.Contains(stopped.Notices, n => n.Contains("suspended"));
        Assert.DoesNotContain(f.Results, r => r.Status == "Create");
        var g = new DurableWorkerTests.Fixture();
        Suspend(g);
        Assert.Equal("RetryWait", g.Claim().Status);
        Assert.DoesNotContain(g.Results, r => r.Status == "Read");
    }

    [Fact]
    public void ACreateAlreadySentFinishesAfterSuspend()
    {
        var f = new DurableWorkerTests.Fixture();
        var prepared = f.Call("PrepareCreate", f.Observe(f.Preflight(f.Claim()), "{}"));
        Assert.Equal("Create", prepared.Status);
        Assert.StartsWith("Suspended", Suspend(f).Status);
        var created = f.Call("CreateResponse", prepared, DurableWorkerTests.CreateBody(), 200);
        Assert.Equal("Read", created.Status);
        var verified = f.ObserveAndFinalize(created, f.Item(Guid.NewGuid(), null));
        Assert.Equal("Applied", f.Call("Complete", verified).Status);
    }
}
