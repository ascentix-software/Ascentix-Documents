using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

/// <summary>A library create whose answer was lost is resolved by looking SharePoint up (spec 6.8).</summary>
public sealed class LookupRecoveryTests
{
    private static LibraryProvisioningTests.Fixture New() => new LibraryProvisioningTests.Fixture();

    private static void SiteFree(LibraryProvisioningTests.Fixture f, string key)
    {
        Assert.Null(f.Claim(key).RunId);
        Assert.False(WorkCoordination.Busy(f.Service));
    }

    [Fact]
    public void ALostCreateIsLookedUpWithReadsOnlyAndFound()
    {
        var f = New();
        var key = f.Lost();
        var result = f.Run(key);
        Assert.Equal("RecoveryRequired", result.Status);
        Assert.Equal(
            "SharePoint has a library named Documents that matches this request.",
            Assert.Single(result.Notices)
        );
        Assert.Equal("Found", f.Finding(key).State);
        Assert.Equal(2, f.LookupReads.Count);
        Assert.StartsWith(
            "_api/web/lists/GetByTitle(@p)?@p='Documents'&$select=Id,Title,BaseTemplate,Created",
            f.LookupReads[0]
        );
        Assert.StartsWith(
            "_api/web/GetList(@p)?@p='%2Fsites%2Ftest%2FDocuments'&$select=",
            f.LookupReads[1]
        );
        Assert.Single(f.Posts, p => p == "_api/web/lists");
        SiteFree(f, key);
        Assert.DoesNotContain(
            key,
            f.Store.Pending("asx_operation", now: DateTime.UtcNow.AddHours(1))
        );
    }

    [Fact]
    public void UseTheLibraryThatWasCreatedFinishesWithoutCreatingAgain()
    {
        var f = New();
        var key = f.Lost();
        f.Run(key);
        Assert.Equal("Pending", f.Choose(key, "UseLibrary", f.List).Status);
        Assert.Equal("AccessPending", f.Run(key).Status);
        Assert.Single(f.Posts, p => p == "_api/web/lists");
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
        Assert.Contains(
            f.Service.Rows.Values,
            r =>
                r.LogicalName == "asx_attempt"
                && r.GetAttributeValue<string>("asx_payload")
                    .Contains("RecoveryUseLibrary:" + f.List.ToString("D"))
        );
    }

    [Fact]
    public void AListThatIsNotACandidateIsRefused()
    {
        var f = New();
        var key = f.Lost();
        f.Run(key);
        var error = Assert.Throws<EvaluationBlockedException>(() =>
            f.Choose(key, "UseLibrary", Guid.NewGuid())
        );
        Assert.Equal("Choose one of the libraries Documents found.", error.Message);
        Assert.Equal("RecoveryRequired", f.Status(key));
    }

    [Fact]
    public void AStaleRowVersionIsRefused()
    {
        var f = New();
        var key = f.Lost();
        f.Run(key);
        var stale = f.Worker.Inspect(key).RowVersion;
        Assert.Equal(
            "Reconciling",
            f.Service.Transaction(() =>
                new CatalogAdministration(f.Service).Execute(
                    new CatalogRequest { Command = "RecheckSetup", Key = key },
                    true
                )
            ).Status
        );
        f.Run(key);
        var error = Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() =>
                new CatalogAdministration(f.Service).Execute(
                    new CatalogRequest
                    {
                        Command = "ResolveSetup",
                        Key = key,
                        Choice = "UseLibrary",
                        ListId = f.List,
                        RowVersion = stale,
                    },
                    true
                )
            )
        );
        Assert.Equal(
            "This setup changed. Look at its latest finding and choose again.",
            error.Message
        );
    }

    [Fact]
    public void NotFoundOffersCreateAgainWhichCreatesOnce()
    {
        var f = New();
        var key = f.Lost();
        f.Exists = false;
        f.Run(key);
        Assert.Equal("NotFound", f.Finding(key).State);
        Assert.Equal("Pending", f.Choose(key, "CreateAgain").Status);
        Assert.Equal("AccessPending", f.Run(key).Status);
        Assert.Equal(2, f.Posts.Count(p => p == "_api/web/lists"));
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    [Fact]
    public void LookupWaitsForTheLeaseAndCreateAgainRereadsByTitle()
    {
        var f = new LibraryProvisioningTests.Fixture { UnknownCreate = true };
        var key = f.Queue().Key;
        Assert.Equal("Quarantined", f.Run(key).Status);
        var waiting = f.Store.Require<LibrarySetup>("asx_operation", key).Value;
        Assert.Equal("Reconciling", waiting.Status);
        Assert.Equal("Checking", waiting.Recovery!.State);
        // The run that sent the create still held a live lease: the first read waits for it to end.
        Assert.True(waiting.NextAttemptUtc > DateTime.UtcNow.AddMinutes(4));
        Assert.DoesNotContain(key, f.Store.Pending("asx_operation"));
        Assert.Contains(key, f.Store.Pending("asx_operation", now: DateTime.UtcNow.AddMinutes(6)));
        // The create had not landed when Documents looked.
        f.UnknownCreate = false;
        f.Exists = false;
        f.Due(key);
        f.Run(key);
        Assert.Equal("NotFound", f.Finding(key).State);
        // It lands afterwards. Create it again reads the title first and stops instead of creating.
        f.Exists = true;
        Assert.Equal("Pending", f.Choose(key, "CreateAgain").Status);
        var after = f.Run(key);
        Assert.Equal("Blocked", after.Status);
        Assert.Contains(
            "A library with this name already exists. Add the existing library instead.",
            after.Notices
        );
        Assert.Single(f.Posts, p => p == "_api/web/lists");
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    [Fact]
    public void ATitleHitWithAnotherAddressIsAmbiguousAndOffersNoCreate()
    {
        var f = New();
        var key = f.Lost();
        var moved = f.Made();
        moved.Root = new FolderObservation { Id = f.Root, Path = "/sites/test/Documents1" };
        f.TitleHit = moved;
        f.Exists = false;
        f.Run(key);
        var finding = f.Finding(key);
        Assert.Equal("Ambiguous", finding.State);
        Assert.Equal("A library has this name but a different address.", finding.Reason);
        Assert.Equal(new[] { "UseCandidate", "CheckAgain", "Cancel" }, finding.Choices);
        var refused = Assert.Throws<EvaluationBlockedException>(() => f.Choose(key, "CreateAgain"));
        Assert.Equal(
            "SharePoint has a library with this name, so Documents won't create another. Use it, or cancel the setup.",
            refused.Message
        );
        Assert.Equal("Pending", f.Choose(key, "UseLibrary", f.List).Status);
        f.Exists = true;
        Assert.Equal("AccessPending", f.Run(key).Status);
        Assert.Single(f.Posts, p => p == "_api/web/lists");
    }

    [Fact]
    public void ACandidateTheCatalogGivesToAnotherSiteCannotBeUsed()
    {
        var f = New();
        var key = f.Lost();
        var other = Guid.NewGuid();
        f.Service.Seed(
            new Entity("asx_library", SiteIdentity.LibraryId(other, f.List))
            {
                ["asx_name"] = "Documents",
                ["asx_siteid"] = new EntityReference("asx_site", other),
                ["asx_listid"] = f.List.ToString("D"),
                ["statecode"] = new OptionSetValue(0),
            }
        );
        f.Run(key);
        var finding = f.Finding(key);
        Assert.Equal("Ambiguous", finding.State);
        Assert.Equal("Conflict", Assert.Single(finding.Candidates).CatalogEntry);
        Assert.DoesNotContain("UseCandidate", finding.Choices);
        Assert.Contains(
            "already has a different entry",
            Assert
                .Throws<EvaluationBlockedException>(() => f.Choose(key, "UseLibrary", f.List))
                .Message
        );
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnEntryAnAdminAddedMeanwhileIsAdopted(bool withAccess)
    {
        var f = New();
        var key = f.Lost();
        var row = SiteIdentity.LibraryId(f.Site, f.List);
        f.Service.Seed(
            new Entity("asx_library", row)
            {
                ["asx_name"] = "Documents",
                ["asx_siteid"] = new EntityReference("asx_site", f.Site),
                ["asx_listid"] = f.List.ToString("D"),
                ["asx_entryid"] = f.Root.ToString("D"),
                ["asx_entryurl"] = "https://example.sharepoint.com/sites/test/Documents",
                ["asx_approved"] = true,
                ["statecode"] = new OptionSetValue(0),
            }
        );
        if (withAccess)
            f.Store.Create(
                "asx_policy",
                new PolicyDocument
                {
                    Key = "policy:" + row.ToString("N"),
                    Status = "Applied",
                    LibraryId = row,
                }
            );
        f.Run(key);
        Assert.Equal("ThisSite", Assert.Single(f.Finding(key).Candidates).CatalogEntry);
        Assert.Equal("Found", f.Finding(key).State);
        f.Choose(key, "UseLibrary", f.List);
        var done = f.Run(key);
        Assert.Equal(withAccess ? "Applied" : "AccessPending", done.Status);
        Assert.Equal(
            row,
            Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_library").Id
        );
        Assert.Single(f.Posts, p => p == "_api/web/lists");
    }

    [Fact]
    public void CheckAgainLooksAgain()
    {
        var f = New();
        var key = f.Lost();
        f.Exists = false;
        f.Run(key);
        Assert.Equal("NotFound", f.Finding(key).State);
        f.Exists = true;
        var again = f.Service.Transaction(() =>
            new CatalogAdministration(f.Service).Execute(
                new CatalogRequest { Command = "RecheckSetup", Key = key },
                true
            )
        );
        Assert.Equal("Reconciling", again.Status);
        Assert.Equal("Checking", again.Recovery!.State);
        f.Run(key);
        Assert.Equal("Found", f.Finding(key).State);
    }

    [Fact]
    public void ATemporaryReadFailureWaitsWithReadsOnly()
    {
        var f = New();
        var key = f.Lost();
        f.LookupStatus = 503;
        Assert.Equal("RetryWait", f.Run(key).Status);
        var op = f.Store.Require<LibrarySetup>("asx_operation", key).Value;
        Assert.Equal("Reconciling", op.Status);
        Assert.Equal(1, op.RetryCount);
        Assert.Contains("HTTP 503", op.ErrorCode);
        Assert.True(op.Reconcile);
        Assert.DoesNotContain(key, f.Store.Pending("asx_operation"));
        f.Due(key);
        f.Run(key);
        Assert.Equal("Found", f.Finding(key).State);
        Assert.Single(f.Posts, p => p == "_api/web/lists");
    }

    [Fact]
    public void ARefusedReadIsAmbiguousWithItsCause()
    {
        var f = New();
        var key = f.Lost();
        f.LookupStatus = 403;
        Assert.Equal("RecoveryRequired", f.Run(key).Status);
        var finding = f.Finding(key);
        Assert.Equal("Ambiguous", finding.State);
        Assert.StartsWith("Documents couldn't read SharePoint: ", finding.Reason);
        Assert.Equal(new[] { "CheckAgain", "Cancel" }, finding.Choices);
        SiteFree(f, key);
    }

    [Fact]
    public void ReconcileReadOnASuspendedSiteIsNotUnsent()
    {
        var f = New();
        var key = f.Lost();
        f.Service.Rows[f.Site]["asx_approved"] = false;
        var read = f.Call("Claim", new WorkerResult { Key = key });
        Assert.Equal("ReconcileTitle", read.ProbeKind);
        // The connection admits the read: a lookup is never stopped like a write.
        Assert.Equal("Permit", f.Permit(read).Status);
        var setup = f.Store.Require<LibrarySetup>("asx_operation", key).Value;
        Assert.True(setup.ExternalSubmitted);
        Assert.False(setup.ExternalResponseKnown);
        Assert.Equal("CreateLibrary", setup.Mutation);
        var next = f.Call("Observe", read, LibraryProvisioningTests.Fixture.Body(f.Made()), 200);
        Assert.Equal("ReconcileUrl", next.ProbeKind);
        var settled = f.Call("Observe", next, LibraryProvisioningTests.Fixture.Body(f.Made()), 200);
        Assert.Equal("RecoveryRequired", settled.Status);
        Assert.Single(f.Posts, p => p == "_api/web/lists");
    }

    [Fact]
    public void CancelDuringTheLookupDeletesNothing()
    {
        var f = New();
        var key = f.Lost();
        var cancelled = f.Manage("Cancel", key);
        Assert.Equal("Cancelled", cancelled.Status);
        Assert.False(f.Store.Require<LibrarySetup>("asx_operation", key).Value.Reconcile);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
        Assert.DoesNotContain(f.Posts, p => p.Contains("delete"));
    }

    [Fact]
    public void RetryOfALostCreateIsRefusedWithTheWayOut()
    {
        var f = New();
        var key = f.Lost();
        var error = Assert.Throws<EvaluationBlockedException>(() => f.Manage("Retry", key));
        Assert.Equal(
            "Documents is checking SharePoint for this library. Use the choice it offers on the setup, or cancel the setup. Cancel deletes nothing in SharePoint.",
            error.Message
        );
    }

    [Fact]
    public void ASetupFrom0103IsPickedUpWhenInspected()
    {
        var f = New();
        var key = f.Lost();
        var legacy = f.Store.Require<LibrarySetup>("asx_operation", key);
        legacy.Value.Status = "RecoveryRequired";
        f.Store.Save(legacy);
        LegacyPayload.Strip(
            f.Service,
            "asx_operation",
            "Reconcile",
            "Recovery",
            "AbsentUtc",
            "TitleHit"
        );
        var inspected = f.Worker.Inspect(key);
        Assert.Equal("Reconciling", inspected.Status);
        Assert.Equal("Checking", inspected.Recovery!.State);
        Assert.Contains(key, f.Store.Pending("asx_operation"));
        // Without AbsentUtc the request time is the setup row's createdon.
        f.Run(key);
        Assert.Equal("Found", f.Finding(key).State);
    }
}
