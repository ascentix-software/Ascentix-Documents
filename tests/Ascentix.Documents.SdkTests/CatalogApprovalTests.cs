using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class CatalogApprovalTests
{
    [Fact]
    public void SiteAndLibraryRequireFreshReadOnlyObservationsBeforeApproval()
    {
        var f = new Fixture();
        var site = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "ProbeSite",
                NativeSiteId = f.NativeSite,
                RequestId = Guid.NewGuid(),
            },
            true
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "Approve",
                    Key = site.Key,
                    RowVersion = site.RowVersion,
                    Name = "Site",
                },
                true
            )
        );
        site = f.Capture(site.Key);
        var approved = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "Approve",
                Key = site.Key,
                RowVersion = site.RowVersion,
                Name = "Site",
            },
            true
        );
        Assert.NotEqual(Guid.Empty, approved.CatalogId);
        var library = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "ProbeLibrary",
                SiteId = approved.CatalogId,
                ListId = f.List,
                NativeParentId = f.NativeParent,
                RequestId = Guid.NewGuid(),
            },
            true
        );
        library = f.Capture(library.Key);
        Assert.Equal(f.Entry, library.Observation!.EntryId);
        Assert.Equal(2, library.Observation.Roles.Length);
        var result = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "Approve",
                Key = library.Key,
                RowVersion = library.RowVersion,
                Name = "Documents",
            },
            true
        );
        Assert.True(f.Service.Rows[result.CatalogId].GetAttributeValue<bool>("asx_approved"));
        Assert.False(f.Service.Rows[result.CatalogId].GetAttributeValue<bool>("asx_policyapplied"));
        // The library ACL is neither read nor recorded: approval never gates on entries
        // Documents does not own.
        Assert.False(f.Service.Rows[result.CatalogId].Contains("asx_aclhash"));
        Assert.DoesNotContain(f.Reads, r => r.Contains("roleassignments"));
    }

    [Fact]
    public void ApprovalProceedsWhileAnotherRunWritesOnTheSite()
    {
        var f = new Fixture();
        var probe = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "ProbeSite",
                NativeSiteId = f.NativeSite,
                RequestId = Guid.NewGuid(),
            },
            true
        );
        probe = f.Capture(probe.Key);
        var store = new DocumentStore(f.Service);
        var writer = WorkCoordination.Operation(f.Service, probe.Key);
        AdminStopTests.HoldWriter(store, writer, f.Now);
        var approved = f.Service.Transaction(() =>
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "Approve",
                    Key = probe.Key,
                    RowVersion = probe.RowVersion,
                    Name = "Site",
                },
                true
            )
        );
        Assert.True(f.Service.Rows[approved.CatalogId].GetAttributeValue<bool>("asx_approved"));
        Assert.Equal(
            "other/run",
            store.Require<DispatcherDocument>("asx_claim", writer).Value.RunId
        );
    }

    [Fact]
    public void ExpiredObservationAndChangedNativeMappingCannotBeApproved()
    {
        var f = new Fixture();
        var probe = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "ProbeSite",
                NativeSiteId = f.NativeSite,
                RequestId = Guid.NewGuid(),
            },
            true
        );
        probe = f.Capture(probe.Key);
        f.Now = f.Now.AddMinutes(16);
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "Approve",
                    Key = probe.Key,
                    RowVersion = probe.RowVersion,
                    Name = "Site",
                },
                true
            )
        );
        f.Now = f.Now.AddMinutes(-16);
        f.Service.Rows[f.NativeSite]["absoluteurl"] = "https://example.sharepoint.com/sites/other";
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "Approve",
                    Key = probe.Key,
                    RowVersion = probe.RowVersion,
                    Name = "Site",
                },
                true
            )
        );
    }

    [Fact]
    public void PhysicalSiteMismatchBlocksCapture()
    {
        var f = new Fixture();
        var probe = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "ProbeSite",
                NativeSiteId = f.NativeSite,
                RequestId = Guid.NewGuid(),
            },
            true
        );
        var work = f.Worker.Execute(
            new WorkerRequest
            {
                Command = "Claim",
                Key = probe.Key,
                RunId = "probe/run",
            },
            true
        );
        var result = f.Worker.Execute(
            new WorkerRequest
            {
                Command = "Observe",
                Key = probe.Key,
                RunId = "probe/run",
                Token = work.Token,
                ProbeId = work.ProbeId,
                ProbeKind = work.ProbeKind,
                HttpStatus = 200,
                ResponseBody = Envelope(
                    new WebObservation
                    {
                        Id = f.Web,
                        Url = "https://example.sharepoint.com/sites/other",
                    }
                ),
            },
            true
        );
        Assert.Equal("Blocked", result.Status);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_site");
    }

    [Fact]
    public void RuntimeTransportGateRejectsUnapprovedSiteAndMutationMethods()
    {
        var f = new Fixture();
        RuntimeSeed.Seed(f.Service, Guid.NewGuid(), "account");
        var profile = RuntimeProfile.Read(f.Service);
        var intent = new WorkerResult
        {
            SiteUrl = "https://example.sharepoint.com/sites/proto",
            Http = new HttpIntent { RelativeUri = "_api/web", Method = "GET" },
        };
        profile.ValidateTransport(intent);
        intent.SiteUrl = "https://example.sharepoint.com/sites/other";
        Assert.Throws<EvaluationBlockedException>(() => profile.ValidateTransport(intent));
        intent.SiteUrl = "https://example.sharepoint.com/sites/proto";
        intent.Http.Method = "DELETE";
        Assert.Throws<EvaluationBlockedException>(() => profile.ValidateTransport(intent));
        Assert.Throws<EvaluationBlockedException>(() =>
            RuntimeProfile.ValidateHosts(new[] { "https://example.com/sites/proto" })
        );
    }

    [Fact]
    public void AddSiteValidatesAndApprovesAutomaticallyIncludingRecheck()
    {
        var f = new Fixture();
        for (int i = 0; i < 2; i++)
        {
            var queued = f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    Key = null!,
                    NativeSiteId = f.NativeSite,
                    Name = "Site",
                    RequestId = Guid.NewGuid(),
                },
                true
            );
            Assert.Equal("Pending", queued.Status);
            var result = f.Capture(queued.Key);
            Assert.Equal("Approved", result.Status);
            Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_site");
        }
    }

    [Fact]
    public void LibraryCaptureDoesNotRequestAnOwnershipColumn()
    {
        var f = new Fixture();
        var site = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    Key = null!,
                    NativeSiteId = f.NativeSite,
                    Name = "Site",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        var library = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddLibrary",
                    SiteId = site.CatalogId,
                    ListId = f.List,
                    NativeParentId = f.NativeParent,
                    Name = "General",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        Assert.Equal("Approved", library.Status);
        Assert.NotNull(f.Service.Rows[library.CatalogId].GetAttributeValue<string>("asx_readrole"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NestedNativeEntryIsApprovedWhateverItsFolderPermissions(bool unique)
    {
        var f = new Fixture();
        f.Nested = true;
        f.UniqueFolders = unique;
        var root = Guid.NewGuid();
        f.Service.Seed(
            new Entity("sharepointdocumentlocation", root)
            {
                ["relativeurl"] = "General",
                ["statecode"] = new OptionSetValue(0),
                ["servicetype"] = new OptionSetValue(0),
                ["parentsiteorlocation"] = new EntityReference("sharepointsite", f.NativeSite),
            }
        );
        f.Service.Rows[f.NativeParent]["relativeurl"] = "Archive/Entry";
        f.Service.Rows[f.NativeParent]["parentsiteorlocation"] = new EntityReference(
            "sharepointdocumentlocation",
            root
        );
        var site = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "ProbeSite",
                    NativeSiteId = f.NativeSite,
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        var approved = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "Approve",
                Key = site.Key,
                RowVersion = site.RowVersion,
                Name = "Site",
            },
            true
        );
        var probe = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "ProbeLibrary",
                    SiteId = approved.CatalogId,
                    ListId = f.List,
                    NativeParentId = f.NativeParent,
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        Assert.Equal("Captured", probe.Status);
        Assert.Equal(f.NestedEntry, probe.Observation!.EntryId);
        Assert.EndsWith("/General/Archive/Entry", probe.Observation.EntryUrl);
        Assert.Equal(4, f.AncestorReads);
        var library = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "Approve",
                Key = probe.Key,
                RowVersion = probe.RowVersion,
                Name = "Library",
            },
            true
        );
        Assert.True(f.Service.Rows[library.CatalogId].GetAttributeValue<bool>("asx_approved"));
    }

    [Fact]
    public void DiscoveryIsReadOnlyPagedAndAutomaticLibraryAdditionCreatesNativeNavigation()
    {
        var f = new Fixture();
        var site = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    Key = null!,
                    NativeSiteId = f.NativeSite,
                    Name = "Site",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        var request = new CatalogRequest
        {
            Command = "DiscoverLibraries",
            Key = null!,
            SiteId = site.CatalogId,
            RequestId = Guid.NewGuid(),
        };
        var first = f.Capture(f.Admin.Execute(request, true).Key);
        Assert.Equal("Discovered", first.Status);
        Assert.Single(first.Observation!.Libraries);
        Assert.NotNull(first.Observation.NextLibraries);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "NextLibraries",
                    Key = first.Key,
                    RowVersion = "stale",
                },
                true
            )
        );
        var next = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "NextLibraries",
                    Key = first.Key,
                    RowVersion = first.RowVersion,
                },
                true
            ).Key
        );
        Assert.Null(next.Observation!.NextLibraries);
        f.Service.Rows.Remove(f.NativeParent);
        var added = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddLibrary",
                    SiteId = site.CatalogId,
                    ListId = f.List,
                    Name = "General",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        Assert.Equal("Approved", added.Status);
        var native = f
            .Service.Rows[added.CatalogId]
            .GetAttributeValue<EntityReference>("asx_nativeparentid");
        Assert.Equal("General", f.Service.Rows[native.Id].GetAttributeValue<string>("relativeurl"));
        Assert.False(f.Service.Rows[added.CatalogId].GetAttributeValue<bool>("asx_policyapplied"));
    }

    [Fact]
    public void ThrottledCatalogReadsWaitUntilRetryAfterWithNoAttemptCap()
    {
        var f = new Fixture();
        var queued = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "AddSite",
                Key = null!,
                NativeSiteId = f.NativeSite,
                Name = "Site",
                RequestId = Guid.NewGuid(),
            },
            true
        );
        for (int i = 0; i < 8; i++)
        {
            var work = f.Worker.Execute(
                new WorkerRequest
                {
                    Command = "Claim",
                    Key = queued.Key,
                    RunId = "run",
                },
                true
            );
            Assert.Equal("Read", work.Status);
            var result = f.Worker.Execute(
                new WorkerRequest
                {
                    Command = "Observe",
                    Key = queued.Key,
                    RunId = "run",
                    Token = work.Token,
                    ProbeId = work.ProbeId,
                    ProbeKind = work.ProbeKind,
                    HttpStatus = 429,
                    RetryAfter = "120",
                },
                true
            );
            Assert.Equal("RetryWait", result.Status);
            Assert.Contains(
                "attempt " + (i + 1) + ".",
                new DocumentStore(f.Service)
                    .Require<CatalogProbe>("asx_operation", queued.Key)
                    .Value.ErrorCode
            );
            Assert.Equal(
                result.Status,
                f.Worker.Execute(
                    new WorkerRequest
                    {
                        Command = "Claim",
                        Key = queued.Key,
                        RunId = "early",
                    },
                    true
                ).Status
            );
            f.Now = f.Now.AddHours(1);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(408)]
    public void CatalogReadTimeoutWaitsAndReleasesTheWriter(int status)
    {
        var f = new Fixture();
        var queued = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "AddSite",
                Key = null!,
                NativeSiteId = f.NativeSite,
                Name = "Site",
                RequestId = Guid.NewGuid(),
            },
            true
        );
        var work = f.Worker.Execute(
            new WorkerRequest
            {
                Command = "Claim",
                Key = queued.Key,
                RunId = "run",
            },
            true
        );
        var result = f.Worker.Execute(
            new WorkerRequest
            {
                Command = "Observe",
                Key = queued.Key,
                RunId = "run",
                Token = work.Token,
                ProbeId = work.ProbeId,
                ProbeKind = work.ProbeKind,
                HttpStatus = status,
            },
            true
        );
        Assert.Equal("RetryWait", result.Status);
        Assert.Null(
            new DocumentStore(f.Service)
                .Require<DispatcherDocument>(
                    "asx_claim",
                    WorkCoordination.Operation(f.Service, queued.Key)
                )
                .Value.RunId
        );
    }

    [Theory]
    [InlineData("Retry", "Pending")]
    [InlineData("Cancel", "Cancelled")]
    public void ProbeRetryAndCancelWorkOnceTheClaimExpires(string command, string status)
    {
        var f = new Fixture();
        var queued = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "AddSite",
                Key = null!,
                NativeSiteId = f.NativeSite,
                Name = "Site",
                RequestId = Guid.NewGuid(),
            },
            true
        );
        var work = f.Worker.Execute(
            new WorkerRequest
            {
                Command = "Claim",
                Key = queued.Key,
                RunId = "run",
            },
            true
        );
        Assert.Equal("Read", work.Status);
        var manage = new WorkerRequest { Command = command, Key = queued.Key };
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() => f.Worker.Execute(manage, true))
        );
        f.Now = f.Now.AddMinutes(6);
        Assert.Equal(status, f.Service.Transaction(() => f.Worker.Execute(manage, true)).Status);
        Assert.Null(
            new DocumentStore(f.Service)
                .Require<DispatcherDocument>(
                    "asx_claim",
                    WorkCoordination.Operation(f.Service, queued.Key)
                )
                .Value.RunId
        );
    }

    private static string Envelope<T>(T value) =>
        JsonWire.Write(new ODataEnvelope<T> { Data = value });

    [Fact]
    public void SameWebAndListIdsInDifferentCollectionsHaveSeparateCatalogAndNavigation()
    {
        var f = new Fixture();
        var first = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    NativeSiteId = f.NativeSite,
                    Name = "First",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        f.Service.Rows.Remove(f.NativeParent);
        var firstLibrary = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddLibrary",
                    SiteId = first.CatalogId,
                    ListId = f.List,
                    Name = "General",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        f.NativeSite = Guid.NewGuid();
        f.Collection = Guid.NewGuid();
        f.Url = "https://example.sharepoint.com/sites/second";
        f.Service.Seed(
            new Entity("sharepointsite", f.NativeSite)
            {
                ["absoluteurl"] = f.Url,
                ["statecode"] = new OptionSetValue(0),
            }
        );
        var second = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    NativeSiteId = f.NativeSite,
                    Name = "Second",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        var secondLibrary = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddLibrary",
                    SiteId = second.CatalogId,
                    ListId = f.List,
                    Name = "General",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        Assert.NotEqual(first.CatalogId, second.CatalogId);
        Assert.NotEqual(firstLibrary.CatalogId, secondLibrary.CatalogId);
        Assert.NotEqual(
            f.Service.Rows[firstLibrary.CatalogId]
                .GetAttributeValue<EntityReference>("asx_nativeparentid")
                .Id,
            f.Service.Rows[secondLibrary.CatalogId]
                .GetAttributeValue<EntityReference>("asx_nativeparentid")
                .Id
        );
        var repeat = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    NativeSiteId = f.NativeSite,
                    Name = "Second",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        Assert.Equal(second.CatalogId, repeat.CatalogId);
    }

    [Fact]
    public void ExplicitIdentityUpgradePreservesExistingRecordAndRequiresFreshMatchingObservation()
    {
        var f = new Fixture();
        var existing = Guid.NewGuid();
        f.Service.Seed(
            new Entity("asx_site", existing)
            {
                ["asx_nativeid"] = new EntityReference("sharepointsite", f.NativeSite),
                ["asx_webid"] = f.Web.ToString("D"),
                ["asx_url"] = f.Url,
                ["asx_approved"] = true,
            }
        );
        var probe = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "ProbeSite",
                    NativeSiteId = f.NativeSite,
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        var upgraded = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "CompleteSiteIdentity",
                CatalogId = existing,
                Key = probe.Key,
                RowVersion = probe.RowVersion,
            },
            true
        );
        Assert.Equal(existing, upgraded.CatalogId);
        Assert.Equal(
            SiteIdentity.Key(f.Url, f.Collection, f.Web),
            f.Service.Rows[existing].GetAttributeValue<string>("asx_identity")
        );
        var repeat = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    NativeSiteId = f.NativeSite,
                    Name = "Site",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        Assert.Equal(existing, repeat.CatalogId);
    }

    [Fact]
    public void ChangedCollectionBlocksLibraryDiscoveryBeforeReadingLibraries()
    {
        var f = new Fixture();
        var site = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    NativeSiteId = f.NativeSite,
                    Name = "Site",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        var queued = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "DiscoverLibraries",
                SiteId = site.CatalogId,
                RequestId = Guid.NewGuid(),
            },
            true
        );
        f.Collection = Guid.NewGuid();
        var result = f.Capture(queued.Key);
        Assert.Equal("Blocked", result.Status);
        Assert.Contains("collection identity changed", result.Issue);
    }

    [Fact]
    public void ApprovalFailureRecordsItsReasonAndReleasesTheLease()
    {
        var f = new Fixture();
        var queued = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "AddSite",
                NativeSiteId = f.NativeSite,
                Name = "Site",
                RequestId = Guid.NewGuid(),
            },
            true
        );
        var work = f.Worker.Execute(
            new WorkerRequest
            {
                Command = "Claim",
                Key = queued.Key,
                RunId = "test",
            },
            true
        );
        work = f.Worker.Execute(
            new WorkerRequest
            {
                Command = "Observe",
                Key = queued.Key,
                RunId = "test",
                Token = work.Token,
                ProbeId = work.ProbeId,
                ProbeKind = work.ProbeKind,
                HttpStatus = 200,
                ResponseBody = Envelope(new WebObservation { Id = f.Web, Url = f.Url }),
            },
            true
        );
        work = f.Worker.Execute(
            new WorkerRequest
            {
                Command = "Observe",
                Key = queued.Key,
                RunId = "test",
                Token = work.Token,
                ProbeId = work.ProbeId,
                ProbeKind = work.ProbeKind,
                HttpStatus = 200,
                ResponseBody = Envelope(new WebObservation { Id = f.Collection }),
            },
            true
        );
        // A separate catalog row with this physical identity must not be silently reassigned.
        f.Service.Seed(
            new Entity("asx_site", Guid.NewGuid())
            {
                ["asx_identity"] = SiteIdentity.Key(f.Url, f.Collection, f.Web),
                ["asx_nativeid"] = new EntityReference("sharepointsite", Guid.NewGuid()),
            }
        );
        var result = f.Worker.Execute(
            new WorkerRequest
            {
                Command = "Complete",
                Key = queued.Key,
                RunId = "test",
                Token = work.Token,
            },
            true
        );
        Assert.Equal("Blocked", result.Status);
        Assert.Contains(
            "no reassignment",
            f.Admin.Execute(
                new CatalogRequest { Command = "Inspect", Key = queued.Key },
                true
            ).Issue
        );
        Assert.Equal(
            "Pending",
            f.Worker.Execute(new WorkerRequest { Command = "Retry", Key = queued.Key }, true).Status
        );
    }

    private const string InheritanceWarning =
        "This library inherits permissions from the site. When you approve it, Documents stops the inheritance, keeps a copy of the current site permissions, and then manages team access on it.";

    [Fact]
    public void InheritingLibraryIsApprovedOnlyWithTheAdminsAcknowledgement()
    {
        var f = new Fixture();
        f.Unique = false;
        var site = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    NativeSiteId = f.NativeSite,
                    Name = "Site",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        var probe = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "ProbeLibrary",
                    SiteId = site.CatalogId,
                    ListId = f.List,
                    NativeParentId = f.NativeParent,
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        Assert.Equal("Captured", probe.Status);
        Assert.True(probe.Observation!.Inherits, "The probe records that the library inherits");
        var refused = Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() =>
                f.Admin.Execute(
                    new CatalogRequest
                    {
                        Command = "Approve",
                        Key = probe.Key,
                        RowVersion = probe.RowVersion,
                        Name = "General",
                    },
                    true
                )
            )
        );
        Assert.StartsWith(InheritanceWarning, refused.Message);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
        var approved = f.Admin.Execute(
            new CatalogRequest
            {
                Command = "Approve",
                Key = probe.Key,
                RowVersion = probe.RowVersion,
                Name = "General",
                BreakInheritance = true,
            },
            true
        );
        Assert.True(f.Service.Rows[approved.CatalogId].GetAttributeValue<bool>("asx_approved"));
        var policy = new DocumentStore(f.Service)
            .Require<PolicyDocument>("asx_policy", "policy:" + approved.CatalogId.ToString("N"))
            .Value;
        Assert.True(policy.BreakInheritance, "Approval records the admin's consent");
        Assert.Equal("Missing", policy.Status);
        Assert.Equal(System.Guid.Empty, policy.Generation);
    }

    [Fact]
    public void AddingAnInheritingLibraryCarriesTheAcknowledgementIntoAutomaticApproval()
    {
        var f = new Fixture();
        f.Unique = false;
        var site = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    NativeSiteId = f.NativeSite,
                    Name = "Site",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        var discovery = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "DiscoverLibraries",
                    SiteId = site.CatalogId,
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        Assert.Contains("HasUniqueRoleAssignments", f.Reads.Last());
        Assert.False(discovery.Observation!.Libraries.Single().Unique);
        var blocked = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddLibrary",
                    SiteId = site.CatalogId,
                    ListId = f.List,
                    NativeParentId = f.NativeParent,
                    Name = "General",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key,
            blocked: true
        );
        Assert.Equal("Blocked", blocked.Status);
        Assert.StartsWith(InheritanceWarning, blocked.Issue);
        Assert.True(blocked.Observation!.Inherits);
        var added = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddLibrary",
                    SiteId = site.CatalogId,
                    ListId = f.List,
                    NativeParentId = f.NativeParent,
                    Name = "General",
                    RequestId = Guid.NewGuid(),
                    BreakInheritance = true,
                },
                true
            ).Key
        );
        Assert.Equal("Approved", added.Status);
        Assert.True(
            new DocumentStore(f.Service)
                .Require<PolicyDocument>("asx_policy", "policy:" + added.CatalogId.ToString("N"))
                .Value.BreakInheritance
        );
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReapprovingASuspendedLibraryProceedsUnlessItsAccessRunHasAWriteOutstanding(
        bool outstanding
    )
    {
        var f = new Fixture();
        var added = Added(f);
        Assert.Equal("Approved", added.Status);
        var store = new DocumentStore(f.Service);
        // An access run is queued for the library and was held while it was suspended.
        string run = "policywork:" + Guid.NewGuid().ToString("N");
        store.Create(
            "asx_operation",
            new SecurityOperation
            {
                Key = run,
                LibraryId = added.CatalogId,
                Status = outstanding ? "ExternalUnknown" : "RetryWait",
                ExternalSubmitted = outstanding,
            }
        );
        string policyKey = "policy:" + added.CatalogId.ToString("N");
        var policy = store.Find<PolicyDocument>("asx_policy", policyKey);
        if (policy == null)
            store.Create(
                "asx_policy",
                new PolicyDocument
                {
                    Key = policyKey,
                    Status = "Queued",
                    LibraryId = added.CatalogId,
                    OperationKey = run,
                }
            );
        else
        {
            policy.Value.OperationKey = run;
            store.Save(policy);
        }
        f.Service.Transaction(() =>
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "SuspendLibrary",
                    CatalogId = added.CatalogId,
                    CatalogRowVersion = f.Service.Rows[added.CatalogId].RowVersion,
                },
                true
            )
        );
        var library = f.Service.Rows[added.CatalogId];
        var probe = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "ProbeLibrary",
                    SiteId = library.GetAttributeValue<EntityReference>("asx_siteid").Id,
                    ListId = f.List,
                    NativeParentId = library
                        .GetAttributeValue<EntityReference>("asx_nativeparentid")
                        .Id,
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        CatalogResult Approve() =>
            f.Service.Transaction(() =>
                f.Admin.Execute(
                    new CatalogRequest
                    {
                        Command = "Approve",
                        Key = probe.Key,
                        RowVersion = probe.RowVersion,
                        CatalogRowVersion = f.Service.Rows[added.CatalogId].RowVersion,
                        Name = "General",
                    },
                    true
                )
            );
        if (outstanding)
        {
            var refused = Assert.Throws<EvaluationBlockedException>(() => Approve());
            Assert.Contains("write to SharePoint whose result is not known yet", refused.Message);
            Assert.False(f.Service.Rows[added.CatalogId].GetAttributeValue<bool>("asx_approved"));
            return;
        }
        Assert.Equal("Approved", Approve().Status);
        Assert.True(f.Service.Rows[added.CatalogId].GetAttributeValue<bool>("asx_approved"));
        // The held access run is still queued and resumes by itself; nothing was cancelled.
        Assert.Equal(
            run,
            store.Require<PolicyDocument>("asx_policy", policyKey).Value.OperationKey
        );
        Assert.Equal(
            "RetryWait",
            store.Require<SecurityOperation>("asx_operation", run).Value.Status
        );
    }

    private static CatalogResult Added(Fixture f)
    {
        var site = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    NativeSiteId = f.NativeSite,
                    Name = "Site",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        f.Service.Rows.Remove(f.NativeParent);
        return f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddLibrary",
                    SiteId = site.CatalogId,
                    ListId = f.List,
                    Name = "General",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
    }

    private static CatalogResult Repoint(
        Fixture f,
        string command,
        Guid id,
        bool blocked = false
    ) =>
        f.Capture(
            f.Service.Transaction(() =>
                f.Admin.Execute(
                    new CatalogRequest
                    {
                        Command = command,
                        CatalogId = Runtime(f, id),
                        RequestId = Guid.NewGuid(),
                    },
                    true
                )
            ).Key,
            blocked
        );

    /// <summary>Seeds the runtime profile (allowed hosts) once.</summary>
    private static Guid Runtime(Fixture f, Guid id)
    {
        if (!f.Service.Rows.Values.Any(r => r.LogicalName == "asx_runtime"))
            RuntimeSeed.Seed(f.Service, Guid.NewGuid());
        return id;
    }

    [Fact]
    public void RepointNamesAHostOrSiteRecordDocumentsCannotCall()
    {
        var f = new Fixture();
        var library = Added(f);
        var siteId = f
            .Service.Rows[library.CatalogId]
            .GetAttributeValue<EntityReference>("asx_siteid")
            .Id;
        Runtime(f, siteId);
        CatalogRequest Request() =>
            new CatalogRequest
            {
                Command = "RepointSite",
                CatalogId = siteId,
                RequestId = Guid.NewGuid(),
            };
        // The site moved to another tenant host that the Runtime panel does not allow.
        f.Service.Rows[f.NativeSite]["absoluteurl"] = "https://contoso.sharepoint.com/sites/proto";
        var host = Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() => f.Admin.Execute(Request(), true))
        );
        Assert.Equal(
            "The host contoso.sharepoint.com is not one of the SharePoint hosts allowed in the Runtime panel. Add it there, then re-point again.",
            host.Message
        );
        // A failed transaction restores copies of the rows, so the row is read again here.
        f.Service.Rows.Values.Single(r => r.LogicalName == "asx_runtime")["asx_sharepointhosts"] =
            "[\"example.sharepoint.com\",\"contoso.sharepoint.com\"]";
        // The site record's text differs from the address Documents calls.
        f.Service.Rows[f.NativeSite]["absoluteurl"] =
            "https://contoso.sharepoint.com/sites/New Site";
        var record = Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() => f.Admin.Execute(Request(), true))
        );
        Assert.Equal(
            "No active SharePoint site record in Dataverse has the address https://contoso.sharepoint.com/sites/New%20Site. Set the site's SharePoint site record in Dataverse to exactly that address, then re-point again.",
            record.Message
        );
        Assert.DoesNotContain(
            f.Service.Rows.Values,
            r => r.GetAttributeValue<string>("asx_workkind") == "Repoint"
        );
        f.Service.Rows[f.NativeSite]["absoluteurl"] = "https://contoso.sharepoint.com/sites/proto";
        Assert.Equal(
            "Pending",
            f.Service.Transaction(() => f.Admin.Execute(Request(), true)).Status
        );
    }

    private static Entity Location(Fixture f, Guid library) =>
        f.Service.Rows[
            f.Service.Rows[library].GetAttributeValue<EntityReference>("asx_nativeparentid").Id
        ];

    [Fact]
    public void RepointFollowsARenamedLibraryByItsListId()
    {
        var f = new Fixture();
        var library = Added(f);
        var before = f.Service.Rows[library.CatalogId].GetAttributeValue<string>("asx_entryurl");
        Assert.Equal(
            "General",
            Location(f, library.CatalogId).GetAttributeValue<string>("relativeurl")
        );
        f.LibraryName = "Shared Documents";
        f.Reads.Clear();
        var result = Repoint(f, "RepointLibrary", library.CatalogId);
        Assert.Equal("Approved", result.Status);
        Assert.Contains(f.Reads, r => r.Contains("lists(guid'" + f.List));
        string after = f.Url + "/Shared Documents";
        Assert.Equal(
            after,
            f.Service.Rows[library.CatalogId].GetAttributeValue<string>("asx_entryurl")
        );
        Assert.Equal(
            "Shared Documents",
            Location(f, library.CatalogId).GetAttributeValue<string>("relativeurl")
        );
        Assert.Contains(before + " → " + after, result.Observation!.Changes);
        Assert.Equal(
            f.List.ToString("D"),
            f.Service.Rows[library.CatalogId].GetAttributeValue<string>("asx_listid")
        );
        Assert.All(f.Reads, r => Assert.DoesNotContain("roleassignments", r));
    }

    [Fact]
    public void RepointThatCannotFollowTheLocationChainWritesNothing()
    {
        var f = new Fixture();
        var library = Added(f);
        // The library's document location now hangs under another site record.
        var other = Guid.NewGuid();
        f.Service.Seed(
            new Entity("sharepointsite", other)
            {
                ["absoluteurl"] = "https://example.sharepoint.com/sites/other",
                ["statecode"] = new OptionSetValue(0),
            }
        );
        var location = Location(f, library.CatalogId);
        location["parentsiteorlocation"] = new EntityReference("sharepointsite", other);
        var libraryVersion = f.Service.Rows[library.CatalogId].RowVersion;
        var locationVersion = location.RowVersion;
        f.LibraryName = "Shared Documents";
        var result = Repoint(f, "RepointLibrary", library.CatalogId, blocked: true);
        Assert.Equal("Blocked", result.Status);
        Assert.Equal("Native site differs from approval.", result.Issue);
        Assert.Equal(libraryVersion, f.Service.Rows[library.CatalogId].RowVersion);
        Assert.Equal(
            f.Url + "/General",
            f.Service.Rows[library.CatalogId].GetAttributeValue<string>("asx_entryurl")
        );
        Assert.Equal(locationVersion, Location(f, library.CatalogId).RowVersion);
        Assert.Equal(
            "General",
            Location(f, library.CatalogId).GetAttributeValue<string>("relativeurl")
        );
    }

    [Fact]
    public void RepointOfALibraryThatNoLongerExistsSaysSoAndChangesNothing()
    {
        var f = new Fixture();
        var library = Added(f);
        var row = f.Service.Rows[library.CatalogId].RowVersion;
        f.ListGone = true;
        var result = Repoint(f, "RepointLibrary", library.CatalogId);
        Assert.Equal("Blocked", result.Status);
        Assert.Equal(
            "This library no longer exists on the site. Remove it, or register the new library.",
            result.Issue
        );
        Assert.Equal(row, f.Service.Rows[library.CatalogId].RowVersion);
    }

    [Fact]
    public void RepointFollowsARenamedEntryFolderByItsUniqueId()
    {
        var f = new Fixture();
        f.Nested = true;
        var root = Guid.NewGuid();
        f.Service.Seed(
            new Entity("sharepointdocumentlocation", root)
            {
                ["relativeurl"] = "General",
                ["statecode"] = new OptionSetValue(0),
                ["servicetype"] = new OptionSetValue(0),
                ["parentsiteorlocation"] = new EntityReference("sharepointsite", f.NativeSite),
            }
        );
        f.Service.Rows[f.NativeParent]["relativeurl"] = "Archive/Entry";
        f.Service.Rows[f.NativeParent]["parentsiteorlocation"] = new EntityReference(
            "sharepointdocumentlocation",
            root
        );
        var site = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    NativeSiteId = f.NativeSite,
                    Name = "Site",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        var library = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddLibrary",
                    SiteId = site.CatalogId,
                    ListId = f.List,
                    NativeParentId = f.NativeParent,
                    Name = "General",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        Assert.Equal("Approved", library.Status);
        f.EntryPath = "/sites/proto/General/Archive/Renamed";
        var result = Repoint(f, "RepointLibrary", library.CatalogId);
        Assert.Equal("Approved", result.Status);
        Assert.Equal(
            f.Url + "/General/Archive/Renamed",
            f.Service.Rows[library.CatalogId].GetAttributeValue<string>("asx_entryurl")
        );
        Assert.Equal(
            "Archive/Renamed",
            f.Service.Rows[f.NativeParent].GetAttributeValue<string>("relativeurl")
        );
        Assert.Equal("General", f.Service.Rows[root].GetAttributeValue<string>("relativeurl"));
        // A deleted entry folder is reported and never recreated.
        f.EntryGone = true;
        var gone = Repoint(f, "RepointLibrary", library.CatalogId);
        Assert.Equal("Approved", gone.Status);
        Assert.Contains(
            gone.Observation!.Changes,
            c => c.Contains("no longer exists") && c.Contains("did not recreate")
        );
        Assert.Equal(
            f.Url + "/General/Archive/Renamed",
            f.Service.Rows[library.CatalogId].GetAttributeValue<string>("asx_entryurl")
        );
    }

    [Fact]
    public void RepointFollowsAMovedSiteAndItsLibrariesByWebId()
    {
        var f = new Fixture();
        var library = Added(f);
        var siteId = f
            .Service.Rows[library.CatalogId]
            .GetAttributeValue<EntityReference>("asx_siteid")
            .Id;
        // Re-pointing a library first is refused while its site's address changed.
        f.Url = "https://example.sharepoint.com/sites/renamed";
        f.Service.Rows[f.NativeSite]["absoluteurl"] = f.Url;
        var early = Assert.Throws<EvaluationBlockedException>(() =>
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "RepointLibrary",
                    CatalogId = library.CatalogId,
                    RequestId = Guid.NewGuid(),
                },
                true
            )
        );
        Assert.Equal("The site's address changed. Re-point the site first.", early.Message);
        var result = Repoint(f, "RepointSite", siteId);
        Assert.Equal("Approved", result.Status);
        var site = f.Service.Rows[siteId];
        Assert.Equal(f.Url, site.GetAttributeValue<string>("asx_url"));
        Assert.Equal(
            SiteIdentity.Key(f.Url, f.Collection, f.Web),
            site.GetAttributeValue<string>("asx_identity")
        );
        Assert.Equal(f.Web.ToString("D"), site.GetAttributeValue<string>("asx_webid"));
        Assert.Equal(
            f.Url + "/General",
            f.Service.Rows[library.CatalogId].GetAttributeValue<string>("asx_entryurl")
        );
        Assert.Contains(
            "https://example.sharepoint.com/sites/proto → " + f.Url,
            result.Observation!.Changes
        );
        // The library's location chain ends at the site record, so it resolves to the new address.
        new NativeLocations(f.Service).ValidateParent(
            Location(f, library.CatalogId).Id,
            f.Url + "/General",
            f.NativeSite
        );
        // A site reporting another address asks the admin to update its Dataverse record.
        f.Service.Rows[f.NativeSite]["absoluteurl"] = "https://example.sharepoint.com/sites/stale";
        var stale = Repoint(f, "RepointSite", siteId);
        Assert.Equal("Blocked", stale.Status);
        Assert.Contains("SharePoint reports this site at " + f.Url, stale.Issue);
        Assert.Equal(f.Url, f.Service.Rows[siteId].GetAttributeValue<string>("asx_url"));
    }

    private static CatalogResult Remove(Fixture f, string command, Guid id) =>
        f.Service.Transaction(() =>
            f.Admin.Execute(new CatalogRequest { Command = command, CatalogId = id }, true)
        );

    /// <summary>Seeds a template revision whose destination uses the library.</summary>
    private static Guid UseIn(Fixture f, Guid library, string template, string status, bool current)
    {
        var templateId = DocumentStore.StableId("template:" + template);
        var revision = Guid.NewGuid();
        if (!f.Service.Rows.ContainsKey(templateId))
            f.Service.Seed(new Entity("asx_template", templateId) { ["asx_name"] = template });
        if (current)
            f.Service.Rows[templateId]["asx_publishedrevisionid"] = new EntityReference(
                "asx_revision",
                revision
            );
        f.Service.Seed(
            new Entity("asx_revision", revision)
            {
                ["asx_templateid"] = new EntityReference("asx_template", templateId),
                ["asx_status"] = status,
            }
        );
        f.Service.Seed(
            new Entity("asx_destination", Guid.NewGuid())
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", revision),
                ["asx_key"] = "general",
                ["asx_libraryid"] = new EntityReference("asx_library", library),
            }
        );
        return revision;
    }

    [Fact]
    public void RemoveIsRefusedWhileADraftOrPublishedTemplateUsesTheLibrary()
    {
        var f = new Fixture();
        var library = Added(f);
        UseIn(f, library.CatalogId, "Accounts", "Published", current: true);
        UseIn(f, library.CatalogId, "Contracts", "Draft", current: false);
        var version = f.Service.Rows[library.CatalogId].RowVersion;
        var refused = Assert.Throws<EvaluationBlockedException>(() =>
            Remove(f, "RemoveLibrary", library.CatalogId)
        );
        Assert.Equal(
            "Used by template 'Accounts' (published). Used by template 'Contracts' (draft). Change the templates first.",
            refused.Message
        );
        Assert.Equal(version, f.Service.Rows[library.CatalogId].RowVersion);
    }

    [Fact]
    public void RemoveReadsEveryDestinationRowAndEveryLibraryHoweverMany()
    {
        var f = new Fixture();
        var library = Added(f);
        // 5,200 rows of an older revision come first; the one that uses the library now is
        // last, past the first 5,000 rows.
        var older = UseIn(f, library.CatalogId, "Archive", "Published", current: false);
        for (int i = 0; i < 5199; i++)
            f.Service.Seed(
                new Entity("asx_destination", Guid.NewGuid())
                {
                    ["asx_revisionid"] = new EntityReference("asx_revision", older),
                    ["asx_key"] = "s" + i,
                    ["asx_libraryid"] = new EntityReference("asx_library", library.CatalogId),
                }
            );
        UseIn(f, library.CatalogId, "Accounts", "Published", current: true);
        var refused = Assert.Throws<EvaluationBlockedException>(() =>
            Remove(f, "RemoveLibrary", library.CatalogId)
        );
        Assert.Equal(
            "Used by template 'Accounts' (published). Change the template first.",
            refused.Message
        );
        // A site with more than 5,000 removed libraries is still removed (kept for history).
        var siteId = f
            .Service.Rows[library.CatalogId]
            .GetAttributeValue<EntityReference>("asx_siteid")
            .Id;
        f.Service.Rows[library.CatalogId]["statecode"] = new OptionSetValue(1);
        for (int i = 0; i < 5001; i++)
            f.Service.Seed(
                new Entity("asx_library", Guid.NewGuid())
                {
                    ["asx_name"] = "Old " + i,
                    ["asx_siteid"] = new EntityReference("asx_site", siteId),
                    ["statecode"] = new OptionSetValue(1),
                }
            );
        Assert.Equal("Removed", Remove(f, "RemoveSite", siteId).Status);
    }

    [Fact]
    public void LibraryUsedOnlyByAnOlderRevisionIsKeptRemovedForHistory()
    {
        var f = new Fixture();
        var library = Added(f);
        UseIn(f, library.CatalogId, "Accounts", "Published", current: false);
        var operations = f.Service.Rows.Values.Count(r => r.LogicalName == "asx_operation");
        var result = Remove(f, "RemoveLibrary", library.CatalogId);
        Assert.Equal("Removed", result.Status);
        Assert.Contains(
            result.Notices,
            n => n.StartsWith("Nothing was deleted or changed in SharePoint")
        );
        var row = f.Service.Rows[library.CatalogId];
        Assert.False(row.GetAttributeValue<bool>("asx_approved"));
        Assert.Equal(1, row.GetAttributeValue<OptionSetValue>("statecode").Value);
        Assert.True(new WorkerCatalog(f.Service).Removed(library.CatalogId));
        // Removal reads and writes only Dataverse: no SharePoint work is queued.
        Assert.Equal(
            operations,
            f.Service.Rows.Values.Count(r => r.LogicalName == "asx_operation")
        );
        // The site still has the removed library, so it is kept Removed too.
        var siteId = row.GetAttributeValue<EntityReference>("asx_siteid").Id;
        Assert.Equal("Removed", Remove(f, "RemoveSite", siteId).Status);
        Assert.Equal(
            1,
            f.Service.Rows[siteId].GetAttributeValue<OptionSetValue>("statecode").Value
        );
        // Adding the library again makes it active.
        var again = f.Capture(
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "AddSite",
                    NativeSiteId = f.NativeSite,
                    Name = "Site",
                    RequestId = Guid.NewGuid(),
                },
                true
            ).Key
        );
        Assert.Equal(siteId, again.CatalogId);
        Assert.Equal(
            0,
            f.Service.Rows[siteId].GetAttributeValue<OptionSetValue>("statecode").Value
        );
        Assert.True(f.Service.Rows[siteId].GetAttributeValue<bool>("asx_approved"));
    }

    [Fact]
    public void UnreferencedLibraryAndThenItsSiteAreDeleted()
    {
        var f = new Fixture();
        var library = Added(f);
        var siteId = f
            .Service.Rows[library.CatalogId]
            .GetAttributeValue<EntityReference>("asx_siteid")
            .Id;
        var refused = Assert.Throws<EvaluationBlockedException>(() =>
            Remove(f, "RemoveSite", siteId)
        );
        Assert.Equal("Remove the site's libraries first: General.", refused.Message);
        var location = f
            .Service.Rows[library.CatalogId]
            .GetAttributeValue<EntityReference>("asx_nativeparentid")
            .Id;
        Assert.Equal("Deleted", Remove(f, "RemoveLibrary", library.CatalogId).Status);
        Assert.False(f.Service.Rows.ContainsKey(library.CatalogId));
        // The library's Dataverse document location is navigation others may use; it stays.
        Assert.True(f.Service.Rows.ContainsKey(location));
        Assert.Equal("Deleted", Remove(f, "RemoveSite", siteId).Status);
        Assert.False(f.Service.Rows.ContainsKey(siteId));
    }

    [Fact]
    public void LibraryWithRecordFoldersIsKeptRemoved()
    {
        var f = new Fixture();
        var library = Added(f);
        var location = f
            .Service.Rows[library.CatalogId]
            .GetAttributeValue<EntityReference>("asx_nativeparentid")
            .Id;
        f.Service.Seed(
            new Entity("sharepointdocumentlocation", Guid.NewGuid())
            {
                ["relativeurl"] = "Example",
                ["description"] = "AscentixDocuments:abc",
                ["parentsiteorlocation"] = new EntityReference(
                    "sharepointdocumentlocation",
                    location
                ),
            }
        );
        Assert.Equal("Removed", Remove(f, "RemoveLibrary", library.CatalogId).Status);
        Assert.True(f.Service.Rows.ContainsKey(library.CatalogId));
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Admin.Execute(
                new CatalogRequest
                {
                    Command = "RepointLibrary",
                    CatalogId = library.CatalogId,
                    RequestId = Guid.NewGuid(),
                },
                true
            )
        );
    }

    [Fact]
    public void SiteIdentityIncludesHostnameAndRejectsIncompleteIdentity()
    {
        var web = Guid.NewGuid();
        var collection = Guid.NewGuid();
        Assert.NotEqual(
            SiteIdentity.Key("https://a.sharepoint.com/sites/test", collection, web),
            SiteIdentity.Key("https://b.sharepoint.com/sites/test", collection, web)
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            SiteIdentity.Key("https://a.sharepoint.com", Guid.Empty, web)
        );
    }

    private sealed class Fixture
    {
        public DurableWorkerTests.MemoryService Service = new DurableWorkerTests.MemoryService();
        public DateTime Now = new DateTime(2026, 9, 8, 12, 0, 0, DateTimeKind.Utc);
        public string Url = "https://example.sharepoint.com/sites/proto";
        public bool Nested,
            UniqueFolders;

        /// <summary>Whether the library has its own permissions or inherits the site's.</summary>
        public bool Unique = true;

        /// <summary>The library's URL segment in SharePoint; Re-point follows a rename.</summary>
        public string LibraryName = "General";

        /// <summary>The library or the nested entry folder no longer exists.</summary>
        public bool ListGone,
            EntryGone;

        /// <summary>Where SharePoint reports the nested entry folder (by its unique ID).</summary>
        public string EntryPath = "/sites/proto/General/Archive/Entry";
        public int AncestorReads;
        public List<string> Reads = new List<string>();
        public Guid NestedEntry = Guid.NewGuid();
        public Guid Collection = Guid.NewGuid();
        public Guid NativeSite = Guid.NewGuid(),
            NativeParent = Guid.NewGuid(),
            Web = Guid.NewGuid(),
            List = Guid.NewGuid(),
            Entry = Guid.NewGuid();
        public CatalogAdministration Admin => new CatalogAdministration(Service, () => Now);
        public CatalogWorker Worker => new CatalogWorker(Service, () => Now);

        public Fixture()
        {
            Service.Seed(
                new Entity("sharepointsite", NativeSite)
                {
                    ["absoluteurl"] = Url,
                    ["statecode"] = new OptionSetValue(0),
                }
            );
            Service.Seed(
                new Entity("sharepointdocumentlocation", NativeParent)
                {
                    ["relativeurl"] = "General",
                    ["statecode"] = new OptionSetValue(0),
                    ["servicetype"] = new OptionSetValue(0),
                    ["parentsiteorlocation"] = new EntityReference("sharepointsite", NativeSite),
                }
            );
        }

        /// <summary>
        /// Drives the probe as the flow does. Completion must succeed unless <paramref name="blocked"/>:
        /// a refused automatic approval, or a re-point whose step failed, rolled back and was
        /// reported with Fail by the flow's failure branch.
        /// </summary>
        public CatalogResult Capture(string key, bool blocked = false)
        {
            var work = Worker.Execute(
                new WorkerRequest
                {
                    Command = "Claim",
                    Key = key,
                    RunId = "probe/run",
                },
                true
            );
            for (int i = 0; i < 20; i++)
            {
                if (work.Status == "Verified")
                {
                    var token = work.Token;
                    try
                    {
                        work = Service.Transaction(() =>
                            Worker.Execute(
                                new WorkerRequest
                                {
                                    Command = "Complete",
                                    Key = key,
                                    RunId = "probe/run",
                                    Token = token,
                                },
                                true
                            )
                        );
                    }
                    catch (EvaluationBlockedException error)
                    {
                        work = Service.Transaction(() =>
                            Worker.Execute(
                                new WorkerRequest
                                {
                                    Command = "Fail",
                                    Key = key,
                                    RunId = "probe/run",
                                    Token = token,
                                    StatusCode = 400,
                                    ErrorCode = "0x80040265",
                                    Error = error.Message,
                                },
                                true
                            )
                        );
                    }
                    var inspected = Admin.Execute(
                        new CatalogRequest { Command = "Inspect", Key = key },
                        true
                    );
                    if (blocked)
                        Assert.Equal("Blocked", work.Status);
                    else
                        Assert.True(
                            new[] { "Captured", "Approved", "Discovered" }.Contains(work.Status),
                            work.Status + ": " + inspected.Issue
                        );
                    return inspected;
                }
                if (work.Status == "Blocked")
                    return Admin.Execute(
                        new CatalogRequest { Command = "Inspect", Key = key },
                        true
                    );
                Assert.Equal("Read", work.Status);
                Assert.Equal("GET", work.Http!.Method);
                Reads.Add(work.Http.RelativeUri);
                Assert.DoesNotContain("AsxdWorkKey", work.Http.RelativeUri);
                string body;
                int status = 200;
                switch (work.ProbeKind)
                {
                    case "CatalogLibraries":
                        body = Envelope(
                            new ODataRows<LibraryChoice>
                            {
                                Rows = new[]
                                {
                                    new LibraryChoice
                                    {
                                        Id = List,
                                        Title = "General",
                                        BaseTemplate = 101,
                                        Unique = Unique,
                                    },
                                },
                                Next = work.Http.RelativeUri.Contains("skiptoken")
                                    ? null
                                    : "https://example.sharepoint.com/sites/proto/"
                                        + work.Http.RelativeUri
                                        + "&$skiptoken=second",
                            }
                        );
                        break;
                    case "CatalogCollection":
                        body = Envelope(new WebObservation { Id = Collection });
                        break;
                    case "CatalogWeb":
                        body = Envelope(new WebObservation { Id = Web, Url = Url });
                        break;
                    case "CatalogLibrary":
                    case "CatalogFinal":
                    case "RepointLibrary":
                        if (work.ProbeKind == "RepointLibrary" && ListGone)
                        {
                            status = 404;
                            body =
                                "{\"error\":{\"code\":\"-1\",\"message\":{\"value\":\"List does not exist.\"}}}";
                            break;
                        }
                        body = Envelope(
                            new CatalogLibraryObservation
                            {
                                Id = List,
                                Unique = Unique,
                                Root = new FolderObservation
                                {
                                    Id = Entry,
                                    Path = new Uri(Url).AbsolutePath + "/" + LibraryName,
                                },
                            }
                        );
                        break;
                    case "RepointEntry":
                        Assert.Contains(NestedEntry.ToString("D"), work.Http.RelativeUri);
                        if (EntryGone)
                        {
                            status = 404;
                            body = "{}";
                            break;
                        }
                        body = Envelope(
                            new FolderObservation { Id = NestedEntry, Path = EntryPath }
                        );
                        break;
                    case "CatalogEntry":
                    case "CatalogEntryFinal":
                    case "CatalogAncestor":
                    case "CatalogAncestorFinal":
                        bool entry = work.ProbeKind.StartsWith(
                            "CatalogEntry",
                            StringComparison.Ordinal
                        );
                        AncestorReads++;
                        // The entry's path is a query-string alias, never in the URL path.
                        Assert.StartsWith(
                            "_api/web/GetFolderByServerRelativePath(decodedUrl=@p)?@p='",
                            work.Http.RelativeUri
                        );
                        body = Envelope(
                            new FolderObservation
                            {
                                Id = entry ? NestedEntry : Guid.NewGuid(),
                                Path = entry
                                    ? "/sites/proto/General/Archive/Entry"
                                    : "/sites/proto/General/Archive",
                            }
                        );
                        if (UniqueFolders)
                            body = DurableWorkerTests.WithUniquePermissions(body);
                        break;
                    case "CatalogRoles":
                        body = Envelope(
                            new ODataRows<SecurityRoleObservation>
                            {
                                Rows = new[]
                                {
                                    new SecurityRoleObservation
                                    {
                                        Id = 2,
                                        Type = 2,
                                        Permissions = new PermissionMask
                                        {
                                            High = "176",
                                            Low = "138612833",
                                        },
                                    },
                                    new SecurityRoleObservation
                                    {
                                        Id = 3,
                                        Type = 3,
                                        Permissions = new PermissionMask
                                        {
                                            High = "432",
                                            Low = "1011028719",
                                        },
                                    },
                                },
                            }
                        );
                        break;
                    default:
                        throw new Exception(work.ProbeKind);
                }
                work = Worker.Execute(
                    new WorkerRequest
                    {
                        Command = "Observe",
                        Key = key,
                        RunId = "probe/run",
                        Token = work.Token,
                        ProbeId = work.ProbeId,
                        ProbeKind = work.ProbeKind,
                        HttpStatus = status,
                        ResponseBody = body,
                    },
                    true
                );
            }
            throw new Exception("Probe did not complete.");
        }
    }
}
