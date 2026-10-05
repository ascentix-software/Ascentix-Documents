using System;
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
        Assert.Equal(
            64,
            f.Service.Rows[result.CatalogId].GetAttributeValue<string>("asx_aclhash").Length
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
    public void NestedNativeEntryRequiresInheritedIntermediateAncestors(bool unique)
    {
        var f = new Fixture();
        f.Nested = true;
        f.UniqueAncestor = unique;
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
        if (unique)
        {
            Assert.Equal("Blocked", probe.Status);
            Assert.Throws<EvaluationBlockedException>(() =>
                f.Admin.Execute(
                    new CatalogRequest
                    {
                        Command = "Approve",
                        Key = probe.Key,
                        RowVersion = probe.RowVersion,
                        Name = "Library",
                    },
                    true
                )
            );
        }
        else
        {
            Assert.Equal("Captured", probe.Status);
            Assert.Equal(f.NestedEntry, probe.Observation!.EntryId);
            Assert.EndsWith("/General/Archive/Entry", probe.Observation.EntryUrl);
            Assert.Equal(4, f.AncestorReads);
        }
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
    public void ThrottledCatalogReadsWaitUntilRetryAfterAndStopAfterFiveRetries()
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
        for (int i = 0; i < 6; i++)
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
            Assert.Equal(i == 5 ? "Blocked" : "RetryWait", result.Status);
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
            UniqueAncestor;
        public int AncestorReads;
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

        public CatalogResult Capture(string key)
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
                    work = Worker.Execute(
                        new WorkerRequest
                        {
                            Command = "Complete",
                            Key = key,
                            RunId = "probe/run",
                            Token = work.Token,
                        },
                        true
                    );
                    Assert.Contains(work.Status, new[] { "Captured", "Approved", "Discovered" });
                    return Admin.Execute(
                        new CatalogRequest { Command = "Inspect", Key = key },
                        true
                    );
                }
                if (work.Status == "Blocked")
                    return Admin.Execute(
                        new CatalogRequest { Command = "Inspect", Key = key },
                        true
                    );
                Assert.Equal("Read", work.Status);
                Assert.Equal("GET", work.Http!.Method);
                Assert.DoesNotContain("AsxdWorkKey", work.Http.RelativeUri);
                string body;
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
                        body = Envelope(
                            new CatalogLibraryObservation
                            {
                                Id = List,
                                Unique = true,
                                Root = new FolderObservation
                                {
                                    Id = Entry,
                                    Path = new Uri(Url).AbsolutePath + "/General",
                                },
                            }
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
                        body = Envelope(
                            new FolderObservation
                            {
                                Id = entry ? NestedEntry : Guid.NewGuid(),
                                Path = entry
                                    ? "/sites/proto/General/Archive/Entry"
                                    : "/sites/proto/General/Archive",
                                Item = new ParentListItem
                                {
                                    UniquePermissions = !entry && UniqueAncestor,
                                },
                            }
                        );
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
                    case "CatalogAcl":
                        body = Envelope(
                            new ODataRows<AclAssignment> { Rows = Array.Empty<AclAssignment>() }
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
                        HttpStatus = 200,
                        ResponseBody = body,
                    },
                    true
                );
            }
            throw new Exception("Probe did not complete.");
        }
    }
}
