using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class LibraryProvisioningTests
{
    [Theory]
    [InlineData(2, 0, null, false)]
    [InlineData(3, 2, null, false)]
    [InlineData(3, 3, "only the guests", true)]
    [InlineData(1, 0, "Access teams", true)]
    [InlineData(3, 1, "The group's guests will also have access to this library.", false)]
    [InlineData(
        2,
        2,
        "All members of the group will have access to this library, not only its owners.",
        false
    )]
    [InlineData(3, 1, null, true)]
    [InlineData(2, 2, null, true)]
    public void NewLibraryAcceptsGroupTeamsAndRefusesTeamsSharePointCannotIdentify(
        int type,
        int membership,
        string? reason,
        bool acknowledged
    )
    {
        var f = new Fixture();
        var team = f.Service.Rows[f.Team];
        team["name"] = "Operations";
        team["teamtype"] = new Microsoft.Xrm.Sdk.OptionSetValue(type);
        team["membershiptype"] = new Microsoft.Xrm.Sdk.OptionSetValue(membership);
        team["azureactivedirectoryobjectid"] = Guid.NewGuid();
        f.Acknowledge = acknowledged;
        if (reason == null)
        {
            var queued = f.Queue();
            Assert.Equal("Pending", queued.Status);
            // The consent travels with the setup to the access policy it applies later.
            Assert.Equal("AccessPending", f.Run(queued.Key).Status);
            return;
        }
        Assert.Contains(
            reason,
            Assert
                .Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() => f.Queue())
                .Message
        );
    }

    [Fact]
    public void ASetupSavedMidWholeListReadReadsTheOwnersGroupInstead()
    {
        var f = new Fixture { LegacyAcl = true };
        var queued = f.Queue();
        Assert.Equal("AccessPending", f.Run(queued.Key).Status);
        Assert.False(f.LegacyAcl);
        Assert.NotEmpty(f.OwnerGrantReads);
        Assert.Contains(
            f.Posts,
            p =>
                p.EndsWith("addroleassignment(principalid=7,roledefid=5)", StringComparison.Ordinal)
        );
    }

    [Fact]
    public void CreatesLibraryBoundaryAndOwnerGrantThenQueuesInitialTeamSync()
    {
        var f = new Fixture();
        var queued = f.Queue();
        var result = f.Run(queued.Key);
        Assert.Equal("AccessPending", result.Status);
        Assert.Equal(3, f.Posts.Count);
        Assert.Contains(f.Posts, p => p == "_api/web/lists");
        Assert.Contains(
            f.Posts,
            p =>
                p.EndsWith(
                    "breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)",
                    StringComparison.Ordinal
                )
        );
        Assert.Contains(
            f.Posts,
            p =>
                p.EndsWith("addroleassignment(principalid=7,roledefid=5)", StringComparison.Ordinal)
        );
        // Only the owners group's own assignment is read, never the library's whole list.
        Assert.NotEmpty(f.OwnerGrantReads);
        Assert.All(
            f.OwnerGrantReads,
            read => Assert.Contains("/roleassignments/getbyprincipalid(7)?", read)
        );
        var catalog = Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
        Assert.True(catalog.GetAttributeValue<bool>("asx_approved"));
        Assert.False(catalog.GetAttributeValue<bool>("asx_policyapplied"));
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "sharepointdocumentlocation");
        var registration = f.Store.Require<TeamRegistration>(
            "asx_teamregistration",
            "team:" + f.Team.ToString("N")
        );
        Assert.True(registration.Value.Enabled);
        var policy = f.Store.Require<PolicyDocument>(
            "asx_policy",
            "policy:" + catalog.Id.ToString("N")
        );
        Assert.Equal(f.Team, Assert.Single(policy.Value.Queued).TeamId);
        Assert.NotNull(policy.Value.OperationKey);
        var operation = f.Store.Require<SecurityOperation>(
            "asx_operation",
            policy.Value.OperationKey!
        );
        operation.Value.Status = "Applied";
        f.Store.Save(operation);
        Assert.Equal("Ready", f.Worker.Inspect(queued.Key).Status);
        Assert.Equal(queued.Key, f.Queue().Key);
        Assert.Equal(3, f.Posts.Count);
    }

    [Fact]
    public void ExistingSameNameLibraryIsNeverChangedByCreate()
    {
        var f = new Fixture { Exists = true };
        var result = f.Run(f.Queue().Key);
        Assert.Equal("Blocked", result.Status);
        Assert.Empty(f.Posts);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    [Fact]
    public void ThrottledSetupReadsKeepWaitingWithNoAttemptCap()
    {
        var f = new Fixture();
        var queued = f.Queue();
        for (int attempt = 1; attempt <= 7; attempt++)
        {
            var work = f.Call("Claim", new WorkerResult { Key = queued.Key });
            Assert.Equal("Read", work.Status);
            Assert.Equal("RetryWait", f.Call("Observe", work, "{}", 503).Status);
            var op = f.Store.Require<LibrarySetup>("asx_operation", queued.Key);
            Assert.Equal(attempt, op.Value.RetryCount);
            Assert.Equal(
                "Waiting to retry after a temporary error (HTTP 503); attempt " + attempt + ".",
                op.Value.ErrorCode
            );
            op.Value.NextAttemptUtc = DateTime.UtcNow.AddSeconds(-1);
            f.Store.Save(op);
        }
    }

    [Fact]
    public void TransientSetupDriveFailureWaitsAndGenuineOneBlocks()
    {
        var f = new Fixture();
        var queued = f.Queue();
        var work = f.Call("Claim", new WorkerResult { Key = queued.Key });
        var fail = new WorkerRequest
        {
            Command = "Fail",
            Key = queued.Key,
            RunId = "library-test",
            Token = work.Token,
            StatusCode = 429,
        };
        Assert.Equal("RetryWait", f.Service.Transaction(() => f.Worker.Execute(fail, true)).Status);
        var op = f.Store.Require<LibrarySetup>("asx_operation", queued.Key);
        op.Value.NextAttemptUtc = DateTime.UtcNow.AddSeconds(-1);
        f.Store.Save(op);
        work = f.Call("Claim", new WorkerResult { Key = queued.Key });
        fail.Token = work.Token;
        fail.StatusCode = 400;
        fail.ErrorCode = "0x80040265";
        Assert.Equal("Blocked", f.Service.Transaction(() => f.Worker.Execute(fail, true)).Status);
    }

    [Fact]
    public void ThrottledLibraryCreateWasNotExecutedSoItWaitsAndCreatesOnce()
    {
        var f = new Fixture { ThrottledCreates = 1 };
        var queued = f.Queue();
        Assert.Equal("RetryWait", f.Run(queued.Key).Status);
        var op = f.Store.Require<LibrarySetup>("asx_operation", queued.Key);
        Assert.False(op.Value.ExternalSubmitted);
        Assert.False(op.Value.ExternalResponseKnown);
        Assert.Equal(Guid.Empty, op.Value.ListId);
        Assert.Contains("HTTP 429", op.Value.ErrorCode);
        Assert.Null(
            f.Store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(f.Service, queued.Key)
            ).Value.RunId
        );
        op.Value.NextAttemptUtc = DateTime.UtcNow.AddSeconds(-1);
        f.Store.Save(op);
        Assert.Equal("AccessPending", f.Run(queued.Key).Status);
        Assert.Equal(2, f.Posts.Count(p => p == "_api/web/lists"));
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    [Fact]
    public void WaitingSetupSavedBy0103RetriesAndCreatesTheLibrary()
    {
        var f = new Fixture { ThrottledCreates = 1 };
        var queued = f.Queue();
        Assert.Equal("RetryWait", f.Run(queued.Key).Status);
        // 0.1.0.3 setups had no consent flag and none of the newer operation markers.
        Assert.True(
            LegacyPayload.Strip(
                f.Service,
                "asx_operation",
                "AcknowledgeBroaderAccess",
                "EntryPath",
                "Reprobe",
                "NameConflict"
            ) > 0
        );
        Assert.Equal("Pending", f.Manage("Retry", queued.Key).Status);
        Assert.Equal("AccessPending", f.Run(queued.Key).Status);
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    /// <summary>Queues a setup and runs it until its create is prepared; the run then stops.</summary>
    private static (Fixture F, string Key, WorkerResult Prepared) PreparedCreate()
    {
        var f = new Fixture { StopBeforePost = true };
        var key = f.Queue().Key;
        var prepared = f.Run(key);
        Assert.Equal("Create", prepared.Status);
        f.StopBeforePost = false;
        return (f, key, prepared);
    }

    private static void CreatedOnce(Fixture f, string key)
    {
        Assert.Equal("AccessPending", f.Run(key).Status);
        Assert.Single(f.Posts, p => p == "_api/web/lists");
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    [Fact]
    public void CreateNeverPermittedIsTakenOverAndReadAgainInsteadOfNeedingRecovery()
    {
        var (f, key, _) = PreparedCreate();
        f.Expire(key);
        // Nothing was sent, so the next run reads again and creates the library once.
        CreatedOnce(f, key);
    }

    [Fact]
    public void ClaimFailureOnACreateNeverPermittedWaitsAndThenCreates()
    {
        var (f, key, _) = PreparedCreate();
        f.Expire(key);
        var failed = f.Manage("FailUnclaimed", key);
        Assert.Equal("RetryWait", failed.Status);
        var op = f.Store.Require<LibrarySetup>("asx_operation", key);
        Assert.False(op.Value.ExternalSubmitted);
        op.Value.NextAttemptUtc = DateTime.UtcNow.AddSeconds(-1);
        f.Store.Save(op);
        CreatedOnce(f, key);
    }

    [Fact]
    public void FlowFailureBeforeThePermitWaitsInsteadOfNeedingRecovery()
    {
        var (f, key, prepared) = PreparedCreate();
        var failed = f.Service.Transaction(() =>
            f.Worker.Execute(
                new WorkerRequest
                {
                    Command = "Fail",
                    Key = key,
                    RunId = "library-test",
                    Token = prepared.Token,
                    StatusCode = 0,
                },
                true
            )
        );
        Assert.Equal("RetryWait", failed.Status);
        var op = f.Store.Require<LibrarySetup>("asx_operation", key);
        op.Value.NextAttemptUtc = DateTime.UtcNow.AddSeconds(-1);
        f.Store.Save(op);
        CreatedOnce(f, key);
    }

    [Fact]
    public void OperatorRetryOfACreateNeverPermittedReadsAgainAndCreatesOnce()
    {
        var (f, key, _) = PreparedCreate();
        // A run is still working on it: Retry waits for its claim to expire.
        Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            f.Manage("Retry", key)
        );
        f.Expire(key);
        Assert.Equal("Pending", f.Manage("Retry", key).Status);
        Assert.Null(f.Claim(key).RunId);
        CreatedOnce(f, key);
    }

    [Fact]
    public void CancelOfACreateNeverPermittedStopsAndTheNameCanBeCreatedAgain()
    {
        var (f, key, _) = PreparedCreate();
        f.Expire(key);
        Assert.Equal("Cancelled", f.Manage("Cancel", key).Status);
        Assert.Null(f.Claim(key).RunId);
        Assert.False(f.Claim(key).HttpOutstanding);
        Assert.Empty(f.Posts);
        // Creating the same library again starts over with fresh reads.
        Assert.Equal("Pending", f.Queue().Status);
        CreatedOnce(f, key);
    }

    [Fact]
    public void PermittedCreateWithAnUnknownOutcomeKeepsRecoveryAndCancelDeletesNothing()
    {
        var (f, key, prepared) = PreparedCreate();
        Assert.Equal("Permit", f.Permit(prepared).Status);
        // The POST may have been sent; the run stopped before its answer was recorded.
        f.Expire(key);
        Assert.Equal("Quarantined", f.Call("Claim", new WorkerResult { Key = key }).Status);
        Assert.Equal("RecoveryRequired", f.Status(key));
        var refused = Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            f.Manage("Retry", key)
        );
        Assert.Contains("original create response", refused.Message);
        Assert.Contains("Cancel", refused.Message);
        var cancelled = f.Manage("Cancel", key);
        Assert.Equal("Cancelled", cancelled.Status);
        Assert.Contains(cancelled.Notices, n => n.Contains("Nothing in SharePoint was deleted"));
        Assert.Null(f.Claim(key).RunId);
        Assert.False(f.Claim(key).HttpOutstanding);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
        Assert.DoesNotContain(f.Posts, p => p.Contains("delete"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SetupStoredBeforeTheMarkerCountsAsPossiblySentAndCanStillBeCancelled(bool permitted)
    {
        var (f, key, prepared) = PreparedCreate();
        if (permitted)
            Assert.Equal("Permit", f.Permit(prepared).Status);
        // An earlier release did not record the permit, and its answer handling cleared the
        // claim's outstanding flag even for a 5xx, so the claim cannot tell either: a stored
        // create counts as possibly sent and keeps evidence-based recovery.
        LegacyPayload.Strip(f.Service, "asx_operation", "WritePermitted");
        f.Expire(key);
        Assert.Equal("Quarantined", f.Call("Claim", new WorkerResult { Key = key }).Status);
        Assert.Equal("RecoveryRequired", f.Status(key));
        Assert.Contains(
            "original create response",
            Assert
                .Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
                    f.Manage("Retry", key)
                )
                .Message
        );
        Assert.Equal("Cancelled", f.Manage("Cancel", key).Status);
        Assert.Empty(f.Posts);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SetupForASuspendedSiteWaitsAndResumesAfterReapproval(bool prepared)
    {
        Fixture f;
        string key;
        if (prepared)
        {
            // Interrupted after preparing its create, before any write permit.
            (f, key, _) = PreparedCreate();
            f.Expire(key);
        }
        else
        {
            f = new Fixture();
            key = f.Queue().Key;
        }
        f.Service.Rows[f.Site]["asx_approved"] = false;
        Assert.Equal("RetryWait", f.Call("Claim", new WorkerResult { Key = key }).Status);
        var op = f.Store.Require<LibrarySetup>("asx_operation", key).Value;
        Assert.Equal("RetryWait", op.Status);
        Assert.Contains("suspended", op.ErrorCode);
        Assert.False(op.ExternalSubmitted);
        // It takes no dispatch slot until its next check is due.
        Assert.DoesNotContain(key, f.Store.Pending("asx_operation"));
        Assert.Contains(key, f.Store.Pending("asx_operation", now: DateTime.UtcNow.AddMinutes(16)));
        Assert.Null(
            f.Store.Find<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(f.Service, key)
            )?.Value.RunId
        );
        // Approved again, with no operator action, it resumes at its next check.
        f.Service.Rows[f.Site]["asx_approved"] = true;
        var due = f.Store.Require<LibrarySetup>("asx_operation", key);
        due.Value.NextAttemptUtc = DateTime.UtcNow.AddSeconds(-1);
        f.Store.Save(due);
        CreatedOnce(f, key);
    }

    [Fact]
    public void SetupOfASuspendedSiteWhoseWriterAnotherRunHoldsBacksOffInsteadOfStayingListed()
    {
        var f = new Fixture();
        var key = f.Queue().Key;
        var other = f.Service.Transaction(() =>
            f.Worker.Queue(
                new CatalogRequest
                {
                    SiteId = f.Site,
                    Name = "Other",
                    RequestId = Guid.NewGuid(),
                    Entries = Array.Empty<PolicyEntry>(),
                }
            )
        );
        Assert.Equal("Read", f.Call("Claim", new WorkerResult { Key = other.Key }).Status);
        f.Service.Rows[f.Site]["asx_approved"] = false;
        Assert.Equal("RetryWait", f.Call("Claim", new WorkerResult { Key = key }).Status);
        var op = f.Store.Require<LibrarySetup>("asx_operation", key).Value;
        Assert.Equal("RetryWait", op.Status);
        Assert.Contains("suspended", op.ErrorCode);
        Assert.True(op.NextAttemptUtc > DateTime.UtcNow);
        Assert.DoesNotContain(key, f.Store.Pending("asx_operation"));
        // The other run keeps its writer.
        Assert.Equal(other.Key, f.Claim(key).OperationKey);
    }

    private static CatalogResult Catalog(Fixture f, string command, string key) =>
        f.Service.Transaction(() =>
            new CatalogAdministration(f.Service).Execute(
                new CatalogRequest { Command = command, Key = key },
                true
            )
        );

    [Fact]
    public void SecurityAdministratorRetriesAndCancelsASetupThroughTheCatalog()
    {
        // A create never permitted is simply prepared again.
        var (f, key, _) = PreparedCreate();
        f.Expire(key);
        Assert.Equal("Pending", Catalog(f, "RetrySetup", key).Status);
        CreatedOnce(f, key);

        // A create that may have reached SharePoint keeps evidence-based recovery; Cancel works
        // and deletes nothing.
        var (g, sent, prepared) = PreparedCreate();
        Assert.Equal("Permit", g.Permit(prepared).Status);
        g.Expire(sent);
        Assert.Equal("Quarantined", g.Call("Claim", new WorkerResult { Key = sent }).Status);
        Assert.Equal("RecoveryRequired", Catalog(g, "Inspect", sent).Status);
        Assert.Contains(
            "original create response",
            Assert
                .Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
                    Catalog(g, "RetrySetup", sent)
                )
                .Message
        );
        var cancelled = Catalog(g, "CancelSetup", sent);
        Assert.Equal("Cancelled", cancelled.Status);
        Assert.Contains(cancelled.Notices, n => n.Contains("Nothing in SharePoint was deleted"));
        Assert.DoesNotContain(g.Service.Rows.Values, r => r.LogicalName == "asx_library");

        Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            Catalog(g, "CancelSetup", "catalogprobe:" + Guid.NewGuid().ToString("N"))
        );
    }

    [Fact]
    public void CancellingANewLibrarysFirstAccessRunEndsItsSetupAndKeepsTheLibrary()
    {
        var f = new Fixture();
        var key = f.Queue().Key;
        Assert.Equal("AccessPending", f.Run(key).Status);
        var access = f.Store.Require<LibrarySetup>("asx_operation", key).Value.PolicyOperation!;
        var done = Catalog(f, "CancelSetup", key);
        Assert.Equal("Ready", done.Status);
        Assert.Equal(SecurityWorker.SetupAccessCancelled, done.Issue);
        Assert.Equal(
            "Cancelled",
            f.Store.Require<SecurityOperation>("asx_operation", access).Value.Status
        );
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    [Fact]
    public void UnknownCreationIsQuarantinedWithoutDeletingOrAdoptingContent()
    {
        var f = new Fixture { UnknownCreate = true };
        var queued = f.Queue();
        Assert.Equal("Quarantined", f.Run(queued.Key).Status);
        Assert.Single(f.Posts);
        Assert.Equal("RecoveryRequired", f.Worker.Inspect(queued.Key).Status);
        // An expired lease does not release an unknown library creation: a second create
        // could make a duplicate library, so it keeps operator recovery with evidence.
        var lease = f.Store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(f.Service, queued.Key)
        );
        lease.Value.LeaseUntilUtc = DateTime.UtcNow.AddMinutes(-1);
        f.Store.Save(lease);
        Assert.Equal("Quarantined", f.Call("Claim", new WorkerResult { Key = queued.Key }).Status);
        Assert.DoesNotContain(
            queued.Key,
            f.Store.Pending("asx_operation", now: DateTime.UtcNow.AddHours(1))
        );
        Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            f.Worker.Execute(new WorkerRequest { Command = "Retry", Key = queued.Key }, true)
        );
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    [Fact]
    public void UnknownLibraryRecoveryRequiresOriginalResponseThenResumesWithoutRecreating()
    {
        var f = new Fixture { UnknownCreate = true };
        var queued = f.Queue();
        f.Run(queued.Key);
        var claim = f.Store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(f.Service, queued.Key)
        );
        claim.Value.LeaseUntilUtc = DateTime.UtcNow.AddMinutes(-1);
        f.Store.Save(claim);
        var request = new WorkerRequest
        {
            Key = queued.Key,
            RunId = claim.Value.RunId!,
            Token = claim.Value.Token,
            Evidence =
                "Original run terminated; successful create response recovered from run history.",
        };
        var recovery = new WorkerCoordinator(f.Service);
        Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            f.Service.Transaction(() => recovery.PermitRecovery(request, true))
        );
        request.ResponseBody = JsonWire.Write(
            new ODataEnvelope<CreatedLibrary>
            {
                Data = new CreatedLibrary { Id = f.List, Title = "Different" },
            }
        );
        Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            f.Service.Transaction(() => recovery.PermitRecovery(request, true))
        );
        request.ResponseBody = JsonWire.Write(
            new ODataEnvelope<CreatedLibrary>
            {
                Data = new CreatedLibrary { Id = f.List, Title = "Documents" },
            }
        );
        Assert.DoesNotContain(
            queued.Key,
            f.Store.Pending("asx_operation", now: DateTime.UtcNow.AddHours(1))
        );
        Assert.Equal(
            "RecoveryPermitted",
            f.Service.Transaction(() => recovery.PermitRecovery(request, true)).Status
        );
        // The permitted operation is listed again so the dispatcher resumes it.
        Assert.Contains(queued.Key, f.Store.Pending("asx_operation"));
        Assert.Equal("AccessPending", f.Run(queued.Key).Status);
        Assert.Single(f.Posts, p => p == "_api/web/lists");
        Assert.DoesNotContain(f.Posts, p => p.Contains("delete"));
    }

    private sealed class Fixture
    {
        public readonly DurableWorkerTests.MemoryService Service =
            new DurableWorkerTests.MemoryService();
        public DocumentStore Store => new DocumentStore(Service);
        public LibraryProvisioning Worker => new LibraryProvisioning(Service);
        public Guid Site = Guid.NewGuid(),
            Native = Guid.NewGuid(),
            Web = Guid.NewGuid(),
            List = Guid.NewGuid(),
            Root = Guid.NewGuid(),
            Team = Guid.NewGuid();
        public bool Exists,
            Unique,
            OwnerAccess,
            UnknownCreate;
        public int ThrottledCreates;
        public bool Acknowledge;
        public bool StopBeforePost;
        public List<string> Posts = new List<string>();

        /// <summary>The owners-group reads, by their request path.</summary>
        public List<string> OwnerGrantReads = new List<string>();

        /// <summary>Answers the first owners-group read as a 0.1.0.3 whole-list read.</summary>
        public bool LegacyAcl;

        public DispatcherDocument Claim(string key) =>
            Store
                .Require<DispatcherDocument>("asx_claim", WorkCoordination.Operation(Service, key))
                .Value;

        /// <summary>Lets the run's 5-minute claim expire.</summary>
        public void Expire(string key)
        {
            var claim = Store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(Service, key)
            );
            claim.Value.LeaseUntilUtc = DateTime.UtcNow.AddMinutes(-1);
            Store.Save(claim);
        }

        /// <summary>The shared connection's admission of the prepared write, as the flow asks.</summary>
        public WorkerResult Permit(WorkerResult work) =>
            Service.Transaction(() =>
                WorkCoordination.BeginHttp(
                    Service,
                    new WorkerRequest
                    {
                        Command = "BeginHttp",
                        Key = work.Key,
                        RunId = "library-test",
                        Token = work.Token,
                    },
                    DateTime.UtcNow
                )
            );

        public string Status(string key) =>
            Store.Require<LibrarySetup>("asx_operation", key).Value.Status;

        public Fixture()
        {
            Service.Seed(
                new Entity("asx_site", Site)
                {
                    ["asx_nativeid"] = new EntityReference("sharepointsite", Native),
                    ["asx_webid"] = Web.ToString("D"),
                    ["asx_url"] = "https://example.sharepoint.com/sites/test",
                    ["asx_approved"] = true,
                }
            );
            Service.Seed(
                new Entity("team", Team)
                {
                    ["teamtype"] = new OptionSetValue(0),
                    ["isdefault"] = false,
                }
            );
        }

        public CatalogResult Queue() =>
            Service.Transaction(() =>
                Worker.Queue(
                    new CatalogRequest
                    {
                        SiteId = Site,
                        Name = "Documents",
                        RequestId = Guid.NewGuid(),
                        AcknowledgeBroaderAccess = Acknowledge,
                        Entries = new[]
                        {
                            new PolicyEntry { TeamId = Team, Access = "Read" },
                        },
                    }
                )
            );

        private static string Body<T>(T data) =>
            JsonWire.Write(new ODataEnvelope<T> { Data = data });

        public WorkerResult Run(string key)
        {
            var work = Call("Claim", new WorkerResult { Key = key });
            for (int i = 0; i < 30; i++)
            {
                if (work.Status == "ReadyToCreate")
                {
                    work = Call("PrepareCreate", work);
                    continue;
                }
                if (work.Status == "Create")
                {
                    // The run stops after preparing the create, before any write permit.
                    if (StopBeforePost && work.Http!.RelativeUri == "_api/web/lists")
                        return work;
                    Posts.Add(work.Http!.RelativeUri);
                    int status = 200;
                    string body = "{}";
                    if (work.Http.RelativeUri == "_api/web/lists")
                    {
                        if (ThrottledCreates > 0)
                        {
                            // SharePoint refuses a throttled request without executing it.
                            ThrottledCreates--;
                            work = Call("CreateResponse", work, "{}", 429);
                            continue;
                        }
                        Exists = true;
                        status = UnknownCreate ? 0 : 201;
                        body = Body(new CreatedLibrary { Id = List, Title = "Documents" });
                    }
                    else if (work.Http.RelativeUri.Contains("breakroleinheritance"))
                        Unique = true;
                    else if (work.Http.RelativeUri.Contains("addroleassignment"))
                        OwnerAccess = true;
                    else
                        throw new Exception("Unexpected mutation");
                    work = Call("CreateResponse", work, body, status);
                    continue;
                }
                if (work.Status != "Read")
                    return work;
                string response;
                int code = 200;
                switch (work.ProbeKind)
                {
                    case "Owner":
                        response = Body(new SiteGroup { Id = 7, Type = 8 });
                        break;
                    case "Roles":
                        response = Body(
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
                                    new SecurityRoleObservation
                                    {
                                        Id = 5,
                                        Type = 5,
                                        Permissions = new PermissionMask
                                        {
                                            High = "2147483647",
                                            Low = "4294967295",
                                        },
                                    },
                                },
                            }
                        );
                        break;
                    case "Library":
                        code = Exists ? 200 : 404;
                        response = Body(
                            new CatalogLibraryObservation
                            {
                                Id = List,
                                Unique = Unique,
                                Root = new FolderObservation
                                {
                                    Id = Root,
                                    Path = "/sites/test/Documents",
                                },
                            }
                        );
                        break;
                    case "OwnerGrant" when LegacyAcl:
                        // A setup saved by 0.1.0.3 mid-read: its probe asked for the whole list.
                        LegacyAcl = false;
                        var legacy = Store.Require<LibrarySetup>("asx_operation", work.Key);
                        legacy.Value.ProbeKind = "Acl";
                        Store.Save(legacy);
                        work.ProbeKind = "Acl";
                        response = Body(
                            new ODataRows<AclAssignment> { Rows = Array.Empty<AclAssignment>() }
                        );
                        break;
                    case "OwnerGrant":
                        // Only the owners group's own assignment is read; SharePoint answers
                        // 404 when the group has none on the list.
                        OwnerGrantReads.Add(work.Http!.RelativeUri);
                        if (!OwnerAccess)
                        {
                            code = 404;
                            response =
                                "{\"error\":{\"code\":\"-2146232832\",\"message\":{\"lang\":\"en-US\",\"value\":\"Can not find the principal with id: 7.\"}}}";
                            break;
                        }
                        response = Body(
                            new AclAssignment
                            {
                                Member = new AclMember { Id = 7, Type = 8 },
                                Roles = new ODataRows<AclRole>
                                {
                                    Rows = new[]
                                    {
                                        new AclRole
                                        {
                                            Id = 5,
                                            Permissions = new PermissionMask
                                            {
                                                High = "2147483647",
                                                Low = "4294967295",
                                            },
                                        },
                                    },
                                },
                            }
                        );
                        break;
                    default:
                        throw new Exception(work.ProbeKind);
                }
                work = Call("Observe", work, response, code);
            }
            throw new Exception("Library setup did not terminate");
        }

        /// <summary>An operator's Retry or Cancel, which carries only the setup key.</summary>
        public WorkerResult Manage(string command, string key) =>
            Service.Transaction(() =>
                Worker.Execute(new WorkerRequest { Command = command, Key = key }, true)
            );

        public WorkerResult Call(
            string command,
            WorkerResult work,
            string? body = null,
            int status = 0
        ) =>
            Service.Transaction(() =>
                Worker.Execute(
                    new WorkerRequest
                    {
                        Command = command,
                        Key = work.Key,
                        RunId = "library-test",
                        Token = work.Token,
                        ProbeId = work.ProbeId,
                        ProbeKind = work.ProbeKind,
                        ResponseBody = body,
                        HttpStatus = status,
                    },
                    true
                )
            );
    }
}
