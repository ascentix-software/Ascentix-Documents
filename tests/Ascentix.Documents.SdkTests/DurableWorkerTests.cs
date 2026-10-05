using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class DurableWorkerTests
{
    [Fact]
    public void CompleteProtocolPersistsOneFolderBindingAndNativeLocation()
    {
        var fixture = new Fixture();
        var read = fixture.Claim();
        read = fixture.Preflight(read);
        var absent = fixture.Observe(read, Rows<ItemObservation>());
        Assert.Equal("ReadyToCreate", absent.Status);
        var prepared = fixture.Call("PrepareCreate", absent);
        Assert.Equal("Create", prepared.Status);
        Assert.Equal("None", prepared.Http!.RetryPolicy);
        var created = fixture.Call("CreateResponse", prepared, CreateBody(), 200);
        Assert.Equal("Read", created.Status);
        var physical = Guid.NewGuid();
        var verified = fixture.ObserveAndFinalize(created, fixture.Item(physical, null));
        var completed = fixture.Call("Complete", verified);
        Assert.Equal("Applied", completed.Status);
        Assert.Equal(physical, completed.PhysicalId);
        Assert.NotEqual(Guid.Empty, completed.LocationId);
        Assert.Equal(completed.LocationId, fixture.Call("Complete", verified).LocationId);
        Assert.Single(
            fixture.Service.Rows.Values,
            e => e.LogicalName == "sharepointdocumentlocation" && e.Id != fixture.NativeParent
        );
        Assert.Equal(
            "Idle",
            fixture
                .Store.Require<DispatcherDocument>(
                    "asx_claim",
                    WorkCoordination.Operation(fixture.Service, fixture.Operation.Key)
                )
                .Value.Status
        );
    }

    [Fact]
    public void NativeLocationAndCompletionRollbackTogetherOnCasConflict()
    {
        var fixture = new Fixture();
        var read = fixture.Preflight(fixture.Claim());
        var verified = fixture.ObserveAndFinalize(read, fixture.Item(Guid.NewGuid(), null));
        fixture.Service.FailUpdateTable = "asx_operation";
        Assert.Throws<InvalidOperationException>(() => fixture.Call("Complete", verified));
        Assert.Single(
            fixture.Service.Rows.Values,
            e => e.LogicalName == "sharepointdocumentlocation"
        );
        Assert.Equal(
            "Verified",
            fixture
                .Store.Require<OperationDocument>("asx_operation", fixture.Operation.Key)
                .Value.Status
        );
        fixture.Service.FailUpdateTable = null;
        Assert.Equal("Applied", fixture.Call("Complete", verified).Status);
        Assert.Equal(
            2,
            fixture.Service.Rows.Values.Count(e => e.LogicalName == "sharepointdocumentlocation")
        );
    }

    [Fact]
    public void OperatorRecoveryPermitStillWorksOnAnExpiredClaim()
    {
        var fixture = new Fixture();
        var claim = fixture.Claim();
        fixture.Now = fixture.Now.AddMinutes(6);
        Assert.Throws<EvaluationBlockedException>(() => fixture.Call("Renew", claim));
        var request = new WorkerRequest
        {
            Key = fixture.Operation.Key,
            RunId = "run-1",
            Token = claim.Token,
            Evidence =
                "Operator verified prior run terminated and outstanding calls resolved: test receipt",
        };
        var permitted = fixture.Service.Transaction(() =>
            fixture.Coordinator.PermitRecovery(request, true)
        );
        Assert.Equal("RecoveryPermitted", permitted.Status);
        var recovered = fixture.Claim("run-2");
        Assert.Equal("Read", recovered.Status);
        Assert.NotEqual(claim.Token, recovered.Token);
    }

    [Fact]
    public void UnknownCreateCannotChooseAnotherCandidateOrComplete()
    {
        var fixture = new Fixture();
        var absent = fixture.Observe(fixture.Preflight(fixture.Claim()), Rows<ItemObservation>());
        var prepared = fixture.Call("PrepareCreate", absent);
        Assert.Throws<EvaluationBlockedException>(() => fixture.Call("PrepareCreate", prepared));
        var resumed = fixture.Claim("run-1", prepared.Token);
        var folderRead = fixture.Preflight(resumed);
        var observed = fixture.Observe(folderRead, Rows(fixture.Item(Guid.NewGuid(), null)));
        Assert.Equal("Quarantined", observed.Status);
        Assert.Equal(
            "Example",
            fixture
                .Store.Require<OperationDocument>("asx_operation", fixture.Operation.Key)
                .Value.Folder.Candidate
        );
        Assert.Throws<EvaluationBlockedException>(() => fixture.Call("Complete", resumed));
    }

    [Fact]
    public void KnownCreateResponseStillRequiresIndependentFolderAndFinalParentReads()
    {
        var fixture = new Fixture();
        var absent = fixture.Observe(fixture.Preflight(fixture.Claim()), Rows<ItemObservation>());
        var prepared = fixture.Call("PrepareCreate", absent);
        var read = fixture.Call("CreateResponse", prepared, CreateBody(), 200);
        Assert.Throws<EvaluationBlockedException>(() => fixture.Call("Complete", read));
        var finalParent = fixture.Observe(read, Rows(fixture.Item(Guid.NewGuid(), null)));
        Assert.Equal("FinalParent", finalParent.ProbeKind);
        Assert.Throws<EvaluationBlockedException>(() => fixture.Call("Complete", finalParent));
    }

    [Fact]
    public void FinalParentIdentityChangeBlocksAppliedAndReleasesKnownWriter()
    {
        var fixture = new Fixture();
        var read = fixture.Preflight(fixture.Claim());
        var finalParent = fixture.Observe(read, Rows(fixture.Item(Guid.NewGuid(), null)));
        Assert.Equal("FinalParent", finalParent.ProbeKind);
        var drift = fixture.Observe(
            finalParent,
            JsonWire.Write(
                new ODataEnvelope<FolderObservation>
                {
                    Data = new FolderObservation
                    {
                        Id = Guid.NewGuid(),
                        Path = "/sites/proto/General",
                    },
                }
            )
        );
        Assert.Equal("Blocked", drift.Status);
        Assert.NotEqual(
            "Applied",
            fixture
                .Store.Require<OperationDocument>("asx_operation", fixture.Operation.Key)
                .Value.Folder.Status
        );
        Assert.Equal(
            "Idle",
            fixture
                .Store.Require<DispatcherDocument>(
                    "asx_claim",
                    WorkCoordination.Operation(fixture.Service, fixture.Operation.Key)
                )
                .Value.Status
        );
    }

    [Fact]
    public void StaleObservationNonceCannotAdvanceAnotherProbe()
    {
        var fixture = new Fixture();
        var claim = fixture.Claim();
        var next = fixture.Observe(claim, fixture.LibraryBody());
        Assert.Throws<EvaluationBlockedException>(() =>
            fixture.Observe(claim, fixture.LibraryBody())
        );
        Assert.Equal("Parent", next.ProbeKind);
    }

    [Fact]
    public void ExistingFolderIsReusedRegardlessOfCreationSource()
    {
        var fixture = new Fixture();
        var read = fixture.Preflight(fixture.Claim());
        var physical = Guid.NewGuid();
        var verified = fixture.ObserveAndFinalize(read, fixture.Item(physical, Guid.NewGuid()));
        var completed = fixture.Call("Complete", verified);
        Assert.Equal(physical, completed.PhysicalId);
        Assert.Equal(
            "Example",
            fixture
                .Store.Require<OperationDocument>("asx_operation", fixture.Operation.Key)
                .Value.Folder.Candidate
        );
        Assert.DoesNotContain(
            fixture.Service.Rows.Values,
            r => r.LogicalName == "asx_binding" || r.LogicalName == "asx_physicalfolder"
        );
    }

    [Fact]
    public void QueueAndPlanAreIdempotentAcrossNewCoordinatorInstances()
    {
        var fixture = new Fixture(seedBinding: false);
        fixture.SeedTemplate();
        var request = new WorkerRequest
        {
            Command = "Queue",
            RequestId = Guid.NewGuid(),
            TemplateId = fixture.TemplateId,
            RecordId = fixture.RecordId,
        };
        var first = fixture.Service.Transaction(() => fixture.Coordinator.Execute(request, true));
        var again = fixture.Service.Transaction(() =>
            new WorkerCoordinator(fixture.Service, () => fixture.Now).Execute(request, true)
        );
        Assert.Equal(first.Key, again.Key);
        var planned = fixture.Service.Transaction(() =>
            fixture.Coordinator.Execute(
                new WorkerRequest { Command = "Plan", Key = first.Key },
                true
            )
        );
        Assert.Single(planned.Keys);
        var repeated = fixture.Service.Transaction(() =>
            fixture.Coordinator.Execute(
                new WorkerRequest { Command = "Plan", Key = first.Key },
                true
            )
        );
        Assert.Equal(planned.Keys, repeated.Keys);
        Assert.DoesNotContain(
            fixture.Service.Rows.Values,
            e => e.LogicalName == "asx_binding" || e.LogicalName == "asx_physicalfolder"
        );
        Assert.Single(fixture.Service.Rows.Values, e => e.LogicalName == "asx_operation");
        request.RecordId = Guid.NewGuid();
        Assert.Throws<EvaluationBlockedException>(() =>
            fixture.Service.Transaction(() => fixture.Coordinator.Execute(request, true))
        );
    }

    [Fact]
    public void DurableFullKeyMismatchFailsEvenIfHashLookupMatches()
    {
        var fixture = new Fixture();
        var row = fixture
            .Store.Require<OperationDocument>("asx_operation", fixture.Operation.Key)
            .Row;
        var changed = fixture.Service.Rows[row.Id];
        var payload = JsonWire.Read<OperationDocument>(
            changed.GetAttributeValue<string>("asx_payload")
        );
        payload.Key = "other";
        changed["asx_payload"] = JsonWire.Write(payload);
        Assert.Throws<EvaluationBlockedException>(() =>
            fixture.Store.Require<OperationDocument>("asx_operation", fixture.Operation.Key)
        );
    }

    [Fact]
    public void RealSdkCasRequestIsMandatoryForEveryDurableSave()
    {
        var fixture = new Fixture();
        var stale = fixture.Store.Require<OperationDocument>(
            "asx_operation",
            fixture.Operation.Key
        );
        var current = fixture.Store.Require<OperationDocument>(
            "asx_operation",
            fixture.Operation.Key
        );
        current.Value.Folder.Candidate = "Changed";
        fixture.Store.Save(current);
        stale.Value.Folder.Candidate = "Stale";
        Assert.Throws<InvalidOperationException>(() => fixture.Store.Save(stale));
        Assert.All(
            fixture.Service.Updates,
            update =>
                Assert.Equal(ConcurrencyBehavior.IfRowVersionMatches, update.ConcurrencyBehavior)
        );
    }

    [Fact]
    public void WorkerRejectsMissingTransactionAndIncompleteAcl()
    {
        var fixture = new Fixture();
        Assert.Throws<EvaluationBlockedException>(() =>
            fixture.Coordinator.Execute(
                new WorkerRequest
                {
                    Command = "Claim",
                    Key = fixture.Operation.Key,
                    RunId = "run",
                },
                false
            )
        );
        var partial = new ODataRows<AclAssignment> { Rows = fixture.Acl.Rows, Next = "next-page" };
        Assert.Throws<EvaluationBlockedException>(() => SharePointObservations.AclHash(partial));
    }

    [Fact]
    public void NativeParentCyclesAreRejectedWithoutCreatingLocations()
    {
        var fixture = new Fixture();
        fixture.Service.Rows[fixture.NativeParent]["parentsiteorlocation"] = new EntityReference(
            "sharepointdocumentlocation",
            fixture.NativeParent
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            new NativeLocations(fixture.Service).ValidateParent(
                fixture.NativeParent,
                fixture.EntryUrl,
                fixture.NativeSite
            )
        );
        Assert.Single(
            fixture.Service.Rows.Values,
            e => e.LogicalName == "sharepointdocumentlocation"
        );
    }

    [Fact]
    public void ReadFailureSchedulesPersistentBackoffBeforeAnotherClaim()
    {
        var fixture = new Fixture();
        var claim = fixture.Claim();
        var waiting = fixture.Call("Observe", claim, "{}", 429);
        Assert.Equal("RetryWait", waiting.Status);
        var operation = fixture
            .Store.Require<OperationDocument>("asx_operation", fixture.Operation.Key)
            .Value;
        Assert.Equal(1, operation.RetryCount);
        Assert.Equal(fixture.Now.AddSeconds(30), operation.NextAttemptUtc);
        Assert.Equal("RetryWait", fixture.Claim("run-2").Status);
        fixture.Now = fixture.Now.AddSeconds(31);
        Assert.Equal("Read", fixture.Claim("run-2").Status);
    }

    [Fact]
    public void InFlightWorkForARemovedTableStopsBeforeItsNextWrite()
    {
        var f = new Fixture();
        f.AllowedTables = new[] { "account" };
        var work = f.Observe(f.Preflight(f.Claim()), Rows<ItemObservation>());
        Assert.Equal("ReadyToCreate", work.Status);
        f.AllowedTables = new[] { "contact" };
        var stopped = Assert.Throws<EvaluationBlockedException>(() =>
            f.Call("PrepareCreate", work)
        );
        Assert.Contains("outside the approved runtime scope", stopped.Message);
        Assert.False(
            f.Store.Require<OperationDocument>(
                "asx_operation",
                f.Operation.Key
            ).Value.ExternalSubmitted
        );
    }

    [Fact]
    public void SentCreateForARemovedTableFinishesAndReleasesItsWriterSlot()
    {
        var f = new Fixture();
        f.AllowedTables = new[] { "account" };
        var work = f.Observe(f.Preflight(f.Claim()), Rows<ItemObservation>());
        work = f.Call("PrepareCreate", work);
        Assert.Equal("Create", work.Status);
        string[] Writers() =>
            f
                .Store.Require<ConnectionBudget>("asx_claim", WorkCoordination.BudgetKey)
                .Value.Writers;
        Assert.NotEmpty(Writers());
        f.AllowedTables = new[] { "contact" };
        work = f.Call("CreateResponse", work, CreateBody(), 200);
        Assert.Equal("Read", work.Status);
        var verified = f.ObserveAndFinalize(work, f.Item(Guid.NewGuid(), null));
        Assert.Equal("Applied", f.Call("Complete", verified).Status);
        Assert.Empty(Writers());
    }

    [Fact]
    public void WorkerSourceScopeIsCheckedBeforeReadingQueuedBusinessRecord()
    {
        var fixture = new Fixture(seedBinding: false);
        fixture.SeedTemplate();
        var scoped = new WorkerCoordinator(fixture.Service, () => fixture.Now, new[] { "contact" });
        Assert.Throws<EvaluationBlockedException>(() =>
            fixture.Service.Transaction(() =>
                scoped.Execute(
                    new WorkerRequest
                    {
                        Command = "Queue",
                        TemplateId = fixture.TemplateId,
                        RecordId = fixture.RecordId,
                        RequestId = Guid.NewGuid(),
                    },
                    true
                )
            )
        );
        Assert.DoesNotContain(fixture.Service.Rows.Values, row => row.LogicalName == "asx_outbox");
    }

    [Fact]
    public void RuntimeProfileRequiresAnExplicitIdentityAndAllowsDisabledSetup()
    {
        var fixture = new Fixture();
        Assert.Throws<EvaluationBlockedException>(() => RuntimeProfile.Read(fixture.Service));
        var worker = Guid.NewGuid();
        RuntimeSeed.Seed(fixture.Service, worker, "account");
        var profile = RuntimeProfile.Read(fixture.Service);
        Assert.Equal(worker, profile.WorkerId);
        Assert.False(profile.Enabled);
        Assert.Equal(new[] { "account" }, profile.Tables);
    }

    /// <summary>
    /// Adds a SharePoint report of hand-set (unique) permissions to a folder or folder-item read.
    /// The product no longer requests this field, so the edit works on the JSON text.
    /// </summary>
    /// <param name="body">A serialized folder observation or folder lookup observation.</param>
    /// <returns>The same body with <c>ListItemAllFields/HasUniqueRoleAssignments</c> set to true.</returns>
    internal static string WithUniquePermissions(string body)
    {
        body = System.Text.RegularExpressions.Regex.Replace(
            body,
            "\"HasUniqueRoleAssignments\":(true|false|null),?",
            ""
        );
        body = body.Replace("\"ListItemAllFields\":null", "\"ListItemAllFields\":{}");
        if (!body.Contains("\"ListItemAllFields\":{"))
            body = body.Replace("{\"d\":{", "{\"d\":{\"ListItemAllFields\":{},");
        return body.Replace(
                "\"ListItemAllFields\":{",
                "\"ListItemAllFields\":{\"HasUniqueRoleAssignments\":true,"
            )
            .Replace(",}", "}");
    }

    private static string Rows<T>(params T[] values) =>
        typeof(T) == typeof(ItemObservation)
            ? (
                values.Length == 0
                    ? "{}"
                    : JsonWire.Write(
                        new ODataEnvelope<FolderLookupObservation>
                        {
                            Data = new FolderLookupObservation
                            {
                                Exists = true,
                                Id = values.Cast<ItemObservation>().Single().Id,
                                Path = values.Cast<ItemObservation>().Single().Path,
                                ListItemAllFields = values.Cast<ItemObservation>().Single(),
                            },
                        }
                    )
            )
            : JsonWire.Write(
                new ODataEnvelope<ODataRows<T>> { Data = new ODataRows<T> { Rows = values } }
            );

    internal static string CreateBody() =>
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
        );

    [Fact]
    public void LoadRevisionCanRepairSuspendedLibraryAndPreservesFieldComparison()
    {
        var f = new Fixture(seedBinding: false);
        f.SeedTemplate();
        f.Service.Rows.Values.Single(r => r.LogicalName == "asx_destination")["asx_name"] =
            "Account onboarding";
        f.Service.Rows[f.LibraryId]["asx_approved"] = false;
        var root = f.Service.Rows.Values.Single(r => r.LogicalName == "asx_folder");
        root["asx_groupkey"] = "group";
        f.Service.Seed(
            new Entity("asx_documentconditiongroup", Guid.NewGuid())
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", f.RevisionId),
                ["asx_key"] = "group",
                ["asx_all"] = true,
            }
        );
        f.Service.Seed(
            new Entity("asx_condition", Guid.NewGuid())
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", f.RevisionId),
                ["asx_groupkey"] = "group",
                ["asx_payload"] = JsonWire.Write(
                    new ConditionDto
                    {
                        Source = "root",
                        Column = "name",
                        Operator = "Equal",
                        RightSource = "customer",
                        RightColumn = "name",
                    }
                ),
            }
        );
        var loaded = new DraftReader(f.Service).Read(f.RevisionId);
        Assert.Equal("Published", loaded.Status);
        Assert.NotEmpty(loaded.RowVersion);
        Assert.Equal("general", loaded.Draft.Destinations[0].Key);
        Assert.Equal("Account onboarding", loaded.Draft.Destinations[0].Name);
        var condition = loaded.Draft.Destinations[0].Folders[0].Condition!.Conditions.Single();
        Assert.Equal("customer", condition.RightSource);
        Assert.Null(condition.LiteralKind);
        Assert.Throws<EvaluationBlockedException>(() =>
            new TemplateStore(f.Service).Read(f.RevisionId)
        );
    }

    [Fact]
    public void LoadRevisionRejectsOrphanedGroupsInsteadOfSilentlyDroppingConditions()
    {
        var f = new Fixture(seedBinding: false);
        f.SeedTemplate();
        f.Service.Seed(
            new Entity("asx_documentconditiongroup", Guid.NewGuid())
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", f.RevisionId),
                ["asx_key"] = "orphan",
                ["asx_all"] = true,
            }
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            new DraftReader(f.Service).Read(f.RevisionId)
        );
    }

    [Fact]
    public void RetryAfterHonorsServerDelayUpToTheFifteenMinuteCap()
    {
        var now = new DateTime(2026, 9, 8, 8, 0, 0, DateTimeKind.Utc);
        Assert.Equal(now.AddMinutes(10), WorkerCoordinator.RetryAt(now, 1, "600"));
        Assert.Equal(now.AddMinutes(15), WorkerCoordinator.RetryAt(now, 1, "3600"));
        Assert.Equal(
            now.AddMinutes(12),
            WorkerCoordinator.RetryAt(now, 1, now.AddMinutes(12).ToString("r"))
        );
        Assert.Equal(
            now.AddMinutes(15),
            WorkerCoordinator.RetryAt(now, 1, now.AddMinutes(30).ToString("r"))
        );
        Assert.Equal(now.AddSeconds(60), WorkerCoordinator.RetryAt(now, 2, "1"));
        Assert.Equal(now.AddMinutes(15), WorkerCoordinator.RetryAt(now, 1, "999999999999"));
    }

    [Fact]
    public void CancellationCannotAbandonAnActiveWriter()
    {
        var f = new Fixture();
        f.Claim();
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() =>
                f.Coordinator.Execute(
                    new WorkerRequest { Command = "Cancel", Key = f.Operation.Key },
                    true
                )
            )
        );
        var idle = new Fixture();
        Assert.Equal(
            "Cancelled",
            idle.Coordinator.Execute(
                new WorkerRequest { Command = "Cancel", Key = idle.Operation.Key },
                true
            ).Status
        );
        Assert.Equal("Cancelled", idle.Claim().Status);
    }

    [Fact]
    public void ExplicitReplanCreatesFreshOperationAndPreservesPhysicalBinding()
    {
        var f = new Fixture(seedBinding: false);
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
        var planned = f.Coordinator.Execute(
            new WorkerRequest { Command = "Plan", Key = queued.Key },
            true
        );
        f.Coordinator.Execute(
            new WorkerRequest { Command = "Cancel", Key = planned.Keys.Single() },
            true
        );
        var replan = f.Coordinator.Execute(
            new WorkerRequest
            {
                Command = "Replan",
                TemplateId = f.TemplateId,
                RecordId = f.RecordId,
                RequestId = Guid.NewGuid(),
            },
            true
        );
        var again = f.Coordinator.Execute(
            new WorkerRequest { Command = "Plan", Key = replan.Key },
            true
        );
        Assert.NotEqual(planned.Keys.Single(), again.Keys.Single());
        Assert.DoesNotContain(
            f.Service.Rows.Values,
            r => r.LogicalName == "asx_binding" || r.LogicalName == "asx_physicalfolder"
        );
    }

    [Fact]
    public void NewPublicationSupersedesUnstartedOldWork()
    {
        var f = new Fixture();
        f.Service.Rows[f.TemplateId]["asx_publishedrevisionid"] = new EntityReference(
            "asx_revision",
            Guid.NewGuid()
        );
        Assert.Equal("Superseded", f.Claim().Status);
        Assert.False(
            f.Store.Require<OperationDocument>(
                "asx_operation",
                f.Operation.Key
            ).Value.ExternalSubmitted
        );
    }

    [Fact]
    public void NewRecordSelectionSupersedesUnsubmittedWorkWithinSameRevision()
    {
        var f = new Fixture();
        f.Store.Create(
            "asx_outbox",
            new RecordPlanDocument
            {
                Key = "recordplan:" + f.TemplateId.ToString("N") + ":" + f.RecordId.ToString("N"),
                Status = "Selection",
                RevisionId = f.RevisionId,
                IncludedSections = Array.Empty<string>(),
            }
        );
        Assert.Equal("Superseded", f.Claim().Status);
        Assert.False(
            f.Store.Require<OperationDocument>(
                "asx_operation",
                f.Operation.Key
            ).Value.ExternalSubmitted
        );
    }

    [Fact]
    public void SelectionChangeAfterClaimPreventsCreatePreparation()
    {
        var f = new Fixture();
        var absent = f.Observe(f.Preflight(f.Claim()), Rows<ItemObservation>());
        f.Store.Create(
            "asx_outbox",
            new RecordPlanDocument
            {
                Key = "recordplan:" + f.TemplateId.ToString("N") + ":" + f.RecordId.ToString("N"),
                Status = "Selection",
                RevisionId = f.RevisionId,
                IncludedSections = Array.Empty<string>(),
            }
        );
        Assert.Throws<EvaluationBlockedException>(() => f.Call("PrepareCreate", absent));
        Assert.False(
            f.Store.Require<OperationDocument>(
                "asx_operation",
                f.Operation.Key
            ).Value.ExternalSubmitted
        );
    }

    [Fact]
    public void DispatchPrioritizesSecurityAndParentsAheadOfOlderChildren()
    {
        var service = new MemoryService();
        var store = new DocumentStore(service);
        for (int i = 0; i < 40; i++)
            store.Create(
                "asx_operation",
                new OperationDocument { Key = "child:" + i, Priority = 2 }
            );
        store.Create("asx_operation", new OperationDocument { Key = "parent", Priority = 1 });
        store.Create("asx_operation", new SecurityOperation { Key = "policywork:priority" });
        var keys = store.Pending("asx_operation");
        Assert.Equal("policywork:priority", keys[0]);
        Assert.Equal("parent", keys[1]);
        Assert.Equal(20, keys.Length);
    }

    [Fact]
    public void FailedUnclaimedSecurityWorkPreservesItsGenerationPayload()
    {
        var service = new MemoryService();
        var store = new DocumentStore(service);
        Guid library = Guid.NewGuid();
        store.Create(
            "asx_operation",
            new SecurityOperation
            {
                Key = "policywork:failure",
                WebUrl = "https://example.sharepoint.com/sites/test",
                LibraryId = library,
                Entries = new[]
                {
                    new PolicyEntry { TeamId = Guid.NewGuid(), Access = "Read" },
                },
            }
        );
        Assert.Equal(
            "Blocked",
            store.FailUnclaimed<SecurityOperation>("policywork:failure").Status
        );
        var value = store.Require<SecurityOperation>("asx_operation", "policywork:failure").Value;
        Assert.Equal(library, value.LibraryId);
        Assert.Equal("Read", value.Entries.Single().Access);
    }

    [Fact]
    public void RetryDispatchIncludesOnlyDueRows()
    {
        var service = new MemoryService();
        var store = new DocumentStore(service);
        var now = new DateTime(2026, 9, 8, 12, 0, 0, DateTimeKind.Utc);
        foreach (var offset in new[] { -1, 1 })
        {
            string key = "retry:" + offset;
            store.Create("asx_operation", new OperationDocument { Key = key });
            var row = store.Require<OperationDocument>("asx_operation", key);
            row.Value.Status = "RetryWait";
            row.Value.NextAttemptUtc = now.AddMinutes(offset);
            store.Save(row);
        }
        Assert.Equal(new[] { "retry:-1" }, store.Pending("asx_operation", now: now));
    }

    [Fact]
    public void DeletionAfterPhysicalWriteKeepsReceiptWithoutCreatingNativeRegarding()
    {
        var f = new Fixture();
        var work = f.Preflight(f.Claim());
        work = f.Observe(work, Rows<ItemObservation>());
        work = f.Call("PrepareCreate", work);
        f.Store.Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = WorkerCoordinator.RetirementKey("account", f.RecordId),
                Status = "DecommissionReview",
                Table = "account",
                RecordId = f.RecordId,
            }
        );
        work = f.Call("CreateResponse", work, CreateBody(), 200);
        work = f.ObserveAndFinalize(work, f.Item(Guid.NewGuid(), null));
        work = f.Call("Complete", work);
        Assert.Equal("Applied", work.Status);
        Assert.Equal(Guid.Empty, work.LocationId);
        Assert.Equal(
            "DecommissionReview",
            f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value.Folder.Status
        );
        Assert.DoesNotContain(
            f.Service.Rows.Values,
            r =>
                r.LogicalName == "sharepointdocumentlocation"
                && r.GetAttributeValue<EntityReference>("regardingobjectid") != null
        );
    }

    [Theory]
    [InlineData("None")]
    [InlineData("EmptyId")]
    [InlineData("OtherPath")]
    public void FolderUnderAncestorsWithHandSetPermissionsIsCreatedUnlessAncestorIdentityChanges(
        string finalChange
    )
    {
        var f = new Fixture();
        string entry = "/sites/proto/General/Archive/Entry";
        f.Service.Rows[f.LibraryId]["asx_entryurl"] = "https://example.sharepoint.com" + entry;
        f.Service.Rows[f.NativeParent]["relativeurl"] = "General/Archive/Entry";
        var work = f.Claim();
        work = f.Observe(
            work,
            JsonWire.Write(
                new ODataEnvelope<LibraryObservation>
                {
                    Data = new LibraryObservation
                    {
                        Id = f.ListId,
                        Root = new FolderObservation
                        {
                            Id = Guid.NewGuid(),
                            Path = "/sites/proto/General",
                        },
                    },
                }
            )
        );
        string parent = WithUniquePermissions(
            JsonWire.Write(
                new ODataEnvelope<FolderObservation>
                {
                    Data = new FolderObservation { Id = f.EntryId, Path = entry },
                }
            )
        );
        string Ancestor(Guid id, string path) =>
            WithUniquePermissions(
                JsonWire.Write(
                    new ODataEnvelope<FolderObservation>
                    {
                        Data = new FolderObservation { Id = id, Path = path },
                    }
                )
            );
        var ancestorId = Guid.NewGuid();
        work = f.Observe(work, parent);
        Assert.Equal("Ancestor", work.ProbeKind);
        work = f.Observe(work, Ancestor(ancestorId, "/sites/proto/General/Archive"));
        Assert.Equal("Folder", work.ProbeKind);
        work = f.Observe(work, Rows<ItemObservation>());
        Assert.Equal("ReadyToCreate", work.Status);
        work = f.Call("PrepareCreate", work);
        Assert.Equal("Create", work.Status);
        work = f.Call("CreateResponse", work, CreateBody(), 200);
        var item = f.Item(Guid.NewGuid(), null);
        item.Path = entry + "/Example";
        work = f.Observe(work, Rows(item));
        Assert.Equal("FinalParent", work.ProbeKind);
        work = f.Observe(work, parent);
        Assert.Equal("FinalAncestor", work.ProbeKind);
        work = f.Observe(
            work,
            Ancestor(
                finalChange == "EmptyId" ? Guid.Empty : ancestorId,
                finalChange == "OtherPath"
                    ? "/sites/proto/General/Moved"
                    : "/sites/proto/General/Archive"
            )
        );
        if (finalChange != "None")
        {
            // The kept identity and path checks still stop the work at the final read.
            Assert.Equal("Blocked", work.Status);
            Assert.Equal(
                "ObservationMismatch",
                f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value.ErrorCode
            );
            return;
        }
        Assert.Equal("Verified", work.Status);
        Assert.Equal("Applied", f.Call("Complete", work).Status);
    }

    [Fact]
    public void ChildFolderUnderAParentWithHandSetPermissionsIsCreatedAndApplied()
    {
        var f = new Fixture(false);
        var child = JsonWire.Read<FolderStep>(JsonWire.Write(f.Binding));
        child.Key += "-child";
        child.Node = "invoices";
        child.ParentBinding = "root";
        child.OriginalName = child.Candidate = "Invoices";
        f.Operation.Folders = new[] { f.Binding, child };
        f.Store.Create("asx_operation", f.Operation);
        var root = Guid.NewGuid();
        var work = f.ObserveAndFinalize(f.Preflight(f.Claim()), f.Item(root, null));
        Assert.Equal("Pending", f.Call("Complete", work).Status);
        work = f.Claim("child-run");
        work = f.Observe(work, f.LibraryBody());
        string parent = WithUniquePermissions(
            JsonWire.Write(
                new ODataEnvelope<FolderObservation>
                {
                    Data = new FolderObservation
                    {
                        Id = root,
                        Path = "/sites/proto/General/Example",
                    },
                }
            )
        );
        work = f.Observe(work, parent);
        Assert.Equal("Folder", work.ProbeKind);
        work = f.Observe(work, Rows<ItemObservation>());
        Assert.Equal("ReadyToCreate", work.Status);
        work = f.Call("PrepareCreate", work);
        Assert.Equal("Create", work.Status);
        work = f.Call("CreateResponse", work, CreateBody(), 200);
        var item = f.Item(Guid.NewGuid(), null);
        item.Name = "Invoices";
        item.Path += "/Invoices";
        // The new folder also reports hand-set permissions, for example from a SharePoint policy.
        work = f.Observe(work, WithUniquePermissions(Rows(item)));
        Assert.Equal("FinalParent", work.ProbeKind);
        work = f.Observe(work, parent);
        Assert.Equal("Verified", work.Status);
        Assert.Equal("Applied", f.Call("Complete", work).Status);
    }

    [Fact]
    public void FileAtExpectedPathBlocksWithoutSuffixOrNativeNavigation()
    {
        var f = new Fixture();
        var work = f.Observe(f.Preflight(f.Claim()), Rows<ItemObservation>());
        work = f.Call("PrepareCreate", work);
        work = f.Call("CreateResponse", work, "{}", 409);
        work = f.Observe(work, Rows<ItemObservation>());
        Assert.Equal("ConflictFile", work.ProbeKind);
        Assert.Equal("GET", work.Http!.Method);
        work = f.Observe(
            work,
            JsonWire.Write(
                new ODataEnvelope<FolderObservation>
                {
                    Data = new FolderObservation
                    {
                        Id = Guid.NewGuid(),
                        Path = "/sites/proto/General/Example",
                    },
                }
            )
        );
        Assert.Equal("Blocked", work.Status);
        Assert.Equal(
            "FileAtExpectedFolderPath",
            f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value.ErrorCode
        );
        Assert.DoesNotContain(
            f.Service.Rows.Values,
            r =>
                r.LogicalName == "sharepointdocumentlocation"
                && r.GetAttributeValue<EntityReference>("regardingobjectid") != null
        );
    }

    [Fact]
    public void ChildStepsResumeWithOneNativeRootAndNoFolderInventory()
    {
        var f = new Fixture(false);
        var child = JsonWire.Read<FolderStep>(JsonWire.Write(f.Binding));
        child.Key += "-child";
        child.Node = "invoices";
        child.ParentBinding = "root";
        child.OriginalName = child.Candidate = "Invoices";
        f.Operation.Folders = new[] { f.Binding, child };
        f.Store.Create("asx_operation", f.Operation);
        var root = Guid.NewGuid();
        var work = f.ObserveAndFinalize(f.Preflight(f.Claim()), f.Item(root, null));
        work = f.Call("Complete", work);
        Assert.Equal("Pending", work.Status);
        Assert.Equal(
            1,
            f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value.Cursor
        );
        work = f.Claim("child-run");
        work = f.Observe(work, f.LibraryBody());
        string parent = JsonWire.Write(
            new ODataEnvelope<FolderObservation>
            {
                Data = new FolderObservation { Id = root, Path = "/sites/proto/General/Example" },
            }
        );
        work = f.Observe(work, parent);
        var item = f.Item(Guid.NewGuid(), null);
        item.Name = "Invoices";
        item.Path += "/Invoices";
        work = f.Observe(work, Rows(item));
        work = f.Observe(work, parent);
        Assert.Equal("Verified", work.Status);
        work = f.Call("Complete", work);
        Assert.Equal("Applied", work.Status);
        Assert.Single(
            f.Service.Rows.Values,
            r =>
                r.LogicalName == "sharepointdocumentlocation"
                && r.GetAttributeValue<EntityReference>("regardingobjectid") != null
        );
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_operation");
        Assert.DoesNotContain(
            f.Service.Rows.Values,
            r => r.LogicalName == "asx_binding" || r.LogicalName == "asx_physicalfolder"
        );
        var renamed = JsonWire.Read<FolderStep>(JsonWire.Write(f.Binding));
        renamed.Candidate = "Changed account name";
        Assert.Equal(
            "Example",
            new NativeLocations(f.Service)
                .Find(renamed, f.NativeParent, f.EntryUrl, f.NativeSite)!
                .GetAttributeValue<string>("relativeurl")
        );
    }

    [Fact]
    public void InFlightFolderJobSurvivesPolicyGenerationChange()
    {
        var f = new Fixture();
        var work = f.Preflight(f.Claim());
        f.StartPolicyUpdate();
        work = f.Observe(work, Rows<ItemObservation>());
        Assert.Equal("ReadyToCreate", work.Status);
        work = f.Call("PrepareCreate", work);
        f.StartPolicyUpdate();
        work = f.Call("CreateResponse", work, CreateBody(), 200);
        work = f.ObserveAndFinalize(work, f.Item(Guid.NewGuid(), null));
        f.StartPolicyUpdate();
        Assert.Equal("Applied", f.Call("Complete", work).Status);
    }

    [Fact]
    public void PlanSucceedsWhileLibraryPolicyIsBeingApplied()
    {
        var f = new Fixture(seedBinding: false);
        f.SeedTemplate();
        f.StartPolicyUpdate();
        var planned = f.PlanRecord();
        Assert.Equal("Planned", planned.Status);
        Assert.Single(planned.Keys);
        Assert.StartsWith("folderjob:", planned.Keys[0]);
        Assert.Equal(
            "Example",
            f.Store.Require<OperationDocument>(
                "asx_operation",
                planned.Keys[0]
            ).Value.Folder.Candidate
        );
    }

    [Fact]
    public void ChildFolderNamedFormsIsCreatedBelowTheRoot()
    {
        var f = new Fixture(false);
        var child = JsonWire.Read<FolderStep>(JsonWire.Write(f.Binding));
        child.Key += "-child";
        child.Node = "forms";
        child.ParentBinding = "root";
        child.OriginalName = child.Candidate = "Forms";
        f.Operation.Folders = new[] { f.Binding, child };
        f.Store.Create("asx_operation", f.Operation);
        var root = Guid.NewGuid();
        var work = f.ObserveAndFinalize(f.Preflight(f.Claim()), f.Item(root, null));
        Assert.Equal("Pending", f.Call("Complete", work).Status);
        work = f.Claim("child-run");
        work = f.Observe(work, f.LibraryBody());
        string parent = JsonWire.Write(
            new ODataEnvelope<FolderObservation>
            {
                Data = new FolderObservation { Id = root, Path = "/sites/proto/General/Example" },
            }
        );
        work = f.Observe(work, parent);
        Assert.Equal("ReadyToCreate", f.Observe(work, Rows<ItemObservation>()).Status);
        work = f.Call("PrepareCreate", f.Results.Last());
        Assert.Equal("Create", work.Status);
        work = f.Call("CreateResponse", work, CreateBody(), 200);
        var item = f.Item(Guid.NewGuid(), null);
        item.Name = "Forms";
        item.Path += "/Forms";
        work = f.Observe(work, Rows(item));
        Assert.Equal("FinalParent", work.ProbeKind);
        work = f.Observe(work, parent);
        Assert.Equal("Verified", work.Status);
        Assert.Equal("Applied", f.Call("Complete", work).Status);
    }

    [Fact]
    public void BlankRecordNameWaitsWithANoticeAndIsPlannedOnceFilledIn()
    {
        var f = new Fixture(seedBinding: false);
        f.SeedTemplate();
        f.Service.Rows[f.RecordId]["name"] = null;
        var waiting = f.PlanRecord();
        Assert.Equal("Planned", waiting.Status);
        Assert.Empty(waiting.Keys);
        const string notice = "Folder 'general/root' is waiting for 'root.name' to have a value.";
        Assert.Equal(new[] { notice }, waiting.Notices);
        var inspected = RecordInspection.Read(
            f.Service,
            new WorkerRequest { TemplateId = f.TemplateId, RecordId = f.RecordId },
            new[] { "account" }
        );
        Assert.Equal("NoCurrentOperations", inspected.Status);
        Assert.Contains(notice, inspected.Notices);
        // The record's Update event queues the record again once the field is filled in.
        f.Service.Rows[f.RecordId]["name"] = "Example";
        var planned = f.PlanRecord();
        Assert.Equal("Planned", planned.Status);
        Assert.Empty(planned.Notices);
        Assert.Equal(
            "Example",
            f.Store.Require<OperationDocument>(
                "asx_operation",
                Assert.Single(planned.Keys)
            ).Value.Folder.Candidate
        );
        Assert.DoesNotContain(
            notice,
            RecordInspection
                .Read(
                    f.Service,
                    new WorkerRequest { TemplateId = f.TemplateId, RecordId = f.RecordId },
                    new[] { "account" }
                )
                .Notices
        );
    }

    [Fact]
    public void RecordNameSharePointForbidsIsCleanedAndApplied()
    {
        var f = new Fixture(seedBinding: false);
        f.SeedTemplate();
        f.Service.Rows[f.RecordId]["name"] = "Smith & Sons: Holdings";
        var planned = f.PlanRecord();
        Assert.Equal("Planned", planned.Status);
        const string notice =
            "Folder name 'Smith & Sons: Holdings' was adjusted to 'Smith & Sons- Holdings' for SharePoint.";
        Assert.Equal(new[] { notice }, planned.Notices);
        f.OperationKey = Assert.Single(planned.Keys);
        Assert.Equal(
            "Smith & Sons- Holdings",
            f.Store.Require<OperationDocument>(
                "asx_operation",
                f.OperationKey
            ).Value.Folder.Candidate
        );
        var work = f.Observe(f.Preflight(f.Claim()), Rows<ItemObservation>());
        Assert.Equal("ReadyToCreate", work.Status);
        work = f.Call("PrepareCreate", work);
        Assert.Contains("Smith & Sons- Holdings", work.Http!.Body, StringComparison.Ordinal);
        work = f.Call("CreateResponse", work, CreateBody(), 200);
        var item = f.Item(Guid.NewGuid(), null);
        item.Name = "Smith & Sons- Holdings";
        item.Path = "/sites/proto/General/Smith & Sons- Holdings";
        work = f.ObserveAndFinalize(work, item);
        Assert.Equal("Applied", f.Call("Complete", work).Status);
        var inspected = RecordInspection.Read(
            f.Service,
            new WorkerRequest { TemplateId = f.TemplateId, RecordId = f.RecordId },
            new[] { "account" }
        );
        Assert.Equal("Applied", inspected.Status);
        Assert.Equal("Smith & Sons- Holdings", Assert.Single(inspected.Record!.Folders).Candidate);
        Assert.Contains(notice, inspected.Notices);
    }

    [Fact]
    public void PlanAndClaimSucceedWhenLibraryHasNoPolicyRevisionOrAclFingerprint()
    {
        foreach (var seeded in new[] { false, true })
        {
            var f = new Fixture(seedBinding: seeded);
            f.SeedTemplate();
            f.Service.Rows[f.LibraryId].Attributes.Remove("asx_policyrevision");
            f.Service.Rows[f.LibraryId].Attributes.Remove("asx_aclhash");
            f.Service.Rows[f.LibraryId]["asx_policyapplied"] = false;
            if (seeded)
                Assert.Equal("Library", f.Claim().ProbeKind);
            else
            {
                var planned = f.PlanRecord();
                Assert.Equal("Planned", planned.Status);
                Assert.Single(planned.Keys);
            }
        }
    }

    [Fact]
    public void ClaimAndDriveSucceedWhilePolicyIsBeingApplied()
    {
        var f = new Fixture();
        f.StartPolicyUpdate();
        var work = f.Observe(f.Preflight(f.Claim()), Rows<ItemObservation>());
        work = f.Call("PrepareCreate", work);
        work = f.Call("CreateResponse", work, CreateBody(), 200);
        work = f.ObserveAndFinalize(work, f.Item(Guid.NewGuid(), null));
        Assert.Equal("Applied", f.Call("Complete", work).Status);
        Assert.NotEmpty(f.Results);
        Assert.DoesNotContain(
            f.Results,
            r => r.Http?.RelativeUri.IndexOf("roleassignments", StringComparison.Ordinal) >= 0
        );
        Assert.DoesNotContain(f.Results, r => r.ProbeKind == "Acl" || r.ProbeKind == "FinalAcl");
    }

    [Fact]
    public void SuspendedLibraryStillBlocksPlanning()
    {
        var f = new Fixture();
        f.SeedTemplate();
        f.Service.Rows[f.LibraryId]["asx_approved"] = false;
        AssertPlanAndClaimBlocked(f);
    }

    [Fact]
    public void SuspendedSiteStillBlocksPlanning()
    {
        var f = new Fixture();
        f.SeedTemplate();
        f.Service.Rows[f.SiteId]["asx_approved"] = false;
        AssertPlanAndClaimBlocked(f);
    }

    private static void AssertPlanAndClaimBlocked(Fixture f)
    {
        var queued = f.Service.Transaction(() =>
            f.Coordinator.Execute(
                new WorkerRequest
                {
                    Command = "Queue",
                    TemplateId = f.TemplateId,
                    RecordId = f.RecordId,
                    RequestId = Guid.NewGuid(),
                },
                true
            )
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() =>
                f.Coordinator.Execute(
                    new WorkerRequest { Command = "Plan", Key = queued.Key },
                    true
                )
            )
        );
        Assert.Equal(
            "Pending",
            f.Store.Require<OutboxDocument>("asx_outbox", queued.Key).Value.Status
        );
        Assert.DoesNotContain(
            f.Service.Rows.Values,
            r => r.LogicalName == "asx_operation" && r.Id != f.OperationRowId
        );
        Assert.Throws<EvaluationBlockedException>(() => f.Claim());
        Assert.Equal(
            "Pending",
            f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value.Status
        );
    }

    [Theory]
    [InlineData("Acl", "Parent")]
    [InlineData("FinalAcl", "FinalParent")]
    public void PersistedAclProbeFromAnEarlierVersionAdvancesWithoutReadingPermissions(
        string legacy,
        string next
    )
    {
        var f = new Fixture();
        // An earlier version issued Acl right after the library read, and FinalAcl right after
        // the folder read; persist the operation in that state.
        var work = f.Observe(f.Claim(), f.LibraryBody());
        if (legacy == "FinalAcl")
            work = f.Observe(f.Observe(work, f.ParentBody()), Rows(f.Item(Guid.NewGuid(), null)));
        var stored = f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key);
        stored.Value.ProbeKind = legacy;
        stored.Value.ProbeId = Guid.NewGuid();
        if (legacy == "FinalAcl")
            stored.Value.Status = "NeedsFinalPolicy";
        f.Store.Save(stored);
        work = new WorkerResult
        {
            Status = "Read",
            Key = work.Key,
            Token = work.Token,
            ProbeId = stored.Value.ProbeId,
            ProbeKind = legacy,
        };
        work = f.Observe(work, "{}");
        Assert.Equal("Read", work.Status);
        Assert.Equal(next, work.ProbeKind);
        work = f.Observe(work, f.ParentBody());
        if (legacy == "Acl")
            Assert.Equal("Folder", work.ProbeKind);
        else
        {
            Assert.Equal("Verified", work.Status);
            Assert.Equal("Applied", f.Call("Complete", work).Status);
        }
    }

    public const string PlanningFailedNotice =
        "Planning failed. Check that the site and library are approved, then use Retry in Blocked records.";

    [Fact]
    public void FailOutboxBlocksOnlyPendingRows()
    {
        var f = new Fixture(seedBinding: false);
        f.SeedTemplate();
        var key = f.BlockRecord();
        var blocked = f.Store.Require<OutboxDocument>("asx_outbox", key).Value;
        Assert.Equal("Blocked", blocked.Status);
        Assert.Equal(new[] { PlanningFailedNotice }, blocked.Notices);
        Assert.Equal(
            "Blocked",
            f.Execute(new WorkerRequest { Command = "FailOutbox", Key = key }).Status
        );
        var planned = f.PlanRecord();
        Assert.Equal("Planned", planned.Status);
        var after = f.Execute(new WorkerRequest { Command = "FailOutbox", Key = planned.Key });
        Assert.Equal("Planned", after.Status);
        var row = f.Store.Require<OutboxDocument>("asx_outbox", planned.Key).Value;
        Assert.Equal("Planned", row.Status);
        Assert.Empty(row.Notices);
    }

    [Fact]
    public void OutboxRowsIndexTheirRecordAndTable()
    {
        var f = new Fixture(seedBinding: false);
        f.SeedTemplate();
        var key = f.BlockRecord();
        var row = f.Service.Rows[f.Store.Require<OutboxDocument>("asx_outbox", key).Row.Id];
        Assert.Equal(f.RecordId.ToString("D"), row.GetAttributeValue<string>("asx_recordid"));
        Assert.Equal("account", row.GetAttributeValue<string>("asx_table"));
        Assert.Contains(f.Service.Updates, u => u.Target.Contains("asx_recordid"));
        f.Store.Create(
            "asx_outbox",
            new OutboxDocument { Key = "team-event:index", SecurityTeamId = Guid.NewGuid() }
        );
        var team = f.Service.Rows[
            f.Store.Require<OutboxDocument>("asx_outbox", "team-event:index").Row.Id
        ];
        Assert.Null(team.GetAttributeValue<string>("asx_recordid"));
        Assert.Null(team.GetAttributeValue<string>("asx_table"));
    }

    [Fact]
    public void RetryOutboxReturnsABlockedRowToPendingAndClearsItsNotices()
    {
        var f = new Fixture(seedBinding: false);
        f.SeedTemplate();
        var key = f.BlockRecord();
        var retried = f.Execute(new WorkerRequest { Command = "RetryOutbox", Key = key });
        Assert.Equal("Pending", retried.Status);
        Assert.Equal(key, retried.Key);
        var row = f.Store.Require<OutboxDocument>("asx_outbox", key).Value;
        Assert.Equal("Pending", row.Status);
        Assert.Empty(row.Notices);
        Assert.Equal(new[] { key }, f.Store.Pending("asx_outbox"));
        var planned = f.Execute(new WorkerRequest { Command = "Plan", Key = key });
        Assert.Equal("Planned", planned.Status);
        Assert.Single(planned.Keys);
    }

    [Fact]
    public void RetryOutboxLeavesAnyOtherStatusUnchanged()
    {
        var f = new Fixture(seedBinding: false);
        f.SeedTemplate();
        var planned = f.PlanRecord();
        var before = f.Store.Require<OutboxDocument>("asx_outbox", planned.Key).Row.RowVersion;
        var result = f.Execute(new WorkerRequest { Command = "RetryOutbox", Key = planned.Key });
        Assert.Equal("Planned", result.Status);
        var row = f.Store.Require<OutboxDocument>("asx_outbox", planned.Key);
        Assert.Equal("Planned", row.Value.Status);
        Assert.Equal(before, row.Row.RowVersion);
        Assert.Equal(planned.Keys, row.Value.Operations);
    }

    [Fact]
    public void RetryOutboxRefusesAnUnknownKey()
    {
        var f = new Fixture(seedBinding: false);
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Execute(new WorkerRequest { Command = "RetryOutbox", Key = "request:unknown" })
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Execute(new WorkerRequest { Command = "RetryOutbox", Key = "" })
        );
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void RetryOutboxRunsOnlyWithTheCallersOwnPermissions()
    {
        var f = new Fixture(seedBinding: false);
        f.SeedTemplate();
        RuntimeSeed.Seed(f.Service, Guid.NewGuid(), "account");
        var key = f.BlockRecord();
        Guid operatorId = Guid.NewGuid(),
            other = Guid.NewGuid();
        var asked = new List<Guid?>();
        // Dataverse refuses asx_ManageWork without prvCreateasx_operatorcommand, and every write
        // runs as the caller. A caller without operator privileges cannot change the row.
        Func<Guid?, IOrganizationService> factory = id =>
        {
            asked.Add(id);
            return id == operatorId ? f.Service : new DenyWrites(f.Service);
        };
        Assert.ThrowsAny<Exception>(() => ManageWork(f, factory, other, "RetryOutbox", key));
        Assert.Equal("Blocked", f.Store.Require<OutboxDocument>("asx_outbox", key).Value.Status);
        var result = ManageWork(f, factory, operatorId, "RetryOutbox", key);
        Assert.Equal("Pending", result.Status);
        Assert.Equal("Pending", f.Store.Require<OutboxDocument>("asx_outbox", key).Value.Status);
        Assert.Equal(new Guid?[] { other, operatorId }, asked.ToArray());
    }

    private static WorkerResult ManageWork(
        Fixture f,
        Func<Guid?, IOrganizationService> factory,
        Guid caller,
        string command,
        string key
    )
    {
        var context = GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["Stage"] = 30,
                ["Mode"] = 0,
                ["IsInTransaction"] = true,
                ["UserId"] = caller,
                ["CorrelationId"] = Guid.NewGuid(),
                ["MessageName"] = "asx_ManageWork",
                ["InputParameters"] = new ParameterCollection
                {
                    ["Request"] = JsonWire.Write(
                        new WorkerRequest { Command = command, Key = key }
                    ),
                },
                ["OutputParameters"] = new ParameterCollection(),
                ["SharedVariables"] = new ParameterCollection(),
            }
        );
        return f.Service.Transaction(() =>
        {
            new Ascentix.Documents.Plugins.ManageWorkApi().Execute(
                new ApiProvider(context, factory)
            );
            return JsonWire.Read<WorkerResult>((string)context.OutputParameters["Result"]);
        });
    }

    private sealed class ApiProvider : IServiceProvider, IOrganizationServiceFactory
    {
        private readonly IPluginExecutionContext context;
        private readonly Func<Guid?, IOrganizationService> factory;

        public ApiProvider(
            IPluginExecutionContext context,
            Func<Guid?, IOrganizationService> factory
        )
        {
            this.context = context;
            this.factory = factory;
        }

        public object GetService(Type type) =>
            type == typeof(IPluginExecutionContext) ? context : this;

        public IOrganizationService CreateOrganizationService(Guid? userId) => factory(userId);
    }

    private sealed class DenyWrites : IOrganizationService
    {
        private readonly IOrganizationService inner;

        public DenyWrites(IOrganizationService inner)
        {
            this.inner = inner;
        }

        private static Exception Denied() =>
            new InvalidOperationException("Principal user is missing the operator privilege.");

        public Guid Create(Entity entity) => throw Denied();

        public Entity Retrieve(string name, Guid id, ColumnSet columns) =>
            inner.Retrieve(name, id, columns);

        public void Update(Entity entity) => throw Denied();

        public void Delete(string name, Guid id) => throw Denied();

        public OrganizationResponse Execute(OrganizationRequest request) =>
            request is CreateRequest || request is UpdateRequest || request is DeleteRequest
                ? throw Denied()
                : inner.Execute(request);

        public void Associate(
            string name,
            Guid id,
            Relationship relationship,
            EntityReferenceCollection entities
        ) => throw Denied();

        public void Disassociate(
            string name,
            Guid id,
            Relationship relationship,
            EntityReferenceCollection entities
        ) => throw Denied();

        public EntityCollection RetrieveMultiple(QueryBase query) => inner.RetrieveMultiple(query);
    }

    internal sealed class Fixture
    {
        public MemoryService Service { get; } = new MemoryService();
        public DocumentStore Store => new DocumentStore(Service);
        public WorkerCoordinator Coordinator =>
            new WorkerCoordinator(Service, () => Now, AllowedTables);

        /// <summary>The enabled tables the worker API passes in; null means no scope check.</summary>
        public string[]? AllowedTables;

        /// <summary>The operation Claim and Call drive; null means the seeded one.</summary>
        public string? OperationKey;
        public DateTime Now = new DateTime(2026, 9, 8, 8, 0, 0, DateTimeKind.Utc);
        public Guid TemplateId { get; } = Guid.NewGuid();
        public Guid RevisionId { get; } = Guid.NewGuid();
        public Guid RecordId { get; } = Guid.NewGuid();
        public Guid LibraryId { get; } = Guid.NewGuid();
        public Guid ListId { get; } = Guid.NewGuid();
        public Guid EntryId { get; } = Guid.NewGuid();
        public Guid NativeParent { get; } = Guid.NewGuid();
        public Guid NativeSite { get; } = Guid.NewGuid();
        public Guid SiteId { get; } = Guid.NewGuid();
        public List<WorkerResult> Results { get; } = new List<WorkerResult>();
        public Guid OperationRowId => DocumentStore.StableId("asx_operation:" + Operation.Key);
        public string EntryUrl => "https://example.sharepoint.com/sites/proto/General";
        public FolderStep Binding { get; }
        public OperationDocument Operation { get; }
        public ODataRows<AclAssignment> Acl { get; } =
            new ODataRows<AclAssignment>
            {
                Rows = new[]
                {
                    new AclAssignment
                    {
                        Member = new AclMember { Id = 1, Type = 8 },
                        Roles = new ODataRows<AclRole>
                        {
                            Rows = new[]
                            {
                                new AclRole
                                {
                                    Id = 1073741829,
                                    Permissions = new PermissionMask
                                    {
                                        High = "2147483647",
                                        Low = "4294967295",
                                    },
                                },
                            },
                        },
                    },
                },
            };
        private string run = "run-1";

        public Fixture(bool seedBinding = true)
        {
            Guid site = SiteId,
                policy = Guid.NewGuid();
            Service.Seed(
                new Entity("asx_site", site)
                {
                    ["asx_nativeid"] = new EntityReference("sharepointsite", NativeSite),
                    ["asx_webid"] = Guid.NewGuid().ToString(),
                    ["asx_url"] = "https://example.sharepoint.com/sites/proto",
                    ["asx_approved"] = true,
                }
            );
            Service.Seed(
                new Entity("asx_library", LibraryId)
                {
                    ["asx_siteid"] = new EntityReference("asx_site", site),
                    ["asx_listid"] = ListId.ToString(),
                    ["asx_entryid"] = EntryId.ToString(),
                    ["asx_entryurl"] = EntryUrl,
                    ["asx_nativeparentid"] = new EntityReference(
                        "sharepointdocumentlocation",
                        NativeParent
                    ),
                    ["asx_policyrevision"] = policy.ToString(),
                    ["asx_approved"] = true,
                    ["asx_policyapplied"] = true,
                    ["asx_aclhash"] = SharePointObservations.AclHash(Acl),
                }
            );
            Service.Seed(
                new Entity("sharepointsite", NativeSite)
                {
                    ["absoluteurl"] = "https://example.sharepoint.com/sites/proto",
                    ["statecode"] = new OptionSetValue(0),
                }
            );
            Service.Seed(
                new Entity("sharepointdocumentlocation", NativeParent)
                {
                    ["relativeurl"] = "General",
                    ["parentsiteorlocation"] = new EntityReference("sharepointsite", NativeSite),
                    ["statecode"] = new OptionSetValue(0),
                    ["servicetype"] = new OptionSetValue(0),
                }
            );
            Binding = new FolderStep
            {
                Key = TemplateId.ToString("N") + ":" + RecordId.ToString("N") + ":general:root",
                TemplateId = TemplateId,
                RecordId = RecordId,
                Table = "account",
                Section = "general",
                Node = "root",
                LibraryId = LibraryId,
                EntryId = EntryId,
                OriginalName = "Example",
                Candidate = "Example",
            };
            Operation = new OperationDocument
            {
                Key = Binding.Key + ":revision:" + RevisionId.ToString("N"),
                Folders = new[] { Binding },
                RevisionId = RevisionId,
            };
            Service.Seed(
                new Entity("asx_template", TemplateId)
                {
                    ["asx_table"] = "account",
                    ["asx_publishedrevisionid"] = new EntityReference("asx_revision", RevisionId),
                }
            );
            if (seedBinding)
            {
                Store.Create("asx_operation", Operation);
            }
        }

        public void SeedTemplate()
        {
            Service.Seed(
                new Entity("asx_template", TemplateId)
                {
                    ["asx_name"] = "Accounts",
                    ["asx_table"] = "account",
                    ["asx_publishedrevisionid"] = new EntityReference("asx_revision", RevisionId),
                }
            );
            Service.Seed(
                new Entity("asx_revision", RevisionId)
                {
                    ["asx_templateid"] = new EntityReference("asx_template", TemplateId),
                    ["asx_version"] = 1,
                    ["asx_status"] = "Published",
                }
            );
            Service.Seed(
                new Entity("asx_source", Guid.NewGuid())
                {
                    ["asx_revisionid"] = new EntityReference("asx_revision", RevisionId),
                    ["asx_payload"] = JsonWire.Write(
                        new SourceDto
                        {
                            Table = "account",
                            Columns = new[]
                            {
                                new ColumnDto { Name = "name", Kind = "Text" },
                            },
                        }
                    ),
                }
            );
            Service.Seed(
                new Entity("asx_destination", Guid.NewGuid())
                {
                    ["asx_revisionid"] = new EntityReference("asx_revision", RevisionId),
                    ["asx_key"] = "general",
                    ["asx_libraryid"] = new EntityReference("asx_library", LibraryId),
                }
            );
            Service.Seed(
                new Entity("asx_folder", Guid.NewGuid())
                {
                    ["asx_revisionid"] = new EntityReference("asx_revision", RevisionId),
                    ["asx_key"] = "root",
                    ["asx_sectionkey"] = "general",
                    ["asx_expression"] = "{root.name}",
                    ["asx_order"] = 0,
                }
            );
            Service.Seed(new Entity("account", RecordId) { ["name"] = "Example" });
        }

        /// <summary>
        /// Starts a new library policy generation the way a team-membership refresh does:
        /// the library stays approved while its policy is marked as not yet applied.
        /// </summary>
        public void StartPolicyUpdate()
        {
            Service.Rows[LibraryId]["asx_policyrevision"] = Guid.NewGuid().ToString();
            Service.Rows[LibraryId]["asx_policyapplied"] = false;
        }

        public WorkerResult Execute(WorkerRequest request) =>
            Service.Transaction(() => Coordinator.Execute(request, true));

        /// <summary>
        /// Queues the record while its library is suspended, lets planning fail and records the
        /// failure the way the dispatch flow does, then lifts the suspension.
        /// </summary>
        public string BlockRecord()
        {
            Service.Rows[LibraryId]["asx_approved"] = false;
            var queued = Execute(
                new WorkerRequest
                {
                    Command = "Queue",
                    TemplateId = TemplateId,
                    RecordId = RecordId,
                    RequestId = Guid.NewGuid(),
                }
            );
            Assert.Throws<EvaluationBlockedException>(() =>
                Execute(new WorkerRequest { Command = "Plan", Key = queued.Key })
            );
            Assert.Equal(
                "Blocked",
                Execute(new WorkerRequest { Command = "FailOutbox", Key = queued.Key }).Status
            );
            Service.Rows[LibraryId]["asx_approved"] = true;
            return queued.Key;
        }

        public WorkerResult PlanRecord()
        {
            var queued = Service.Transaction(() =>
                Coordinator.Execute(
                    new WorkerRequest
                    {
                        Command = "Queue",
                        TemplateId = TemplateId,
                        RecordId = RecordId,
                        RequestId = Guid.NewGuid(),
                    },
                    true
                )
            );
            return Service.Transaction(() =>
                Coordinator.Execute(new WorkerRequest { Command = "Plan", Key = queued.Key }, true)
            );
        }

        public WorkerResult Claim(string runId = "run-1", Guid token = default)
        {
            run = runId;
            var result = Service.Transaction(() =>
                Coordinator.Execute(
                    new WorkerRequest
                    {
                        Command = "Claim",
                        Key = OperationKey ?? Operation.Key,
                        RunId = runId,
                        Token = token,
                    },
                    true
                )
            );
            Results.Add(result);
            return result;
        }

        public WorkerResult Call(
            string command,
            WorkerResult previous,
            string? body = null,
            int status = 0
        )
        {
            var result = Service.Transaction(() =>
                Coordinator.Execute(
                    new WorkerRequest
                    {
                        Command = command,
                        Key = OperationKey ?? Operation.Key,
                        RunId = run,
                        Token = previous.Token,
                        ProbeId = previous.ProbeId,
                        ProbeKind = previous.ProbeKind,
                        ResponseBody = body,
                        HttpStatus = status,
                    },
                    true
                )
            );
            Results.Add(result);
            return result;
        }

        public WorkerResult Observe(WorkerResult read, string body) =>
            Call("Observe", read, body, read.ProbeKind == "Folder" && body == "{}" ? 404 : 200);

        public string LibraryBody() =>
            JsonWire.Write(
                new ODataEnvelope<LibraryObservation>
                {
                    Data = new LibraryObservation
                    {
                        Id = ListId,
                        Root = new FolderObservation
                        {
                            Id = EntryId,
                            Path = "/sites/proto/General",
                        },
                    },
                }
            );

        public string ParentBody() =>
            JsonWire.Write(
                new ODataEnvelope<FolderObservation>
                {
                    Data = new FolderObservation { Id = EntryId, Path = "/sites/proto/General" },
                }
            );

        public string AclBody() =>
            JsonWire.Write(new ODataEnvelope<ODataRows<AclAssignment>> { Data = Acl });

        public WorkerResult Preflight(WorkerResult read)
        {
            Assert.Equal("Library", read.ProbeKind);
            read = Observe(read, LibraryBody());
            Assert.Equal("Parent", read.ProbeKind);
            read = Observe(read, ParentBody());
            Assert.Equal("Folder", read.ProbeKind);
            return read;
        }

        public ItemObservation Item(Guid id, Guid? marker) =>
            new ItemObservation
            {
                ItemId = 7,
                Id = id,
                Name = "Example",
                Path = "/sites/proto/General/Example",
                Type = 1,
            };

        public WorkerResult ObserveAndFinalize(WorkerResult read, ItemObservation item)
        {
            read = Observe(read, Rows(item));
            Assert.Equal("FinalParent", read.ProbeKind);
            read = Observe(read, ParentBody());
            Assert.Equal("Verified", read.Status);
            return read;
        }
    }

    internal sealed class MemoryService : IOrganizationService
    {
        public Dictionary<Guid, Entity> Rows { get; private set; } = new Dictionary<Guid, Entity>();
        public List<UpdateRequest> Updates { get; } = new List<UpdateRequest>();
        public string? FailUpdateTable;
        public Func<QueryExpression, EntityCollection?>? QueryHook;

        /// <summary>Lookup columns by "table.column", with the table each one targets.</summary>
        public Dictionary<string, string> Lookups { get; } = new Dictionary<string, string>();
        private long version = 1;

        private static Entity Copy(Entity row, ColumnSet? columns = null)
        {
            var copy = new Entity(row.LogicalName, row.Id) { RowVersion = row.RowVersion };
            foreach (var pair in row.Attributes)
                if (columns == null || columns.AllColumns || columns.Columns.Contains(pair.Key))
                    copy[pair.Key] = pair.Value;
            return copy;
        }

        public void Seed(Entity row)
        {
            var copy = Copy(row);
            copy.RowVersion = (++version).ToString();
            Rows[copy.Id] = copy;
        }

        public T Transaction<T>(Func<T> action)
        {
            var snapshot = Rows.ToDictionary(p => p.Key, p => Copy(p.Value));
            var oldVersion = version;
            try
            {
                return action();
            }
            catch
            {
                Rows = snapshot;
                version = oldVersion;
                throw;
            }
        }

        public Entity Retrieve(string name, Guid id, ColumnSet columns)
        {
            var row = Rows[id];
            if (row.LogicalName != name)
                throw new InvalidOperationException("Wrong entity identity.");
            return Copy(row, columns);
        }

        public Guid Create(Entity row)
        {
            if (row.Id == Guid.Empty)
                row.Id = Guid.NewGuid();
            if (
                Rows.ContainsKey(row.Id)
                || (
                    row.Contains("asx_key")
                    && Rows.Values.Any(r =>
                        r.LogicalName == row.LogicalName
                        && r.GetAttributeValue<string>("asx_key")
                            == row.GetAttributeValue<string>("asx_key")
                        && (
                            !row.Contains("asx_revisionid")
                            || r.GetAttributeValue<EntityReference>("asx_revisionid")?.Id
                                == row.GetAttributeValue<EntityReference>("asx_revisionid")?.Id
                        )
                        && (
                            row.LogicalName != "asx_folder"
                            || r.GetAttributeValue<string>("asx_sectionkey")
                                == row.GetAttributeValue<string>("asx_sectionkey")
                        )
                    )
                )
            )
                throw new InvalidOperationException("Unique key conflict.");
            var copy = Copy(row);
            if (!copy.Contains("createdon"))
                copy["createdon"] = new DateTime(2026, 9, 8).AddMilliseconds(version);
            if (copy.LogicalName == "sharepointdocumentlocation")
                copy["statecode"] = new OptionSetValue(0);
            Seed(copy);
            return copy.Id;
        }

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            if (request is DeleteRequest delete)
            {
                var row = Rows[delete.Target.Id];
                if (
                    delete.ConcurrencyBehavior != ConcurrencyBehavior.IfRowVersionMatches
                    || row.RowVersion != delete.Target.RowVersion
                )
                    throw new InvalidOperationException("Delete version conflict.");
                Delete(delete.Target.LogicalName, delete.Target.Id);
                return new DeleteResponse();
            }
            if (request is CreateRequest create)
            {
                var response = new CreateResponse();
                response.Results["id"] = Create(create.Target);
                return response;
            }
            if (request is UpdateRequest update)
            {
                Updates.Add(update);
                if (update.Target.LogicalName == FailUpdateTable)
                    throw new InvalidOperationException("Injected compare-and-swap failure.");
                var row = Rows[update.Target.Id];
                if (
                    update.ConcurrencyBehavior != ConcurrencyBehavior.IfRowVersionMatches
                    || string.IsNullOrEmpty(update.Target.RowVersion)
                    || row.RowVersion != update.Target.RowVersion
                )
                    throw new InvalidOperationException("Version conflict.");
                var copy = Copy(row);
                foreach (var value in update.Target.Attributes)
                    copy[value.Key] = value.Value;
                Seed(copy);
                return new UpdateResponse();
            }
            if (request is RetrieveEntityRequest entityRequest)
            {
                var metadata = new EntityMetadata { LogicalName = entityRequest.LogicalName };
                typeof(EntityMetadata)
                    .GetProperty("PrimaryIdAttribute")!
                    .SetValue(metadata, entityRequest.LogicalName + "id", null);
                var response = new RetrieveEntityResponse();
                response.Results["EntityMetadata"] = metadata;
                return response;
            }
            if (request is RetrieveAttributeRequest attribute)
            {
                AttributeMetadata metadata = Lookups.TryGetValue(
                    attribute.EntityLogicalName + "." + attribute.LogicalName,
                    out var target
                )
                    ? new LookupAttributeMetadata
                    {
                        LogicalName = attribute.LogicalName,
                        IsSecured = false,
                        Targets = new[] { target },
                    }
                    : new StringAttributeMetadata
                    {
                        LogicalName = attribute.LogicalName,
                        IsSecured = false,
                    };
                typeof(AttributeMetadata)
                    .GetProperty("IsValidForRead")!
                    .SetValue(metadata, true, null);
                var response = new RetrieveAttributeResponse();
                response.Results["AttributeMetadata"] = metadata;
                return response;
            }
            throw new NotSupportedException(request.RequestName);
        }

        public EntityCollection RetrieveMultiple(QueryBase raw)
        {
            var query = (QueryExpression)raw;
            var intercepted = QueryHook?.Invoke(query);
            if (intercepted != null)
                return intercepted;
            IEnumerable<Entity> values = Rows.Values.Where(r => r.LogicalName == query.EntityName);
            bool Condition(Entity row, ConditionExpression condition)
            {
                object? actual =
                    condition.AttributeName == row.LogicalName + "id" ? (object)row.Id
                    : row.Contains(condition.AttributeName) ? row[condition.AttributeName]
                    : null;
                if (actual is EntityReference reference)
                    actual = reference.Id;
                if (actual is OptionSetValue option)
                    actual = option.Value;
                if (condition.Operator == ConditionOperator.Null)
                    return actual == null;
                if (condition.Operator == ConditionOperator.NotNull)
                    return actual != null;
                if (condition.Operator == ConditionOperator.In)
                    return condition.Values.Any(v =>
                        Equals(actual, v is EntityReference r ? r.Id : v)
                    );
                object expected = condition.Values[0];
                if (condition.Operator == ConditionOperator.Equal)
                    return Equals(actual, expected);
                if (condition.Operator == ConditionOperator.LessEqual)
                    return actual is IComparable comparable && comparable.CompareTo(expected) <= 0;
                throw new NotSupportedException(condition.Operator.ToString());
            }
            bool Matches(Entity row, FilterExpression filter)
            {
                var conditions = filter
                    .Conditions.Select(c => Condition(row, c))
                    .Concat(filter.Filters.Select(f => Matches(row, f)))
                    .ToArray();
                return filter.FilterOperator == LogicalOperator.And
                    ? conditions.All(v => v)
                    : conditions.Any(v => v);
            }
            values = values.Where(row => Matches(row, query.Criteria));
            IOrderedEnumerable<Entity>? sorted = null;
            foreach (var order in query.Orders)
            {
                Func<Entity, object?> key = row =>
                    row.Contains(order.AttributeName) ? row[order.AttributeName] : null;
                sorted =
                    sorted == null
                        ? (
                            order.OrderType == OrderType.Descending
                                ? values.OrderByDescending(key)
                                : values.OrderBy(key)
                        )
                        : (
                            order.OrderType == OrderType.Descending
                                ? sorted.ThenByDescending(key)
                                : sorted.ThenBy(key)
                        );
            }
            if (sorted != null)
                values = sorted;
            if (query.TopCount.HasValue)
                values = values.Take(query.TopCount.Value);
            if (query.PageInfo != null && query.PageInfo.Count > 0)
            {
                var all = values.ToList();
                int skip = (Math.Max(1, query.PageInfo.PageNumber) - 1) * query.PageInfo.Count;
                var rows = all.Skip(skip).Take(query.PageInfo.Count).ToList();
                var page = new EntityCollection(rows.Select(r => Copy(r, query.ColumnSet)).ToList())
                {
                    MoreRecords = skip + rows.Count < all.Count,
                };
                page.PagingCookie = page.MoreRecords
                    ? "page-" + (query.PageInfo.PageNumber + 1)
                    : null;
                return page;
            }
            return new EntityCollection(values.Select(r => Copy(r, query.ColumnSet)).ToList());
        }

        public void Update(Entity row) => throw new NotSupportedException("Use UpdateRequest CAS.");

        public void Delete(string name, Guid id)
        {
            if (Rows[id].LogicalName != name)
                throw new InvalidOperationException("Wrong delete table.");
            Rows.Remove(id);
        }

        public void Associate(
            string name,
            Guid id,
            Relationship relationship,
            EntityReferenceCollection entities
        ) => throw new NotSupportedException();

        public void Disassociate(
            string name,
            Guid id,
            Relationship relationship,
            EntityReferenceCollection entities
        ) => throw new NotSupportedException();
    }
}
