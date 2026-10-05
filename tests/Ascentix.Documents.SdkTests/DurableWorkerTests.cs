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
    public void ExpiredClaimsNeverTakeOverWithoutSeparateRecoveryPermit()
    {
        var fixture = new Fixture();
        var claim = fixture.Claim();
        fixture.Now = fixture.Now.AddMinutes(6);
        Assert.Equal("Quarantined", fixture.Claim("run-2").Status);
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
    public void KnownCreateResponseStillRequiresIndependentFolderAndFinalPolicyReads()
    {
        var fixture = new Fixture();
        var absent = fixture.Observe(fixture.Preflight(fixture.Claim()), Rows<ItemObservation>());
        var prepared = fixture.Call("PrepareCreate", absent);
        var read = fixture.Call("CreateResponse", prepared, CreateBody(), 200);
        Assert.Throws<EvaluationBlockedException>(() => fixture.Call("Complete", read));
        var finalAcl = fixture.Observe(read, Rows(fixture.Item(Guid.NewGuid(), null)));
        Assert.Equal("FinalAcl", finalAcl.ProbeKind);
        Assert.Throws<EvaluationBlockedException>(() => fixture.Call("Complete", finalAcl));
    }

    [Fact]
    public void FinalAclDriftBlocksAppliedAndReleasesKnownWriter()
    {
        var fixture = new Fixture();
        var read = fixture.Preflight(fixture.Claim());
        var finalAcl = fixture.Observe(read, Rows(fixture.Item(Guid.NewGuid(), null)));
        var drift = fixture.Observe(finalAcl, Rows<AclAssignment>());
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
        Assert.Equal("Acl", next.ProbeKind);
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
        fixture.Service.Seed(
            new Entity("asx_runtime", Guid.NewGuid())
            {
                ["asx_name"] = "Default",
                ["asx_workeruserid"] = worker.ToString(),
                ["asx_enabled"] = false,
                ["asx_allowedtables"] = "[\"account\"]",
                ["asx_sharepointhosts"] = "[\"example.sharepoint.com\"]",
            }
        );
        var profile = RuntimeProfile.Read(fixture.Service);
        Assert.Equal(worker, profile.WorkerId);
        Assert.False(profile.Enabled);
        Assert.Equal(new[] { "account" }, profile.Tables);
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

    private static string CreateBody() =>
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
    public void RetryAfterHonorsServerDelayAndRejectsUnboundedInput()
    {
        var now = new DateTime(2026, 9, 8, 8, 0, 0, DateTimeKind.Utc);
        Assert.Equal(now.AddHours(1), WorkerCoordinator.RetryAt(now, 1, "3600"));
        Assert.Equal(
            now.AddMinutes(30),
            WorkerCoordinator.RetryAt(now, 1, now.AddMinutes(30).ToString("r"))
        );
        Assert.Equal(now.AddSeconds(60), WorkerCoordinator.RetryAt(now, 2, "1"));
        Assert.Throws<EvaluationBlockedException>(() =>
            WorkerCoordinator.RetryAt(now, 1, "999999999999")
        );
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
    [InlineData(false)]
    [InlineData(true)]
    public void NestedEntryCannotInheritAnUnreviewedIntermediateScope(bool driftAtFinal)
    {
        var f = new Fixture();
        string entry = "/sites/proto/General/Archive/Entry";
        f.Service.Rows[f.LibraryId]["asx_entryurl"] = "https://example.sharepoint.com" + entry;
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
        work = f.Observe(work, f.AclBody());
        string parent = JsonWire.Write(
            new ODataEnvelope<FolderObservation>
            {
                Data = new FolderObservation
                {
                    Id = f.EntryId,
                    Path = entry,
                    Item = new ParentListItem { UniquePermissions = false },
                },
            }
        );
        string ancestor(bool unique) =>
            JsonWire.Write(
                new ODataEnvelope<FolderObservation>
                {
                    Data = new FolderObservation
                    {
                        Id = Guid.NewGuid(),
                        Path = "/sites/proto/General/Archive",
                        Item = new ParentListItem { UniquePermissions = unique },
                    },
                }
            );
        work = f.Observe(work, parent);
        Assert.Equal("Ancestor", work.ProbeKind);
        work = f.Observe(work, ancestor(!driftAtFinal));
        if (driftAtFinal)
        {
            Assert.Equal("Folder", work.ProbeKind);
            var item = f.Item(Guid.NewGuid(), null);
            item.Path = entry + "/Example";
            work = f.Observe(work, Rows(item));
            work = f.Observe(work, f.AclBody());
            work = f.Observe(work, parent);
            Assert.Equal("FinalAncestor", work.ProbeKind);
            work = f.Observe(work, ancestor(true));
        }
        Assert.Equal("Blocked", work.Status);
        Assert.DoesNotContain(
            f.Service.Rows.Values,
            r =>
                r.LogicalName == "sharepointdocumentlocation"
                && r.GetAttributeValue<EntityReference>("regardingobjectid") != null
        );
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
        work = f.Observe(work, f.AclBody());
        string parent = JsonWire.Write(
            new ODataEnvelope<FolderObservation>
            {
                Data = new FolderObservation
                {
                    Id = root,
                    Path = "/sites/proto/General/Example",
                    Item = new ParentListItem { UniquePermissions = false },
                },
            }
        );
        work = f.Observe(work, parent);
        var item = f.Item(Guid.NewGuid(), null);
        item.Name = "Invoices";
        item.Path += "/Invoices";
        work = f.Observe(work, Rows(item));
        work = f.Observe(work, f.AclBody());
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
    public void FreshClaimRevalidatesCurrentApprovedPolicyButInFlightGenerationChangesBlock()
    {
        var f = new Fixture();
        var next = Guid.NewGuid();
        f.Service.Rows[f.LibraryId]["asx_policyrevision"] = next.ToString();
        var work = f.Claim();
        Assert.Equal("Library", work.ProbeKind);
        Assert.Equal(
            next,
            f.Store.Require<OperationDocument>(
                "asx_operation",
                f.Operation.Key
            ).Value.PolicyRevision
        );
        work = f.Preflight(work);
        f.Service.Rows[f.LibraryId]["asx_policyrevision"] = Guid.NewGuid().ToString();
        Assert.Throws<EvaluationBlockedException>(() => f.Observe(work, Rows<ItemObservation>()));
        Assert.False(
            f.Store.Require<OperationDocument>(
                "asx_operation",
                f.Operation.Key
            ).Value.ExternalSubmitted
        );
    }

    internal sealed class Fixture
    {
        public MemoryService Service { get; } = new MemoryService();
        public DocumentStore Store => new DocumentStore(Service);
        public WorkerCoordinator Coordinator => new WorkerCoordinator(Service, () => Now);
        public DateTime Now = new DateTime(2026, 9, 8, 8, 0, 0, DateTimeKind.Utc);
        public Guid TemplateId { get; } = Guid.NewGuid();
        public Guid RevisionId { get; } = Guid.NewGuid();
        public Guid RecordId { get; } = Guid.NewGuid();
        public Guid LibraryId { get; } = Guid.NewGuid();
        public Guid ListId { get; } = Guid.NewGuid();
        public Guid EntryId { get; } = Guid.NewGuid();
        public Guid NativeParent { get; } = Guid.NewGuid();
        public Guid NativeSite { get; } = Guid.NewGuid();
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
            Guid site = Guid.NewGuid(),
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
                PolicyRevision = policy,
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

        public WorkerResult Claim(string runId = "run-1", Guid token = default)
        {
            run = runId;
            return Service.Transaction(() =>
                Coordinator.Execute(
                    new WorkerRequest
                    {
                        Command = "Claim",
                        Key = Operation.Key,
                        RunId = runId,
                        Token = token,
                    },
                    true
                )
            );
        }

        public WorkerResult Call(
            string command,
            WorkerResult previous,
            string? body = null,
            int status = 0
        ) =>
            Service.Transaction(() =>
                Coordinator.Execute(
                    new WorkerRequest
                    {
                        Command = command,
                        Key = Operation.Key,
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
            Assert.Equal("Acl", read.ProbeKind);
            read = Observe(read, AclBody());
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
                UniquePermissions = false,
            };

        public WorkerResult ObserveAndFinalize(WorkerResult read, ItemObservation item)
        {
            read = Observe(read, Rows(item));
            Assert.Equal("FinalAcl", read.ProbeKind);
            read = Observe(read, AclBody());
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
                AttributeMetadata metadata = new StringAttributeMetadata
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
