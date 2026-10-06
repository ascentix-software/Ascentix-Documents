using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class SecurityWorkerTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BrowserApplyPayloadWithoutRoleObjectsUsesVerifiedLibraryRoles(bool explicitNull)
    {
        var f = new Fixture();
        f.Service.Rows[f.Library]["asx_readrole"] = JsonWire.Write(f.Read);
        f.Service.Rows[f.Library]["asx_contributerole"] = JsonWire.Write(f.Contribute);
        var json =
            "{\"Command\":\"ApplyPolicy\",\"LibraryId\":\""
            + f.Library
            + "\",\"RowVersion\":null,\"Entries\":[{\"TeamId\":\""
            + f.Team
            + "\",\"Access\":\"Read\"}]}";
        if (explicitNull)
            json =
                json.Substring(0, json.Length - 1) + ",\"ReadRole\":null,\"ContributeRole\":null}";
        var request = JsonWire.Read<SecurityRequest>(json);
        var result = f.Service.Transaction(() => f.Admin.Execute(request, true));
        Assert.Equal("Queued", result.Status);
        f.Key = result.Policy!.OperationKey!;
        f.Drive();
        Assert.Equal("Read", f.Policy().Applied.Single().Access);
    }

    [Fact]
    public void FirstSyncCreatesGroupWithExactTeamNameAndEscapesLookup()
    {
        var f = new Fixture();
        f.Service.Rows[f.Team]["name"] = "Director's Operations";
        f.Queue("Read");
        var work = f.Start();
        while (work.Status == "Read" && work.ProbeKind != "SecurityGroup")
            work = f.Observe(work);
        Assert.Contains(
            "Title eq 'Director''s Operations'",
            Uri.UnescapeDataString(work.Http!.RelativeUri)
        );
        work = f.Observe(work);
        var operation = f.Operation();
        Assert.Contains("Director's Operations", operation.Mutation!.Body);
        f.Call("PrepareCreate", work);
        f.Apply();
        Assert.Equal("Director's Operations", f.Group!.Title);
    }

    [Fact]
    public void SameNamedUnownedGroupIsNeverAdopted()
    {
        var f = new Fixture();
        f.Group = new SiteGroup
        {
            Id = 42,
            Title = "Operations",
            Description = "Unrelated group",
            Type = 8,
        };
        f.Queue("Read");
        Assert.Equal("Blocked", f.Drive(false).Status);
        Assert.Empty(f.Writes);
    }

    [Fact]
    public void SubsequentSyncUsesOwnedIdDespiteDisplayNameChanges()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        f.Group!.Title = "Renamed in SharePoint";
        f.Service.Rows[f.Team]["name"] = "Renamed team";
        f.Writes.Clear();
        f.Queue("Read");
        var work = f.Start();
        while (work.Status == "Read" && work.ProbeKind != "SecurityGroup")
            work = f.Observe(work);
        Assert.Contains("Id eq 42", Uri.UnescapeDataString(work.Http!.RelativeUri));
        while (work.Status == "Read")
            work = f.Observe(work);
        Assert.DoesNotContain("GroupCreate", f.Writes);
        Assert.NotEqual("Blocked", work.Status);
    }

    [Fact]
    public void ApplyOnboardsTeamAndSynchronizesMembersWithoutSeparateRegistration()
    {
        var f = new Fixture();
        f.AddUser();
        var registration = f.Store.Require<TeamRegistration>(
            "asx_teamregistration",
            "team:" + f.Team.ToString("N")
        );
        f.Service.Rows.Remove(registration.Row.Id);
        var result = f.Service.Transaction(() =>
            f.Admin.Execute(
                new SecurityRequest
                {
                    Command = "ApplyPolicy",
                    LibraryId = f.Library,
                    Entries = new[]
                    {
                        new PolicyEntry { TeamId = f.Team, Access = "Read" },
                    },
                    ReadRole = f.Read,
                    ContributeRole = f.Contribute,
                },
                true
            )
        );
        Assert.True(
            f.Store.Require<TeamRegistration>(
                "asx_teamregistration",
                "team:" + f.Team.ToString("N")
            ).Value.Enabled
        );
        f.Key = result.Policy!.OperationKey!;
        f.Drive();
        Assert.Single(f.Members);
        Assert.Equal("Read", f.Policy().Applied.Single().Access);
    }

    [Fact]
    public void DraftDoesNotOnboardTeamAndFailedApplyRollsBackRegistration()
    {
        var f = new Fixture();
        var key = "team:" + f.Team.ToString("N");
        f.Service.Rows.Remove(
            f.Store.Require<TeamRegistration>("asx_teamregistration", key).Row.Id
        );
        var draft = f.Service.Transaction(() =>
            f.Admin.Execute(
                new SecurityRequest
                {
                    Command = "SavePolicy",
                    LibraryId = f.Library,
                    Entries = new[]
                    {
                        new PolicyEntry { TeamId = f.Team, Access = "Read" },
                    },
                    ReadRole = f.Read,
                    ContributeRole = f.Contribute,
                },
                true
            )
        );
        Assert.Null(f.Store.Find<TeamRegistration>("asx_teamregistration", key));
        f.FailTeamRead();
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() =>
                f.Admin.Execute(
                    new SecurityRequest
                    {
                        Command = "ApplyPolicy",
                        LibraryId = f.Library,
                        RowVersion = draft.RowVersion,
                        Entries = draft.Policy!.Desired,
                        ReadRole = f.Read,
                        ContributeRole = f.Contribute,
                    },
                    true
                )
            )
        );
        Assert.Null(f.Store.Find<TeamRegistration>("asx_teamregistration", key));
    }

    [Fact]
    public void LibraryGrantRemovalDoesNotOptTeamOutOfMembershipSync()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        var row = f.Store.Require<PolicyDocument>(
            "asx_policy",
            "policy:" + f.Library.ToString("N")
        );
        var result = f.Service.Transaction(() =>
            f.Admin.Execute(
                new SecurityRequest
                {
                    Command = "ApplyPolicy",
                    LibraryId = f.Library,
                    RowVersion = row.Row.RowVersion,
                    Entries = Array.Empty<PolicyEntry>(),
                    ReadRole = f.Read,
                    ContributeRole = f.Contribute,
                },
                true
            )
        );
        f.Key = result.Policy!.OperationKey!;
        f.Drive();
        Assert.Empty(f.Acl);
        Assert.True(
            f.Store.Require<TeamRegistration>(
                "asx_teamregistration",
                "team:" + f.Team.ToString("N")
            ).Value.Enabled
        );
        Assert.Single(f.Members);
    }

    [Fact]
    public void ApplyPolicySavesAndQueuesInOneRequestThenWorkerVerifiesAccess()
    {
        var f = new Fixture();
        var result = f.Service.Transaction(() =>
            f.Admin.Execute(
                new SecurityRequest
                {
                    Command = "ApplyPolicy",
                    LibraryId = f.Library,
                    Entries = new[]
                    {
                        new PolicyEntry { TeamId = f.Team, Access = "Read" },
                    },
                    ReadRole = f.Read,
                    ContributeRole = f.Contribute,
                },
                true
            )
        );
        Assert.Equal("Queued", result.Status);
        Assert.Empty(f.Acl);
        Assert.Empty(f.Writes);
        f.Key = f.Policy().OperationKey!;
        f.Drive();
        Assert.Equal("Read", f.Policy().Applied.Single().Access);
    }

    [Fact]
    public void ApplyPolicyRollsBackDraftWhenQueueValidationFails()
    {
        var f = new Fixture();
        f.FailTeamRead();
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() =>
                f.Admin.Execute(
                    new SecurityRequest
                    {
                        Command = "ApplyPolicy",
                        LibraryId = f.Library,
                        Entries = new[]
                        {
                            new PolicyEntry { TeamId = f.Team, Access = "Read" },
                        },
                        ReadRole = f.Read,
                        ContributeRole = f.Contribute,
                    },
                    true
                )
            )
        );
        Assert.Null(
            f.Store.Find<PolicyDocument>("asx_policy", "policy:" + f.Library.ToString("N"))
        );
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_operation");
    }

    [Fact]
    public void ApplyPolicyRejectsStaleEditsAndCannotReplaceQueuedWork()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        var row = f.Store.Require<PolicyDocument>(
            "asx_policy",
            "policy:" + f.Library.ToString("N")
        );
        var request = new SecurityRequest
        {
            Command = "ApplyPolicy",
            LibraryId = f.Library,
            RowVersion = "stale",
            Entries = new[]
            {
                new PolicyEntry { TeamId = f.Team, Access = "Contribute" },
            },
            ReadRole = f.Read,
            ContributeRole = f.Contribute,
        };
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() => f.Admin.Execute(request, true))
        );
        Assert.Equal("Read", f.Policy().Desired.Single().Access);
        request.RowVersion = row.Row.RowVersion;
        f.Service.Transaction(() => f.Admin.Execute(request, true));
        var key = f.Policy().OperationKey;
        request.RowVersion = f
            .Store.Require<PolicyDocument>("asx_policy", row.Value.Key)
            .Row.RowVersion;
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() => f.Admin.Execute(request, true))
        );
        Assert.Equal(key, f.Policy().OperationKey);
    }

    [Fact]
    public void SecurityGrantDowngradeAndRevokeUseDurableOwnedReceipts()
    {
        var f = new Fixture();
        f.Queue("Contribute");
        f.Drive();
        Assert.Equal(3, f.Acl.Single().Roles.Rows.Single().Id);
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Equal(new[] { "GrantRemove", "GrantAdd" }, f.Writes);
        Assert.Equal(2, f.Acl.Single().Roles.Rows.Single().Id);
        f.Writes.Clear();
        f.Queue("None");
        f.Drive();
        Assert.Equal(new[] { "GrantRemove" }, f.Writes);
        Assert.Empty(f.Acl);
        var policy = f.Policy();
        Assert.Equal("Applied", policy.Status);
        Assert.Equal("None", policy.Applied.Single().Access);
        Assert.NotEmpty(policy.ResidualAccess);
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_managedgroup");
    }

    [Fact]
    public void MembershipIsIndependentlyVerifiedAndOptOutRevokesManagedAccess()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        Assert.Single(f.Members);
        Assert.Contains("MemberAdd", f.Writes);
        var registration = f.Store.Require<TeamRegistration>(
            "asx_teamregistration",
            "team:" + f.Team.ToString("N")
        );
        registration.Value.Enabled = false;
        f.Store.Save(registration);
        f.Queue("Read");
        f.Drive();
        Assert.Empty(f.Members);
        Assert.Empty(f.Acl);
        Assert.Equal("None", f.Policy().Applied.Single().Access);
    }

    [Fact]
    public void AccessQueuesWhileAnotherRunWritesOnTheSite()
    {
        var f = new Fixture();
        AdminStopTests.HoldWriter(
            f.Store,
            WorkCoordination.Library(f.Service, f.Library),
            DateTime.UtcNow
        );
        f.Queue("Read");
        Assert.NotNull(f.Policy().OperationKey);
        Assert.Equal("Pending", f.Operation().Status);
    }

    [Fact]
    public void ThrottledSecurityReadsKeepWaitingWithNoAttemptCap()
    {
        var f = new Fixture();
        f.Queue("Read");
        for (int attempt = 1; attempt <= 7; attempt++)
        {
            var work = f.Start();
            Assert.Equal("Read", work.Status);
            Assert.Equal("RetryWait", f.Call("Observe", work, status: 429).Status);
            var op = f.Store.Require<SecurityOperation>("asx_operation", f.Key);
            Assert.Equal(attempt, op.Value.RetryCount);
            Assert.Equal(
                "Waiting to retry after a temporary error (HTTP 429); attempt " + attempt + ".",
                op.Value.ErrorCode
            );
            op.Value.NextAttemptUtc = DateTime.UtcNow.AddSeconds(-1);
            f.Store.Save(op);
        }
        Assert.Equal("Applied", f.Drive().Status);
    }

    [Fact]
    public void TransientSecurityFailuresWaitAndGenuineOnesBlock()
    {
        var f = new Fixture();
        f.Queue("Read");
        var unclaimed = f.Service.Transaction(() =>
            new SecurityWorker(f.Service).Execute(
                new WorkerRequest
                {
                    Command = "FailUnclaimed",
                    Key = f.Key,
                    StatusCode = 503,
                },
                true
            )
        );
        Assert.Equal("RetryWait", unclaimed.Status);
        var op = f.Store.Require<SecurityOperation>("asx_operation", f.Key);
        op.Value.NextAttemptUtc = DateTime.UtcNow.AddSeconds(-1);
        f.Store.Save(op);
        var work = f.Start();
        var fail = new WorkerRequest
        {
            Command = "Fail",
            Key = f.Key,
            RunId = "security/run-1",
            Token = work.Token,
            StatusCode = 0,
        };
        Assert.Equal(
            "RetryWait",
            f.Service.Transaction(() => new SecurityWorker(f.Service).Execute(fail, true)).Status
        );
        op = f.Store.Require<SecurityOperation>("asx_operation", f.Key);
        op.Value.NextAttemptUtc = DateTime.UtcNow.AddSeconds(-1);
        f.Store.Save(op);
        fail.Token = f.Start().Token;
        fail.StatusCode = 400;
        fail.ErrorCode = "0x80040265";
        Assert.Equal(
            "Blocked",
            f.Service.Transaction(() => new SecurityWorker(f.Service).Execute(fail, true)).Status
        );
    }

    [Fact]
    public void ThrottledSecurityWriteWasNotExecutedSoItWaitsAndRetries()
    {
        var f = new Fixture();
        f.Queue("Read");
        var work = f.Start();
        while (work.Status == "Read")
            work = f.Observe(work);
        work = f.Call("PrepareCreate", work);
        Assert.Equal("Create", work.Status);
        var throttled = f.Call("CreateResponse", work, status: 429);
        Assert.Equal("RetryWait", throttled.Status);
        var op = f.Operation();
        Assert.False(op.ExternalSubmitted);
        Assert.False(op.ExternalResponseKnown);
        Assert.Null(op.Mutation);
        Assert.Contains("HTTP 429", op.ErrorCode);
        Assert.Null(
            f.Store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(f.Service, f.Key)
            ).Value.RunId
        );
        Assert.Empty(f.Writes);
        var stored = f.Store.Require<SecurityOperation>("asx_operation", f.Key);
        stored.Value.NextAttemptUtc = DateTime.UtcNow.AddSeconds(-1);
        f.Store.Save(stored);
        Assert.Equal("Applied", f.Drive().Status);
    }

    private static void ExpireClaim(Fixture f)
    {
        var claim = f.Store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(f.Service, f.Key)
        );
        claim.Value.LeaseUntilUtc = DateTime.UtcNow.AddMinutes(-1);
        f.Store.Save(claim);
    }

    [Theory]
    [InlineData("GroupCreate", false)]
    [InlineData("GroupCreate", true)]
    [InlineData("MemberAdd", false)]
    [InlineData("MemberAdd", true)]
    [InlineData("GrantAdd", false)]
    [InlineData("GrantAdd", true)]
    public void UnknownSecurityWriteIsReadBackAndReconciledAfterTheLease(string kind, bool applied)
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        var work = f.Start();
        for (int i = 0; ; i++)
        {
            Assert.True(i < 200, "Mutation " + kind + " never prepared.");
            if (work.Status == "Read")
                work = f.Observe(work);
            else if (work.Status == "ReadyToCreate" && f.Operation().MutationKind == kind)
                break;
            else if (work.Status == "ReadyToCreate")
                work = f.Call("PrepareCreate", work);
            else if (work.Status == "Create")
            {
                f.Apply();
                work = f.Call("CreateResponse", work, status: 200);
            }
            else
                throw new Exception(work.Status);
        }
        work = f.Call("PrepareCreate", work);
        Assert.Equal("Create", work.Status);
        if (applied)
            f.Apply();
        Assert.Equal("Quarantined", f.Call("CreateResponse", work, status: 0).Status);
        Assert.Equal("Quarantined", f.Start().Status);
        ExpireClaim(f);
        Assert.Equal("Applied", f.Drive().Status);
        Assert.Equal(1, f.Writes.Count(w => w == kind));
        Assert.Single(f.Members);
        Assert.Equal(2, f.Acl.Single(a => a.Member.Id == 42).Roles.Rows.Single().Id);
    }

    private static WorkerResult Manage(Fixture f, string command) =>
        f.Service.Transaction(() =>
            new SecurityWorker(f.Service).Execute(
                new WorkerRequest { Command = command, Key = f.Key },
                true
            )
        );

    [Fact]
    public void CancelWaitsForALiveClaimAndWorksAfterItExpires()
    {
        var f = new Fixture();
        f.Queue("Read");
        Assert.Equal("Read", f.Start().Status);
        Assert.ThrowsAny<Exception>(() => Manage(f, "Cancel"));
        ExpireClaim(f);
        Assert.Equal("Cancelled", Manage(f, "Cancel").Status);
        Assert.Null(
            f.Store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(f.Service, f.Key)
            ).Value.RunId
        );
    }

    [Fact]
    public void RetryAndCancelWorkAfterAnUnknownWriteOnceTheClaimExpires()
    {
        var f = new Fixture();
        f.Queue("Read");
        var work = f.Start();
        while (work.Status == "Read")
            work = f.Observe(work);
        work = f.Call("PrepareCreate", work);
        var kind = f.Operation().MutationKind;
        f.Apply();
        Assert.Equal("Quarantined", f.Call("CreateResponse", work, status: 0).Status);
        Assert.ThrowsAny<Exception>(() => Manage(f, "Retry"));
        ExpireClaim(f);
        Assert.Equal("Pending", Manage(f, "Retry").Status);
        Assert.Equal("Applied", f.Drive().Status);
        Assert.Equal(1, f.Writes.Count(w => w == kind));

        var g = new Fixture();
        g.Queue("Read");
        work = g.Start();
        while (work.Status == "Read")
            work = g.Observe(work);
        work = g.Call("PrepareCreate", work);
        Assert.Equal("Quarantined", g.Call("CreateResponse", work, status: 0).Status);
        ExpireClaim(g);
        Assert.Equal("Cancelled", Manage(g, "Cancel").Status);
        Assert.Equal("NeedsReview", g.Policy().Status);
    }

    [Fact]
    public void UnknownSecurityWriteCannotRepeatOrReleaseTheDispatcher()
    {
        var f = new Fixture();
        f.Queue("Read");
        var work = f.Start();
        while (work.Status == "Read")
            work = f.Observe(work);
        Assert.Equal("ReadyToCreate", work.Status);
        work = f.Call("PrepareCreate", work);
        Assert.Equal("Create", work.Status);
        var unknown = f.Call("CreateResponse", work, status: 0);
        Assert.Equal("Quarantined", unknown.Status);
        Assert.Equal(
            "Claimed",
            f.Store.Require<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(f.Service, work.Key)
            ).Value.Status
        );
        Assert.Throws<EvaluationBlockedException>(() => f.Call("PrepareCreate", work));
    }

    [Fact]
    public void SharingEntriesAndHandAddedPeopleOnTheLibraryNeverBlockTeamSync()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        // A file share adds a Limited Access entry and an admin adds a person by hand.
        var limited = new AclAssignment
        {
            Member = new AclMember { Id = 501, Type = 1 },
            Roles = new ODataRows<AclRole>
            {
                Rows = new[]
                {
                    new AclRole
                    {
                        Id = 1073741825,
                        Permissions = new PermissionMask { High = "0", Low = "0" },
                    },
                },
            },
        };
        var handAdded = f.Grant(502, f.Contribute.Id);
        handAdded.Member.Type = 1;
        f.Acl.Add(limited);
        f.Acl.Add(handAdded);
        f.Service.Seed(
            new Entity("systemuser", Guid.NewGuid())
            {
                ["domainname"] = "second@example.com",
                ["azureactivedirectoryobjectid"] = Guid.NewGuid(),
                ["isdisabled"] = false,
            }
        );
        f.Writes.Clear();
        f.Queue("Contribute");
        var work = f.Start();
        for (int i = 0; i < 200 && work.Status != "Applied" && work.Status != "Blocked"; i++)
        {
            if (work.Status == "Read")
                work = f.Observe(work);
            else if (work.Status == "ReadyToCreate")
                work = f.Call("PrepareCreate", work);
            else if (work.Status == "Create")
            {
                string kind = f.Operation().MutationKind;
                f.Apply();
                // Another share lands while Documents is writing its own grant.
                if (kind == "GrantAdd")
                    f.Acl.Add(f.Grant(503, f.Read.Id));
                work = f.Call("CreateResponse", work, status: 200);
            }
            else if (work.Status == "Verified")
                work = f.Call("Complete", work);
        }
        Assert.True(work.Status == "Applied", string.Join(" ", work.Notices));
        Assert.Equal(2, f.Members.Count);
        Assert.Equal(3, f.Acl.Single(a => a.Member.Id == 42).Roles.Rows.Single().Id);
        Assert.Contains(f.Acl, a => a.Member.Id == 501);
        Assert.Contains(f.Acl, a => a.Member.Id == 502);
        Assert.Contains(f.Acl, a => a.Member.Id == 503);
        Assert.Equal(new[] { "MemberAdd", "GrantRemove", "GrantAdd" }, f.Writes);
        Assert.True(f.Service.Rows[f.Library].GetAttributeValue<bool>("asx_policyapplied"));
    }

    [Fact]
    public void GrantLevelChangedByHandIsReconciledBackWithANotice()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        f.SetRoles(42, f.Contribute.Id);
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Equal(new[] { f.Read.Id }, f.Roles(42));
        Assert.Equal(new[] { "GrantRemove", "GrantAdd" }, f.Writes);
        var notice = f.Policy().Notices.Single();
        Assert.Contains("Operations", notice);
        Assert.Contains("changed outside Documents", notice);
        Assert.Contains("Contribute", notice);
        // The next run finds the configured level and has nothing to report.
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Empty(f.Writes);
        Assert.Empty(f.Policy().Notices);
    }

    [Fact]
    public void GrantRemovedByHandIsGrantedAgainWithANotice()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        f.SetRoles(42);
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Equal(new[] { f.Read.Id }, f.Roles(42));
        Assert.Equal(new[] { "GrantAdd" }, f.Writes);
        Assert.Contains("changed outside Documents", f.Policy().Notices.Single());
    }

    [Fact]
    public void ExtraBindingOnTheDocumentsGroupIsRemovedWithANotice()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        f.SetRoles(42, f.Read.Id, 1073741829);
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Equal(new[] { f.Read.Id }, f.Roles(42));
        Assert.Equal(new[] { "GrantRemove" }, f.Writes);
        Assert.Contains("1073741829", f.Policy().Notices.Single());
    }

    [Fact]
    public void ExistingGrantOnANewDocumentsGroupIsReconciledToTheConfiguredLevel()
    {
        var f = new Fixture();
        f.SetRoles(42, f.Contribute.Id);
        f.Queue("Read");
        f.Drive();
        Assert.Equal(new[] { f.Read.Id }, f.Roles(42));
        Assert.Equal(new[] { "GroupCreate", "GrantRemove", "GrantAdd" }, f.Writes);
        Assert.Contains("changed outside Documents", f.Policy().Notices.Single());
    }

    [Fact]
    public void EntraGroupInsideTheDocumentsGroupIsLeftInPlaceWithANotice()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        var entraGroup = new SitePerson
        {
            Id = 900,
            Login = "c:0t.c|tenant|5f0c3a2e-0000-0000-0000-000000000001",
            Type = 4,
        };
        f.Members.Add(entraGroup);
        f.Members.Add(
            new SitePerson
            {
                Id = 901,
                Login = "i:0#.f|membership|handadded@example.com",
                Type = 1,
            }
        );
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Contains(entraGroup, f.Members);
        Assert.DoesNotContain(f.Members, m => m.Id == 901);
        Assert.Equal(new[] { "MemberRemove" }, f.Writes);
        var notice = f.Policy().Notices.Single();
        Assert.Contains("5f0c3a2e-0000-0000-0000-000000000001", notice);
        Assert.Contains("left in place", notice);
    }

    [Fact]
    public void CancelAfterAnUnknownGrantThatLandedIsReconciledByTheNextRun()
    {
        var f = new Fixture();
        f.Queue("Read");
        var work = f.Start();
        for (int i = 0; ; i++)
        {
            Assert.True(i < 200, "GrantAdd never prepared.");
            if (work.Status == "Read")
                work = f.Observe(work);
            else if (work.Status == "ReadyToCreate" && f.Operation().MutationKind == "GrantAdd")
                break;
            else if (work.Status == "ReadyToCreate")
                work = f.Call("PrepareCreate", work);
            else if (work.Status == "Create")
            {
                f.Apply();
                work = f.Call("CreateResponse", work, status: 200);
            }
            else
                throw new Exception(work.Status);
        }
        work = f.Call("PrepareCreate", work);
        f.Apply();
        Assert.Equal("Quarantined", f.Call("CreateResponse", work, status: 0).Status);
        ExpireClaim(f);
        Assert.Equal("Cancelled", Manage(f, "Cancel").Status);
        f.Writes.Clear();
        f.Queue("Read");
        Assert.Equal("Applied", f.Drive().Status);
        Assert.Empty(f.Writes);
        Assert.Equal(new[] { f.Read.Id }, f.Roles(42));
        Assert.Empty(f.Policy().Notices);
    }

    [Fact]
    public void GuestTeamMemberIsAddedToTheGroupWithItsExternalSignInName()
    {
        var f = new Fixture();
        f.AddPerson("Partner_Fabrikam.com#EXT#@contoso.onmicrosoft.com", "Pat Partner");
        f.Queue("Read");
        f.Drive();
        Assert.Equal(
            "i:0#.f|membership|partner_fabrikam.com#ext#@contoso.onmicrosoft.com",
            f.Members.Single().Login
        );
        Assert.Empty(f.Policy().Notices);
    }

    [Fact]
    public void MembersDocumentsCannotResolveAreSkippedWithANoticeAndTheRestSync()
    {
        var f = new Fixture();
        f.AddUser();
        f.AddPerson("app@example.com", "Integration App", application: Guid.NewGuid());
        f.AddPerson("noentra@example.com", "No Entra", entra: Guid.Empty);
        f.AddPerson("bad|claim@example.com", "Claim Breaker");
        f.Queue("Read");
        var result = f.Drive();
        Assert.Equal("person@example.com", f.Members.Single().Login.Split('|').Last());
        var notices = f.Policy().Notices;
        Assert.Equal(3, notices.Length);
        Assert.Contains(notices, n => n.Contains("Integration App") && n.Contains("application"));
        Assert.Contains(notices, n => n.Contains("No Entra") && n.Contains("Entra object ID"));
        Assert.Contains(notices, n => n.Contains("Claim Breaker") && n.Contains("sign-in name"));
        Assert.All(notices, n => Assert.Contains("Operations", n));
        Assert.All(notices, n => Assert.Contains(n, result.Notices));
    }

    [Fact]
    public void TeamSnapshotSkipsUnusableMembersInsteadOfRejectingTheTeam()
    {
        var f = new Fixture();
        f.AddUser();
        f.AddPerson("app@example.com", "Integration App", application: Guid.NewGuid());
        var snapshot = new TeamSnapshotReader(f.Service).Snapshot(f.Team);
        Assert.Single(snapshot.People);
        Assert.Contains("Integration App", snapshot.Skipped.Single());
        Assert.Single(new TeamSnapshotReader(f.Service).Read(f.Team));
    }

    [Fact]
    public void OneMemberSharePointRejectsIsSkippedWithItsMessageAndOthersAreAdded()
    {
        var f = new Fixture();
        f.AddUser();
        f.AddPerson("ghost@example.com", "Ghost");
        f.AddPerson("third@example.com", "Third");
        f.Queue("Read");
        var work = f.Start();
        for (int i = 0; i < 200 && work.Status != "Applied" && work.Status != "Blocked"; i++)
        {
            if (work.Status == "Read")
                work = f.Observe(work);
            else if (work.Status == "ReadyToCreate")
                work = f.Call("PrepareCreate", work);
            else if (work.Status == "Create")
            {
                var op = f.Operation();
                if (op.MutationKind == "MemberAdd" && op.MutationLogin.Contains("ghost"))
                    work = f.Call(
                        "CreateResponse",
                        work,
                        "{\"error\":{\"code\":\"-2146232832, Microsoft.SharePoint.SPException\",\"message\":{\"lang\":\"en-US\",\"value\":\"The user does not exist or is not unique.\"}}}",
                        400
                    );
                else
                {
                    f.Apply();
                    work = f.Call("CreateResponse", work, status: 200);
                }
            }
            else if (work.Status == "Verified")
                work = f.Call("Complete", work);
        }
        Assert.True(work.Status == "Applied", string.Join(" ", work.Notices));
        Assert.Equal(2, f.Members.Count);
        Assert.DoesNotContain(f.Members, m => m.Login.Contains("ghost"));
        var notice = f.Policy().Notices.Single();
        Assert.Contains("ghost@example.com", notice);
        Assert.Contains("The user does not exist or is not unique.", notice);
        Assert.Equal(1, f.Writes.Count(w => w == "GrantAdd"));
    }

    [Theory]
    [InlineData(
        "{\"odata.error\":{\"code\":\"-1\",\"message\":{\"lang\":\"en-US\",\"value\":\"User not found.\"}}}",
        "User not found."
    )]
    [InlineData("{\"error\":{\"code\":\"-1\",\"message\":\"Plain message\"}}", "HTTP 404.")]
    [InlineData("<html>Not found</html>", "HTTP 404.")]
    [InlineData(null, "HTTP 404.")]
    public void SharePointErrorMessageFallsBackToTheStatus(string? body, string expected) =>
        Assert.Equal(
            expected,
            SharePointObservations.ErrorMessage(
                new WorkerRequest { HttpStatus = 404, ResponseBody = body }
            )
        );

    [Fact]
    public void LimitedAccessOnTheDocumentsGroupIsLeftAloneWithNoWriteOrBlock()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        // SharePoint adds Limited Access when an item is shared with the Documents group.
        f.SetRoles(42, f.Read.Id, Fixture.LimitedAccess);
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Empty(f.Writes);
        Assert.Empty(f.Policy().Notices);
        Assert.Equal(new[] { f.Read.Id, Fixture.LimitedAccess }, f.Roles(42));
    }

    [Fact]
    public void OnlyTheDocumentsGroupAssignmentIsReadSoLargeLibrariesNeedNoPaging()
    {
        var f = new Fixture();
        for (int i = 0; i < 25000; i++)
        {
            var other = f.Grant(1000 + i, f.Read.Id);
            other.Member.Type = 1;
            f.Acl.Add(other);
        }
        f.Queue("Read");
        f.Drive();
        Assert.NotEmpty(f.AclReads);
        Assert.All(
            f.AclReads,
            uri =>
            {
                Assert.Contains("/roleassignments/getbyprincipalid(42)?", uri);
                Assert.Contains("RoleDefinitionBindings/RoleTypeKind", uri);
                Assert.DoesNotContain("skiptoken", uri);
            }
        );
        // The first read finds no assignment (404) and Documents grants it.
        Assert.Equal(404, f.AclStatuses.First());
        Assert.Equal(new[] { f.Read.Id }, f.Roles(42));
    }

    [Theory]
    [InlineData(401)]
    [InlineData(403)]
    public void PermissionLossOnAMemberWriteBlocks(int status)
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Reject = op => op.MutationKind == "MemberAdd" ? status : (int?)null;
        var result = f.Drive(false);
        Assert.Equal("Blocked", result.Status);
        Assert.Contains("SecurityWriteRejected", result.Notices);
    }

    [Fact]
    public void RejectedMemberIsRetriedByTheNextScheduledRefresh()
    {
        var f = new Fixture();
        f.AddUser();
        f.AddPerson("ghost@example.com", "Ghost");
        f.Queue("Read");
        f.Reject = op =>
            op.MutationKind == "MemberAdd" && op.MutationLogin.Contains("ghost") ? 400 : (int?)null;
        f.Drive();
        Assert.Single(f.Members);
        Assert.Contains("ghost@example.com", f.Policy().Notices.Single());
        f.Reject = null;
        f.Writes.Clear();
        var refreshed = f.Service.Transaction(() =>
            new SecurityRefresh(f.Service, () => DateTime.UtcNow.AddDays(2)).Scan()
        );
        f.Key = refreshed.Keys.Single();
        f.Drive();
        Assert.Equal(new[] { "MemberAdd" }, f.Writes);
        Assert.Equal(2, f.Members.Count);
        Assert.Empty(f.Policy().Notices);
        // Nothing is skipped now, so the next refresh has nothing to do.
        Assert.Empty(
            f.Service.Transaction(() =>
                new SecurityRefresh(f.Service, () => DateTime.UtcNow.AddDays(4)).Scan()
            ).Keys
        );
    }

    [Fact]
    public void RejectedMemberRemovalIsSkippedWithANotice()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        f.Members.Add(
            new SitePerson
            {
                Id = 901,
                Login = "i:0#.f|membership|handadded@example.com",
                Type = 1,
            }
        );
        f.Reject = op => op.MutationKind == "MemberRemove" ? 404 : (int?)null;
        f.Queue("Read");
        f.Drive();
        Assert.Contains(f.Members, m => m.Id == 901);
        var notice = f.Policy().Notices.Single();
        Assert.Contains("did not remove handadded@example.com", notice);
        Assert.Contains("The user does not exist or is not unique.", notice);
    }

    [Fact]
    public void MemberSharePointStoresUnderAnotherSignInNameIsSkippedAndKept()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        // SharePoint accepts the add but lists the person under a renamed sign-in name.
        f.Mutate = op =>
        {
            if (op.MutationKind != "MemberAdd")
                return false;
            f.Members.Add(
                new SitePerson
                {
                    Id = 77,
                    Login = "i:0#.f|membership|person.renamed@example.com",
                    Type = 1,
                }
            );
            return true;
        };
        f.Drive();
        Assert.Contains(f.Members, m => m.Id == 77);
        Assert.Equal(new[] { "MemberAdd" }, f.Writes.Where(w => w.StartsWith("Member")));
        Assert.Contains("person@example.com", f.Policy().Notices.Single());
    }

    [Fact]
    public void GrantMissingAfterASuccessfulWriteWaitsAndReadsAgain()
    {
        var f = new Fixture();
        f.Queue("Read");
        bool delayed = true;
        f.Mutate = op => op.MutationKind == "GrantAdd" && delayed;
        var work = f.Drive(false);
        Assert.Equal("RetryWait", work.Status);
        // SharePoint shows the grant on the next read.
        delayed = false;
        f.SetRoles(42, f.Read.Id);
        Due(f);
        Assert.Equal("Applied", f.Drive().Status);
        Assert.Equal(1, f.Writes.Count(w => w == "GrantAdd"));
    }

    [Fact]
    public void GrantStillMissingOnTheThirdReadBackBlocks()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Mutate = op => op.MutationKind == "GrantAdd";
        Assert.Equal("RetryWait", f.Drive(false).Status);
        Due(f);
        Assert.Equal("RetryWait", f.Drive(false).Status);
        Due(f);
        var blocked = f.Drive(false);
        Assert.Equal("Blocked", blocked.Status);
        Assert.Contains("Grant independent readback", blocked.Notices.Single());
        Assert.Equal(1, f.Writes.Count(w => w == "GrantAdd"));
    }

    [Fact]
    public void RetryAfterReadBackMissesReadsAgainAndWritesTheGrantAgain()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Mutate = op => op.MutationKind == "GrantAdd";
        Assert.Equal("RetryWait", f.Drive(false).Status);
        Due(f);
        Assert.Equal("RetryWait", f.Drive(false).Status);
        Due(f);
        Assert.Equal("Blocked", f.Drive(false).Status);
        f.Mutate = null;
        Assert.Equal("Pending", Manage(f, "Retry").Status);
        Assert.Equal(0, f.Operation().ReadbackMisses);
        Assert.Equal("Applied", f.Drive().Status);
        Assert.Equal(2, f.Writes.Count(w => w == "GrantAdd"));
        Assert.Equal(new[] { f.Read.Id }, f.Roles(42));
    }

    private static void Due(Fixture f)
    {
        var op = f.Store.Require<SecurityOperation>("asx_operation", f.Key);
        op.Value.NextAttemptUtc = DateTime.UtcNow.AddSeconds(-1);
        f.Store.Save(op);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TeamMembersSharingAnIdentityAreAllSkipped(bool sameEntra)
    {
        var f = new Fixture();
        f.AddUser();
        var entra = Guid.NewGuid();
        f.AddPerson(
            sameEntra ? "one@example.com" : "twin@example.com",
            "Twin One",
            sameEntra ? entra : (Guid?)null
        );
        f.AddPerson("twin@example.com", "Twin Two", sameEntra ? entra : (Guid?)null);
        var snapshot = new TeamSnapshotReader(f.Service).Snapshot(f.Team);
        Assert.Equal("i:0#.f|membership|person@example.com", snapshot.People.Single().Login);
        Assert.Equal(2, snapshot.Skipped.Length);
        Assert.Contains(snapshot.Skipped, n => n.Contains("Twin One"));
        Assert.Contains(snapshot.Skipped, n => n.Contains("Twin Two"));
    }

    [Fact]
    public void DocumentsGroupDeletedInSharePointIsRecreatedWithANotice()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        f.Group = null;
        f.Members.Clear();
        f.SetRoles(42);
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Equal(new[] { "GroupCreate", "MemberAdd", "GrantAdd" }, f.Writes);
        Assert.Equal(43, f.Group!.Id);
        Assert.Equal(new[] { f.Read.Id }, f.Roles(43));
        Assert.Single(f.Members);
        Assert.Contains("deleted", f.Policy().Notices.Single());
    }

    [Fact]
    public void DocumentsGroupDescriptionEditedInSharePointKeepsTheGroupWithANotice()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        f.Group!.Description = "Edited by an admin";
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Empty(f.Writes);
        Assert.Contains("description", f.Policy().Notices.Single());
    }

    [Fact]
    public void SkipDetailsStopAtThePayloadSizeAndTheRestWaitForTheNextRefresh()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        // People added by hand with long sign-in names make the run's stored operation large.
        for (int i = 0; i < 800; i++)
            f.Members.Add(
                new SitePerson
                {
                    Id = 1000 + i,
                    Login = "i:0#.f|membership|" + new string('a', 420) + i + "@example.com",
                    Type = 1,
                }
            );
        f.Reject = op => op.MutationKind == "MemberRemove" ? 400 : (int?)null;
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Equal(new[] { "RejectedMemberRemove" }, f.Writes);
        Assert.Contains(f.Policy().Notices, n => n.Contains("more than one run can record"));
        Assert.True(f.Operation().SkippedCount == 1);
        var group = f
            .Store.Require<ManagedGroup>(
                "asx_managedgroup",
                "group:" + f.Site.ToString("N") + ":" + f.Team.ToString("N")
            )
            .Value;
        Assert.StartsWith("retry:", group.MembershipHash);
    }

    [Fact]
    public void ChangedTeamSnapshotBlocksBeforeMembershipRemoval()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        var work = f.Start();
        while (work.ProbeKind != "SecurityMembers")
        {
            if (work.Status == "Read")
                work = f.Observe(work);
            else if (work.Status == "ReadyToCreate")
                work = f.Call("PrepareCreate", work);
            else
            {
                f.Apply();
                work = f.Call("CreateResponse", work, status: 200);
            }
        }
        f.Service.Rows[f.User]["domainname"] = "changed@example.com";
        var result = f.Observe(work);
        Assert.Equal("Blocked", result.Status);
        Assert.DoesNotContain("MemberRemove", f.Writes);
    }

    [Fact]
    public void SecurityPagingCannotEscapeOrChangeTheReviewedQuery()
    {
        var web = new Uri("https://example.sharepoint.com/sites/proto");
        string first = "_api/web/sitegroups(42)/users?$select=Id,LoginName&$top=500";
        string next = web.AbsoluteUri + "/" + first + "&$skiptoken=Paged%3DTRUE%26p_ID%3D500";
        var relative = SecurityPaging.Next(web, first, next, new[] { first });
        Assert.Contains("skiptoken", relative);
        Assert.Throws<EvaluationBlockedException>(() =>
            SecurityPaging.Next(web, first, next, new[] { first, relative })
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            SecurityPaging.Next(
                web,
                first,
                next.Replace("sitegroups(42)", "sitegroups(43)"),
                new[] { first }
            )
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            SecurityPaging.Next(
                web,
                first,
                next.Replace("example.sharepoint.com", "other.sharepoint.com"),
                new[] { first }
            )
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            SecurityPaging.Next(web, first, next.Replace("$top=500", "$top=1"), new[] { first })
        );
    }

    [Fact]
    public void SecurityQueueRequiresCurrentDiffAndIdleWriter()
    {
        var f = new Fixture();
        f.Queue("Read");
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Admin.Execute(
                new SecurityRequest
                {
                    Command = "QueuePolicy",
                    LibraryId = f.Library,
                    RowVersion = "stale",
                },
                true
            )
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Admin.Execute(
                new SecurityRequest
                {
                    Command = "SavePolicy",
                    LibraryId = f.Library,
                    Entries = Array.Empty<PolicyEntry>(),
                    RowVersion = f
                        .Store.Require<PolicyDocument>(
                            "asx_policy",
                            "policy:" + f.Library.ToString("N")
                        )
                        .Row.RowVersion,
                    ReadRole = f.Read,
                    ContributeRole = f.Contribute,
                },
                true
            )
        );
    }

    [Fact]
    public void AutomaticRefreshCannotPublishDraftAccessOrRoleChanges()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        var row = f.Store.Require<PolicyDocument>(
            "asx_policy",
            "policy:" + f.Library.ToString("N")
        );
        f.Admin.Execute(
            new SecurityRequest
            {
                Command = "SavePolicy",
                LibraryId = f.Library,
                RowVersion = row.Row.RowVersion,
                Entries = new[]
                {
                    new PolicyEntry { TeamId = f.Team, Access = "Contribute" },
                },
                ReadRole = new PolicyRole
                {
                    Id = 22,
                    High = "176",
                    Low = "138612833",
                },
                ContributeRole = new PolicyRole
                {
                    Id = 33,
                    High = "432",
                    Low = "1011028719",
                },
            },
            true
        );
        f.AddUser();
        new SecurityRefresh(f.Service, () => DateTime.UtcNow.AddDays(2)).Scan();
        f.Key = f.Policy().OperationKey!;
        Assert.Equal("Read", f.Operation().Entries.Single().Access);
        Assert.Equal(2, f.Operation().ReadRole.Id);
        Assert.Equal("Contribute", f.Policy().Desired.Single().Access);
        f.Drive();
        Assert.Equal("Read", f.Policy().Applied.Single().Access);
    }

    [Fact]
    public void TeamReaderConsumesEveryCookiePageAndRejectsIncompleteContinuation()
    {
        var f = new Fixture();
        var a = new Entity("systemuser", Guid.NewGuid())
        {
            ["domainname"] = "a@example.com",
            ["azureactivedirectoryobjectid"] = Guid.NewGuid(),
        };
        var b = new Entity("systemuser", Guid.NewGuid())
        {
            ["domainname"] = "b@example.com",
            ["azureactivedirectoryobjectid"] = Guid.NewGuid(),
        };
        int calls = 0;
        f.Service.QueryHook = query =>
        {
            if (query.EntityName != "systemuser")
                return null;
            calls++;
            Assert.False(query.ColumnSet.AllColumns);
            Assert.Single(query.LinkEntities);
            if (query.PageInfo.PageNumber == 1)
                return new EntityCollection(new[] { a })
                {
                    MoreRecords = true,
                    PagingCookie = "cookie-1",
                };
            Assert.Equal("cookie-1", query.PageInfo.PagingCookie);
            return new EntityCollection(new[] { b });
        };
        Assert.Equal(2, new TeamSnapshotReader(f.Service).Read(f.Team).Length);
        Assert.Equal(2, calls);
        f.Service.QueryHook = query =>
            query.EntityName == "systemuser"
                ? new EntityCollection(new[] { a }) { MoreRecords = true }
                : null;
        Assert.Throws<EvaluationBlockedException>(() =>
            new TeamSnapshotReader(f.Service).Read(f.Team)
        );
    }

    [Fact]
    public void HashSupportsFullBoundedTeamsAndDoesNotDependOnPageOrder()
    {
        var users = Enumerable
            .Range(0, 2000)
            .Select(i => new TeamPerson
            {
                UserId = Guid.NewGuid(),
                EntraId = Guid.NewGuid(),
                Login = "i:0#.f|membership|user" + i + "@example.com",
            })
            .ToArray();
        Assert.Equal(
            TeamSnapshotReader.Hash(users),
            TeamSnapshotReader.Hash(Enumerable.Reverse(users).ToArray())
        );
        Assert.Equal(64, TeamSnapshotReader.Hash(users).Length);
    }

    [Fact]
    public void ForeignOrIncompleteMembershipPageNeverProducesRemovals()
    {
        var f = new Fixture();
        f.Queue("Read");
        var work = f.Start();
        while (work.ProbeKind != "SecurityMembers")
        {
            if (work.Status == "Read")
                work = f.Observe(work);
            else if (work.Status == "ReadyToCreate")
                work = f.Call("PrepareCreate", work);
            else
            {
                f.Apply();
                work = f.Call("CreateResponse", work, status: 200);
            }
        }
        var page = JsonWire.Write(
            new ODataEnvelope<ODataRows<SitePerson>>
            {
                Data = new ODataRows<SitePerson>
                {
                    Rows = new[]
                    {
                        new SitePerson
                        {
                            Id = 7,
                            Login = "i:0#.f|membership|unwanted@example.com",
                            Type = 1,
                        },
                    },
                    Next = "https://other.sharepoint.com/_api/web/sitegroups(42)/users",
                },
            }
        );
        var result = f.Call("Observe", work, page);
        Assert.Equal("Blocked", result.Status);
        Assert.DoesNotContain("MemberRemove", f.Writes);
    }

    [Theory]
    [InlineData(true, 4u)]
    [InlineData(false, 2048u)]
    [InlineData(false, 33554432u)]
    [InlineData(false, 1073741824u)]
    public void NamedRolesRejectWriteOrAdministrativePermissionsOutsideTheirMeaning(
        bool read,
        uint added
    )
    {
        var f = new Fixture();
        var role = read ? f.Read : f.Contribute;
        role.Low = (uint.Parse(role.Low) | added).ToString(
            System.Globalization.CultureInfo.InvariantCulture
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            SecurityAdministration.ValidateRole(role, read)
        );
    }

    [Fact]
    public void RolesRequireDocumentAccessAndCanonicalMasks()
    {
        Assert.Throws<EvaluationBlockedException>(() =>
            SecurityAdministration.ValidateRole(
                new PolicyRole
                {
                    Id = 2,
                    High = "0",
                    Low = "1",
                },
                true
            )
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            SecurityAdministration.ValidateRole(
                new PolicyRole
                {
                    Id = 2,
                    High = "0176",
                    Low = "138612833",
                },
                true
            )
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            SecurityAdministration.ValidateRole(
                new PolicyRole
                {
                    Id = 3,
                    High = "176",
                    Low = "138612833",
                },
                false
            )
        );
        SecurityAdministration.ValidateRole(
            new PolicyRole
            {
                Id = 2,
                High = "0",
                Low = "200737",
            },
            true
        );
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TeamChangeReplacesOnlyIdleUnsubmittedSnapshots(bool submitted)
    {
        var f = new Fixture();
        f.Queue("Read");
        string old = f.Key;
        var operation = f.Store.Require<SecurityOperation>("asx_operation", old);
        f.Store.Create(
            "asx_membership",
            new MembershipDocument
            {
                Key =
                    "membership:"
                    + operation.Value.PolicyRevision.ToString("N")
                    + ":"
                    + f.Team.ToString("N"),
                Generation = operation.Value.PolicyRevision,
                Desired = Array.Empty<TeamPerson>(),
            }
        );
        operation.Value.Status = "Blocked";
        operation.Value.ExternalSubmitted = submitted;
        f.Store.Save(operation);
        f.AddUser();
        var result = f.Service.Transaction(() =>
            new SecurityRefresh(f.Service, () => DateTime.UtcNow.AddDays(2)).Scan()
        );
        if (submitted)
        {
            Assert.Empty(result.Keys);
            Assert.Equal(old, f.Policy().OperationKey);
            Assert.Equal("Blocked", f.Operation().Status);
        }
        else
        {
            Assert.Single(result.Keys);
            Assert.Equal("Cancelled", f.Operation().Status);
            Assert.NotEqual(old, f.Policy().OperationKey);
            f.Key = f.Policy().OperationKey!;
            f.Drive();
            Assert.Single(f.Members);
        }
    }

    [Fact]
    public void TeamOutboxPersistsPageUntilAllAffectedPoliciesAreInspected()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        f.AddUser();
        f.Store.Create(
            "asx_policy",
            new PolicyDocument
            {
                Key = "unused",
                Status = "Draft",
                LibraryId = Guid.NewGuid(),
            }
        );
        f.Store.Create(
            "asx_policyentry",
            new PolicyTeamReference
            {
                Key = "unused-index",
                TeamId = f.Team,
                PolicyKey = "unused",
                Status = "Active",
            }
        );
        var actual = f
            .Store.Require<PolicyTeamReference>(
                "asx_policyentry",
                "policyteam:" + f.Library.ToString("N") + ":" + f.Team.ToString("N")
            )
            .Row;
        f.Service.QueryHook = q =>
        {
            if (q.EntityName != "asx_policyentry" || q.PageInfo.Count != 5)
                return null;
            if (q.PageInfo.PageNumber == 1)
                return new EntityCollection(
                    new[]
                    {
                        f.Store.Require<PolicyTeamReference>("asx_policyentry", "unused-index").Row,
                    }
                )
                {
                    MoreRecords = true,
                    PagingCookie = "policy-page-2",
                };
            Assert.Equal(2, q.PageInfo.PageNumber);
            Assert.Equal("policy-page-2", q.PageInfo.PagingCookie);
            return new EntityCollection(new[] { actual });
        };
        f.Store.Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = "team-change",
                SecurityTeamId = f.Team,
                SecurityPage = 1,
            }
        );
        var coordinator = new WorkerCoordinator(f.Service);
        var first = f.Service.Transaction(() =>
            coordinator.Execute(new WorkerRequest { Command = "Plan", Key = "team-change" }, true)
        );
        Assert.Equal("Pending", first.Status);
        Assert.Equal(
            2,
            f.Store.Require<OutboxDocument>("asx_outbox", "team-change").Value.SecurityPage
        );
        var second = f.Service.Transaction(() =>
            new WorkerCoordinator(f.Service).Execute(
                new WorkerRequest { Command = "Plan", Key = "team-change" },
                true
            )
        );
        Assert.Equal("Planned", second.Status);
        Assert.Single(second.Keys);
        Assert.Equal(
            second.Keys,
            f.Store.Require<OutboxDocument>("asx_outbox", "team-change").Value.Operations
        );
        Assert.Equal(
            second.Keys,
            coordinator
                .Execute(new WorkerRequest { Command = "Plan", Key = "team-change" }, true)
                .Keys
        );
    }

    [Fact]
    public void RecordPlannedInTheSameRunAsATeamChangeIsNotBlocked()
    {
        // TEST 2026-10-05: a user joined a registered team; the team event issued a new policy
        // generation (asx_policyapplied=false) and every record planned before it applied Blocked.
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        Assert.True(f.Service.Rows[f.Library].GetAttributeValue<bool>("asx_policyapplied"));
        var applied = f.Service.Rows[f.Library].GetAttributeValue<string>("asx_policyrevision");
        Guid template = Guid.NewGuid(),
            record = Guid.NewGuid();
        SeedRecordTemplate(f, template, record);
        f.AddUser();
        f.Store.Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = "team-event",
                SecurityTeamId = f.Team,
                SecurityPage = 1,
            }
        );
        var refreshed = f.Service.Transaction(() =>
            new WorkerCoordinator(f.Service).Execute(
                new WorkerRequest { Command = "Plan", Key = "team-event" },
                true
            )
        );
        Assert.Equal("Planned", refreshed.Status);
        Assert.StartsWith("policywork:", Assert.Single(refreshed.Keys));
        Assert.False(f.Service.Rows[f.Library].GetAttributeValue<bool>("asx_policyapplied"));
        Assert.NotEqual(
            applied,
            f.Service.Rows[f.Library].GetAttributeValue<string>("asx_policyrevision")
        );
        var queued = f.Service.Transaction(() =>
            new WorkerCoordinator(f.Service).Execute(
                new WorkerRequest
                {
                    Command = "Queue",
                    TemplateId = template,
                    RecordId = record,
                    RequestId = Guid.NewGuid(),
                },
                true
            )
        );
        var planned = f.Service.Transaction(() =>
            new WorkerCoordinator(f.Service).Execute(
                new WorkerRequest { Command = "Plan", Key = queued.Key },
                true
            )
        );
        Assert.Equal("Planned", planned.Status);
        Assert.StartsWith("folderjob:", Assert.Single(planned.Keys));
    }

    private static void SeedRecordTemplate(Fixture f, Guid template, Guid record)
    {
        Guid revision = Guid.NewGuid(),
            nativeSite = Guid.NewGuid(),
            nativeParent = Guid.NewGuid();
        f.Service.Rows[f.Site]["asx_nativeid"] = new EntityReference("sharepointsite", nativeSite);
        f.Service.Rows[f.Library]["asx_entryid"] = Guid.NewGuid().ToString();
        f.Service.Rows[f.Library]["asx_nativeparentid"] = new EntityReference(
            "sharepointdocumentlocation",
            nativeParent
        );
        f.Service.Seed(
            new Entity("sharepointsite", nativeSite)
            {
                ["absoluteurl"] = "https://example.sharepoint.com/sites/proto",
                ["statecode"] = new OptionSetValue(0),
            }
        );
        f.Service.Seed(
            new Entity("sharepointdocumentlocation", nativeParent)
            {
                ["relativeurl"] = "General",
                ["parentsiteorlocation"] = new EntityReference("sharepointsite", nativeSite),
                ["statecode"] = new OptionSetValue(0),
                ["servicetype"] = new OptionSetValue(0),
            }
        );
        f.Service.Seed(
            new Entity("asx_template", template)
            {
                ["asx_name"] = "Accounts",
                ["asx_table"] = "account",
                ["asx_publishedrevisionid"] = new EntityReference("asx_revision", revision),
            }
        );
        f.Service.Seed(
            new Entity("asx_revision", revision)
            {
                ["asx_templateid"] = new EntityReference("asx_template", template),
                ["asx_version"] = 1,
                ["asx_status"] = "Published",
            }
        );
        f.Service.Seed(
            new Entity("asx_source", Guid.NewGuid())
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", revision),
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
        f.Service.Seed(
            new Entity("asx_destination", Guid.NewGuid())
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", revision),
                ["asx_key"] = "general",
                ["asx_libraryid"] = new EntityReference("asx_library", f.Library),
            }
        );
        f.Service.Seed(
            new Entity("asx_folder", Guid.NewGuid())
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", revision),
                ["asx_key"] = "root",
                ["asx_sectionkey"] = "general",
                ["asx_expression"] = "{root.name}",
                ["asx_order"] = 0,
            }
        );
        f.Service.Seed(new Entity("account", record) { ["name"] = "Example" });
    }

    [Fact]
    public void IncompleteTeamPolicyPageNeverConsumesInvalidation()
    {
        var f = new Fixture();
        f.Store.Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = "team-change",
                SecurityTeamId = f.Team,
                SecurityPage = 1,
            }
        );
        f.Service.QueryHook = q =>
            q.EntityName == "asx_policyentry" && q.PageInfo.Count == 5
                ? new EntityCollection { MoreRecords = true }
                : null;
        Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() =>
                new WorkerCoordinator(f.Service).Execute(
                    new WorkerRequest { Command = "Plan", Key = "team-change" },
                    true
                )
            )
        );
        Assert.Equal(
            "Pending",
            f.Store.Require<OutboxDocument>("asx_outbox", "team-change").Value.Status
        );
    }

    [Fact]
    public void SecurityGroupTeamGrantsItsGroupThroughTheDocumentsGroupWithoutReadingMembers()
    {
        var f = new Fixture();
        f.MakeGroupTeam(2);
        // A Dataverse member the policy run must never read: the group itself is granted.
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        string claim = "c:0t.c|tenant|" + f.GroupObject.ToString("D");
        Assert.Equal(new[] { "GroupCreate", "PrincipalEnsure", "MemberAdd", "GrantAdd" }, f.Writes);
        var ensure = f.Posts.Single(p => p.RelativeUri == "_api/web/ensureuser");
        Assert.Equal("POST", ensure.Method);
        Assert.Equal("{\"logonName\":\"" + claim + "\"}", ensure.Body);
        Assert.Equal(claim, f.Members.Single().Login);
        Assert.Equal(new[] { f.Read.Id }, f.Roles(42));
        Assert.Equal(0, f.SystemUserReads);
        Assert.Empty(f.Policy().Notices);
        // The next run finds the group in place and writes nothing.
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Empty(f.Writes);
        Assert.Equal(0, f.SystemUserReads);
    }

    [Theory]
    [InlineData(0, "")]
    [InlineData(1, "")]
    [InlineData(2, "_o")]
    public void Microsoft365GroupTeamUsesTheMembersOrOwnersClaim(int membership, string suffix)
    {
        var f = new Fixture();
        f.MakeGroupTeam(3, membership);
        f.Queue("Contribute");
        f.Drive();
        Assert.Equal(
            "c:0o.c|federateddirectoryclaimprovider|" + f.GroupObject.ToString("D") + suffix,
            f.Members.Single().Login
        );
        Assert.Equal(new[] { f.Contribute.Id }, f.Roles(42));
        Assert.Equal(0, f.SystemUserReads);
    }

    [Theory]
    [InlineData(3, 3, true, "only the guests")]
    [InlineData(2, 3, true, "only the guests")]
    [InlineData(2, 2, true, "only the owners")]
    [InlineData(2, 0, false, "object ID")]
    [InlineData(3, 1, false, "object ID")]
    public void GroupTeamsSharePointCannotIdentifyAreRefusedWithTheReason(
        int type,
        int membership,
        bool objectId,
        string reason
    )
    {
        var f = new Fixture();
        f.MakeGroupTeam(type, membership, objectId);
        var applied = Assert.Throws<EvaluationBlockedException>(() => f.Queue("Read"));
        Assert.Contains(reason, applied.Message);
        Assert.Contains("Operations", applied.Message);
        if (objectId)
            Assert.Contains("SharePoint has no sign-in claim", applied.Message);
        var registered = Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() =>
                f.Admin.Execute(
                    new SecurityRequest
                    {
                        Command = "RegisterTeam",
                        TeamId = f.Team,
                        Enabled = true,
                    },
                    true
                )
            )
        );
        Assert.Equal(applied.Message, registered.Message);
    }

    [Fact]
    public void AccessAndDefaultTeamsAreStillRefused()
    {
        var f = new Fixture();
        f.Service.Rows[f.Team]["teamtype"] = new OptionSetValue(1);
        Assert.Contains(
            "Access teams",
            Assert.Throws<EvaluationBlockedException>(() => f.Queue("Read")).Message
        );
        f.Service.Rows[f.Team]["teamtype"] = new OptionSetValue(2);
        f.Service.Rows[f.Team]["azureactivedirectoryobjectid"] = f.GroupObject;
        f.Service.Rows[f.Team]["isdefault"] = true;
        Assert.Contains(
            "default team",
            Assert.Throws<EvaluationBlockedException>(() => f.Queue("Read")).Message
        );
    }

    [Fact]
    public void HandEditsToAGroupTeamsDocumentsGroupKeepOtherPrincipalsAndRestoreTheGroup()
    {
        var f = new Fixture();
        f.MakeGroupTeam(2);
        f.Queue("Read");
        f.Drive();
        string claim = f.Members.Single().Login;
        // An admin removes the team's group and adds another Entra group by hand.
        f.Members.Clear();
        var other = new SitePerson
        {
            Id = 900,
            Login = "c:0t.c|tenant|5f0c3a2e-0000-0000-0000-000000000001",
            Type = 4,
        };
        f.Members.Add(other);
        f.Writes.Clear();
        f.Queue("Read");
        f.Drive();
        Assert.Equal(new[] { "PrincipalEnsure", "MemberAdd" }, f.Writes);
        Assert.Contains(other, f.Members);
        Assert.Contains(f.Members, m => m.Login == claim);
        var notice = f.Policy().Notices.Single();
        Assert.Contains("5f0c3a2e-0000-0000-0000-000000000001", notice);
        Assert.Contains("left in place", notice);
        Assert.Equal(0, f.SystemUserReads);
    }

    [Fact]
    public void RemovingAGroupTeamFromTheLibraryRemovesItsGrant()
    {
        var f = new Fixture();
        f.MakeGroupTeam(3, 1);
        f.Queue("Read");
        f.Drive();
        f.Writes.Clear();
        f.Queue("None");
        f.Drive();
        Assert.Empty(f.Roles(42));
        Assert.Equal(new[] { "GrantRemove" }, f.Writes);
        Assert.Equal(0, f.SystemUserReads);
    }

    [Fact]
    public void TeamMembershipEventOnAGroupTeamReadsNoMembersAndQueuesNothing()
    {
        var f = new Fixture();
        f.MakeGroupTeam(2);
        f.Queue("Read");
        f.Drive();
        // A person signs in and Dataverse adds them to the group team.
        f.AddUser();
        f.Store.Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = "team-change",
                SecurityTeamId = f.Team,
                SecurityPage = 1,
            }
        );
        var planned = f.Service.Transaction(() =>
            new WorkerCoordinator(f.Service).Execute(
                new WorkerRequest { Command = "Plan", Key = "team-change" },
                true
            )
        );
        Assert.Equal("Planned", planned.Status);
        Assert.Empty(planned.Keys);
        Assert.Equal(0, f.SystemUserReads);
        Assert.Null(f.Policy().OperationKey);
    }

    [Fact]
    public void ChangedGroupClaimReplacesThePrincipalOnTheNextRefresh()
    {
        var f = new Fixture();
        f.MakeGroupTeam(3, 1);
        f.Queue("Read");
        f.Drive();
        f.Service.Rows[f.Team]["membershiptype"] = new OptionSetValue(2);
        f.Writes.Clear();
        var refreshed = f.Service.Transaction(() =>
            new SecurityRefresh(f.Service, () => DateTime.UtcNow.AddDays(2)).Scan()
        );
        f.Key = refreshed.Keys.Single();
        f.Drive();
        Assert.Equal(new[] { "MemberRemove", "PrincipalEnsure", "MemberAdd" }, f.Writes);
        Assert.Equal(
            "c:0o.c|federateddirectoryclaimprovider|" + f.GroupObject.ToString("D") + "_o",
            f.Members.Single().Login
        );
        Assert.Empty(f.Policy().Notices);
        Assert.Equal(0, f.SystemUserReads);
    }

    [Fact]
    public void GroupTeamThatCanNoLongerBeIdentifiedLosesItsPrincipalWithANotice()
    {
        var f = new Fixture();
        f.MakeGroupTeam(3, 0);
        f.Queue("Read");
        f.Drive();
        f.Service.Rows[f.Team]["membershiptype"] = new OptionSetValue(3);
        f.Writes.Clear();
        var refreshed = f.Service.Transaction(() =>
            new SecurityRefresh(f.Service, () => DateTime.UtcNow.AddDays(2)).Scan()
        );
        f.Key = refreshed.Keys.Single();
        f.Drive();
        Assert.Equal(new[] { "MemberRemove" }, f.Writes);
        Assert.Empty(f.Members);
        Assert.Contains("only the guests", f.Policy().Notices.Single());
    }

    [Fact]
    public void GroupSharePointCannotResolveIsSkippedWithItsMessage()
    {
        var f = new Fixture();
        f.MakeGroupTeam(2);
        f.Queue("Read");
        f.Reject = op => op.MutationKind == "PrincipalEnsure" ? 404 : (int?)null;
        var result = f.Drive();
        Assert.Empty(f.Members);
        Assert.Equal(new[] { f.Read.Id }, f.Roles(42));
        var notice = f.Policy().Notices.Single();
        Assert.Contains("c:0t.c|tenant|" + f.GroupObject.ToString("D"), notice);
        Assert.Contains("The user does not exist or is not unique.", notice);
        Assert.Contains(notice, result.Notices);
    }

    [Fact]
    public void UnknownGroupResolveOutcomeWaitsAndResolvesAgain()
    {
        var f = new Fixture();
        f.MakeGroupTeam(2);
        f.Queue("Read");
        var work = f.Start();
        for (int i = 0; ; i++)
        {
            Assert.True(i < 200, "PrincipalEnsure never prepared.");
            if (work.Status == "Read")
                work = f.Observe(work);
            else if (
                work.Status == "ReadyToCreate"
                && f.Operation().MutationKind == "PrincipalEnsure"
            )
                break;
            else if (work.Status == "ReadyToCreate")
                work = f.Call("PrepareCreate", work);
            else if (work.Status == "Create")
            {
                f.Apply();
                work = f.Call("CreateResponse", work, status: 200);
            }
            else
                throw new Exception(work.Status);
        }
        work = f.Call("PrepareCreate", work);
        // Resolving the group is repeatable, so a lost response waits and resolves again.
        Assert.Equal("RetryWait", f.Call("CreateResponse", work, status: 504).Status);
        Due(f);
        f.Writes.Clear();
        f.Drive();
        Assert.Equal(new[] { "PrincipalEnsure", "MemberAdd", "GrantAdd" }, f.Writes);
        Assert.Single(f.Members);
    }

    private sealed class Fixture
    {
        public DurableWorkerTests.MemoryService Service { get; } =
            new DurableWorkerTests.MemoryService();
        public DocumentStore Store => new DocumentStore(Service);
        public SecurityAdministration Admin => new SecurityAdministration(Service);
        public Guid Library = Guid.NewGuid(),
            Site = Guid.NewGuid(),
            Team = Guid.NewGuid(),
            List = Guid.NewGuid(),
            Web = Guid.NewGuid(),
            User = Guid.NewGuid();
        public string Key = "";
        public PolicyRole Read = new PolicyRole
            {
                Id = 2,
                High = "176",
                Low = "138612833",
            },
            Contribute = new PolicyRole
            {
                Id = 3,
                High = "432",
                Low = "1011028719",
            };
        public List<AclAssignment> Acl = new List<AclAssignment>();
        public List<SitePerson> Members = new List<SitePerson>();
        public SiteGroup? Group;
        public List<string> Writes = new List<string>();
        public List<string> AclReads = new List<string>();
        public List<int> AclStatuses = new List<int>();
        public int NextGroupId = 42;

        /// <summary>Answers a member or grant write with this HTTP status instead of applying it.</summary>
        public Func<SecurityOperation, int?>? Reject;

        /// <summary>Applies a write differently from SharePoint's normal behavior; true when handled.</summary>
        public Func<SecurityOperation, bool>? Mutate;
        public const int LimitedAccess = 1073741825;
        public const string RejectedBody =
            "{\"error\":{\"code\":\"-2146232832, Microsoft.SharePoint.SPException\",\"message\":{\"lang\":\"en-US\",\"value\":\"The user does not exist or is not unique.\"}}}";

        public Fixture()
        {
            Service.Seed(
                new Entity("asx_site", Site)
                {
                    ["asx_webid"] = Web.ToString(),
                    ["asx_url"] = "https://example.sharepoint.com/sites/proto",
                    ["asx_approved"] = true,
                }
            );
            Service.Seed(
                new Entity("asx_library", Library)
                {
                    ["asx_siteid"] = new EntityReference("asx_site", Site),
                    ["asx_listid"] = List.ToString(),
                    ["asx_entryurl"] = "https://example.sharepoint.com/sites/proto/General",
                    ["asx_approved"] = true,
                }
            );
            Service.Seed(
                new Entity("team", Team)
                {
                    ["name"] = "Operations",
                    ["teamtype"] = new OptionSetValue(0),
                    ["isdefault"] = false,
                }
            );
            Store.Create(
                "asx_teamregistration",
                new TeamRegistration
                {
                    Key = "team:" + Team.ToString("N"),
                    TeamId = Team,
                    Enabled = true,
                    Status = "Enabled",
                }
            );
        }

        public void AddPerson(
            string domain,
            string name,
            Guid? entra = null,
            Guid? application = null
        )
        {
            var row = new Entity("systemuser", Guid.NewGuid())
            {
                ["domainname"] = domain,
                ["fullname"] = name,
                ["azureactivedirectoryobjectid"] = entra ?? Guid.NewGuid(),
                ["isdisabled"] = false,
            };
            if (application != null)
                row["applicationid"] = application.Value;
            Service.Seed(row);
        }

        public Guid GroupObject = Guid.Parse("7d1e0c55-2f4b-4a77-9c1d-0b6a5e2f9a31");
        public List<HttpIntent> Posts = new List<HttpIntent>();

        /// <summary>Dataverse team-member reads (systemuser queries) since MakeGroupTeam.</summary>
        public int SystemUserReads;

        /// <summary>Turns the fixture team into an Entra (2) or Microsoft 365 (3) group team.</summary>
        public void MakeGroupTeam(int type, int membership = 0, bool objectId = true)
        {
            var team = Service.Rows[Team];
            team["teamtype"] = new OptionSetValue(type);
            team["membershiptype"] = new OptionSetValue(membership);
            if (objectId)
                team["azureactivedirectoryobjectid"] = GroupObject;
            Service.QueryHook = query =>
            {
                if (query.EntityName == "systemuser")
                    SystemUserReads++;
                return null;
            };
        }

        public void AddUser() =>
            Service.Seed(
                new Entity("systemuser", User)
                {
                    ["domainname"] = "person@example.com",
                    ["azureactivedirectoryobjectid"] = Guid.NewGuid(),
                    ["isdisabled"] = false,
                }
            );

        // Makes the team read fail after registration so the whole request must roll back.
        public void FailTeamRead() =>
            Service.QueryHook = query =>
                query.EntityName == "systemuser"
                    ? new EntityCollection { MoreRecords = true }
                    : null;

        public PolicyDocument Policy() =>
            Store.Require<PolicyDocument>("asx_policy", "policy:" + Library.ToString("N")).Value;

        public SecurityOperation Operation() =>
            Store.Require<SecurityOperation>("asx_operation", Key).Value;

        public void Queue(string access)
        {
            var existing = Store.Find<PolicyDocument>(
                "asx_policy",
                "policy:" + Library.ToString("N")
            );
            var saved = Service.Transaction(() =>
                Admin.Execute(
                    new SecurityRequest
                    {
                        Command = "SavePolicy",
                        LibraryId = Library,
                        Entries = new[]
                        {
                            new PolicyEntry { TeamId = Team, Access = access },
                        },
                        ReadRole = Read,
                        ContributeRole = Contribute,
                        RowVersion = existing?.Row.RowVersion,
                    },
                    true
                )
            );
            Service.Transaction(() =>
                Admin.Execute(
                    new SecurityRequest
                    {
                        Command = "QueuePolicy",
                        LibraryId = Library,
                        RowVersion = saved.RowVersion,
                    },
                    true
                )
            );
            Key = Policy().OperationKey!;
        }

        public WorkerResult Start() =>
            Service.Transaction(() =>
                new SecurityWorker(Service).Execute(
                    new WorkerRequest
                    {
                        Command = "Claim",
                        Key = Key,
                        RunId = "security/run-1",
                    },
                    true
                )
            );

        public WorkerResult Call(
            string command,
            WorkerResult work,
            string? body = null,
            int status = 200
        ) =>
            Service.Transaction(() =>
                new SecurityWorker(Service).Execute(
                    new WorkerRequest
                    {
                        Command = command,
                        Key = Key,
                        RunId = "security/run-1",
                        Token = work.Token,
                        ProbeId = work.ProbeId,
                        ProbeKind = work.ProbeKind,
                        HttpStatus = status,
                        ResponseBody = body,
                    },
                    true
                )
            );

        public WorkerResult Observe(WorkerResult work)
        {
            string body;
            switch (work.ProbeKind)
            {
                case "SecurityLibrary":
                    body = Envelope(
                        new SecurityLibraryObservation { Id = List, UniquePermissions = true }
                    );
                    break;
                case "SecurityReadRole":
                case "SecurityContributeRole":
                    var role = work.ProbeKind == "SecurityReadRole" ? Read : Contribute;
                    body = Envelope(
                        new SecurityRoleObservation
                        {
                            Id = role.Id,
                            Type = role == Read ? 2 : 3,
                            Permissions = new PermissionMask { High = role.High, Low = role.Low },
                        }
                    );
                    break;
                case "SecurityAcl":
                    AclReads.Add(work.Http!.RelativeUri);
                    var principal = System.Text.RegularExpressions.Regex.Match(
                        work.Http.RelativeUri,
                        @"getbyprincipalid\((\d+)\)"
                    );
                    if (!principal.Success)
                    {
                        body = Envelope(new ODataRows<AclAssignment> { Rows = Acl.ToArray() });
                        break;
                    }
                    var assignment = Acl.SingleOrDefault(a =>
                        a.Member.Id == int.Parse(principal.Groups[1].Value)
                    );
                    // SharePoint answers 404 when the principal has no assignment on the list.
                    AclStatuses.Add(assignment == null ? 404 : 200);
                    if (assignment == null)
                        return Call(
                            "Observe",
                            work,
                            "{\"error\":{\"code\":\"-2146232832\",\"message\":{\"lang\":\"en-US\",\"value\":\"Can not find the principal with id.\"}}}",
                            404
                        );
                    body = Envelope(assignment);
                    break;
                case "SecurityGroup":
                    body = Envelope(
                        new ODataRows<SiteGroup>
                        {
                            Rows = Group == null ? Array.Empty<SiteGroup>() : new[] { Group },
                        }
                    );
                    break;
                case "SecurityMembers":
                    body = Envelope(new ODataRows<SitePerson> { Rows = Members.ToArray() });
                    break;
                default:
                    throw new Exception(work.ProbeKind);
            }
            return Call("Observe", work, body);
        }

        public int TypeOf(int role) =>
            role == Read.Id ? 2
            : role == Contribute.Id ? 3
            : role == LimitedAccess ? 1
            : 0;

        public AclAssignment Grant(int group, int role)
        {
            var permission = role == Read.Id ? Read : Contribute;
            return new AclAssignment
            {
                Member = new AclMember { Id = group, Type = 8 },
                Roles = new ODataRows<AclRole>
                {
                    Rows = new[]
                    {
                        new AclRole
                        {
                            Id = role,
                            Type = TypeOf(role),
                            Permissions = new PermissionMask
                            {
                                High = permission.High,
                                Low = permission.Low,
                            },
                        },
                    },
                },
            };
        }

        public int[] Roles(int principal) =>
            Acl.Where(a => a.Member.Id == principal)
                .SelectMany(a => a.Roles.Rows.Select(r => r.Id))
                .ToArray();

        /// <summary>Sets a principal's role bindings on the library, as SharePoint would hold them.</summary>
        public void SetRoles(int principal, params int[] roles)
        {
            Acl.RemoveAll(a => a.Member.Id == principal);
            if (roles.Length == 0)
                return;
            var assignment = Grant(principal, roles[0]);
            assignment.Roles.Rows = roles
                .Select(role => new AclRole
                {
                    Id = role,
                    Type = TypeOf(role),
                    Permissions = new PermissionMask { High = "0", Low = "0" },
                })
                .ToArray();
            Acl.Add(assignment);
        }

        public void Apply()
        {
            var operation = Operation();
            Writes.Add(operation.MutationKind);
            switch (operation.MutationKind)
            {
                case "GroupCreate":
                    var owned = Store
                        .Require<ManagedGroup>("asx_managedgroup", operation.GroupKey)
                        .Value;
                    Group = new SiteGroup
                    {
                        Id = NextGroupId++,
                        Title = owned.Title,
                        Description = owned.Marker,
                        Type = 8,
                    };
                    break;
                case "PrincipalEnsure":
                    // ensureuser adds the principal to the site, not to any group.
                    break;
                case "MemberAdd":
                    Members.Add(
                        new SitePerson
                        {
                            Id = Members.Count == 0 ? 7 : Members.Max(m => m.Id) + 1,
                            Login = operation.MutationLogin,
                            // SharePoint lists an Entra or Microsoft 365 group as a security group.
                            Type = operation.MutationLogin.StartsWith(
                                "c:0",
                                StringComparison.Ordinal
                            )
                                ? 4
                                : 1,
                        }
                    );
                    break;
                case "MemberRemove":
                    Members.RemoveAll(p => p.Id == operation.MutationMemberId);
                    break;
                case "GrantAdd":
                    SetRoles(
                        Group!.Id,
                        Roles(Group.Id).Concat(new[] { operation.MutationRole }).ToArray()
                    );
                    break;
                case "GrantRemove":
                    SetRoles(
                        Group!.Id,
                        Roles(Group.Id).Where(r => r != operation.MutationRole).ToArray()
                    );
                    break;
                default:
                    throw new Exception(operation.MutationKind);
            }
        }

        public WorkerResult Drive(bool expectApplied = true)
        {
            var work = Start();
            for (int i = 0; i < 200; i++)
            {
                switch (work.Status)
                {
                    case "Read":
                        work = Observe(work);
                        break;
                    case "ReadyToCreate":
                        work = Call("PrepareCreate", work);
                        break;
                    case "Create":
                        var op = Operation();
                        Posts.Add(work.Http!);
                        var rejected = Reject?.Invoke(op);
                        if (rejected != null)
                        {
                            Writes.Add("Rejected" + op.MutationKind);
                            work = Call("CreateResponse", work, RejectedBody, rejected.Value);
                            break;
                        }
                        if (Mutate?.Invoke(op) == true)
                            Writes.Add(op.MutationKind);
                        else
                            Apply();
                        work = Call("CreateResponse", work, status: 200);
                        break;
                    case "Verified":
                        work = Call("Complete", work);
                        break;
                    case "Pending":
                        work = Start();
                        break;
                    default:
                        if (expectApplied)
                            Assert.True(
                                work.Status == "Applied",
                                work.Status + ": " + string.Join(" ", work.Notices)
                            );
                        return work;
                }
            }
            throw new Exception("Protocol did not converge.");
        }

        private static string Envelope<T>(T value) =>
            JsonWire.Write(new ODataEnvelope<T> { Data = value });
    }
}
