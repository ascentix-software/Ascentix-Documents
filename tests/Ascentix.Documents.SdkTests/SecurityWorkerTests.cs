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
        f.Service.Rows[f.Library]["asx_aclhash"] = "invalid";
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
        f.Service.Rows[f.Library]["asx_aclhash"] = "invalid";
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
    public void ForeignGrantOnOwnedPrincipalIsNeverAdopted()
    {
        var f = new Fixture();
        f.Acl.Add(f.Grant(42, 2));
        f.ResetBaseline();
        f.Queue("Read");
        var result = f.Drive(false);
        Assert.Equal("Blocked", result.Status);
        Assert.Contains("Unowned", result.Notices.Single());
        Assert.DoesNotContain("GrantRemove", f.Writes);
        Assert.DoesNotContain("GrantAdd", f.Writes);
    }

    [Fact]
    public void OtherAclMutationDuringOurGrantBlocksApplied()
    {
        var f = new Fixture();
        f.Queue("Read");
        var work = f.Start();
        for (int i = 0; i < 50; i++)
        {
            if (work.Status == "Read")
                work = f.Observe(work);
            else if (work.Status == "ReadyToCreate")
                work = f.Call("PrepareCreate", work);
            else if (work.Status == "Create")
            {
                string kind = f.Operation().MutationKind;
                f.Apply();
                if (kind == "GrantAdd")
                    f.Acl.Add(f.Grant(99, 2));
                work = f.Call("CreateResponse", work, status: 200);
            }
            else
                break;
        }
        Assert.Equal("Blocked", work.Status);
        Assert.Contains("beyond", work.Notices.Single());
        Assert.False(f.Service.Rows[f.Library].GetAttributeValue<bool>("asx_policyapplied"));
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
            ResetBaseline();
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

        public void AddUser() =>
            Service.Seed(
                new Entity("systemuser", User)
                {
                    ["domainname"] = "person@example.com",
                    ["azureactivedirectoryobjectid"] = Guid.NewGuid(),
                    ["isdisabled"] = false,
                }
            );

        public void ResetBaseline() =>
            Service.Rows[Library]["asx_aclhash"] = SharePointObservations.AclHash(
                new ODataRows<AclAssignment> { Rows = Acl.ToArray() }
            );

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
                case "SecurityBaseline":
                case "SecurityAcl":
                case "SecurityFinal":
                    body = Envelope(new ODataRows<AclAssignment> { Rows = Acl.ToArray() });
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
                        Id = 42,
                        Title = owned.Title,
                        Description = owned.Marker,
                        Type = 8,
                    };
                    break;
                case "MemberAdd":
                    Members.Add(
                        new SitePerson
                        {
                            Id = 7,
                            Login = operation.MutationLogin,
                            Type = 1,
                        }
                    );
                    break;
                case "MemberRemove":
                    Members.RemoveAll(p => p.Id == operation.MutationMemberId);
                    break;
                case "GrantAdd":
                    Acl.Add(Grant(42, operation.MutationRole));
                    break;
                case "GrantRemove":
                    Acl.RemoveAll(a => a.Member.Id == 42);
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
                            Assert.Equal("Applied", work.Status);
                        return work;
                }
            }
            throw new Exception("Protocol did not converge.");
        }

        private static string Envelope<T>(T value) =>
            JsonWire.Write(new ODataEnvelope<T> { Data = value });
    }
}
