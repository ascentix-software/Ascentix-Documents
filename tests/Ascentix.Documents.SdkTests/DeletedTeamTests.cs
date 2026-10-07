using System;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;
using Fixture = Ascentix.Documents.SdkTests.SecurityWorkerTests.Fixture;

namespace Ascentix.Documents.SdkTests;

/// <summary>
/// A Dataverse team deleted while it has access to a library never blocks that library: its
/// access is shown as a deleted team's, never read as a live team, and the next access run, an
/// admin's Apply or the scheduled refresh, removes Documents' grant for its group and finishes
/// its registration. The group itself stays in SharePoint.
/// </summary>
public sealed class DeletedTeamTests
{
    /// <summary>Deletes the fixture team in Dataverse. When the delete event was seen, its
    /// registration is retired the way TeamRetirementPlugin retires it.</summary>
    private static void DeleteTeam(Fixture f, bool seen)
    {
        f.Service.Rows.Remove(f.Team);
        if (!seen)
            return;
        var registration = Registration(f, f.Team);
        registration.Value.Enabled = false;
        registration.Value.Status = "Revoking";
        f.Store.Save(registration);
    }

    private static StoredRow<TeamRegistration> Registration(Fixture f, Guid team) =>
        f.Store.Require<TeamRegistration>("asx_teamregistration", "team:" + team.ToString("N"));

    private static Guid AddTeam(Fixture f, string name)
    {
        var team = Guid.NewGuid();
        f.Service.Seed(
            new Entity("team", team)
            {
                ["name"] = name,
                ["teamtype"] = new OptionSetValue(0),
                ["isdefault"] = false,
            }
        );
        return team;
    }

    private static SecurityResult Apply(Fixture f, params (Guid Team, string Access)[] entries) =>
        f.Service.Transaction(() =>
            f.Admin.Execute(
                new SecurityRequest
                {
                    Command = "ApplyPolicy",
                    LibraryId = f.Library,
                    RowVersion = f
                        .Store.Require<PolicyDocument>(
                            "asx_policy",
                            "policy:" + f.Library.ToString("N")
                        )
                        .Row.RowVersion,
                    Entries = entries
                        .Select(e => new PolicyEntry { TeamId = e.Team, Access = e.Access })
                        .ToArray(),
                    ReadRole = f.Read,
                    ContributeRole = f.Contribute,
                },
                true
            )
        );

    private static SecurityResult GetPolicy(Fixture f) =>
        f.Service.Transaction(() =>
            f.Admin.Execute(
                new SecurityRequest { Command = "GetPolicy", LibraryId = f.Library },
                true
            )
        );

    private static WorkerResult Scan(Fixture f) =>
        f.Service.Transaction(() =>
            new SecurityRefresh(f.Service, () => DateTime.UtcNow.AddDays(2)).Scan()
        );

    /// <summary>The library's access shows the team was deleted and that its group stays.</summary>
    private static void AssertCleanedUp(Fixture f, int group)
    {
        Assert.Empty(f.Roles(group));
        Assert.DoesNotContain(f.Writes, w => w.Contains("Delete"));
        var registration = Registration(f, f.Team).Value;
        Assert.False(registration.Enabled);
        Assert.Equal("Revoked", registration.Status);
        var policy = f.Policy();
        Assert.Equal("None", policy.Applied.Single(a => a.TeamId == f.Team).Access);
        var notice = Assert.Single(policy.Notices, n => n.Contains("deleted in Dataverse"));
        Assert.StartsWith("Team 'Operations' was deleted in Dataverse.", notice);
        Assert.Contains("removed its access to this library", notice);
        Assert.Contains("Documents does not delete SharePoint groups", notice);
    }

    [Fact]
    public void GetPolicyShowsADeletedTeamByItsLastKnownNameAndNeverFails()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        DeleteTeam(f, seen: true);
        var result = GetPolicy(f);
        Assert.Equal("Applied", result.Status);
        var deleted = Assert.Single(result.Teams);
        Assert.Equal(f.Team, deleted.TeamId);
        Assert.True(deleted.Deleted);
        Assert.Equal("Operations", deleted.Name);
        // A team deleted before Documents stored its name is shown by its ID.
        var registration = Registration(f, f.Team);
        registration.Value.Name = null;
        f.Store.Save(registration);
        Assert.Null(Assert.Single(GetPolicy(f).Teams).Name);
    }

    [Fact]
    public void LiveTeamsAreListedByTheirCurrentName()
    {
        var f = new Fixture();
        f.Queue("Read");
        f.Drive();
        f.Service.Rows[f.Team]["name"] = "Operations EU";
        var team = Assert.Single(GetPolicy(f).Teams);
        Assert.False(team.Deleted);
        Assert.Equal("Operations EU", team.Name);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ApplyWithADeletedTeamAndAChangeToAnotherTeamRemovesOnlyTheDeletedTeamsAccess(
        bool seen
    )
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        var other = AddTeam(f, "Finance");
        f.Key = Apply(f, (f.Team, "Read"), (other, "Read")).Policy!.OperationKey!;
        f.Drive();
        int deletedGroup = f.GroupOf(f.Team),
            otherGroup = f.GroupOf(other);
        Assert.Equal(new[] { f.Read.Id }, f.Roles(deletedGroup));
        DeleteTeam(f, seen);
        f.Writes.Clear();
        // The admin changes only the other team; the deleted team's row is sent as it was.
        var applied = Apply(f, (f.Team, "Read"), (other, "Contribute"));
        Assert.DoesNotContain(applied.Policy!.Desired, e => e.TeamId == f.Team);
        f.Key = applied.Policy.OperationKey!;
        f.Drive();
        Assert.Equal(new[] { f.Contribute.Id }, f.Roles(otherGroup));
        AssertCleanedUp(f, deletedGroup);
        Assert.True(Registration(f, other).Value.Enabled);
        // The next Apply has nothing left of the deleted team.
        Assert.Equal(new[] { other }, GetPolicy(f).Policy!.Desired.Select(e => e.TeamId));
    }

    [Fact]
    public void ApplyThatOnlyTouchesTheDeletedTeamRemovesItsAccess()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        int group = f.GroupOf(f.Team);
        DeleteTeam(f, seen: true);
        f.Key = Apply(f, (f.Team, "Read")).Policy!.OperationKey!;
        f.Drive();
        AssertCleanedUp(f, group);
        Assert.Empty(f.Policy().Desired);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ScheduledRefreshRemovesADeletedTeamsAccessAndFinishesItsRegistration(bool seen)
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        int group = f.GroupOf(f.Team);
        DeleteTeam(f, seen);
        f.Key = Assert.Single(Scan(f).Keys);
        f.Drive();
        AssertCleanedUp(f, group);
        // Once removed, later reviews have nothing to do for it.
        Assert.Empty(Scan(f).Keys);
    }

    [Fact]
    public void ATeamDeletedWhileItsLibraryWasRemovedLosesItsAccessOnceTheLibraryIsBack()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        int group = f.GroupOf(f.Team);
        f.Service.Transaction(() =>
            new CatalogAdministration(f.Service).Execute(
                new CatalogRequest { Command = "RemoveLibrary", CatalogId = f.Library },
                true
            )
        );
        // The team is deleted while the library is out of Documents: its delete event finds no
        // library to refresh.
        DeleteTeam(f, seen: true);
        f.Store.Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = "team-event:delete",
                SecurityTeamId = f.Team,
                SecurityPage = 1,
            }
        );
        var planned = f.Service.Transaction(() =>
            new WorkerCoordinator(f.Service).Execute(
                new WorkerRequest { Command = "Plan", Key = "team-event:delete" },
                true
            )
        );
        Assert.Equal("Planned", planned.Status);
        Assert.Empty(planned.Keys);
        // The library is added again; the scheduled refresh picks it up again.
        f.Service.Rows[f.Library]["statecode"] = new OptionSetValue(0);
        f.Service.Rows[f.Library]["statuscode"] = new OptionSetValue(1);
        f.Service.Rows[f.Library]["asx_approved"] = true;
        f.Key = Assert.Single(Scan(f).Keys);
        f.Drive();
        AssertCleanedUp(f, group);
    }

    [Fact]
    public void ADeletedTeamWithNoLibraryLeftHasItsRegistrationFinishedByItsDeleteEvent()
    {
        var f = new Fixture();
        DeleteTeam(f, seen: true);
        f.Store.Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = "team-event:delete",
                SecurityTeamId = f.Team,
                SecurityPage = 1,
            }
        );
        f.Service.Transaction(() =>
            new WorkerCoordinator(f.Service).Execute(
                new WorkerRequest { Command = "Plan", Key = "team-event:delete" },
                true
            )
        );
        Assert.Equal("Revoked", Registration(f, f.Team).Value.Status);
    }

    [Fact]
    public void AccessRunQueuedBeforeTheTeamWasDeletedRemovesItsAccessInsteadOfStopping()
    {
        var f = new Fixture();
        f.AddUser();
        f.Queue("Read");
        f.Drive();
        int group = f.GroupOf(f.Team);
        f.Queue("Contribute");
        // Deleted before its delete event is processed: the registration is still enabled.
        DeleteTeam(f, seen: false);
        // Polling the library and its run reads no team by ID.
        var status = GetPolicy(f);
        Assert.Equal(f.Key, status.Policy!.OperationKey);
        Assert.True(Assert.Single(status.Teams).Deleted);
        f.Drive();
        AssertCleanedUp(f, group);
    }

    [Fact]
    public void ARunReachingADeletedTeamWithNoGroupYetSkipsIt()
    {
        var f = new Fixture();
        f.Queue("Read");
        DeleteTeam(f, seen: false);
        f.Drive();
        Assert.Empty(f.Acl);
        Assert.Empty(f.Writes);
        Assert.Equal("None", f.Policy().Applied.Single().Access);
        Assert.Equal("Revoked", Registration(f, f.Team).Value.Status);
    }

    [Fact]
    public void RegisteringADeletedTeamSaysItWasDeleted()
    {
        var f = new Fixture();
        DeleteTeam(f, seen: false);
        var refused = Assert.Throws<Conditions.EvaluationBlockedException>(() =>
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
        Assert.Equal(TeamDirectory.DeletedRefusal, refused.Message);
    }
}
