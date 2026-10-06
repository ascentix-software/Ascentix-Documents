using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Plugins;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class LifecycleTests
{
    [Fact]
    public void RelatedEventsQueueOnlyMatchingRecordSelectionsAndDedupe()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        Guid related = Guid.NewGuid(),
            other = Guid.NewGuid();
        var selection = new RecordPlanDocument
        {
            Key = "recordplan:test",
            TemplateId = f.TemplateId,
            RecordId = f.RecordId,
            Table = "account",
            Status = "Selection",
            Sources = new[]
            {
                new SourceVersion { Table = "account", Id = f.RecordId },
                new SourceVersion { Table = "contact", Id = related },
            },
        };
        f.Store.Create("asx_outbox", selection);
        var correlation = Guid.NewGuid();
        TargetedReplan.Append(f.Service, f.Store, "contact", other, correlation);
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
        TargetedReplan.Append(f.Service, f.Store, "contact", related, correlation);
        TargetedReplan.Append(f.Service, f.Store, "contact", related, correlation);
        var key = f.Store.Pending("asx_outbox").Single();
        var worker = new TargetedReplan(f.Service, new[] { "account", "contact" });
        Assert.Equal(
            "Planned",
            worker.Consume(f.Store.Require<OutboxDocument>("asx_outbox", key)).Status
        );
        Assert.Single(f.Store.Pending("asx_outbox"));
        Assert.Equal(3, f.Service.Rows.Values.Count(r => r.LogicalName == "asx_outbox"));
        worker.Consume(f.Store.Require<OutboxDocument>("asx_outbox", key));
        Assert.Equal(3, f.Service.Rows.Values.Count(r => r.LogicalName == "asx_outbox"));
    }

    [Fact]
    public void RootUpdateFiltersAgainstCurrentPublishedDependencies()
    {
        var f = Setup();
        var unrelated = new Entity("account", f.RecordId) { ["telephone1"] = "555" };
        new RecordInvalidationPlugin().Execute(
            new Provider(f, Event(f, "Update", new ParameterCollection { ["Target"] = unrelated }))
        );
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
        var related = new Entity("account", f.RecordId) { ["name"] = "Changed" };
        var context = Event(f, "Update", new ParameterCollection { ["Target"] = related });
        new RecordInvalidationPlugin().Execute(new Provider(f, context));
        new RecordInvalidationPlugin().Execute(new Provider(f, context));
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void BulkRootEventsUseExactWorkerIdentity()
    {
        var f = Setup();
        var targets = new EntityCollection(
            new[]
            {
                new Entity("account", f.RecordId) { ["name"] = "Changed" },
                new Entity("account", Guid.NewGuid()) { ["name"] = "Other" },
            }
        );
        new RecordInvalidationPlugin().Execute(
            new Provider(
                f,
                Event(f, "UpdateMultiple", new ParameterCollection { ["Targets"] = targets })
            )
        );
        Assert.Equal(2, f.Service.Rows.Values.Count(r => r.LogicalName == "asx_outbox"));
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new RecordInvalidationPlugin().Execute(
                new Provider(f, Event(f, "Create", new ParameterCollection(), Guid.NewGuid()))
            )
        );
    }

    [Theory]
    [InlineData("Update")]
    [InlineData("UpdateMultiple")]
    public void DisabledUpdatesExitBeforeDependencyOrTemplateReadsAndLeaveCreateWorking(
        string message
    )
    {
        var f = Setup();
        f.Service.Rows.Values.Single(r => r.LogicalName == "asx_runtime")[
            "asx_processrecordupdates"
        ] = false;
        var reads = new List<string>();
        f.Service.QueryHook = query =>
        {
            reads.Add(query.EntityName);
            return null;
        };
        new RecordInvalidationPlugin().Execute(
            new Provider(f, Event(f, message, new ParameterCollection()))
        );
        Assert.Equal(new[] { "asx_runtime", "asx_runtimetable" }, reads);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
        new RecordInvalidationPlugin().Execute(
            new Provider(f, Event(f, "Create", new ParameterCollection()))
        );
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void TeamDeletionKeepsRevocationTombstoneWithoutReadingDeletedNativeTeam()
    {
        var f = Setup();
        Guid team = Guid.NewGuid();
        f.Store.Create(
            "asx_teamregistration",
            new TeamRegistration
            {
                Key = "team:" + team.ToString("N"),
                TeamId = team,
                Enabled = true,
                Status = "Enabled",
            }
        );
        var context = GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["OwningExtension"] = new EntityReference(
                    "sdkmessageprocessingstep",
                    Guid.NewGuid()
                ),
                ["Stage"] = 40,
                ["Mode"] = 0,
                ["IsInTransaction"] = true,
                ["MessageName"] = "Delete",
                ["PrimaryEntityName"] = "team",
                ["PrimaryEntityId"] = team,
                ["CorrelationId"] = Guid.NewGuid(),
                ["UserId"] = f.RecordId,
                ["SharedVariables"] = new ParameterCollection(),
            }
        );
        var provider = new Provider(f, context);
        new TeamRetirementPlugin().Execute(provider);
        var update = Assert.Single(f.Service.Updates);
        update.Target["modifiedon"] = DateTime.UtcNow;
        update.Target["modifiedby"] = new EntityReference("systemuser", f.RecordId);
        var nested = GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["Stage"] = 20,
                ["IsInTransaction"] = true,
                ["MessageName"] = "Update",
                ["PrimaryEntityName"] = "asx_teamregistration",
                ["UserId"] = f.RecordId,
                ["InputParameters"] = new ParameterCollection { ["Target"] = update.Target },
            }
        );
        new CatalogGuard().Execute(new Provider(f, nested));
        var registration = f
            .Store.Require<TeamRegistration>("asx_teamregistration", "team:" + team.ToString("N"))
            .Value;
        Assert.False(registration.Enabled);
        Assert.Equal("Revoking", registration.Status);
        Assert.Empty(new TeamSnapshotReader(f.Service).Read(team));
    }

    [Theory]
    [InlineData("caller")]
    [InlineData("enabled")]
    [InlineData("status")]
    [InlineData("identity")]
    [InlineData("schema")]
    public void WorkerRetirementRejectsOtherWritersAndInvalidState(string variant)
    {
        var f = Setup();
        var team = Guid.NewGuid();
        var registration = new TeamRegistration
        {
            Key = "team:" + team.ToString("N"),
            TeamId = team,
            Enabled = variant == "enabled",
            Status = variant == "status" ? "Enabled" : "Revoking",
            SchemaVersion = variant == "schema" ? 2 : 1,
        };
        var target = new Entity(
            "asx_teamregistration",
            variant == "identity"
                ? Guid.NewGuid()
                : DocumentStore.StableId("asx_teamregistration:" + registration.Key)
        )
        {
            ["asx_status"] = "Revoking",
            ["asx_payload"] = JsonWire.Write(registration),
        };
        var context = GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["Stage"] = 20,
                ["IsInTransaction"] = true,
                ["MessageName"] = "Update",
                ["PrimaryEntityName"] = "asx_teamregistration",
                ["UserId"] = variant == "caller" ? Guid.NewGuid() : f.RecordId,
                ["InputParameters"] = new ParameterCollection { ["Target"] = target },
            }
        );
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new CatalogGuard().Execute(new Provider(f, context))
        );
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MembershipEventsSupportBothDirectionsAndDedupe(bool reverse)
    {
        var f = Setup();
        Guid team = Guid.NewGuid(),
            person = Guid.NewGuid();
        f.Store.Create(
            "asx_teamregistration",
            new TeamRegistration
            {
                Key = "team:" + team.ToString("N"),
                TeamId = team,
                Enabled = true,
                Status = "Enabled",
            }
        );
        var inputs = new ParameterCollection
        {
            ["Target"] = new EntityReference(
                reverse ? "systemuser" : "team",
                reverse ? person : team
            ),
            ["RelatedEntities"] = new EntityReferenceCollection
            {
                new EntityReference(reverse ? "team" : "systemuser", reverse ? team : person),
            },
            ["Relationship"] = new Relationship("teammembership_association"),
        };
        var context = Event(f, "Disassociate", inputs);
        var provider = new Provider(f, context);
        new TeamMembershipInvalidationPlugin().Execute(provider);
        new TeamMembershipInvalidationPlugin().Execute(provider);
        var row = Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
        var job = JsonWire.Read<OutboxDocument>(row.GetAttributeValue<string>("asx_payload"));
        Assert.Equal(team, job.SecurityTeamId);
        Assert.Equal("Pending", job.Status);
    }

    [Fact]
    public void SignInToAGroupTeamCreatesNoWork()
    {
        // Dataverse adds a person to a group team when they first sign in. The group itself is
        // granted on the library, so there is nothing to sync.
        var f = Setup();
        Guid team = Guid.NewGuid();
        f.Store.Create(
            "asx_teamregistration",
            new TeamRegistration
            {
                Key = "team:" + team.ToString("N"),
                TeamId = team,
                Enabled = true,
                Group = true,
                Status = "Enabled",
            }
        );
        var inputs = new ParameterCollection
        {
            ["Target"] = new EntityReference("team", team),
            ["RelatedEntities"] = new EntityReferenceCollection
            {
                new EntityReference("systemuser", Guid.NewGuid()),
            },
            ["Relationship"] = new Relationship("teammembership_association"),
        };
        new TeamMembershipInvalidationPlugin().Execute(
            new Provider(f, Event(f, "Associate", inputs))
        );
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void OtherRelationshipsProduceNoProductWork()
    {
        var f = Setup();
        new TeamMembershipInvalidationPlugin().Execute(
            new Provider(
                f,
                Event(
                    f,
                    "Associate",
                    new ParameterCollection
                    {
                        ["Relationship"] = new Relationship("unrelated_relationship"),
                    }
                )
            )
        );
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void RecordDeletionCreatesDurableReviewMarkerAndStopsUnstartedWork()
    {
        var f = Setup();
        new RecordInvalidationPlugin().Execute(
            new Provider(f, Event(f, "Delete", new ParameterCollection()))
        );
        Assert.Equal(
            "DecommissionReview",
            f.Store.Require<OutboxDocument>(
                "asx_outbox",
                WorkerCoordinator.RetirementKey("account", f.RecordId)
            ).Value.Status
        );
        f.Store.Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = "old-request",
                Table = "account",
                TemplateId = f.TemplateId,
                RevisionId = f.RevisionId,
                RecordId = f.RecordId,
            }
        );
        f.Service.Rows.Remove(f.RecordId);
        Assert.Equal(
            "DecommissionReview",
            f.Coordinator.Execute(
                new WorkerRequest { Command = "Plan", Key = "old-request" },
                true
            ).Status
        );
    }

    private static DurableWorkerTests.Fixture Setup()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        RuntimeSeed.Seed(f.Service, f.RecordId, "account");
        return f;
    }

    private static IPluginExecutionContext Event(
        DurableWorkerTests.Fixture f,
        string message,
        ParameterCollection inputs,
        Guid? user = null
    ) =>
        GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["OwningExtension"] = new EntityReference(
                    "sdkmessageprocessingstep",
                    Guid.NewGuid()
                ),
                ["Stage"] = 40,
                ["Mode"] = 0,
                ["IsInTransaction"] = true,
                ["MessageName"] = message,
                ["PrimaryEntityName"] = "account",
                ["PrimaryEntityId"] = f.RecordId,
                ["UserId"] = user ?? f.RecordId,
                ["InputParameters"] = inputs,
                ["SharedVariables"] = new ParameterCollection(),
                ["CorrelationId"] = Guid.NewGuid(),
            }
        );

    private sealed class Provider : IServiceProvider, IOrganizationServiceFactory
    {
        private readonly DurableWorkerTests.Fixture fixture;
        private readonly IPluginExecutionContext context;

        public Provider(DurableWorkerTests.Fixture fixture, IPluginExecutionContext context)
        {
            this.fixture = fixture;
            this.context = context;
        }

        public object GetService(Type type) =>
            type == typeof(IPluginExecutionContext) ? context : this;

        public IOrganizationService CreateOrganizationService(Guid? userId)
        {
            Assert.Equal(context.UserId, userId);
            return fixture.Service;
        }
    }
}
