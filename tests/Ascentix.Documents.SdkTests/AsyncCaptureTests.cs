using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Plugins;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class AsyncCaptureTests
{
    private static DurableWorkerTests.Fixture Setup(out Guid runtime)
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        runtime = RuntimeSeed.Seed(f.Service, f.RecordId, "account");
        return f;
    }

    private static IPluginExecutionContext Context(
        DurableWorkerTests.Fixture f,
        string message,
        string entity,
        ParameterCollection inputs,
        Guid? user = null,
        Guid? primaryId = null
    ) =>
        GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["OwningExtension"] = new EntityReference(
                    "sdkmessageprocessingstep",
                    Guid.NewGuid()
                ),
                ["Stage"] = 40,
                ["Mode"] = 1,
                ["IsInTransaction"] = false,
                ["MessageName"] = message,
                ["PrimaryEntityName"] = entity,
                ["PrimaryEntityId"] = primaryId ?? f.RecordId,
                ["UserId"] = user ?? f.RecordId,
                ["InputParameters"] = inputs,
                ["OutputParameters"] = new ParameterCollection(),
                ["SharedVariables"] = new ParameterCollection(),
                ["CorrelationId"] = Guid.NewGuid(),
            }
        );

    private sealed class Provider : IServiceProvider, IOrganizationServiceFactory
    {
        private readonly DurableWorkerTests.Fixture f;
        private readonly IPluginExecutionContext c;

        public Provider(DurableWorkerTests.Fixture f, IPluginExecutionContext c)
        {
            this.f = f;
            this.c = c;
        }

        public object GetService(Type type) => type == typeof(IPluginExecutionContext) ? c : this;

        public IOrganizationService CreateOrganizationService(Guid? userId) => f.Service;
    }

    [Fact]
    public void AsyncCreateIsAcceptedAndCorruptHostsDoNotStopCapture()
    {
        var f = Setup(out var runtime);
        f.Service.Rows[runtime]["asx_sharepointhosts"] = "not json";
        new RecordInvalidationPlugin().Execute(
            new Provider(f, Context(f, "Create", "account", new ParameterCollection()))
        );
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void TablesOutsideTheAllowlistAreSkippedSilently()
    {
        var f = Setup(out _);
        new RecordInvalidationPlugin().Execute(
            new Provider(f, Context(f, "Create", "contact", new ParameterCollection()))
        );
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void BulkTargetsAboveOneHundredAreProcessed()
    {
        var f = Setup(out _);
        var targets = new EntityCollection(
            Enumerable.Range(0, 150).Select(_ => new Entity("account", Guid.NewGuid())).ToList()
        );
        new RecordInvalidationPlugin().Execute(
            new Provider(
                f,
                Context(
                    f,
                    "CreateMultiple",
                    "account",
                    new ParameterCollection { ["Targets"] = targets }
                )
            )
        );
        Assert.Equal(150, f.Service.Rows.Values.Count(r => r.LogicalName == "asx_outbox"));
    }

    [Fact]
    public void MembershipForUnregisteredTeamsSkipsBeforeReadingRuntime()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        var reads = new List<string>();
        f.Service.QueryHook = q =>
        {
            reads.Add(q.EntityName);
            return null;
        };
        var related = new EntityReferenceCollection(
            Enumerable
                .Range(0, 150)
                .Select(_ => new EntityReference("systemuser", Guid.NewGuid()))
                .ToList()
        );
        new TeamMembershipInvalidationPlugin().Execute(
            new Provider(
                f,
                Context(
                    f,
                    "Associate",
                    "team",
                    new ParameterCollection
                    {
                        ["Relationship"] = new Relationship("teammembership_association"),
                        ["Target"] = new EntityReference("team", Guid.NewGuid()),
                        ["RelatedEntities"] = related,
                    }
                )
            )
        );
        Assert.DoesNotContain("asx_runtime", reads);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void RetirementOfUnregisteredTeamSkipsBeforeReadingRuntime()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        var reads = new List<string>();
        f.Service.QueryHook = q =>
        {
            reads.Add(q.EntityName);
            return null;
        };
        new TeamRetirementPlugin().Execute(
            new Provider(
                f,
                Context(f, "Delete", "team", new ParameterCollection(), primaryId: Guid.NewGuid())
            )
        );
        Assert.DoesNotContain("asx_runtime", reads);
    }
}
