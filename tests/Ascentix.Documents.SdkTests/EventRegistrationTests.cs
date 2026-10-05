using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class EventRegistrationTests
{
    internal sealed class Org
    {
        internal readonly StepService S = new();
        internal readonly Guid Worker = Guid.NewGuid();

        internal Org(params string[] tables)
        {
            foreach (
                var type in new[]
                {
                    "Ascentix.Documents.Plugins.RecordInvalidationPlugin",
                    "Ascentix.Documents.Plugins.TeamRetirementPlugin",
                    "Ascentix.Documents.Plugins.TeamMembershipInvalidationPlugin",
                }
            )
                S.Memory.Seed(new Entity("plugintype", Guid.NewGuid()) { ["typename"] = type });
            foreach (
                var m in new[]
                {
                    "Create",
                    "Update",
                    "Delete",
                    "CreateMultiple",
                    "UpdateMultiple",
                    "Associate",
                    "Disassociate",
                }
            )
            {
                var sdk = new Entity("sdkmessage", Guid.NewGuid()) { ["name"] = m };
                S.Memory.Seed(sdk);
                foreach (var t in tables.Concat(new[] { "team" }))
                    S.Memory.Seed(
                        new Entity("sdkmessagefilter", Guid.NewGuid())
                        {
                            ["sdkmessageid"] = sdk.ToEntityReference(),
                            ["primaryobjecttypecode"] = t,
                            ["iscustomprocessingstepallowed"] = true,
                        }
                    );
            }
        }

        internal Entity[] Steps =>
            S.Memory.Rows.Values.Where(r => r.LogicalName == "sdkmessageprocessingstep").ToArray();
    }

    [Fact]
    public void ReconcileRegistersAsyncWorkerStepsAndIsIdempotent()
    {
        var o = new Org("account", "contact");
        var summary = EventRegistrations.Reconcile(
            o.S,
            o.Worker,
            new[] { "account", "contact" },
            false
        );
        Assert.All(summary.Readiness, r => Assert.Equal("Ready", r.Status));
        Assert.Equal(2 * EventRegistrationPlan.RecordMessages.Length + 3, o.Steps.Length);
        Assert.All(
            o.Steps,
            s =>
            {
                Assert.Equal(1, s.GetAttributeValue<OptionSetValue>("mode").Value);
                Assert.Equal(40, s.GetAttributeValue<OptionSetValue>("stage").Value);
                Assert.True(s.GetAttributeValue<bool>("asyncautodelete"));
                Assert.Equal(
                    o.Worker,
                    s.GetAttributeValue<EntityReference>("impersonatinguserid").Id
                );
            }
        );
        Assert.All(
            o.Steps.Where(s => s.GetAttributeValue<string>("name").Contains(" Update ")),
            s => Assert.Equal(1, s.GetAttributeValue<OptionSetValue>("statecode").Value)
        );
        int writes = o.S.StepWrites;
        EventRegistrations.Reconcile(o.S, o.Worker, new[] { "account", "contact" }, false);
        Assert.Equal(writes, o.S.StepWrites);
    }

    [Fact]
    public void WorkerChangeAndTableRemovalRewriteOwnedSteps()
    {
        var o = new Org("account", "contact");
        EventRegistrations.Reconcile(o.S, o.Worker, new[] { "account", "contact" }, true);
        var next = Guid.NewGuid();
        EventRegistrations.Reconcile(o.S, next, new[] { "account" }, true);
        Assert.Equal(EventRegistrationPlan.RecordMessages.Length + 3, o.Steps.Length);
        Assert.All(
            o.Steps,
            s => Assert.Equal(next, s.GetAttributeValue<EntityReference>("impersonatinguserid").Id)
        );
        Assert.DoesNotContain(
            o.Steps,
            s => s.GetAttributeValue<string>("name").EndsWith(" contact")
        );
    }

    [Fact]
    public void UnsupportedTableIsNamedAndInspectReportsInsteadOfThrowing()
    {
        var o = new Org("account");
        var ex = Assert.Throws<EvaluationBlockedException>(() =>
            EventRegistrations.Reconcile(o.S, o.Worker, new[] { "account", "cr123_gone" }, false)
        );
        Assert.Contains("cr123_gone", ex.Message);
        var summary = EventRegistrations.Inspect(
            o.S,
            o.Worker,
            new[] { "account", "cr123_gone" },
            false
        );
        Assert.Contains("cr123_gone", summary.Error);
    }

    [Fact]
    public void ReadbackMismatchFailsAndUnregisterRemovesEverything()
    {
        var o = new Org("account");
        EventRegistrations.Reconcile(o.S, o.Worker, new[] { "account" }, false);
        o.S.IgnoreStepWrites = true;
        Assert.Throws<EvaluationBlockedException>(() =>
            EventRegistrations.Reconcile(o.S, Guid.NewGuid(), new[] { "account" }, false)
        );
        o.S.IgnoreStepWrites = false;
        Assert.Equal(
            EventRegistrationPlan.RecordMessages.Length + 3,
            EventRegistrations.Unregister(o.S)
        );
        Assert.Empty(o.Steps);
    }
}
