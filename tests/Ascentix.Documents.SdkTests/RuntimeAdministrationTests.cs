using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class RuntimeAdministrationTests
{
    private sealed class Env
    {
        internal readonly EventRegistrationTests.Org Org;
        internal readonly Guid Runtime;
        internal readonly Guid Admin = Guid.NewGuid();

        internal Env(params string[] tables)
        {
            Org = new EventRegistrationTests.Org("account", "contact", "lead");
            Runtime = RuntimeSeed.Seed(Org.S.Memory, Org.Worker, tables);
            var role = Guid.NewGuid();
            Org.S.Memory.Seed(
                new Entity("role", role)
                {
                    ["roletemplateid"] = new EntityReference(
                        "roletemplate",
                        AdministratorCheck.SystemAdministratorTemplate
                    ),
                }
            );
            Org.S.Memory.Seed(
                new Entity("systemuserroles", Guid.NewGuid())
                {
                    ["systemuserid"] = Admin,
                    ["roleid"] = role,
                }
            );
        }

        internal RuntimeRequest Get() =>
            RuntimeAdministration.Execute(Org.S, new RuntimeRequest(), true, Admin);

        internal RuntimeRequest Save(RuntimeRequest r, Guid? caller = null)
        {
            r.Command = "Save";
            return RuntimeAdministration.Execute(Org.S, r, true, caller ?? Admin);
        }
    }

    [Fact]
    public void LegacyProfileShowsPendingThenSaveMigratesAndRegisters()
    {
        var e = new Env("account");
        foreach (
            var row in e
                .Org.S.Memory.Rows.Values.Where(r => r.LogicalName == "asx_runtimetable")
                .ToList()
        )
            e.Org.S.Memory.Rows.Remove(row.Id);
        var got = e.Get();
        Assert.False(got.Migrated);
        Assert.Equal(new[] { "account" }, got.Tables);
        Assert.Equal(
            "Pending",
            got.Registration!.Readiness.Single(r => r.Scope == "account").Status
        );
        var saved = e.Save(got);
        Assert.True(saved.Migrated);
        Assert.All(saved.Registration!.Readiness, r => Assert.Equal("Ready", r.Status));
        Assert.Equal(
            "[]",
            e.Org.S.Memory.Rows[e.Runtime].GetAttributeValue<string>("asx_allowedtables")
        );
    }

    [Fact]
    public void NonAdministratorSaveIsRefusedBeforeAnyWrite()
    {
        var e = new Env("account");
        var got = e.Get();
        var ex = Assert.Throws<EvaluationBlockedException>(() => e.Save(got, Guid.NewGuid()));
        Assert.Contains("System Administrator", ex.Message);
        Assert.Empty(e.Org.S.Writes);
    }

    [Fact]
    public void RepeatedSaveWithoutChangesMakesNoStepWrites()
    {
        var e = new Env("account", "contact");
        var saved = e.Save(e.Get());
        int writes = e.Org.S.StepWrites;
        e.Save(saved);
        Assert.Equal(writes, e.Org.S.StepWrites);
    }

    [Fact]
    public void TooManyNewTablesInOneSaveIsRefusedWithReason()
    {
        var e = new Env("account");
        var got = e.Get();
        got.Tables = Enumerable
            .Range(0, EventRegistrations.MaxNewTablesPerSave + 1)
            .Select(i => "cr123_t" + i)
            .Concat(new[] { "account" })
            .ToArray();
        var ex = Assert.Throws<EvaluationBlockedException>(() => e.Save(got));
        Assert.Contains(EventRegistrations.MaxNewTablesPerSave.ToString(), ex.Message);
        Assert.Contains("2 minutes", ex.Message);
    }

    [Fact]
    public void WorkerMustBeAnEnabledApplicationUser()
    {
        var e = new Env("account");
        var got = e.Get();
        var human = Guid.NewGuid();
        e.Org.S.Memory.Seed(new Entity("systemuser", human) { ["isdisabled"] = false });
        got.WorkerId = human;
        var ex = Assert.Throws<EvaluationBlockedException>(() => e.Save(got));
        Assert.Contains("application user", ex.Message);
    }

    [Fact]
    public void GetReportsUnsupportedTableInsteadOfFailing()
    {
        var e = new Env("account", "cr123_gone");
        var got = e.Get();
        Assert.Contains("cr123_gone", got.Registration!.Error);
    }

    [Fact]
    public void UnregisterRemovesAllOwnedSteps()
    {
        var e = new Env("account");
        e.Save(e.Get());
        var r = e.Get();
        r.Command = "Unregister";
        var after = RuntimeAdministration.Execute(e.Org.S, r, true, e.Admin);
        Assert.Empty(e.Org.Steps);
        Assert.All(after.Registration!.Readiness, x => Assert.Equal("Pending", x.Status));
    }
}
