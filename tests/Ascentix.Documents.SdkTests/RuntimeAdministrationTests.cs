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
            : this(tables, Array.Empty<string>()) { }

        internal Env(string[] tables, string[] orgTables)
        {
            Org = new EventRegistrationTests.Org(
                new[] { "account", "contact", "lead" }.Concat(orgTables).ToArray()
            );
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

        internal RuntimeRequest Change(string command, string table, Guid? caller = null) =>
            RuntimeAdministration.Execute(
                Org.S,
                new RuntimeRequest { Command = command, Table = table },
                true,
                caller ?? Admin
            );

        internal string[] Rows() =>
            Org
                .S.Memory.Rows.Values.Where(r => r.LogicalName == "asx_runtimetable")
                .Select(r => r.GetAttributeValue<string>("asx_logicalname"))
                .OrderBy(n => n, StringComparer.Ordinal)
                .ToArray();

        internal int StepsFor(string table) =>
            Org.Steps.Count(s => s.GetAttributeValue<string>("name").EndsWith(" " + table));

        internal void ClearRows()
        {
            foreach (
                var row in Org
                    .S.Memory.Rows.Values.Where(r => r.LogicalName == "asx_runtimetable")
                    .ToList()
            )
                Org.S.Memory.Rows.Remove(row.Id);
        }
    }

    [Fact]
    public void AddTableRegistersOnlyTheNewTableAndAddsItsRow()
    {
        var e = new Env("account");
        e.Save(e.Get());
        int accountSteps = e.StepsFor("account");
        e.Org.S.Writes.Clear();
        var got = e.Change("AddTable", "contact");
        Assert.Equal(new[] { "account", "contact" }, got.Tables.OrderBy(t => t).ToArray());
        Assert.Equal(new[] { "account", "contact" }, e.Rows());
        Assert.Equal(accountSteps, e.StepsFor("account"));
        Assert.Equal(EventRegistrationPlan.RecordMessages.Length, e.StepsFor("contact"));
        Assert.DoesNotContain(
            e.Org.S.Writes,
            w => w.EndsWith("sdkmessageprocessingstep") && w.StartsWith("Delete")
        );
        Assert.All(got.Registration!.Readiness, r => Assert.Equal("Ready", r.Status));
    }

    [Fact]
    public void AddTableForAnEnabledTableMakesNoWrites()
    {
        var e = new Env("account");
        e.Save(e.Get());
        e.Org.S.Writes.Clear();
        int updates = e.Org.S.Memory.Updates.Count;
        e.Change("AddTable", "account");
        Assert.Empty(e.Org.S.Writes);
        Assert.Equal(updates, e.Org.S.Memory.Updates.Count);
    }

    [Fact]
    public void NonAdministratorAddTableIsRefusedBeforeAnyWrite()
    {
        var e = new Env("account");
        var ex = Assert.Throws<EvaluationBlockedException>(() =>
            e.Change("AddTable", "contact", Guid.NewGuid())
        );
        Assert.Contains("System Administrator", ex.Message);
        Assert.Empty(e.Org.S.Writes);
        Assert.Empty(e.Org.S.Memory.Updates);
    }

    [Fact]
    public void AddTableOfAnUnknownTableNamesItAndWritesNothing()
    {
        var e = new Env(new[] { "account" }, new[] { "cr123_gone" });
        e.Org.S.MissingTables.Add("cr123_gone");
        var ex = Assert.Throws<EvaluationBlockedException>(() =>
            e.Change("AddTable", "cr123_gone")
        );
        Assert.Contains("cr123_gone", ex.Message);
        Assert.Empty(e.Org.S.Writes);
        Assert.Equal(new[] { "account" }, e.Rows());
        Assert.Equal(0, e.StepsFor("cr123_gone"));
    }

    [Fact]
    public void GetSucceedsWithoutAWorkerAndReportsAnUnknownWorker()
    {
        var e = new Env("account");
        e.Org.S.Memory.Rows[e.Runtime]["asx_workeruserid"] = Guid.Empty.ToString();
        var none = e.Get();
        Assert.Null(none.Registration!.Error);
        var unknown = Guid.NewGuid();
        e.Org.S.Memory.Rows[e.Runtime]["asx_workeruserid"] = unknown.ToString();
        e.Org.S.UnknownUsers.Add(unknown);
        var got = e.Get();
        Assert.Contains("could not be checked", got.Registration!.Error);
    }

    [Fact]
    public void SaveWithoutSystemJobsReadIsRefusedBeforeAnyWrite()
    {
        var e = new Env("account");
        var request = e.Get();
        e.Org.S.WorkerLacksSystemJobs = true;
        e.Org.S.Writes.Clear();
        int updates = e.Org.S.Memory.Updates.Count;
        var ex = Assert.Throws<EvaluationBlockedException>(() => e.Save(request));
        Assert.Contains("prvReadAsyncOperation", ex.Message);
        Assert.Empty(e.Org.S.Writes);
        Assert.Equal(updates, e.Org.S.Memory.Updates.Count);
        Assert.Equal(new[] { "account" }, e.Rows());
        Assert.Empty(e.Org.Steps);
    }

    [Fact]
    public void AddTableTheWorkerCannotReadGloballyIsRefusedBeforeAnyWrite()
    {
        var e = new Env("account");
        e.Org.S.WorkerReadDepth["contact"] = Microsoft.Crm.Sdk.Messages.PrivilegeDepth.Basic;
        int updates = e.Org.S.Memory.Updates.Count;
        var ex = Assert.Throws<EvaluationBlockedException>(() => e.Change("AddTable", "contact"));
        Assert.Contains("'contact'", ex.Message);
        Assert.Empty(e.Org.S.Writes);
        Assert.Equal(updates, e.Org.S.Memory.Updates.Count);
        Assert.Equal(new[] { "account" }, e.Rows());
        Assert.Empty(e.Org.Steps);
    }

    [Fact]
    public void AddTableOnALegacyProfileMigratesRowsAndClearsTheJson()
    {
        var e = new Env("account");
        e.ClearRows();
        Assert.False(e.Get().Migrated);
        var got = e.Change("AddTable", "contact");
        Assert.True(got.Migrated);
        Assert.Equal(new[] { "account", "contact" }, e.Rows());
        Assert.Equal(
            "[]",
            e.Org.S.Memory.Rows[e.Runtime].GetAttributeValue<string>("asx_allowedtables")
        );
    }

    [Fact]
    public void RemoveTableDeletesItsStepsAndRowOnly()
    {
        var e = new Env("account", "contact");
        e.Save(e.Get());
        int accountSteps = e.StepsFor("account");
        Assert.NotEqual(0, e.StepsFor("contact"));
        var got = e.Change("RemoveTable", "contact");
        Assert.Equal(new[] { "account" }, got.Tables);
        Assert.Equal(new[] { "account" }, e.Rows());
        Assert.Equal(0, e.StepsFor("contact"));
        Assert.Equal(accountSteps, e.StepsFor("account"));
    }

    [Fact]
    public void RemoveTableOfAnAbsentTableMakesNoWrites()
    {
        var e = new Env("account");
        e.Org.S.Writes.Clear();
        e.Change("RemoveTable", "contact");
        Assert.Empty(e.Org.S.Writes);
    }

    [Fact]
    public void RemoveTableIsRefusedWhileAWriterIsActive()
    {
        var e = new Env("account", "contact");
        e.Save(e.Get());
        var store = new DocumentStore(e.Org.S.Memory);
        var budget = store.Require<ConnectionBudget>("asx_claim", WorkCoordination.BudgetKey);
        budget.Value.Writers = new[] { "site" };
        store.Save(budget);
        e.Org.S.Writes.Clear();
        Assert.Throws<EvaluationBlockedException>(() => e.Change("RemoveTable", "contact"));
        Assert.Empty(e.Org.S.Writes);
        Assert.Equal(new[] { "account", "contact" }, e.Rows());
    }

    [Fact]
    public void NonAdministratorRemoveTableIsRefusedBeforeAnyWrite()
    {
        var e = new Env("account", "contact");
        var ex = Assert.Throws<EvaluationBlockedException>(() =>
            e.Change("RemoveTable", "contact", Guid.NewGuid())
        );
        Assert.Contains("System Administrator", ex.Message);
        Assert.Empty(e.Org.S.Writes);
        Assert.Empty(e.Org.S.Memory.Updates);
    }

    [Fact]
    public void SaveIgnoresTheTablesInTheRequest()
    {
        var e = new Env("account");
        var got = e.Get();
        got.Tables = new[] { "contact", "lead" };
        var saved = e.Save(got);
        Assert.Equal(new[] { "account" }, saved.Tables);
        Assert.Equal(new[] { "account" }, e.Rows());
        Assert.Equal(0, e.StepsFor("contact"));
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
    public void BatchedSaveRegistersAtMostTheLimitAndLeavesTheRestPending()
    {
        var added = Enumerable
            .Range(0, EventRegistrations.MaxNewTablesPerSave + 2)
            .Select(i => "cr123_t" + i)
            .ToArray();
        var e = new Env(new[] { "account" }.Concat(added).ToArray(), added);
        var saved = e.Save(e.Get());
        var newScopes = saved
            .Registration!.Readiness.Where(r => r.Scope == "account" || added.Contains(r.Scope))
            .ToList();
        Assert.Equal(
            EventRegistrations.MaxNewTablesPerSave,
            newScopes.Count(r => r.Status == "Ready")
        );
        Assert.Equal(3, newScopes.Count(r => r.Status == "Pending"));
        Assert.Equal(
            added.Length + 1,
            e.Org.S.Memory.Rows.Values.Count(r => r.LogicalName == "asx_runtimetable")
        );
        var second = e.Save(saved);
        Assert.All(second.Registration!.Readiness, r => Assert.Equal("Ready", r.Status));
    }

    [Fact]
    public void LegacyMigrationAboveLimitIsBatched()
    {
        var legacy = Enumerable
            .Range(0, EventRegistrations.MaxNewTablesPerSave + 1)
            .Select(i => "cr123_t" + i)
            .ToArray();
        var e = new Env(legacy, legacy);
        foreach (
            var row in e
                .Org.S.Memory.Rows.Values.Where(r => r.LogicalName == "asx_runtimetable")
                .ToList()
        )
            e.Org.S.Memory.Rows.Remove(row.Id);
        var got = e.Get();
        Assert.False(got.Migrated);
        var saved = e.Save(got);
        Assert.Equal(1, saved.Registration!.Readiness.Count(r => r.Status == "Pending"));
        var second = e.Save(saved);
        Assert.All(second.Registration!.Readiness, r => Assert.Equal("Ready", r.Status));
    }

    [Fact]
    public void StaleRuntimeSaveIsRefusedWithoutWrites()
    {
        var e = new Env("account");
        var got = e.Get();
        got.RowVersion = "stale";
        Assert.Throws<EvaluationBlockedException>(() => e.Save(got));
        Assert.Empty(e.Org.S.Writes);
        Assert.Empty(e.Org.S.Memory.Updates);
    }

    [Fact]
    public void NonAdministratorUnregisterIsRefused()
    {
        var e = new Env("account");
        e.Save(e.Get());
        int steps = e.Org.Steps.Length;
        Assert.NotEqual(0, steps);
        var r = e.Get();
        r.Command = "Unregister";
        Assert.Throws<EvaluationBlockedException>(() =>
            RuntimeAdministration.Execute(e.Org.S, r, true, Guid.NewGuid())
        );
        Assert.Equal(steps, e.Org.Steps.Length);
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

    [Fact]
    public void AddTableOnALegacyProfileBatchesRegistrationAndAlwaysIncludesTheNewTable()
    {
        var legacy = Enumerable
            .Range(0, EventRegistrations.MaxNewTablesPerSave + 1)
            .Select(i => "cr123_t" + i)
            .ToArray();
        var e = new Env(legacy, legacy.Concat(new[] { "cr123_new" }).ToArray());
        e.ClearRows();
        var got = e.Change("AddTable", "cr123_new");
        var scopes = got.Registration!.Readiness.Where(r =>
                r.Scope == "cr123_new" || legacy.Contains(r.Scope)
            )
            .ToList();
        Assert.Equal(
            EventRegistrations.MaxNewTablesPerSave,
            scopes.Count(r => r.Status == "Ready")
        );
        Assert.Equal(2, scopes.Count(r => r.Status == "Pending"));
        Assert.Equal("Ready", scopes.Single(r => r.Scope == "cr123_new").Status);
        Assert.Equal(legacy.Length + 1, e.Rows().Length);
    }

    [Fact]
    public void AddTableWithoutAWorkerIsRefusedAndWritesNothing()
    {
        var e = new Env("account");
        e.Org.S.Memory.Rows[e.Runtime]["asx_workeruserid"] = "";
        var ex = Assert.Throws<EvaluationBlockedException>(() => e.Change("AddTable", "contact"));
        Assert.Contains("Select the worker application user.", ex.Message);
        Assert.Empty(e.Org.S.Writes);
    }

    [Fact]
    public void RemoveTableWorksForTablesDeletedFromTheEnvironment()
    {
        var e = new Env(new[] { "account", "cr123_a", "cr123_b" }, new[] { "cr123_a", "cr123_b" });
        e.Save(e.Get());
        foreach (var gone in new[] { "cr123_a", "cr123_b" })
            e.Org.S.MissingTables.Add(gone);
        int accountSteps = e.StepsFor("account");
        e.Change("RemoveTable", "cr123_a");
        Assert.Equal(new[] { "account", "cr123_b" }, e.Rows());
        e.Change("RemoveTable", "cr123_b");
        Assert.Equal(new[] { "account" }, e.Rows());
        Assert.Equal(accountSteps, e.StepsFor("account"));
        Assert.Equal(0, e.StepsFor("cr123_a"));
        Assert.Equal(0, e.StepsFor("cr123_b"));
    }

    [Fact]
    public void AddTableOnAMigratedProfileChecksTheRuntimeRowVersion()
    {
        var e = new Env("account");
        var before = e.Org.S.Memory.Updates.Count;
        e.Change("AddTable", "contact");
        Assert.True(e.Org.S.Memory.Updates.Count > before);
    }
}
