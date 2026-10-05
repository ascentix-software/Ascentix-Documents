using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class RuntimeTablesTests
{
    [Fact]
    public void ProfileReadsRowsWithoutACountCap()
    {
        var s = new DurableWorkerTests.MemoryService();
        var tables = Enumerable.Range(0, 120).Select(i => "cr123_table" + i).ToArray();
        RuntimeSeed.Seed(s, Guid.NewGuid(), tables);
        Assert.Equal(120, RuntimeProfile.Read(s).Tables.Length);
    }

    [Fact]
    public void UnmigratedProfileFallsBackToLegacyJson()
    {
        var s = new DurableWorkerTests.MemoryService();
        var runtime = RuntimeSeed.Seed(s, Guid.NewGuid(), "account", "contact");
        foreach (var row in s.Rows.Values.Where(r => r.LogicalName == "asx_runtimetable").ToList())
            s.Rows.Remove(row.Id);
        var effective = RuntimeTables.Effective(s, s.Rows[runtime]);
        Assert.Equal(new[] { "account", "contact" }, effective.Tables);
        Assert.False(effective.Migrated);
        Assert.Equal(new[] { "account", "contact" }, RuntimeProfile.Read(s).Tables);
    }

    [Fact]
    public void ReplaceAddsAndRemovesRowsAndValidatesNames()
    {
        var s = new StepService();
        var runtime = RuntimeSeed.Seed(s.Memory, Guid.NewGuid(), "account", "contact");
        RuntimeTables.Replace(s, runtime, new[] { "contact", "cr123_project" });
        Assert.Equal(
            new[] { "contact", "cr123_project" },
            RuntimeTables.Effective(s, s.Memory.Rows[runtime]).Tables
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            RuntimeTables.Replace(s, runtime, new[] { "Bad Name" })
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            RuntimeTables.Replace(s, runtime, new[] { "account", "account" })
        );
        s.Memory.Rows[runtime]["asx_allowedtables"] = "[]";
        RuntimeTables.Replace(s, runtime, Array.Empty<string>());
        Assert.Empty(RuntimeTables.Effective(s, s.Memory.Rows[runtime]).Tables);
    }

    [Fact]
    public void CaptureProfileIgnoresCorruptHosts()
    {
        var s = new DurableWorkerTests.MemoryService();
        var worker = Guid.NewGuid();
        var runtime = RuntimeSeed.Seed(s, worker, "account");
        s.Rows[runtime]["asx_sharepointhosts"] = "not json";
        Assert.ThrowsAny<Exception>(() => RuntimeProfile.Read(s));
        var capture = RuntimeProfile.ReadCapture(s);
        Assert.Equal(worker, capture.WorkerId);
        Assert.Equal(new[] { "account" }, capture.Tables);
    }

    [Theory]
    [InlineData("direct", true)]
    [InlineData("team", true)]
    [InlineData("customizer", false)]
    [InlineData("none", false)]
    public void AdministratorCheckHonoursDirectAndTeamRoles(string grant, bool expected)
    {
        var s = new DurableWorkerTests.MemoryService();
        var user = Guid.NewGuid();
        var admin = Guid.NewGuid();
        var customizer = Guid.NewGuid();
        s.Seed(
            new Entity("role", admin)
            {
                ["roletemplateid"] = new EntityReference(
                    "roletemplate",
                    AdministratorCheck.SystemAdministratorTemplate
                ),
            }
        );
        s.Seed(
            new Entity("role", customizer)
            {
                ["roletemplateid"] = new EntityReference("roletemplate", Guid.NewGuid()),
            }
        );
        if (grant == "direct")
            s.Seed(
                new Entity("systemuserroles", Guid.NewGuid())
                {
                    ["systemuserid"] = user,
                    ["roleid"] = admin,
                }
            );
        if (grant == "customizer")
            s.Seed(
                new Entity("systemuserroles", Guid.NewGuid())
                {
                    ["systemuserid"] = user,
                    ["roleid"] = customizer,
                }
            );
        if (grant == "team")
        {
            var team = Guid.NewGuid();
            s.Seed(
                new Entity("teammembership", Guid.NewGuid())
                {
                    ["systemuserid"] = user,
                    ["teamid"] = team,
                }
            );
            s.Seed(
                new Entity("teamroles", Guid.NewGuid()) { ["teamid"] = team, ["roleid"] = admin }
            );
        }
        Assert.Equal(expected, AdministratorCheck.IsSystemAdministrator(s, user));
    }
}
