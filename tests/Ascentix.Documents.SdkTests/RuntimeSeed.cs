using System;
using System.Linq;
using Microsoft.Xrm.Sdk;

namespace Ascentix.Documents.SdkTests;

/// <summary>Seeds the Default runtime profile in both the 0.1.0.4 row form and the legacy JSON form.</summary>
internal static class RuntimeSeed
{
    internal static Guid Seed(
        DurableWorkerTests.MemoryService service,
        Guid worker,
        params string[] tables
    )
    {
        if (!service.Rows.ContainsKey(worker))
            service.Seed(
                new Entity("systemuser", worker)
                {
                    ["isdisabled"] = false,
                    ["applicationid"] = Guid.NewGuid(),
                }
            );
        var runtime = Guid.NewGuid();
        service.Seed(
            new Entity("asx_runtime", runtime)
            {
                ["asx_name"] = "Default",
                ["asx_enabled"] = false,
                // Off, as 0.1.0.4 stores it; tests of an environment upgraded with the
                // setting on store true themselves (RecordUpdatesStoredOn).
                ["asx_processrecordupdates"] = false,
                ["asx_workeruserid"] = worker.ToString(),
                ["asx_allowedtables"] =
                    "[" + string.Join(",", tables.Select(t => "\"" + t + "\"")) + "]",
                ["asx_sharepointhosts"] = "[\"example.sharepoint.com\"]",
            }
        );
        foreach (var table in tables)
            service.Seed(
                new Entity("asx_runtimetable", Guid.NewGuid())
                {
                    ["asx_name"] = table,
                    ["asx_logicalname"] = table,
                    ["asx_runtimeid"] = new EntityReference("asx_runtime", runtime),
                }
            );
        return runtime;
    }

    /// <summary>
    /// Stores "Update folders when records change" as on in the Default runtime profile, as an
    /// environment that had it on before 0.1.0.4 has it.
    /// </summary>
    internal static void RecordUpdatesStoredOn(DurableWorkerTests.MemoryService service) =>
        service.Rows.Values.Single(r => r.LogicalName == "asx_runtime")[
            "asx_processrecordupdates"
        ] = true;
}
