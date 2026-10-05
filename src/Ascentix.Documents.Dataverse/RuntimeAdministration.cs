using System;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class RuntimeRequest
{
    [DataMember]
    public string Command { get; set; } = "Get";

    [DataMember]
    public string RowVersion { get; set; } = "";

    [DataMember]
    public Guid WorkerId { get; set; }

    [DataMember]
    public bool Enabled { get; set; }

    [DataMember]
    public bool ProcessRecordUpdates { get; set; }

    [DataMember]
    public string[] SharePointHosts { get; set; } = Array.Empty<string>();

    [DataMember]
    public string[] Tables { get; set; } = Array.Empty<string>();

    [DataMember]
    public bool Migrated { get; set; }

    [DataMember]
    public RegistrationSummary? Registration { get; set; }
}

public static class RuntimeAdministration
{
    public static RuntimeRequest Execute(
        IOrganizationService service,
        RuntimeRequest request,
        bool transaction,
        Guid caller
    )
    {
        if (!transaction)
            throw new EvaluationBlockedException("Runtime configuration requires a transaction.");
        var old = Profile(service);
        if (request.Command == "Get")
            return Get(service, old);
        if (request.Command != "Save" && request.Command != "Unregister")
            throw new EvaluationBlockedException("Unsupported runtime command.");
        if (!AdministratorCheck.IsSystemAdministrator(service, caller))
            throw new EvaluationBlockedException(
                "Registering events requires System Administrator."
            );
        if (request.Command == "Unregister")
        {
            EventRegistrations.Unregister(service);
            return Get(service, old);
        }
        if (string.IsNullOrEmpty(request.RowVersion) || old.RowVersion != request.RowVersion)
            throw new EvaluationBlockedException("Refresh the runtime profile before saving.");
        RuntimeTables.Validate(request.Tables);
        RuntimeProfile.ValidateHosts(request.SharePointHosts);
        ValidateWorker(service, request.WorkerId);
        WorkCoordination.RequireIdle(service);
        RuntimeTables.Replace(service, old.Id, request.Tables);
        // Register in batches: each new table's steps are created inside this one call, which Dataverse stops after 2 minutes.
        var probe = EventRegistrations.Inspect(
            service,
            request.WorkerId,
            request.Tables,
            request.ProcessRecordUpdates
        );
        var registerNow = request.Tables;
        if (probe.Error == null)
        {
            var ready = new System.Collections.Generic.HashSet<string>(
                probe.Readiness.Where(r => r.Status == "Ready").Select(r => r.Scope),
                StringComparer.Ordinal
            );
            var firstPending = new System.Collections.Generic.HashSet<string>(
                request
                    .Tables.Where(t => !ready.Contains(t))
                    .Take(EventRegistrations.MaxNewTablesPerSave),
                StringComparer.Ordinal
            );
            registerNow = request
                .Tables.Where(t => ready.Contains(t) || firstPending.Contains(t))
                .ToArray();
        }
        EventRegistrations.Reconcile(
            service,
            request.WorkerId,
            registerNow,
            request.ProcessRecordUpdates
        );
        service.Execute(
            new UpdateRequest
            {
                Target = new Entity("asx_runtime", old.Id)
                {
                    RowVersion = old.RowVersion,
                    ["asx_workeruserid"] = request.WorkerId.ToString("D"),
                    ["asx_enabled"] = request.Enabled,
                    ["asx_processrecordupdates"] = request.ProcessRecordUpdates,
                    ["asx_allowedtables"] = "[]",
                    ["asx_sharepointhosts"] = JsonWire.Write(request.SharePointHosts),
                },
                ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
            }
        );
        return Get(service, Profile(service));
    }

    private static Entity Profile(IOrganizationService service)
    {
        var query = new QueryExpression("asx_runtime")
        {
            ColumnSet = new ColumnSet(
                "asx_workeruserid",
                "asx_enabled",
                "asx_processrecordupdates",
                "asx_allowedtables",
                "asx_sharepointhosts"
            ),
            TopCount = 2,
        };
        query.Criteria.AddCondition("asx_name", ConditionOperator.Equal, "Default");
        var rows = service.RetrieveMultiple(query);
        if (rows.MoreRecords || rows.Entities.Count != 1)
            throw new EvaluationBlockedException(
                "Exactly one installed runtime profile is required."
            );
        return rows.Entities[0];
    }

    private static RuntimeRequest Get(IOrganizationService service, Entity old)
    {
        var worker = Guid.TryParse(TemplateStore.Text(old, "asx_workeruserid"), out var id)
            ? id
            : Guid.Empty;
        var tables = RuntimeTables.Effective(service, old);
        bool updates = old.GetAttributeValue<bool>("asx_processrecordupdates");
        return new RuntimeRequest
        {
            Command = "Get",
            RowVersion = old.RowVersion,
            WorkerId = worker,
            Enabled = old.GetAttributeValue<bool>("asx_enabled"),
            ProcessRecordUpdates = updates,
            SharePointHosts = JsonWire.Read<string[]>(
                old.GetAttributeValue<string>("asx_sharepointhosts") ?? "[]"
            ),
            Tables = tables.Tables,
            Migrated = tables.Migrated,
            Registration = EventRegistrations.Inspect(service, worker, tables.Tables, updates),
        };
    }

    private static void ValidateWorker(IOrganizationService service, Guid worker)
    {
        if (worker == Guid.Empty)
            throw new EvaluationBlockedException("Select the worker application user.");
        var user = service.Retrieve(
            "systemuser",
            worker,
            new ColumnSet("isdisabled", "applicationid")
        );
        if (user.GetAttributeValue<bool>("isdisabled"))
            throw new EvaluationBlockedException("Worker identity is disabled.");
        if (user.GetAttributeValue<Guid>("applicationid") == Guid.Empty)
            throw new EvaluationBlockedException(
                "The worker must be an application user; human users are not accepted."
            );
    }
}
