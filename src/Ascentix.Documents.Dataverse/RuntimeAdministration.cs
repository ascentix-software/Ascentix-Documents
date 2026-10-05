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
    public string Table { get; set; } = "";

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
        if (
            request.Command != "Save"
            && request.Command != "Unregister"
            && request.Command != "AddTable"
            && request.Command != "RemoveTable"
        )
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
        if (request.Command == "AddTable" || request.Command == "RemoveTable")
            return ChangeTable(service, old, request);
        if (string.IsNullOrEmpty(request.RowVersion) || old.RowVersion != request.RowVersion)
            throw new EvaluationBlockedException("Refresh the runtime profile before saving.");
        RuntimeProfile.ValidateHosts(request.SharePointHosts);
        ValidateWorker(service, request.WorkerId);
        // Save never changes the table list; tables are added and removed one at a time.
        var tables = RuntimeTables.Effective(service, old).Tables;
        // Refuse a worker that cannot own the steps before anything is written.
        var probe = EventRegistrations.Inspect(
            service,
            request.WorkerId,
            tables,
            request.ProcessRecordUpdates
        );
        if (probe.Error != null)
            throw new EvaluationBlockedException(probe.Error);
        var unreadable = probe.Readiness.FirstOrDefault(r => r.Status == "WorkerCannotRead");
        if (unreadable != null)
            throw new EvaluationBlockedException(
                EventRegistrations.CannotReadMessage(unreadable.Scope)
            );
        WorkCoordination.RequireIdle(service);
        RuntimeTables.Replace(service, old.Id, tables);
        var registerNow = BatchedTables(
            service,
            request.WorkerId,
            tables,
            request.ProcessRecordUpdates,
            null
        );
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

    /// <summary>
    /// Registers in batches: each new table's steps are created inside one call, which Dataverse stops after 2 minutes.
    /// Registers every Ready table, the table being added (always, within the limit) and then the
    /// first not-Ready tables in list order; the rest stay Pending for a later Save.
    /// </summary>
    private static string[] BatchedTables(
        IOrganizationService service,
        Guid worker,
        string[] tables,
        bool processUpdates,
        string? include
    )
    {
        var probe = EventRegistrations.Inspect(service, worker, tables, processUpdates);
        if (probe.Error != null)
            return tables;
        var ready = new System.Collections.Generic.HashSet<string>(
            probe.Readiness.Where(r => r.Status == "Ready").Select(r => r.Scope),
            StringComparer.Ordinal
        );
        int slots =
            EventRegistrations.MaxNewTablesPerSave
            - (include != null && !ready.Contains(include) ? 1 : 0);
        var firstPending = new System.Collections.Generic.HashSet<string>(
            tables
                .Where(t =>
                    !ready.Contains(t) && !string.Equals(t, include, StringComparison.Ordinal)
                )
                .Take(slots),
            StringComparer.Ordinal
        );
        return tables
            .Where(t =>
                ready.Contains(t)
                || firstPending.Contains(t)
                || string.Equals(t, include, StringComparison.Ordinal)
            )
            .ToArray();
    }

    private static RuntimeRequest ChangeTable(
        IOrganizationService service,
        Entity old,
        RuntimeRequest request
    )
    {
        bool add = request.Command == "AddTable";
        var table = request.Table ?? "";
        RuntimeTables.Validate(new[] { table });
        var current = RuntimeTables.Effective(service, old);
        bool present = current.Tables.Contains(table, StringComparer.Ordinal);
        if (present == add)
            return Get(service, old);
        var worker = Guid.TryParse(TemplateStore.Text(old, "asx_workeruserid"), out var id)
            ? id
            : Guid.Empty;
        string[] tables;
        if (add)
        {
            ValidateWorker(service, worker);
            tables = current.Tables.Concat(new[] { table }).ToArray();
            bool updates = old.GetAttributeValue<bool>("asx_processrecordupdates");
            // Refuse unknown or unsupported tables before anything is written.
            var probe = EventRegistrations.Inspect(service, worker, tables, updates);
            if (probe.Error != null)
                throw new EvaluationBlockedException(probe.Error);
            if (probe.Readiness.Any(r => r.Scope == table && r.Status == "WorkerCannotRead"))
                throw new EvaluationBlockedException(EventRegistrations.CannotReadMessage(table));
            RuntimeTables.Replace(service, old.Id, tables);
            EventRegistrations.Reconcile(
                service,
                worker,
                BatchedTables(service, worker, tables, updates, table),
                updates
            );
        }
        else
        {
            WorkCoordination.RequireIdle(service);
            tables = current
                .Tables.Where(t => !string.Equals(t, table, StringComparison.Ordinal))
                .ToArray();
            // Only this table's own steps are touched, so tables already dropped from the environment stay removable.
            RuntimeTables.Replace(service, old.Id, tables);
            EventRegistrations.RemoveTableSteps(service, table);
        }
        // Always version-check the runtime row so concurrent table changes cannot overwrite each other.
        service.Execute(
            new UpdateRequest
            {
                Target = new Entity("asx_runtime", old.Id)
                {
                    RowVersion = old.RowVersion,
                    ["asx_allowedtables"] = "[]",
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
