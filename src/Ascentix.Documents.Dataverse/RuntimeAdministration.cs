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
        if (TogglesOnly(old, request))
            return Toggle(service, old, request);
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
    /// Whether a Save changes only Enabled and/or ProcessRecordUpdates: the worker and the
    /// SharePoint hosts are unchanged. Save never changes the table list.
    /// </summary>
    /// <param name="old">The stored runtime row.</param>
    /// <param name="request">The Save request.</param>
    /// <returns>True when at least one of the two settings changes and nothing else does.</returns>
    private static bool TogglesOnly(Entity old, RuntimeRequest request)
    {
        bool changed =
            request.Enabled != old.GetAttributeValue<bool>("asx_enabled")
            || request.ProcessRecordUpdates
                != old.GetAttributeValue<bool>("asx_processrecordupdates");
        if (!changed)
            return false;
        if (
            !Guid.TryParse(old.GetAttributeValue<string>("asx_workeruserid"), out var worker)
            || worker != request.WorkerId
        )
            return false;
        string[] hosts;
        try
        {
            hosts = JsonWire.Read<string[]>(
                old.GetAttributeValue<string>("asx_sharepointhosts") ?? "[]"
            );
        }
        catch (EvaluationBlockedException)
        {
            return false;
        }
        catch (System.Runtime.Serialization.SerializationException)
        {
            return false;
        }
        return (request.SharePointHosts ?? Array.Empty<string>()).SequenceEqual(
            hosts ?? Array.Empty<string>(),
            StringComparer.OrdinalIgnoreCase
        );
    }

    /// <summary>
    /// Pauses, resumes or switches record-update processing at once. It does not wait for active
    /// writers and does not check the worker, hosts or tables first, so an administrator can always
    /// stop processing. Only the record Update steps change, to follow ProcessRecordUpdates. A
    /// resume reports worker, host and table problems in the result instead of refusing.
    /// </summary>
    /// <param name="service">The administrator's organization service.</param>
    /// <param name="old">The stored runtime row.</param>
    /// <param name="request">The Save request.</param>
    /// <returns>The saved runtime profile with its registration readiness.</returns>
    private static RuntimeRequest Toggle(
        IOrganizationService service,
        Entity old,
        RuntimeRequest request
    )
    {
        bool resuming = request.Enabled && !old.GetAttributeValue<bool>("asx_enabled");
        if (request.ProcessRecordUpdates != old.GetAttributeValue<bool>("asx_processrecordupdates"))
            EventRegistrations.SetUpdateSteps(service, request.ProcessRecordUpdates);
        service.Execute(
            new UpdateRequest
            {
                Target = new Entity("asx_runtime", old.Id)
                {
                    RowVersion = old.RowVersion,
                    ["asx_enabled"] = request.Enabled,
                    ["asx_processrecordupdates"] = request.ProcessRecordUpdates,
                },
                ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
            }
        );
        var result = Get(service, Profile(service));
        if (resuming)
            ReportProblems(service, result);
        return result;
    }

    /// <summary>
    /// Adds worker and host problems to the readiness result's error text without refusing.
    /// Table readiness is already part of the result.
    /// </summary>
    /// <param name="service">The administrator's organization service.</param>
    /// <param name="result">The runtime profile returned to the administrator.</param>
    private static void ReportProblems(IOrganizationService service, RuntimeRequest result)
    {
        var problems = new System.Collections.Generic.List<string>();
        if (!string.IsNullOrEmpty(result.Registration?.Error))
            problems.Add(result.Registration!.Error!);
        try
        {
            RuntimeProfile.ValidateHosts(result.SharePointHosts);
        }
        catch (EvaluationBlockedException ex)
        {
            problems.Add(ex.Message);
        }
        var worker = WorkerProblem(service, result.WorkerId);
        if (worker != null && !problems.Contains(worker))
            problems.Add(worker);
        if (problems.Count == 0)
            return;
        result.Registration ??= new RegistrationSummary();
        result.Registration.Error = string.Join(" ", problems);
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
            // No wait for active writers: in-flight work for a removed table stops at its next
            // worker step, before any SharePoint write, because the table is out of scope.
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
        var problem = WorkerProblem(service, worker);
        if (problem != null)
            throw new EvaluationBlockedException(problem);
    }

    /// <summary>Describes why a worker cannot run Documents work, without throwing.</summary>
    /// <param name="service">The administrator's organization service.</param>
    /// <param name="worker">The configured worker user ID.</param>
    /// <returns>The problem, or null when the worker is an enabled application user.</returns>
    private static string? WorkerProblem(IOrganizationService service, Guid worker)
    {
        if (worker == Guid.Empty)
            return "Select the worker application user.";
        // A query, unlike Retrieve, returns no row instead of faulting for an unknown user.
        var query = new QueryExpression("systemuser")
        {
            ColumnSet = new ColumnSet("isdisabled", "applicationid"),
            TopCount = 1,
        };
        query.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, worker);
        var user = service.RetrieveMultiple(query).Entities.FirstOrDefault();
        if (user == null)
            return "The worker user was not found. Select an enabled worker application user.";
        if (user.GetAttributeValue<bool>("isdisabled"))
            return "Worker identity is disabled.";
        if (user.GetAttributeValue<Guid>("applicationid") == Guid.Empty)
            return "The worker must be an application user; human users are not accepted.";
        return null;
    }
}
