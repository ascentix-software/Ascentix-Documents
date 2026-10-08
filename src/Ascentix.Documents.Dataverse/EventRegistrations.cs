using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class RegistrationSummary
{
    [DataMember]
    public List<TableReadiness> Readiness { get; set; } = new();

    [DataMember]
    public int ExtraSteps { get; set; }

    [DataMember]
    public string? Error { get; set; }

    [DataMember]
    public Guid[] StepIds { get; set; } = Array.Empty<Guid>();

    /// <summary>Scopes Register can fix: every status except Ready and WorkerCannotRead.</summary>
    [DataMember]
    public int Pending { get; set; }
}

/// <summary>Dataverse I/O for application-owned event steps. Ownership is by event handler.</summary>
public static class EventRegistrations
{
    public const string RecordHandler = "Ascentix.Documents.Plugins.RecordInvalidationPlugin";
    public const string RetirementHandler = "Ascentix.Documents.Plugins.TeamRetirementPlugin";
    public const string MembershipHandler =
        "Ascentix.Documents.Plugins.TeamMembershipInvalidationPlugin";

    /// <summary>
    /// Newly registered tables per call. Each new table creates and verifies its event steps
    /// inside one custom API call, which Dataverse stops after 2 minutes. Measured in DEV on
    /// 2026-10-05: about 0.6 s per new table plus about 5 s fixed (10 tables 11.1 s, 25 tables
    /// 20.1 s), so 90 tables take about 59 s, half of the limit.
    /// </summary>
    public const int MaxNewTablesPerSave = 90;

    private const string Step = "sdkmessageprocessingstep";
    private const string SystemJobsRead = "prvReadAsyncOperation";
    private const string SystemJobsMessage =
        "The worker application user cannot read System Jobs (prvReadAsyncOperation). Assign the Documents Worker role from this release before registering events.";

    public static RegistrationSummary Reconcile(
        IOrganizationService service,
        Guid worker,
        string[] tables,
        bool processUpdates
    )
    {
        var catalog = Catalog(service, tables);
        var privileges = WorkerPrivileges(service, worker);
        if (!privileges.ContainsKey(SystemJobsRead))
            throw new EvaluationBlockedException(SystemJobsMessage);
        foreach (var table in tables)
            if (!CanReadAll(catalog, privileges, table))
                throw new EvaluationBlockedException(CannotReadMessage(table));
        var desired = EventRegistrationPlan.Desired(catalog, tables, processUpdates);
        var changes = EventRegistrationPlan.Diff(desired, Steps(service, catalog), worker);
        foreach (var id in changes.Delete)
            service.Delete(Step, id);
        foreach (var pair in changes.Update)
            Write(service, pair.Value, worker, pair.Key);
        foreach (var spec in changes.Create)
            Write(service, spec, worker, null);
        return Summary(desired, Steps(service, catalog), worker);
    }

    public static RegistrationSummary Inspect(
        IOrganizationService service,
        Guid worker,
        string[] tables,
        bool processUpdates
    )
    {
        try
        {
            var catalog = Catalog(service, tables);
            var desired = EventRegistrationPlan.Desired(catalog, tables, processUpdates);
            var summary = Summary(desired, Steps(service, catalog), worker);
            if (worker == Guid.Empty)
                return summary;
            var privileges = WorkerPrivileges(service, worker);
            if (!privileges.ContainsKey(SystemJobsRead))
                summary.Error = SystemJobsMessage;
            foreach (var row in summary.Readiness)
                if (
                    row.Scope != EventRegistrationPlan.TeamScope
                    && !CanReadAll(catalog, privileges, row.Scope)
                )
                    row.Status = "WorkerCannotRead";
            return summary;
        }
        catch (EvaluationBlockedException ex)
        {
            return new RegistrationSummary { Error = ex.Message };
        }
    }

    public static int Unregister(IOrganizationService service)
    {
        var steps = Steps(service, Handlers(service, new RegistrationCatalog()));
        foreach (var step in steps)
            service.Delete(Step, step.Id);
        return steps.Count;
    }

    /// <summary>
    /// Deletes only the record-invalidation steps of one table, matched by the exact step name the
    /// plan gives them. No sdkmessagefilter query is made: Dataverse faults on a filter condition
    /// for a table that no longer exists, and a deleted table must remain removable.
    /// </summary>
    public static int RemoveTableSteps(IOrganizationService service, string table)
    {
        var names = new HashSet<string>(
            EventRegistrationPlan.RecordMessages.Select(m =>
                "Ascentix Documents: event " + m + " " + table
            ),
            StringComparer.Ordinal
        );
        var owned = RecordSteps(service)
            .Where(s => names.Contains(s.GetAttributeValue<string>("name") ?? ""))
            .ToList();
        foreach (var step in owned)
            service.Delete(Step, step.Id);
        return owned.Count;
    }

    /// <summary>
    /// Turns the record Update event steps on or off to match the record-updates setting. Steps are
    /// matched by the name the plan gives them. Nothing is registered or deleted and the worker is
    /// not checked, so the setting can change while the worker or a table has a problem.
    /// </summary>
    /// <param name="service">The administrator's organization service.</param>
    /// <param name="active">True when record updates are processed.</param>
    /// <returns>The number of steps whose state changed.</returns>
    public static int SetUpdateSteps(IOrganizationService service, bool active)
    {
        const string prefix = "Ascentix Documents: event Update ";
        var changed = RecordSteps(service, "statecode")
            .Select(Map)
            .Where(s => s.Name.StartsWith(prefix, StringComparison.Ordinal) && s.Active != active)
            .ToList();
        foreach (var step in changed)
            service.Update(
                new Entity(Step, step.Id)
                {
                    ["statecode"] = new OptionSetValue(active ? 0 : 1),
                    ["statuscode"] = new OptionSetValue(active ? 1 : 2),
                }
            );
        return changed.Count;
    }

    /// <summary>
    /// The steps of the record-invalidation handler, found without an sdkmessagefilter query.
    /// </summary>
    /// <param name="service">The organization service.</param>
    /// <param name="columns">Step columns to read in addition to the name.</param>
    /// <returns>The steps, or none when the handler is not installed exactly once.</returns>
    private static List<Entity> RecordSteps(IOrganizationService service, params string[] columns)
    {
        var handlers = All(
            service,
            In(
                new QueryExpression("plugintype") { ColumnSet = new ColumnSet("typename") },
                "typename",
                RecordHandler
            )
        );
        if (handlers.Count != 1)
            return new List<Entity>();
        var query = new QueryExpression(Step)
        {
            ColumnSet = new ColumnSet(new[] { "name" }.Concat(columns).ToArray()),
        };
        query.Criteria.AddCondition("eventhandler", ConditionOperator.Equal, handlers[0].Id);
        return All(service, query);
    }

    private static RegistrationSummary Summary(
        List<StepSpec> desired,
        List<ExistingStep> existing,
        Guid worker
    ) =>
        new RegistrationSummary
        {
            Readiness = EventRegistrationPlan.Readiness(desired, existing, worker),
            ExtraSteps = EventRegistrationPlan.Extra(desired, existing),
            StepIds = existing.Select(s => s.Id).ToArray(),
        };

    private static RegistrationCatalog Handlers(
        IOrganizationService service,
        RegistrationCatalog catalog
    )
    {
        var rows = All(
            service,
            In(
                new QueryExpression("plugintype") { ColumnSet = new ColumnSet("typename") },
                "typename",
                RecordHandler,
                RetirementHandler,
                MembershipHandler
            )
        );
        Guid One(string name)
        {
            var match = rows.Where(r => r.GetAttributeValue<string>("typename") == name).ToList();
            if (match.Count != 1)
                throw new EvaluationBlockedException(
                    "Exactly one installed plug-in type " + name + " is required."
                );
            return match[0].Id;
        }
        catalog.RecordHandler = One(RecordHandler);
        catalog.RetirementHandler = One(RetirementHandler);
        catalog.MembershipHandler = One(MembershipHandler);
        return catalog;
    }

    private static RegistrationCatalog Catalog(IOrganizationService service, string[] tables)
    {
        var catalog = Handlers(service, new RegistrationCatalog());
        var names = EventRegistrationPlan
            .RecordMessages.Concat(new[] { "Delete", "Associate", "Disassociate" })
            .Distinct()
            .ToArray();
        foreach (
            var row in All(
                service,
                In(
                    new QueryExpression("sdkmessage") { ColumnSet = new ColumnSet("name") },
                    "name",
                    names
                )
            )
        )
            catalog.Messages[row.GetAttributeValue<string>("name")] = row.Id;
        foreach (var name in names.Where(n => !catalog.Messages.ContainsKey(n)))
            throw new EvaluationBlockedException("Platform message " + name + " is missing.");
        foreach (var table in tables)
            catalog.ReadPrivileges[table] = RequireTable(service, table);
        var recordIds = EventRegistrationPlan
            .RecordMessages.Concat(new[] { "Delete" })
            .Distinct()
            .Select(n => (object)catalog.Messages[n])
            .ToArray();
        var filters = In(
            new QueryExpression("sdkmessagefilter")
            {
                ColumnSet = new ColumnSet(
                    "sdkmessageid",
                    "primaryobjecttypecode",
                    "iscustomprocessingstepallowed"
                ),
            },
            "primaryobjecttypecode",
            tables.Concat(new[] { EventRegistrationPlan.TeamScope }).Cast<object>().ToArray()
        );
        filters.Criteria.AddCondition("sdkmessageid", ConditionOperator.In, recordIds);
        foreach (var row in All(service, filters))
        {
            if (!row.GetAttributeValue<bool>("iscustomprocessingstepallowed"))
                continue;
            var messageId = row.GetAttributeValue<EntityReference>("sdkmessageid").Id;
            var key =
                catalog.Messages.First(p => p.Value == messageId).Key
                + "|"
                + row.GetAttributeValue<string>("primaryobjecttypecode");
            if (catalog.Filters.ContainsKey(key))
                throw new EvaluationBlockedException("Ambiguous event filter " + key + ".");
            catalog.Filters[key] = row.Id;
        }
        foreach (var table in tables)
        foreach (var message in EventRegistrationPlan.RecordMessages)
            if (!catalog.Filters.ContainsKey(message + "|" + table))
                throw new EvaluationBlockedException(
                    "Table '" + table + "' does not support " + message + " event registration."
                );
        if (!catalog.Filters.ContainsKey("Delete|team"))
            throw new EvaluationBlockedException("Team Delete event registration is unavailable.");
        return catalog;
    }

    public static string CannotReadMessage(string table) =>
        "The worker application user cannot read every '"
        + table
        + "' record (organization-level Read required). Grant it before enabling this table.";

    /// <summary>The worker's effective privileges by name; the value is the highest depth held.</summary>
    private static Dictionary<string, PrivilegeDepth> WorkerPrivileges(
        IOrganizationService service,
        Guid worker
    )
    {
        RetrieveUserPrivilegesResponse response;
        try
        {
            response = (RetrieveUserPrivilegesResponse)
                service.Execute(new RetrieveUserPrivilegesRequest { UserId = worker });
        }
        catch (EvaluationBlockedException)
        {
            throw;
        }
        catch (Exception)
        {
            throw new EvaluationBlockedException(
                "The worker application user could not be checked. Select an enabled worker application user."
            );
        }
        var held = new Dictionary<string, PrivilegeDepth>(StringComparer.OrdinalIgnoreCase);
        foreach (var privilege in response.RolePrivileges ?? Array.Empty<RolePrivilege>())
        {
            var name = privilege.PrivilegeName ?? privilege.PrivilegeId.ToString();
            if (!held.TryGetValue(name, out var depth) || privilege.Depth > depth)
                held[name] = privilege.Depth;
            var byId = privilege.PrivilegeId.ToString();
            if (!held.TryGetValue(byId, out depth) || privilege.Depth > depth)
                held[byId] = privilege.Depth;
        }
        return held;
    }

    private static bool CanReadAll(
        RegistrationCatalog catalog,
        Dictionary<string, PrivilegeDepth> privileges,
        string table
    ) =>
        catalog.ReadPrivileges.TryGetValue(table, out var id)
        && id.HasValue
        && privileges.TryGetValue(id.Value.ToString(), out var depth)
        && depth == PrivilegeDepth.Global;

    /// <summary>
    /// The table's Read privilege. A table deleted from the organization is refused without a
    /// faulting call: RetrieveEntity faults for it, and inside a plug-in that fault would end the
    /// transaction even when caught, so Get could not show the problem (see TableInfo.Find).
    /// </summary>
    private static Guid? RequireTable(IOrganizationService service, string table)
    {
        var metadata =
            TableInfo.Find(service, table)
            ?? throw new EvaluationBlockedException(
                "Table '" + table + "' does not exist in this environment."
            );
        return metadata
            .Privileges?.FirstOrDefault(p => p.PrivilegeType == PrivilegeType.Read)
            ?.PrivilegeId;
    }

    private static List<ExistingStep> Steps(
        IOrganizationService service,
        RegistrationCatalog catalog
    ) =>
        All(
                service,
                In(
                    new QueryExpression(Step)
                    {
                        ColumnSet = new ColumnSet(
                            "name",
                            "description",
                            "eventhandler",
                            "sdkmessageid",
                            "sdkmessagefilterid",
                            "stage",
                            "mode",
                            "impersonatinguserid",
                            "statecode",
                            "asyncautodelete"
                        ),
                    },
                    "eventhandler",
                    catalog.RecordHandler,
                    catalog.RetirementHandler,
                    catalog.MembershipHandler
                )
            )
            .Select(Map)
            .ToList();

    private static ExistingStep Map(Entity row) =>
        new ExistingStep
        {
            Id = row.Id,
            Name = row.GetAttributeValue<string>("name") ?? "",
            Description = row.GetAttributeValue<string>("description") ?? "",
            HandlerId = row.GetAttributeValue<EntityReference>("eventhandler")?.Id ?? Guid.Empty,
            MessageId = row.GetAttributeValue<EntityReference>("sdkmessageid")?.Id ?? Guid.Empty,
            FilterId = row.GetAttributeValue<EntityReference>("sdkmessagefilterid")?.Id,
            Stage = row.GetAttributeValue<OptionSetValue>("stage")?.Value ?? 0,
            Mode = row.GetAttributeValue<OptionSetValue>("mode")?.Value ?? 0,
            ImpersonatingUserId = row.GetAttributeValue<EntityReference>("impersonatinguserid")?.Id,
            Active = (row.GetAttributeValue<OptionSetValue>("statecode")?.Value ?? 0) == 0,
            AsyncAutoDelete = row.GetAttributeValue<bool>("asyncautodelete"),
        };

    private static void Write(IOrganizationService service, StepSpec spec, Guid worker, Guid? id)
    {
        var row = new Entity(Step)
        {
            ["name"] = spec.Name,
            ["description"] = EventRegistrationPlan.Label,
            ["impersonatinguserid"] = new EntityReference("systemuser", worker),
            ["stage"] = new OptionSetValue(40),
            ["mode"] = new OptionSetValue(1),
            ["asyncautodelete"] = true,
            ["rank"] = 1,
        };
        Guid stepId;
        if (id.HasValue)
        {
            row.Id = id.Value;
            service.Update(row);
            stepId = id.Value;
        }
        else
        {
            row["eventhandler"] = new EntityReference("plugintype", spec.HandlerId);
            row["sdkmessageid"] = new EntityReference("sdkmessage", spec.MessageId);
            row["sdkmessagefilterid"] = spec.FilterId.HasValue
                ? new EntityReference("sdkmessagefilter", spec.FilterId.Value)
                : null;
            row["supporteddeployment"] = new OptionSetValue(0);
            stepId = service.Create(row);
        }
        var current = Map(service.Retrieve(Step, stepId, new ColumnSet("statecode")));
        if (current.Active != spec.Active)
            service.Update(
                new Entity(Step, stepId)
                {
                    ["statecode"] = new OptionSetValue(spec.Active ? 0 : 1),
                    ["statuscode"] = new OptionSetValue(spec.Active ? 1 : 2),
                }
            );
        var actual = Map(
            service.Retrieve(
                Step,
                stepId,
                new ColumnSet(
                    "name",
                    "description",
                    "eventhandler",
                    "sdkmessageid",
                    "sdkmessagefilterid",
                    "stage",
                    "mode",
                    "impersonatinguserid",
                    "statecode",
                    "asyncautodelete"
                )
            )
        );
        if (!EventRegistrationPlan.Correct(actual, spec, worker))
            throw new EvaluationBlockedException(
                "Event registration readback differs for '" + spec.Name + "'."
            );
    }

    private static QueryExpression In(QueryExpression query, string column, params object[] values)
    {
        query.Criteria.AddCondition(column, ConditionOperator.In, values);
        return query;
    }

    private static List<Entity> All(IOrganizationService service, QueryExpression query)
    {
        // 5,000 = Dataverse maximum RetrieveMultiple page size.
        query.PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 };
        var rows = new List<Entity>();
        while (true)
        {
            var page = service.RetrieveMultiple(query);
            rows.AddRange(page.Entities);
            if (!page.MoreRecords)
                return rows;
            if (string.IsNullOrEmpty(page.PagingCookie))
                throw new EvaluationBlockedException(
                    "Registration inventory continuation missing."
                );
            query.PageInfo.PageNumber++;
            query.PageInfo.PagingCookie = page.PagingCookie;
        }
    }
}
