using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
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
}

/// <summary>Dataverse I/O for application-owned event steps. Ownership is by event handler.</summary>
public static class EventRegistrations
{
    public const string RecordHandler = "Ascentix.Documents.Plugins.RecordInvalidationPlugin";
    public const string RetirementHandler = "Ascentix.Documents.Plugins.TeamRetirementPlugin";
    public const string MembershipHandler =
        "Ascentix.Documents.Plugins.TeamMembershipInvalidationPlugin";

    /// <summary>
    /// Newly added tables per Save. Each new table creates and verifies its steps inside one
    /// custom API call, which Dataverse stops after 2 minutes. Value measured in DEV (Task 13).
    /// </summary>
    public const int MaxNewTablesPerSave = 25;

    private const string Step = "sdkmessageprocessingstep";

    public static RegistrationSummary Reconcile(
        IOrganizationService service,
        Guid worker,
        string[] tables,
        bool processUpdates
    )
    {
        var catalog = Catalog(service, tables);
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
            return Summary(desired, Steps(service, catalog), worker);
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
        filters.Criteria.AddCondition(
            "sdkmessageid",
            ConditionOperator.In,
            catalog.Messages.Values.Cast<object>().ToArray()
        );
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
