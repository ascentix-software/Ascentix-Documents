using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public static class RecordUpdateRegistrations
{
    private const string Ownership = "Owned by standalone Ascentix Documents development setup v1";
    private const string Handler = "Ascentix.Documents.Plugins.RecordInvalidationPlugin";

    public static void Apply(
        IOrganizationService service,
        Guid worker,
        string[] tables,
        bool enabled
    )
    {
        var handlers = new QueryExpression("plugintype")
        {
            ColumnSet = new ColumnSet(false),
            TopCount = 2,
        };
        handlers.Criteria.AddCondition("typename", ConditionOperator.Equal, Handler);
        var plugins = service.RetrieveMultiple(handlers);
        if (plugins.MoreRecords || plugins.Entities.Count != 1)
            throw new InvalidPluginExecutionException(
                "Exactly one installed record-event handler is required."
            );
        var query = new QueryExpression("sdkmessageprocessingstep")
        {
            ColumnSet = new ColumnSet(
                "name",
                "description",
                "eventhandler",
                "sdkmessageid",
                "sdkmessagefilterid",
                "impersonatinguserid",
                "stage",
                "mode",
                "statecode"
            ),
            TopCount = 251,
        };
        query.Criteria.AddCondition("description", ConditionOperator.Equal, Ownership);
        query.Criteria.AddCondition(
            "eventhandler",
            ConditionOperator.Equal,
            plugins.Entities[0].Id
        );
        var rows = service.RetrieveMultiple(query);
        if (rows.MoreRecords || rows.Entities.Count > 250)
            throw new InvalidPluginExecutionException(
                "Record-update registration inventory exceeds its bound."
            );
        var changes = new List<Entity>();
        var identities = new HashSet<string>(StringComparer.Ordinal);
        foreach (var step in rows.Entities)
        {
            var messageRef = step.GetAttributeValue<EntityReference>("sdkmessageid");
            if (messageRef?.LogicalName != "sdkmessage")
                throw new InvalidPluginExecutionException(
                    "Record-event message reference is incomplete."
                );
            var message = service
                .Retrieve("sdkmessage", messageRef.Id, new ColumnSet("name"))
                .GetAttributeValue<string>("name");
            if (message != "Update" && message != "UpdateMultiple")
                continue;
            var handlerRef = step.GetAttributeValue<EntityReference>("eventhandler");
            var filterRef = step.GetAttributeValue<EntityReference>("sdkmessagefilterid");
            if (
                messageRef?.LogicalName != "sdkmessage"
                || handlerRef?.LogicalName != "plugintype"
                || filterRef?.LogicalName != "sdkmessagefilter"
            )
                throw new InvalidPluginExecutionException(
                    "Record-update registration references are incomplete."
                );
            var handler = service
                .Retrieve("plugintype", handlerRef.Id, new ColumnSet("typename"))
                .GetAttributeValue<string>("typename");
            var filter = service.Retrieve(
                "sdkmessagefilter",
                filterRef.Id,
                new ColumnSet("primaryobjecttypecode", "sdkmessageid")
            );
            var table = filter.GetAttributeValue<string>("primaryobjecttypecode");
            if (
                (message != "Update" && message != "UpdateMultiple")
                || handler != Handler
                || string.IsNullOrEmpty(table)
                || filter.GetAttributeValue<EntityReference>("sdkmessageid")?.Id != messageRef.Id
                || step.GetAttributeValue<string>("name")
                    != "Ascentix Documents: event " + message + " " + table
                || step.GetAttributeValue<OptionSetValue>("stage")?.Value != 40
                || step.GetAttributeValue<OptionSetValue>("mode")?.Value != 0
                || !identities.Add(message + ":" + table)
            )
                throw new InvalidPluginExecutionException(
                    "Record-update registration differs from the installed product contract."
                );
            bool activate = enabled && tables.Contains(table, StringComparer.Ordinal);
            if (
                activate
                && step.GetAttributeValue<EntityReference>("impersonatinguserid")?.Id != worker
            )
                throw new InvalidPluginExecutionException(
                    "Record-update registration uses a different worker. Repair event setup first."
                );
            int state = activate ? 0 : 1;
            if (step.GetAttributeValue<OptionSetValue>("statecode")?.Value != state)
                changes.Add(
                    new Entity("sdkmessageprocessingstep", step.Id)
                    {
                        ["statecode"] = new OptionSetValue(state),
                        ["statuscode"] = new OptionSetValue(activate ? 1 : 2),
                    }
                );
        }
        if (enabled && !tables.Any(table => identities.Contains("Update:" + table)))
            throw new InvalidPluginExecutionException(
                "Install record-update event steps before enabling update processing."
            );
        // Validate the complete inventory before changing any registration.
        foreach (var change in changes)
        {
            service.Update(change);
            var actual = service.Retrieve(
                "sdkmessageprocessingstep",
                change.Id,
                new ColumnSet("statecode")
            );
            if (
                actual.GetAttributeValue<OptionSetValue>("statecode")?.Value
                != change.GetAttributeValue<OptionSetValue>("statecode").Value
            )
                throw new InvalidPluginExecutionException(
                    "Record-update registration state did not match the requested setting."
                );
        }
    }
}
