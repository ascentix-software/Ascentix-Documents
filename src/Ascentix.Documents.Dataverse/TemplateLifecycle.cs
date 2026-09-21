using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public static class TemplateLifecycle
{
    public static ColumnSet Columns() =>
        new ColumnSet(
            "asx_table",
            "asx_publishedrevisionid",
            "asx_disabled",
            "asx_startsutc",
            "asx_endsutc"
        );

    public static Entity? Find(IOrganizationService service, Guid id)
    {
        var query = new QueryExpression("asx_template") { ColumnSet = Columns(), TopCount = 2 };
        query.Criteria.AddCondition("asx_templateid", ConditionOperator.Equal, id);
        return service.RetrieveMultiple(query).Entities.SingleOrDefault();
    }

    public static bool Active(Entity? row, DateTime utc)
    {
        if (row == null || row.GetAttributeValue<bool>("asx_disabled"))
            return false;
        var start = row.GetAttributeValue<DateTime?>("asx_startsutc");
        var end = row.GetAttributeValue<DateTime?>("asx_endsutc");
        return (!start.HasValue || start.Value <= utc) && (!end.HasValue || utc < end.Value);
    }

    public static void Validate(Entity target, Entity? existing)
    {
        DateTime? Read(string key) =>
            target.Contains(key)
                ? target.GetAttributeValue<DateTime?>(key)
                : existing?.GetAttributeValue<DateTime?>(key);
        var start = Read("asx_startsutc");
        var end = Read("asx_endsutc");
        if (start.HasValue && end.HasValue && end <= start)
            throw new EvaluationBlockedException("End date must be after start date.");
    }
}
