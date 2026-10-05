using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

/// <summary>System Administrator detection by role template, directly or through team membership.</summary>
public static class AdministratorCheck
{
    public static readonly Guid SystemAdministratorTemplate = new Guid(
        "627090FF-40A3-4053-8790-584EDC5BE201"
    );

    public static bool IsSystemAdministrator(IOrganizationService service, Guid user)
    {
        var roles = Values(
            service,
            "systemuserroles",
            "systemuserid",
            new object[] { user },
            "roleid"
        );
        var teams = Values(
            service,
            "teammembership",
            "systemuserid",
            new object[] { user },
            "teamid"
        );
        if (teams.Count > 0)
            roles.AddRange(
                Values(service, "teamroles", "teamid", teams.Cast<object>().ToArray(), "roleid")
            );
        if (roles.Count == 0)
            return false;
        var query = new QueryExpression("role") { ColumnSet = new ColumnSet(false), TopCount = 1 };
        query.Criteria.AddCondition(
            "roleid",
            ConditionOperator.In,
            roles.Distinct().Cast<object>().ToArray()
        );
        query.Criteria.AddCondition(
            "roletemplateid",
            ConditionOperator.Equal,
            SystemAdministratorTemplate
        );
        return service.RetrieveMultiple(query).Entities.Count > 0;
    }

    private static List<Guid> Values(
        IOrganizationService service,
        string table,
        string filter,
        object[] keys,
        string column
    )
    {
        var query = new QueryExpression(table)
        {
            ColumnSet = new ColumnSet(column),
            PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 },
        };
        query.Criteria.AddCondition(filter, ConditionOperator.In, keys);
        var values = new List<Guid>();
        while (true)
        {
            var page = service.RetrieveMultiple(query);
            values.AddRange(page.Entities.Select(e => e.GetAttributeValue<Guid>(column)));
            if (!page.MoreRecords || string.IsNullOrEmpty(page.PagingCookie))
                return values;
            query.PageInfo.PageNumber++;
            query.PageInfo.PagingCookie = page.PagingCookie;
        }
    }
}
