using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public sealed class EffectiveTables
{
    public string[] Tables { get; set; } = Array.Empty<string>();

    /// <summary>False while the allowlist still lives only in the legacy JSON column.</summary>
    public bool Migrated { get; set; }
}

/// <summary>The relational runtime allowlist (asx_runtimetable), one row per business table.</summary>
public static class RuntimeTables
{
    public const string Entity = "asx_runtimetable";

    // Dataverse logical names are lowercase, start with a letter and are at most 100 characters.
    private static readonly Regex LogicalName = new Regex("\\A[a-z][a-z0-9_]{0,99}\\z");

    public static void Validate(string[]? tables)
    {
        if (
            tables == null
            || tables.Distinct(StringComparer.Ordinal).Count() != tables.Length
            || tables.Any(t => t == null || !LogicalName.IsMatch(t))
        )
            throw new EvaluationBlockedException(
                "Source tables must be distinct Dataverse logical names (lowercase, at most 100 characters)."
            );
    }

    public static EffectiveTables Effective(IOrganizationService service, Entity runtime)
    {
        var rows = Rows(service, runtime.Id).Select(Name).ToArray();
        if (rows.Length > 0)
            return new EffectiveTables { Tables = rows, Migrated = true };
        var legacy = Legacy(runtime);
        return new EffectiveTables { Tables = legacy, Migrated = legacy.Length == 0 };
    }

    public static void Replace(IOrganizationService service, Guid runtimeId, string[] tables)
    {
        Validate(tables);
        var existing = Rows(service, runtimeId);
        foreach (var row in existing.Where(r => !tables.Contains(Name(r), StringComparer.Ordinal)))
            service.Delete(Entity, row.Id);
        var kept = existing.Select(Name).ToArray();
        foreach (var table in tables.Where(t => !kept.Contains(t, StringComparer.Ordinal)))
            service.Create(
                new Entity(Entity)
                {
                    ["asx_name"] = table,
                    ["asx_logicalname"] = table,
                    ["asx_runtimeid"] = new EntityReference("asx_runtime", runtimeId),
                }
            );
    }

    private static string Name(Entity row) =>
        row.GetAttributeValue<string>("asx_logicalname") ?? "";

    private static string[] Legacy(Entity runtime)
    {
        var raw = runtime.GetAttributeValue<string>("asx_allowedtables");
        return string.IsNullOrWhiteSpace(raw)
            ? Array.Empty<string>()
            : JsonWire.Read<string[]>(raw) ?? Array.Empty<string>();
    }

    private static List<Entity> Rows(IOrganizationService service, Guid runtimeId)
    {
        var query = new QueryExpression(Entity)
        {
            ColumnSet = new ColumnSet("asx_logicalname"),
            PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 },
        };
        query.Criteria.AddCondition("asx_runtimeid", ConditionOperator.Equal, runtimeId);
        query.AddOrder("asx_logicalname", OrderType.Ascending);
        var rows = new List<Entity>();
        while (true)
        {
            var page = service.RetrieveMultiple(query);
            rows.AddRange(page.Entities);
            if (!page.MoreRecords)
                return rows;
            if (string.IsNullOrEmpty(page.PagingCookie))
                throw new EvaluationBlockedException("Runtime table continuation missing.");
            query.PageInfo.PageNumber++;
            query.PageInfo.PagingCookie = page.PagingCookie;
        }
    }
}
