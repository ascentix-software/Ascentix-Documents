using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Domain;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public sealed class RuntimeProfile
{
    public Guid WorkerId { get; private set; }
    public bool Enabled { get; private set; }
    public bool ProcessRecordUpdates { get; private set; }
    private IOrganizationService service = null!;
    public string[] SharePointHosts { get; private set; } = Array.Empty<string>();
    public string[] Tables { get; private set; } = Array.Empty<string>();

    public static RuntimeProfile Read(IOrganizationService service) => Load(service, true);

    /// <summary>
    /// Minimal profile for event capture and the outbox guard: worker, tables and update flag only.
    /// SharePoint hosts are neither read nor validated, so host configuration cannot stop capture.
    /// </summary>
    public static RuntimeProfile ReadCapture(IOrganizationService service) => Load(service, false);

    private static RuntimeProfile Load(IOrganizationService service, bool hosts)
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
                "An explicit approved runtime profile is required."
            );
        var row = rows.Entities[0];
        var siteHosts = Array.Empty<string>();
        if (hosts)
        {
            siteHosts = JsonWire.Read<string[]>(
                row.GetAttributeValue<string>("asx_sharepointhosts")
            );
            ValidateHosts(siteHosts);
        }
        var tables = RuntimeTables.Effective(service, row).Tables;
        if (
            !Guid.TryParse(row.GetAttributeValue<string>("asx_workeruserid"), out var worker)
            || worker == Guid.Empty
            || tables.Distinct(StringComparer.Ordinal).Count() != tables.Length
        )
            throw new EvaluationBlockedException("Runtime identity/source scope is incomplete.");
        foreach (var table in tables)
            new FieldReference("root", table);
        return new RuntimeProfile
        {
            WorkerId = worker,
            Enabled = row.GetAttributeValue<bool>("asx_enabled"),
            ProcessRecordUpdates = row.GetAttributeValue<bool>("asx_processrecordupdates"),
            Tables = tables,
            SharePointHosts = siteHosts,
            service = service,
        };
    }

    /// <summary>The SharePoint hosts the Runtime panel allows, without the rest of the profile.</summary>
    public static string[] ReadHosts(IOrganizationService service)
    {
        var query = new QueryExpression("asx_runtime")
        {
            ColumnSet = new ColumnSet("asx_sharepointhosts"),
            TopCount = 2,
        };
        query.Criteria.AddCondition("asx_name", ConditionOperator.Equal, "Default");
        var rows = service.RetrieveMultiple(query);
        if (rows.MoreRecords || rows.Entities.Count != 1)
            throw new EvaluationBlockedException(
                "An explicit approved runtime profile is required."
            );
        return JsonWire.Read<string[]>(
                rows.Entities[0].GetAttributeValue<string>("asx_sharepointhosts")
            ) ?? Array.Empty<string>();
    }

    /// <summary>
    /// Whether record updates are processed, read alone. Used only to word what a waiting
    /// folder needs, so a missing profile reads as Off: a replan works either way.
    /// </summary>
    public static bool RecordUpdates(IOrganizationService service)
    {
        var query = new QueryExpression("asx_runtime")
        {
            ColumnSet = new ColumnSet("asx_processrecordupdates"),
            TopCount = 2,
        };
        query.Criteria.AddCondition("asx_name", ConditionOperator.Equal, "Default");
        var rows = service.RetrieveMultiple(query);
        return rows.Entities.Count == 1
            && rows.Entities[0].GetAttributeValue<bool>("asx_processrecordupdates");
    }

    /// <summary>
    /// Whether an active SharePoint site record has exactly this address, as the transport gate
    /// (ValidateTransport) requires before Documents calls the site.
    /// </summary>
    public static bool Registered(IOrganizationService service, string siteUrl)
    {
        var native = new QueryExpression("sharepointsite")
        {
            ColumnSet = new ColumnSet(false),
            TopCount = 1,
        };
        var urls = new FilterExpression(LogicalOperator.Or);
        urls.AddCondition("absoluteurl", ConditionOperator.Equal, siteUrl);
        urls.AddCondition("absoluteurl", ConditionOperator.Equal, siteUrl + "/");
        native.Criteria.AddFilter(urls);
        native.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
        return service.RetrieveMultiple(native).Entities.Count > 0;
    }

    public static void ValidateHosts(string[] hosts)
    {
        if (
            hosts == null
            || hosts.Length < 1
            || hosts.Length > 20
            || hosts.Distinct(StringComparer.OrdinalIgnoreCase).Count() != hosts.Length
            || hosts.Any(host =>
                !System.Text.RegularExpressions.Regex.IsMatch(
                    host ?? "",
                    @"\A[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.sharepoint\.com\z"
                )
            )
        )
            throw new EvaluationBlockedException(
                "Explicit canonical SharePoint tenant hosts are required."
            );
    }

    public void ValidateTransport(WorkerResult result)
    {
        if (result.Http == null)
            return;
        if (
            !Uri.TryCreate(result.SiteUrl, UriKind.Absolute, out var url)
            || url.Scheme != "https"
            || !url.IsDefaultPort
            || url.UserInfo.Length != 0
            || url.Query.Length != 0
            || url.Fragment.Length != 0
            || url.AbsoluteUri.TrimEnd('/') != result.SiteUrl
            || !SharePointHosts.Contains(url.Host, StringComparer.OrdinalIgnoreCase)
            || (result.Http.Method != "GET" && result.Http.Method != "POST")
            || !result.Http.RelativeUri.StartsWith("_api/", StringComparison.Ordinal)
            || result.Http.RelativeUri.Any(char.IsControl)
        )
            throw new EvaluationBlockedException(
                "Issued transport is outside the configured tenant or supported API methods."
            );
        var native = new QueryExpression("sharepointsite")
        {
            ColumnSet = new ColumnSet(false),
            TopCount = 1,
        };
        var urls = new FilterExpression(LogicalOperator.Or);
        urls.AddCondition("absoluteurl", ConditionOperator.Equal, result.SiteUrl);
        urls.AddCondition("absoluteurl", ConditionOperator.Equal, result.SiteUrl + "/");
        native.Criteria.AddFilter(urls);
        native.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
        if (service.RetrieveMultiple(native).Entities.Count == 0)
            throw new EvaluationBlockedException(
                "Transport requires an active administrator-registered SharePoint site."
            );
    }

    /// <summary>
    /// Checks a template's source tables against the enabled tables before it is published.
    /// The template's own table must be enabled. Lookup source tables need not be.
    /// </summary>
    /// <param name="template">The template revision being published.</param>
    /// <returns>One notice for each lookup source table that is not enabled.</returns>
    public string[] ValidateSources(DocumentTemplate template)
    {
        if (!Tables.Contains(template.Table, StringComparer.Ordinal))
            throw new EvaluationBlockedException(
                "Table '"
                    + template.Table
                    + "' is not enabled. Enable it in the Tables panel before publishing its template."
            );
        return template
            .Sources.Select(s => s.Table)
            .Where(table => !Tables.Contains(table, StringComparer.Ordinal))
            .Distinct(StringComparer.Ordinal)
            .Select(table => RelatedTableNotice(table, template.Table))
            .ToArray();
    }

    /// <summary>Explains that a lookup source table which is not enabled triggers no replans.</summary>
    /// <param name="table">The lookup source table that is not enabled.</param>
    /// <param name="root">The template's own table.</param>
    /// <returns>The notice shown to the template author.</returns>
    public static string RelatedTableNotice(string table, string root) =>
        "Changes to "
        + table
        + " records don't update folders until the "
        + root
        + " record changes or is replanned. Enable "
        + table
        + " in the Tables panel to react to its changes.";
}
