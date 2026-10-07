using System;
using System.Linq;
using System.Security;
using System.ServiceModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

/// <summary>Row counts through FetchXML aggregates, with Dataverse's aggregate limit.</summary>
public static class AggregateCount
{
    /// <summary>
    /// Dataverse refuses an aggregate over more rows than this (AggregateQueryRecordLimit, error
    /// 0x8004E023). A count that hits it is reported as this number and marked capped; the page
    /// shows "50,000+". This is a platform limit, not a Documents one.
    /// </summary>
    public const int Limit = 50000;

    /// <summary>Runs a FetchXML aggregate whose single value has the alias "count".</summary>
    /// <param name="service">The service to count as.</param>
    /// <param name="fetchXml">An aggregate fetch with one count attribute aliased "count".</param>
    /// <param name="capped">True when Dataverse refused the count for its size.</param>
    /// <returns>The count, or Limit when capped.</returns>
    public static int Count(IOrganizationService service, string fetchXml, out bool capped)
    {
        capped = false;
        try
        {
            var row = service
                .RetrieveMultiple(new FetchExpression(fetchXml))
                .Entities.FirstOrDefault();
            return row?.GetAttributeValue<AliasedValue>("count")?.Value is int n ? n : 0;
        }
        catch (FaultException<OrganizationServiceFault> fault) when (OverLimit(fault.Detail))
        {
            capped = true;
            return Limit;
        }
    }

    /// <summary>
    /// Every row of a table. A table the platform's daily row-count snapshot already puts at the
    /// limit is capped without running the aggregate, so the count does not depend on catching
    /// the refusal inside a plug-in transaction; the aggregate gives the exact count below it.
    /// </summary>
    public static int Records(IOrganizationService service, string table, out bool capped)
    {
        if (Snapshot(service, table) >= Limit)
        {
            capped = true;
            return Limit;
        }
        return Count(
            service,
            "<fetch aggregate=\"true\"><entity name=\""
                + SecurityElement.Escape(table)
                + "\"><attribute name=\""
                + SecurityElement.Escape(PrimaryId(service, table))
                + "\" alias=\"count\" aggregate=\"count\" /></entity></fetch>",
            out capped
        );
    }

    public static string PrimaryId(IOrganizationService service, string table) =>
        Metadata(service, table).PrimaryIdAttribute;

    /// <summary>The table's display name in the caller's language, else its logical name.</summary>
    public static string Label(IOrganizationService service, string table)
    {
        try
        {
            var name = Metadata(service, table).DisplayName;
            return name?.UserLocalizedLabel?.Label
                ?? name?.LocalizedLabels.FirstOrDefault()?.Label
                ?? table;
        }
        catch (FaultException<OrganizationServiceFault>)
        {
            return table;
        }
    }

    /// <summary>The platform's row count of the table from its last daily snapshot.</summary>
    private static long Snapshot(IOrganizationService service, string table) =>
        (
            (RetrieveTotalRecordCountResponse)
                service.Execute(
                    new RetrieveTotalRecordCountRequest { EntityNames = new[] { table } }
                )
        ).EntityRecordCountCollection.TryGetValue(table, out long count)
            ? count
            : 0;

    private static EntityMetadata Metadata(IOrganizationService service, string table) =>
        (
            (RetrieveEntityResponse)
                service.Execute(
                    new RetrieveEntityRequest
                    {
                        LogicalName = table,
                        EntityFilters = EntityFilters.Entity,
                    }
                )
        ).EntityMetadata;

    private static bool OverLimit(OrganizationServiceFault fault) =>
        fault.ErrorCode == unchecked((int)0x8004E023)
        || (fault.Message ?? "").IndexOf(
            "AggregateQueryRecordLimit",
            StringComparison.OrdinalIgnoreCase
        ) >= 0;
}
