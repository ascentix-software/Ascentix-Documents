using System;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

/// <summary>
/// What a template re-run needs to know about a table, through calls that do not fault on a
/// table's size or absence. Inside a plug-in, a service call that faults ends the transaction even
/// when the fault is caught, so none of these relies on catching one.
/// </summary>
public static class TableInfo
{
    /// <summary>
    /// The table's metadata, or null when the organization has no such table. Unlike
    /// RetrieveEntity, RetrieveMetadataChanges answers an empty list for a missing table.
    /// </summary>
    public static EntityMetadata? Find(IOrganizationService service, string table)
    {
        var query = new EntityQueryExpression
        {
            Criteria = new MetadataFilterExpression(LogicalOperator.And),
            Properties = new MetadataPropertiesExpression(
                "LogicalName",
                "PrimaryIdAttribute",
                "PrimaryNameAttribute",
                "DisplayName",
                // The table's privileges: EventRegistrations reads its Read privilege.
                "Privileges"
            ),
        };
        query.Criteria.Conditions.Add(
            new MetadataConditionExpression("LogicalName", MetadataConditionOperator.Equals, table)
        );
        return (
            (RetrieveMetadataChangesResponse)
                service.Execute(new RetrieveMetadataChangesRequest { Query = query })
        ).EntityMetadata.FirstOrDefault();
    }

    /// <summary>The table's display name in the caller's language, else its logical name.</summary>
    public static string Label(IOrganizationService service, string table) =>
        Label(Find(service, table), table);

    /// <summary>The display name of metadata Find returned, else the logical name.</summary>
    public static string Label(EntityMetadata? metadata, string table)
    {
        var name = metadata?.DisplayName;
        return name?.UserLocalizedLabel?.Label
            ?? name?.LocalizedLabels.FirstOrDefault()?.Label
            ?? table;
    }

    /// <summary>
    /// About how many rows the table has: the platform's row-count snapshot, taken within the
    /// last 24 hours (RetrieveTotalRecordCount). It has no size limit, unlike a FetchXML aggregate
    /// (AggregateQueryRecordLimit, 50,000 rows), whose refusal would end a plug-in's transaction.
    /// The table must exist (see Find).
    /// </summary>
    public static int Records(IOrganizationService service, string table) =>
        (
            (RetrieveTotalRecordCountResponse)
                service.Execute(
                    new RetrieveTotalRecordCountRequest { EntityNames = new[] { table } }
                )
        ).EntityRecordCountCollection.TryGetValue(table, out long count)
            ? (int)Math.Min(int.MaxValue, count)
            : 0;
}
