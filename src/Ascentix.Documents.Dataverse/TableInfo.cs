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

    /// <summary>
    /// The tables a lookup column targets, or none when the table or column no longer exists or
    /// the column is not a lookup. RetrieveAttribute would fault for a deleted column.
    /// </summary>
    public static string[] LookupTargets(IOrganizationService service, string table, string column)
    {
        var query = new EntityQueryExpression
        {
            Criteria = new MetadataFilterExpression(LogicalOperator.And),
            Properties = new MetadataPropertiesExpression("LogicalName", "Attributes"),
            AttributeQuery = new AttributeQueryExpression
            {
                Criteria = new MetadataFilterExpression(LogicalOperator.And),
                Properties = new MetadataPropertiesExpression("LogicalName", "Targets"),
            },
        };
        query.Criteria.Conditions.Add(
            new MetadataConditionExpression("LogicalName", MetadataConditionOperator.Equals, table)
        );
        query.AttributeQuery.Criteria.Conditions.Add(
            new MetadataConditionExpression("LogicalName", MetadataConditionOperator.Equals, column)
        );
        var metadata = (
            (RetrieveMetadataChangesResponse)
                service.Execute(new RetrieveMetadataChangesRequest { Query = query })
        ).EntityMetadata.FirstOrDefault();
        return (
                metadata?.Attributes?.FirstOrDefault(a => a.LogicalName == column)
                as LookupAttributeMetadata
            )?.Targets
            ?? Array.Empty<string>();
    }

    /// <summary>
    /// Whether the user holds the table's Read privilege at any depth, so that a query of it
    /// returns the rows they may see instead of faulting. Find must have returned the metadata.
    /// </summary>
    public static bool CanRead(IOrganizationService service, EntityMetadata metadata, Guid user)
    {
        var read = metadata
            .Privileges?.FirstOrDefault(p => p.PrivilegeType == PrivilegeType.Read)
            ?.PrivilegeId;
        return read.HasValue
            && (
                (RetrieveUserPrivilegeByPrivilegeIdResponse)
                    service.Execute(
                        new RetrieveUserPrivilegeByPrivilegeIdRequest
                        {
                            UserId = user,
                            PrivilegeId = read.Value,
                        }
                    )
            )
                .RolePrivileges
                ?.Length > 0;
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
