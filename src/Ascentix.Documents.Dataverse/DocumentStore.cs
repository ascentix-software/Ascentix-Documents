using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public sealed class StoredRow<T>
    where T : StoredDocument
{
    public Entity Row { get; }
    public T Value { get; }

    public StoredRow(Entity row, T value)
    {
        Row = row;
        Value = value;
    }
}

public sealed class DocumentStore
{
    private readonly IOrganizationService service;
    private static readonly string[] Allowed =
    {
        "asx_outbox",
        "asx_operation",
        "asx_claim",
        "asx_attempt",
        "asx_managedgroup",
        "asx_managedgrant",
        "asx_membership",
        "asx_policy",
        "asx_policyentry",
        "asx_teamregistration",
    };

    public DocumentStore(IOrganizationService service)
    {
        this.service = service;
    }

    public static string Hash(string key)
    {
        if (string.IsNullOrWhiteSpace(key) || key.Length > 2000)
            throw new EvaluationBlockedException("Invalid stable operation key.");
        using (var hash = SHA256.Create())
            return BitConverter
                .ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(key)))
                .Replace("-", "")
                .ToLowerInvariant();
    }

    public static Guid StableId(string key)
    {
        var hex = Hash(key);
        var bytes = Enumerable
            .Range(0, 16)
            .Select(i => Convert.ToByte(hex.Substring(i * 2, 2), 16))
            .ToArray();
        return new Guid(bytes);
    }

    private static void Table(string table)
    {
        if (!Allowed.Contains(table))
            throw new EvaluationBlockedException("Unsupported state table.");
    }

    /// <summary>
    /// Looks up a hashed key and verifies the full stored key and payload schema.
    /// </summary>
    /// <typeparam name="T">The persisted document type to deserialize.</typeparam>
    /// <param name="table">The logical name of an allowed durable-state table.</param>
    /// <param name="key">The complete, unhashed document key.</param>
    /// <returns>The document and versioned row, or null when no matching row exists.</returns>
    public StoredRow<T>? Find<T>(string table, string key)
        where T : StoredDocument
    {
        Table(table);
        var query = new QueryExpression(table)
        {
            ColumnSet = new ColumnSet("asx_key", "asx_status", "asx_payload"),
            TopCount = 2,
        };
        query.Criteria.AddCondition("asx_key", ConditionOperator.Equal, Hash(key));
        var rows = service.RetrieveMultiple(query);
        if (rows.MoreRecords || rows.Entities.Count > 1)
            throw new EvaluationBlockedException("Ambiguous or incomplete durable state.");
        if (rows.Entities.Count == 0)
            return null;
        var row = rows.Entities[0];
        var value = JsonWire.Read<T>(row.GetAttributeValue<string>("asx_payload"));
        if (
            value.SchemaVersion != 1
            || value.Key != key
            || row.GetAttributeValue<string>("asx_key") != Hash(key)
            || value.Status != row.GetAttributeValue<string>("asx_status")
            || string.IsNullOrEmpty(row.RowVersion)
        )
            throw new EvaluationBlockedException("Durable state identity/version mismatch.");
        return new StoredRow<T>(row, value);
    }

    public StoredRow<T> Require<T>(string table, string key)
        where T : StoredDocument =>
        Find<T>(table, key) ?? throw new EvaluationBlockedException("Durable state was not found.");

    public void Create<T>(string table, T value)
        where T : StoredDocument
    {
        Table(table);
        WorkRetention.Stamp(table, value);
        var entity = new Entity(table, StableId(table + ":" + value.Key))
        {
            ["asx_name"] = Hash(value.Key),
            ["asx_key"] = Hash(value.Key),
            ["asx_status"] = value.Status,
            ["asx_payload"] = JsonWire.Write(value),
        };
        if (WorkRetention.Tables.Contains(table))
            entity["asx_retireafter"] = value.RetireAfterUtc;
        if (value is OperationDocument operation)
            entity["asx_priority"] =
                operation is SecurityOperation || operation is CatalogProbe
                    ? 0
                    : operation.Priority;
        if (value is CatalogProbe probe)
        {
            entity["asx_workkey"] = probe.Key;
            entity["asx_workkind"] =
                probe.DiscoverLibraries ? "LibraryDiscovery"
                : probe.ListId == Guid.Empty ? "SiteValidation"
                : "LibraryValidation";
            entity["asx_displayname"] = probe.DisplayName;
            entity["asx_siteurl"] = probe.WebUrl;
        }
        if (value is LibrarySetup setup)
        {
            entity["asx_workkey"] = setup.Key;
            entity["asx_workkind"] = "LibrarySetup";
            entity["asx_displayname"] = setup.Name.Substring(0, Math.Min(200, setup.Name.Length));
            entity["asx_siteurl"] = setup.WebUrl;
        }
        if (value is PolicyTeamReference reference)
            entity["asx_teamkey"] = reference.TeamId.ToString("N");
        if (value is PolicyDocument policy)
        {
            policy.NextReviewUtc = policy.NextReviewUtc ?? DateTime.UtcNow.AddDays(1);
            entity["asx_reviewafter"] = policy.NextReviewUtc;
            entity["asx_payload"] = JsonWire.Write(value);
        }
        if (value is RecordPlanDocument selection)
            Index(selection, entity);
        if (value is OutboxDocument work)
            Index(work, entity);
        service.Create(entity);
    }

    /// <summary>
    /// Writes payload and indexed fields using the previously read Dataverse row version.
    /// </summary>
    /// <typeparam name="T">The persisted document type.</typeparam>
    /// <param name="row">The updated document and its previously read row identity and version.</param>
    public void Save<T>(StoredRow<T> row)
        where T : StoredDocument
    {
        Table(row.Row.LogicalName);
        WorkRetention.Stamp(row.Row.LogicalName, row.Value);
        if (
            string.IsNullOrEmpty(row.Row.RowVersion)
            || row.Row.GetAttributeValue<string>("asx_key") != Hash(row.Value.Key)
        )
            throw new EvaluationBlockedException(
                "CAS requires an unchanged full stable key and row version."
            );
        var target = new Entity(row.Row.LogicalName, row.Row.Id)
        {
            RowVersion = row.Row.RowVersion,
            ["asx_status"] = row.Value.Status,
            ["asx_payload"] = JsonWire.Write(row.Value),
        };
        if (WorkRetention.Tables.Contains(row.Row.LogicalName))
            target["asx_retireafter"] = row.Value.RetireAfterUtc;
        if (row.Value is OperationDocument operation)
        {
            target["asx_nextattempt"] = operation.NextAttemptUtc;
            target["asx_retrycount"] = operation.RetryCount;
        }
        if (row.Value is PolicyTeamReference reference)
            target["asx_teamkey"] = reference.TeamId.ToString("N");
        if (row.Value is PolicyDocument policy)
        {
            policy.NextReviewUtc = policy.NextReviewUtc ?? DateTime.UtcNow.AddDays(1);
            target["asx_reviewafter"] = policy.NextReviewUtc;
            target["asx_payload"] = JsonWire.Write(row.Value);
        }
        if (row.Value is RecordPlanDocument selection)
            Index(selection, target);
        if (row.Value is OutboxDocument work)
            Index(work, target);
        if (row.Value is DispatcherDocument writer)
            WorkCoordination.SaveWriter(service, writer);
        service.Execute(
            new UpdateRequest
            {
                Target = target,
                ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
            }
        );
    }

    /// <summary>
    /// Indexes an outbox row by its business record so a record's Blocked or Pending rows can be
    /// found without reading every payload. Team and targeted rows carry no record.
    /// </summary>
    private static void Index(OutboxDocument work, Entity target)
    {
        target["asx_recordid"] = work.RecordId == Guid.Empty ? null : work.RecordId.ToString("D");
        target["asx_table"] = string.IsNullOrEmpty(work.Table) ? null : work.Table;
    }

    private static void Index(RecordPlanDocument selection, Entity target)
    {
        if (selection.Sources.Length > 6)
            throw new EvaluationBlockedException("Record source index exceeds the template limit.");
        var related = selection
            .Sources.Where(s => s.Table != selection.Table || s.Id != selection.RecordId)
            .ToArray();
        for (int i = 0; i < 6; i++)
            target["asx_source" + i] =
                i < related.Length
                    ? TargetedReplan.SourceKey(related[i].Table, related[i].Id)
                    : null;
    }

    public WorkerResult FailUnclaimed<T>(string key)
        where T : OperationDocument
    {
        var op = Require<T>("asx_operation", key);
        var claim = Find<DispatcherDocument>("asx_claim", WorkCoordination.Operation(service, key));
        if (claim?.Value.OperationKey == key || op.Value.ExternalSubmitted)
            return new WorkerResult { Status = "Quarantined", Key = key };
        if (op.Value.Status == "Pending" || op.Value.Status == "RetryWait")
        {
            op.Value.Status = "Blocked";
            op.Value.ErrorCode =
                "Claim failed before any external call; inspect scope, policy and configuration before retry.";
            Save(op);
        }
        return new WorkerResult { Status = op.Value.Status, Key = key };
    }

    /// <summary>
    /// Returns a bounded dispatch page of pending work and due retries.
    /// </summary>
    /// <param name="table">The logical name of the durable-state table to query.</param>
    /// <param name="limit">The maximum number of work keys to return, from 1 through 100.</param>
    /// <param name="now">The UTC retry cutoff; null uses the current UTC time.</param>
    /// <returns>The selected work keys in dispatch order.</returns>
    public string[] Pending(string table, int limit = 20, DateTime? now = null)
    {
        Table(table);
        if (limit < 1 || limit > 100)
            throw new EvaluationBlockedException("Invalid worker batch size.");
        var query = new QueryExpression(table)
        {
            ColumnSet = new ColumnSet("asx_payload"),
            TopCount = limit,
        };
        if (table == "asx_operation")
        {
            query.Criteria.FilterOperator = LogicalOperator.Or;
            query.Criteria.AddCondition("asx_status", ConditionOperator.Equal, "Pending");
            var retry = new FilterExpression(LogicalOperator.And);
            retry.AddCondition("asx_status", ConditionOperator.Equal, "RetryWait");
            retry.AddCondition(
                "asx_nextattempt",
                ConditionOperator.LessEqual,
                now ?? DateTime.UtcNow
            );
            query.Criteria.AddFilter(retry);
        }
        else
            query.Criteria.AddCondition("asx_status", ConditionOperator.Equal, "Pending");
        if (table == "asx_operation")
            query.AddOrder("asx_priority", OrderType.Ascending);
        query.AddOrder("createdon", OrderType.Ascending);
        // This is a bounded dispatch page, never a complete membership/security snapshot.
        return service
            .RetrieveMultiple(query)
            .Entities.Select(row =>
                JsonWire.Read<StoredDocument>(row.GetAttributeValue<string>("asx_payload")).Key
            )
            .ToArray();
    }
}
