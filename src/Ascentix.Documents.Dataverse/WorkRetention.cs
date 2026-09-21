using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public static class WorkRetention
{
    public static readonly string[] Tables =
    {
        "asx_operation",
        "asx_outbox",
        "asx_attempt",
        "asx_membership",
    };

    /// <summary>
    /// Sets the expiry of completed work; membership expiry is set by policy completion.
    /// </summary>
    /// <param name="table">The logical name of the table storing the document.</param>
    /// <param name="value">The document whose retirement timestamp is updated.</param>
    public static void Stamp(string table, StoredDocument value)
    {
        bool terminal =
            table == "asx_attempt"
            || table == "asx_outbox" && new[] { "Planned", "Cancelled" }.Contains(value.Status)
            || table == "asx_operation" && IsTerminalOperation(value.Status);
        if (
            value is OperationDocument op
            && (op.HistoryCompacted || op.ExternalSubmitted && !op.ExternalResponseKnown)
        )
            terminal = false;
        if (table == "asx_membership")
            return;
        value.RetireAfterUtc = terminal
            ? (value.RetireAfterUtc ?? DateTime.UtcNow.AddDays(30))
            : null;
    }

    /// <summary>
    /// Processes one expiry page per table and compacts results still referenced by a record selection.
    /// </summary>
    /// <param name="service">The Dataverse service used to read, compact, and delete work history.</param>
    /// <param name="now">The UTC cutoff for expired work.</param>
    /// <returns>A history-pruning result containing the removed and compacted row counts.</returns>
    public static WorkerResult Purge(IOrganizationService service, DateTime now)
    {
        int removed = 0,
            compacted = 0;
        var store = new DocumentStore(service);
        foreach (var table in Tables)
        {
            var cursor = store.Find<ScanDocument>("asx_claim", "retention:" + table);
            if (cursor == null)
            {
                store.Create(
                    "asx_claim",
                    new ScanDocument { Key = "retention:" + table, Status = "Retention" }
                );
                cursor = store.Require<ScanDocument>("asx_claim", "retention:" + table);
            }
            var query = new QueryExpression(table)
            {
                ColumnSet = new ColumnSet(
                    "asx_key",
                    "asx_status",
                    "asx_payload",
                    "asx_retireafter"
                ),
                PageInfo = new PagingInfo
                {
                    Count = 10,
                    PageNumber = cursor.Value.Page,
                    PagingCookie = cursor.Value.Cookie,
                },
            };
            query.Criteria.AddCondition("asx_retireafter", ConditionOperator.LessEqual, now);
            query.AddOrder("asx_retireafter", OrderType.Ascending);
            query.AddOrder(table + "id", OrderType.Ascending);
            var page = service.RetrieveMultiple(query);
            if (
                page.Entities.Count > 10
                || page.MoreRecords && string.IsNullOrEmpty(page.PagingCookie)
            )
                throw new EvaluationBlockedException("Retention page is incomplete.");
            foreach (var row in page.Entities)
            {
                try
                {
                    ValidateDelete(service, row, now);
                }
                catch (EvaluationBlockedException)
                {
                    continue;
                }
                if (table == "asx_operation")
                {
                    var op = store.Require<OperationDocument>(
                        table,
                        JsonWire.Read<StoredDocument>(TemplateStore.Text(row, "asx_payload")).Key
                    );
                    if (op.Value.Folders.Length > 0)
                    {
                        var root = op.Value.Folders[0];
                        var current = store.Find<RecordPlanDocument>(
                            "asx_outbox",
                            "recordplan:"
                                + root.TemplateId.ToString("N")
                                + ":"
                                + root.RecordId.ToString("N")
                        );
                        if (current?.Value.Operations.Contains(op.Value.Key) == true)
                        {
                            op.Value.ResultLibraryId = root.LibraryId;
                            op.Value.ResultLocationId = root.LocationId;
                            op.Value.ResultPhysicalId = root.PhysicalId;
                            op.Value.Folders = Array.Empty<FolderStep>();
                            op.Value.Cursor = 0;
                            op.Value.HistoryCompacted = true;
                            store.Save(op);
                            compacted++;
                            continue;
                        }
                    }
                }
                service.Execute(
                    new DeleteRequest
                    {
                        Target = new EntityReference(table, row.Id) { RowVersion = row.RowVersion },
                        ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                    }
                );
                removed++;
            }
            cursor.Value.Page = page.MoreRecords ? cursor.Value.Page + 1 : 1;
            cursor.Value.Cookie = page.MoreRecords ? page.PagingCookie : null;
            store.Save(cursor);
        }
        return new WorkerResult
        {
            Status = "HistoryPruned",
            Notices = new[]
            {
                removed
                    + " expired work rows removed; "
                    + compacted
                    + " current destination results compacted. SharePoint content and access mappings are unchanged.",
            },
        };
    }

    private static bool IsTerminalOperation(string status) =>
        status is "Applied" or "Approved" or "Discovered" or "Cancelled" or "Superseded";

    /// <summary>
    /// Rejects deletion of unexpired, unfinished, or actively claimed work and access evidence.
    /// </summary>
    /// <param name="service">The Dataverse service used to check owning operations and active claims.</param>
    /// <param name="row">The versioned work row with its payload, status, and retirement timestamp.</param>
    /// <param name="now">The UTC cutoff used to determine whether the row has expired.</param>
    public static void ValidateDelete(IOrganizationService service, Entity row, DateTime now)
    {
        if (!Tables.Contains(row.LogicalName) || string.IsNullOrEmpty(row.RowVersion))
            throw new EvaluationBlockedException(
                "Retention can remove only versioned work history."
            );
        var payload = JsonWire.Read<StoredDocument>(TemplateStore.Text(row, "asx_payload"));
        var due = row.GetAttributeValue<DateTime?>("asx_retireafter");
        if (
            !due.HasValue
            || due > now
            || !payload.RetireAfterUtc.HasValue
            || due.HasValue && Math.Abs((payload.RetireAfterUtc.Value - due.Value).TotalSeconds) > 1
            || payload.Status != TemplateStore.Text(row, "asx_status")
        )
            throw new EvaluationBlockedException("Work history is not due for retention.");
        if (
            row.LogicalName == "asx_outbox"
            && !new[] { "Planned", "Cancelled" }.Contains(payload.Status)
        )
            throw new EvaluationBlockedException(
                "Active record selections and unfinished requests are retained."
            );
        if (row.LogicalName == "asx_operation")
        {
            var op = JsonWire.Read<OperationDocument>(TemplateStore.Text(row, "asx_payload"));
            if (
                !IsTerminalOperation(op.Status)
                || op.ExternalSubmitted && !op.ExternalResponseKnown
            )
                throw new EvaluationBlockedException("Unfinished or uncertain work is retained.");
            var claim = new DocumentStore(service).Find<DispatcherDocument>(
                "asx_claim",
                WorkCoordination.Operation(service, op.Key)
            );
            if (claim?.Value.OperationKey == op.Key && claim.Value.RunId != null)
                throw new EvaluationBlockedException("Active writer history cannot be removed.");
        }
        if (row.LogicalName == "asx_membership")
        {
            var member = JsonWire.Read<MembershipDocument>(TemplateStore.Text(row, "asx_payload"));
            var owner = new DocumentStore(service).Find<OperationDocument>(
                "asx_operation",
                "policywork:" + member.Generation.ToString("N")
            );
            if (member.Status != "Applied" || owner != null && owner.Value.Status != "Applied")
                throw new EvaluationBlockedException("Active membership evidence is retained.");
        }
        if (row.LogicalName == "asx_attempt")
        {
            var attempt = JsonWire.Read<AttemptDocument>(TemplateStore.Text(row, "asx_payload"));
            var owner = new DocumentStore(service).Find<OperationDocument>(
                "asx_operation",
                attempt.OperationKey
            );
            if (owner != null && !IsTerminalOperation(owner.Value.Status))
                throw new EvaluationBlockedException("Unresolved request evidence is retained.");
        }
    }
}
