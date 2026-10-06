using System;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class RecordInspection
{
    [DataMember]
    public FolderStep[] Folders { get; set; } = Array.Empty<FolderStep>();

    [DataMember]
    public string[] OperationStates { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Reads the current record selection and summarizes its destination operations.
    /// </summary>
    /// <param name="service">The caller-scoped Dataverse service used to read configuration and work.</param>
    /// <param name="request">The template and business-record identities to inspect.</param>
    /// <param name="allowed">The logical names of business tables permitted by the runtime scope.</param>
    /// <returns>The record's processing status and available destination-operation details.</returns>
    public static WorkerResult Read(
        IOrganizationService service,
        WorkerRequest request,
        string[] allowed
    )
    {
        var store = new DocumentStore(service);
        var template = TemplateLifecycle.Find(service, request.TemplateId);
        if (template == null)
            return new WorkerResult
            {
                Status = "TemplateDeleted",
                Notices = new[]
                {
                    "Template deleted. Existing SharePoint content and historical work receipts are retained.",
                },
            };
        string table = TemplateStore.Text(template, "asx_table");
        if (!allowed.Contains(table) || request.RecordId == Guid.Empty)
            throw new EvaluationBlockedException(
                "Record inspection is outside the approved source scope."
            );
        if (
            store.Find<OutboxDocument>(
                "asx_outbox",
                WorkerCoordinator.RetirementKey(table, request.RecordId)
            ) != null
        )
            return new WorkerResult
            {
                Status = "DecommissionReview",
                Notices = new[]
                {
                    "Record deleted; physical documents and work receipts are retained.",
                },
            };
        service.Retrieve(table, request.RecordId, new ColumnSet(false));
        var selection = store.Find<RecordPlanDocument>(
            "asx_outbox",
            "recordplan:" + template.Id.ToString("N") + ":" + request.RecordId.ToString("N")
        );
        if (selection == null)
            return Blocked(service, template.Id, request.RecordId)
                ?? new WorkerResult { Status = "NotPlanned" };
        var operations = selection
            .Value.Operations.Select(key =>
                store.Require<OperationDocument>("asx_operation", key).Value
            )
            .ToArray();
        var states = operations.Select(job => job.Status).ToArray();
        var folders = operations.SelectMany(job => job.Folders).ToArray();
        // A record with folders waiting for a value or a usable name reports that it waits;
        // its folder jobs are still listed below.
        var status =
            selection.Value.Status == WorkerCoordinator.WaitingStatus
                ? WorkerCoordinator.WaitingStatus
                : Summarize(states, folders, selection.Value.Status == "SelectionNeedsReview");
        return new WorkerResult
        {
            Status = status,
            Record = new RecordInspection { Folders = folders, OperationStates = states },
            Keys = selection.Value.Operations,
            Notices = new[]
            {
                "Current evaluation only. Existing folders remain intact. Paths shared by records contain shared documents; no permanent child-folder inventory is maintained.",
            }
                .Concat(selection.Value.Notices)
                .ToArray(),
        };
    }

    /// <summary>
    /// Reports a never-planned record as Blocked when its newest Blocked or Pending outbox row for
    /// this template is Blocked, so a newer Pending request wins, and as WaitingToRetry when that
    /// row is Pending with a next attempt after a temporary failure. Rows are found through the
    /// indexed asx_recordid column; rows written before that column was populated are found once
    /// they are retried.
    /// </summary>
    private static WorkerResult? Blocked(IOrganizationService service, Guid template, Guid record)
    {
        var query = new QueryExpression("asx_outbox")
        {
            ColumnSet = new ColumnSet("asx_payload", "createdon"),
            PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 },
        };
        query.Criteria.AddCondition("asx_recordid", ConditionOperator.Equal, record.ToString("D"));
        query.Criteria.AddCondition("asx_status", ConditionOperator.In, "Blocked", "Pending");
        query.AddOrder("createdon", OrderType.Descending);
        while (true)
        {
            var page = service.RetrieveMultiple(query);
            foreach (var row in page.Entities)
            {
                var job = JsonWire.Read<OutboxDocument>(
                    row.GetAttributeValue<string>("asx_payload")
                );
                if (
                    job.TemplateId != template
                    || job.RecordId != record
                    || job.RelatedRecordId != Guid.Empty
                    || job.SecurityTeamId != Guid.Empty
                )
                    continue;
                // A Pending row with a next attempt is waiting after a temporary failure.
                if (job.Status == "Pending" && job.NextAttemptUtc != null)
                    return new WorkerResult
                    {
                        Status = "WaitingToRetry",
                        Key = job.Key,
                        Notices = job.Notices,
                    };
                return job.Status != "Blocked"
                    ? null
                    : new WorkerResult
                    {
                        Status = "Blocked",
                        Key = job.Key,
                        Notices = job.Notices,
                    };
            }
            if (!page.MoreRecords)
                return null;
            query.PageInfo.PageNumber++;
            query.PageInfo.PagingCookie = page.PagingCookie;
        }
    }

    private static string Summarize(
        string[] states,
        FolderStep[] folders,
        bool selectionNeedsReview
    )
    {
        bool hasAppliedFolders = folders.Any(folder => folder.Status == "Applied");
        if (selectionNeedsReview)
            return hasAppliedFolders ? "Partial" : "NeedsReview";
        if (states.Length == 0)
            return "NoCurrentOperations";
        if (states.All(status => status == "Applied"))
            return "Applied";
        if (hasAppliedFolders)
            return "Partial";
        if (
            states.Any(status =>
                status == "Blocked" || status == "Cancelled" || status == "Superseded"
            )
        )
            return "NeedsReview";
        return "Pending";
    }
}
