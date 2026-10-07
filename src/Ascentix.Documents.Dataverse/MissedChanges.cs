using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

/// <summary>Changes Documents did not capture: re-run their record, or dismiss the failed job (spec 6.7).</summary>
public static class MissedChanges
{
    private static readonly string[] Handlers =
    {
        EventRegistrations.RecordHandler,
        EventRegistrations.RetirementHandler,
        EventRegistrations.MembershipHandler,
    };

    /// <summary>
    /// Queues the record for every template of its table that has a published version and is
    /// active. A table that is not enabled is refused as a template re-run refuses it.
    /// </summary>
    /// <param name="service">The caller's service; the queue writes go through the guarded transport.</param>
    /// <param name="request">Table, RecordId and a RequestId that makes the call idempotent.</param>
    /// <param name="allowed">The tables enabled in Documents.</param>
    /// <param name="clock">The UTC clock.</param>
    /// <returns>Queued with each template's request key, or Inactive when no template applies.</returns>
    public static WorkerResult Rerun(
        IOrganizationService service,
        WorkerRequest request,
        string[] allowed,
        Func<DateTime> clock
    )
    {
        if (
            string.IsNullOrEmpty(request.Table)
            || request.RecordId == Guid.Empty
            || request.RequestId == Guid.Empty
        )
            throw new EvaluationBlockedException("Choose a record to re-run.");
        if (!allowed.Contains(request.Table, StringComparer.Ordinal))
            throw new EvaluationBlockedException(WorkerCoordinator.TableNotEnabled(request.Table));
        var query = new QueryExpression("asx_template")
        {
            ColumnSet = new ColumnSet(
                "asx_name",
                "asx_table",
                "asx_publishedrevisionid",
                "asx_disabled",
                "asx_startsutc",
                "asx_endsutc"
            ),
        };
        query.Criteria.AddCondition("asx_table", ConditionOperator.Equal, request.Table);
        var coordinator = new WorkerCoordinator(service, clock, allowed);
        var keys = new List<string>();
        var notices = new List<string>();
        foreach (var template in service.RetrieveMultiple(query).Entities)
        {
            if (template.GetAttributeValue<EntityReference>("asx_publishedrevisionid") == null)
                continue;
            if (!TemplateLifecycle.Active(template, clock()))
            {
                notices.Add(
                    "Template '"
                        + template.GetAttributeValue<string>("asx_name")
                        + "' is off; skipped."
                );
                continue;
            }
            // Each template's request ID follows the caller's, so a repeated call queues nothing twice.
            var queued = coordinator.Queue(
                new WorkerRequest
                {
                    TemplateId = template.Id,
                    RecordId = request.RecordId,
                    RequestId = DocumentStore.StableId(
                        request.RequestId.ToString("N") + ":" + template.Id.ToString("N")
                    ),
                }
            );
            keys.Add(queued.Key);
        }
        return new WorkerResult
        {
            Status = keys.Count > 0 ? "Queued" : "Inactive",
            Keys = keys.ToArray(),
            Notices = notices.ToArray(),
        };
    }

    /// <summary>Deletes one failed system job of a Documents capture handler, as the caller.</summary>
    /// <param name="service">The caller's service.</param>
    /// <param name="jobId">The asyncoperation to delete.</param>
    /// <returns>Dismissed.</returns>
    public static WorkerResult Dismiss(IOrganizationService service, Guid jobId)
    {
        if (jobId == Guid.Empty)
            throw new EvaluationBlockedException("Choose a missed change to dismiss.");
        var job = service.Retrieve(
            "asyncoperation",
            jobId,
            new ColumnSet("statuscode", "owningextensionid")
        );
        var step = job.GetAttributeValue<EntityReference>("owningextensionid");
        string? handler = null;
        if (step != null)
        {
            var type = service
                .Retrieve("sdkmessageprocessingstep", step.Id, new ColumnSet("eventhandler"))
                .GetAttributeValue<EntityReference>("eventhandler");
            if (type != null)
                handler = service
                    .Retrieve("plugintype", type.Id, new ColumnSet("typename"))
                    .GetAttributeValue<string>("typename");
        }
        if (
            job.GetAttributeValue<OptionSetValue>("statuscode")?.Value != 31
            || handler == null
            || !Handlers.Contains(handler, StringComparer.Ordinal)
        )
            throw new EvaluationBlockedException(
                "Only a failed Documents capture job can be dismissed."
            );
        try
        {
            // The guarded transport passes deletes of rows that are not Documents' own straight
            // through (ApiWriteService.Delete); the caller's own privilege decides (decision D12).
            service.Delete("asyncoperation", jobId);
        }
        catch (FaultException<OrganizationServiceFault>)
        {
            // Rethrown at once, so nothing runs after the refused delete in this transaction.
            throw new EvaluationBlockedException(
                "You don't have permission to remove system jobs."
            );
        }
        return new WorkerResult { Status = "Dismissed", Key = jobId.ToString("D") };
    }
}
