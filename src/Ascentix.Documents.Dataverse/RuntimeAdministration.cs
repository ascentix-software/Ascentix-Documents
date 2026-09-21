using System;
using System.Linq;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class RuntimeRequest
{
    [DataMember]
    public string Command { get; set; } = "Get";

    [DataMember]
    public string RowVersion { get; set; } = "";

    [DataMember]
    public Guid WorkerId { get; set; }

    [DataMember]
    public bool Enabled { get; set; }

    [DataMember]
    public bool ProcessRecordUpdates { get; set; }

    [DataMember]
    public string[] SharePointHosts { get; set; } = Array.Empty<string>();

    [DataMember]
    public string[] Tables { get; set; } = Array.Empty<string>();
}

public static class RuntimeAdministration
{
    public static RuntimeRequest Execute(
        IOrganizationService service,
        RuntimeRequest request,
        bool transaction
    )
    {
        if (!transaction)
            throw new EvaluationBlockedException("Runtime configuration requires a transaction.");
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
                "Exactly one installed runtime profile is required."
            );
        var old = rows.Entities[0];
        if (request.Command == "Get")
            return new RuntimeRequest
            {
                Command = "Get",
                RowVersion = old.RowVersion,
                WorkerId = Guid.Parse(TemplateStore.Text(old, "asx_workeruserid")),
                Enabled = old.GetAttributeValue<bool>("asx_enabled"),
                ProcessRecordUpdates = old.GetAttributeValue<bool>("asx_processrecordupdates"),
                SharePointHosts = JsonWire.Read<string[]>(
                    old.GetAttributeValue<string>("asx_sharepointhosts") ?? "[]"
                ),
                Tables = JsonWire.Read<string[]>(TemplateStore.Text(old, "asx_allowedtables")),
            };
        if (
            request.Command != "Save"
            || string.IsNullOrEmpty(request.RowVersion)
            || old.RowVersion != request.RowVersion
        )
            throw new EvaluationBlockedException("Refresh the runtime profile before saving.");
        if (
            request.WorkerId == Guid.Empty
            || request.Tables == null
            || request.Tables.Length < 1
            || request.Tables.Length > 50
            || request.Tables.Distinct().Count() != request.Tables.Length
            || request.Tables.Any(t => !Regex.IsMatch(t, "\\A[a-z][a-z0-9_]{0,99}\\z"))
        )
            throw new EvaluationBlockedException(
                "Explicit worker and source table allowlist required."
            );
        RuntimeProfile.ValidateHosts(request.SharePointHosts);
        var worker = service.Retrieve("systemuser", request.WorkerId, new ColumnSet("isdisabled"));
        if (worker.GetAttributeValue<bool>("isdisabled"))
            throw new EvaluationBlockedException("Worker identity is disabled.");
        WorkCoordination.RequireIdle(service);
        foreach (var table in request.Tables)
            service.Execute(
                new RetrieveEntityRequest
                {
                    LogicalName = table,
                    EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity,
                }
            );
        RecordUpdateRegistrations.Apply(
            service,
            request.WorkerId,
            request.Tables,
            request.ProcessRecordUpdates
        );
        service.Execute(
            new UpdateRequest
            {
                Target = new Entity("asx_runtime", old.Id)
                {
                    RowVersion = old.RowVersion,
                    ["asx_workeruserid"] = request.WorkerId.ToString("D"),
                    ["asx_enabled"] = request.Enabled,
                    ["asx_processrecordupdates"] = request.ProcessRecordUpdates,
                    ["asx_allowedtables"] = JsonWire.Write(request.Tables),
                    ["asx_sharepointhosts"] = JsonWire.Write(request.SharePointHosts),
                },
                ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
            }
        );
        return Execute(service, new RuntimeRequest(), true);
    }
}
