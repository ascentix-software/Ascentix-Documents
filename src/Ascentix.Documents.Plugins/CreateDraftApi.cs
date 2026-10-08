using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Plugins;

public sealed class CreateDraftApi : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        if (!context.IsInTransaction)
            throw new InvalidPluginExecutionException("Draft save requires a transaction.");
        var service = (
            (IOrganizationServiceFactory)provider.GetService(typeof(IOrganizationServiceFactory))
        ).CreateOrganizationService(context.InitiatingUserId);
        var draft = JsonWire.Read<DraftDto>((string)context.InputParameters["Request"]);
        // Validates the whole draft before anything is written, as Preview does (spec 6.3).
        try
        {
            DraftTemplate.Build(service, draft);
        }
        catch (EvaluationBlockedException error)
        {
            throw new InvalidPluginExecutionException(error.Message);
        }
        Guid templateId;
        if (!string.IsNullOrEmpty(draft.TemplateId))
        {
            if (!Guid.TryParse(draft.TemplateId, out templateId) || templateId == Guid.Empty)
                throw new InvalidPluginExecutionException("Select a valid template.");
            var template = service.Retrieve("asx_template", templateId, new ColumnSet("asx_table"));
            if (template.GetAttributeValue<string>("asx_table") != draft.Table)
                throw new InvalidPluginExecutionException("Template belongs to another table.");
        }
        else
        {
            if (!string.IsNullOrEmpty(draft.RevisionId))
                throw new InvalidPluginExecutionException(
                    "Select the template that owns this revision."
                );
            if (string.IsNullOrWhiteSpace(draft.Name) || draft.Name!.Trim().Length > 200)
                throw new InvalidPluginExecutionException(
                    "Enter a template name of 200 characters or fewer."
                );
            templateId = service.Create(
                new Entity("asx_template")
                {
                    ["asx_name"] = draft.Name.Trim(),
                    ["asx_table"] = draft.Table,
                }
            );
        }
        var revisions = new QueryExpression("asx_revision")
        {
            ColumnSet = new ColumnSet("asx_version", "asx_status"),
            TopCount = 1,
        };
        revisions.Criteria.AddCondition("asx_templateid", ConditionOperator.Equal, templateId);
        revisions.AddOrder("asx_version", OrderType.Descending);
        var latest = service.RetrieveMultiple(revisions).Entities.SingleOrDefault();
        Entity? basis = null;
        if (!string.IsNullOrEmpty(draft.RevisionId))
        {
            if (
                !Guid.TryParse(draft.RevisionId, out var basisId)
                || string.IsNullOrEmpty(draft.RowVersion)
            )
                throw new InvalidPluginExecutionException("Reload the revision before saving.");
            basis = service.Retrieve(
                "asx_revision",
                basisId,
                new ColumnSet("asx_templateid", "asx_status")
            );
            if (
                basis.GetAttributeValue<EntityReference>("asx_templateid")?.Id != templateId
                || basis.RowVersion != draft.RowVersion
            )
                throw new InvalidPluginExecutionException(
                    "This revision changed. Reload it before saving to avoid overwriting another edit."
                );
        }
        Guid revisionId;
        if (latest?.GetAttributeValue<string>("asx_status") == "Draft")
        {
            if (
                basis == null
                || basis.Id != latest.Id
                || basis.GetAttributeValue<string>("asx_status") != "Draft"
            )
                throw new InvalidPluginExecutionException(
                    "A working draft already exists. Load the latest revision to continue editing it."
                );
            revisionId = basis.Id;
            service.Execute(
                new UpdateRequest
                {
                    Target = new Entity("asx_revision", revisionId)
                    {
                        RowVersion = draft.RowVersion,
                        ["asx_editstamp"] = Guid.NewGuid().ToString("N"),
                    },
                    ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                }
            );
            // Replace only this mutable draft's normalized tree in the same transaction.
            var store = new TemplateStore(service);
            foreach (
                var table in new[]
                {
                    "asx_condition",
                    "asx_documentconditiongroup",
                    "asx_folder",
                    "asx_destination",
                    "asx_source",
                }
            )
            foreach (var child in store.Children(table, revisionId, table + "id"))
                service.Delete(table, child.Id);
        }
        else
        {
            if (basis?.GetAttributeValue<string>("asx_status") == "Draft")
                throw new InvalidPluginExecutionException(
                    "This is an older draft. Load the latest revision before saving."
                );
            int version = checked((latest?.GetAttributeValue<int>("asx_version") ?? 0) + 1);
            // The revision key rejects competing attempts to start the next draft.
            revisionId = service.Create(
                new Entity("asx_revision")
                {
                    ["asx_name"] = draft.Table + " v" + version,
                    ["asx_templateid"] = new EntityReference("asx_template", templateId),
                    ["asx_version"] = version,
                    ["asx_status"] = "Draft",
                }
            );
        }
        Guid Row(string table, string key, Action<Entity> fill)
        {
            var row = new Entity(table)
            {
                ["asx_name"] = key,
                ["asx_key"] = key,
                ["asx_revisionid"] = new EntityReference("asx_revision", revisionId),
            };
            fill(row);
            return service.Create(row);
        }
        int groupIndex = 0,
            conditionIndex = 0;
        string Group(GroupDto group, string? parent)
        {
            if (
                group == null
                || group.Conditions == null
                || group.Groups == null
                || group.Conditions.Length + group.Groups.Length == 0
            )
                throw new InvalidPluginExecutionException(DraftTemplate.GroupRefusal);
            var key = "group_" + ++groupIndex;
            Row(
                "asx_documentconditiongroup",
                key,
                row =>
                {
                    row["asx_parentkey"] = parent;
                    row["asx_all"] = group.All;
                }
            );
            foreach (var condition in group.Conditions)
            {
                condition.ToModel();
                // The lookup record's name and table are read on load, never stored.
                condition.LiteralLabel = null;
                condition.LiteralTable = null;
                Row(
                    "asx_condition",
                    "condition_" + ++conditionIndex,
                    row =>
                    {
                        row["asx_groupkey"] = key;
                        row["asx_payload"] = JsonWire.Write(condition);
                    }
                );
            }
            foreach (var child in group.Groups)
                Group(child, key);
            return key;
        }
        foreach (var source in draft.Sources)
        {
            source.ToModel();
            Row("asx_source", source.Alias, row => row["asx_payload"] = JsonWire.Write(source));
        }
        foreach (var section in draft.Destinations)
        {
            Row(
                "asx_destination",
                section.Key,
                row =>
                {
                    row["asx_name"] = string.IsNullOrWhiteSpace(section.Name)
                        ? section.Key
                        : section.Name!.Trim();
                    row["asx_libraryid"] = new EntityReference(
                        "asx_library",
                        Guid.Parse(section.LibraryId)
                    );
                }
            );
            int order = 0;
            foreach (var folder in section.Folders)
            {
                var group = folder.Condition == null ? null : Group(folder.Condition, null);
                Row(
                    "asx_folder",
                    folder.Key,
                    row =>
                    {
                        row["asx_sectionkey"] = section.Key;
                        row["asx_parentkey"] = folder.Parent;
                        row["asx_expression"] = folder.Name;
                        row["asx_groupkey"] = group;
                        row["asx_order"] = order++;
                    }
                );
            }
        }
        new SnapshotReader(service).ValidateMetadata(new TemplateStore(service).Read(revisionId));
        var saved = service.Retrieve("asx_revision", revisionId, new ColumnSet("asx_status"));
        context.OutputParameters["Result"] = JsonWire.Write(
            new DraftResult
            {
                TemplateId = templateId.ToString("D"),
                RevisionId = revisionId.ToString("D"),
                RowVersion = saved.RowVersion,
            }
        );
    }
}
