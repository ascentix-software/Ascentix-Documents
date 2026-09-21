using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class LoadedDraft
{
    [DataMember]
    public DraftDto Draft { get; set; } = new DraftDto();

    [DataMember]
    public string RevisionId { get; set; } = "";

    [DataMember]
    public string RowVersion { get; set; } = "";

    [DataMember]
    public string Status { get; set; } = "";
}

public sealed class DraftReader
{
    private readonly IOrganizationService service;

    public DraftReader(IOrganizationService service)
    {
        this.service = service;
    }

    /// <summary>
    /// Loads editable configuration and its row version without validating current destination health.
    /// </summary>
    /// <param name="revisionId">The Dataverse revision row to load.</param>
    /// <returns>The revision's authoring configuration, status, and row version.</returns>
    public LoadedDraft Read(Guid revisionId)
    {
        var revision = service.Retrieve(
            "asx_revision",
            revisionId,
            new ColumnSet("asx_templateid", "asx_status")
        );
        var templateRef =
            revision.GetAttributeValue<EntityReference>("asx_templateid")
            ?? throw new EvaluationBlockedException("Revision template is missing.");
        var template = service.Retrieve(
            "asx_template",
            templateRef.Id,
            new ColumnSet("asx_table", "asx_name")
        );
        var store = new TemplateStore(service);
        var sources = store
            .Children("asx_source", revisionId, "asx_payload")
            .Select(row => JsonWire.Read<SourceDto>(TemplateStore.Text(row, "asx_payload")))
            .ToArray();
        var destinations = store.Children(
            "asx_destination",
            revisionId,
            "asx_name",
            "asx_key",
            "asx_libraryid"
        );
        var folders = store.Children(
            "asx_folder",
            revisionId,
            "asx_key",
            "asx_sectionkey",
            "asx_parentkey",
            "asx_expression",
            "asx_groupkey",
            "asx_order"
        );
        var groups = store.Children(
            "asx_documentconditiongroup",
            revisionId,
            "asx_key",
            "asx_parentkey",
            "asx_all"
        );
        var conditions = store.Children("asx_condition", revisionId, "asx_groupkey", "asx_payload");
        var indexed = groups.ToDictionary(
            row => TemplateStore.Text(row, "asx_key"),
            StringComparer.Ordinal
        );
        var used = new HashSet<string>();
        GroupDto ReadGroup(string key, int depth)
        {
            if (depth > 10 || !used.Add(key) || !indexed.TryGetValue(key, out var row))
                throw new EvaluationBlockedException("Invalid/shared condition group ownership.");
            return new GroupDto
            {
                All = row.GetAttributeValue<bool>("asx_all"),
                Conditions = conditions
                    .Where(c => TemplateStore.Text(c, "asx_groupkey") == key)
                    .Select(c => JsonWire.Read<ConditionDto>(TemplateStore.Text(c, "asx_payload")))
                    .ToArray(),
                Groups = groups
                    .Where(g => g.GetAttributeValue<string>("asx_parentkey") == key)
                    .Select(g => ReadGroup(TemplateStore.Text(g, "asx_key"), depth + 1))
                    .ToArray(),
            };
        }
        var sections = destinations
            .Select(destination =>
            {
                string section = TemplateStore.Text(destination, "asx_key");
                return new DestinationDto
                {
                    Key = section,
                    Name = destination.GetAttributeValue<string>("asx_name"),
                    LibraryId = (
                        destination.GetAttributeValue<EntityReference>("asx_libraryid")
                        ?? throw new EvaluationBlockedException(
                            "Destination library reference missing."
                        )
                    ).Id.ToString("D"),
                    Folders = folders
                        .Where(f => TemplateStore.Text(f, "asx_sectionkey") == section)
                        .OrderBy(f => f.GetAttributeValue<string>("asx_parentkey") == null ? 0 : 1)
                        .ThenBy(f => f.GetAttributeValue<int>("asx_order"))
                        .Select(f =>
                        {
                            var group = f.GetAttributeValue<string>("asx_groupkey");
                            return new FolderDto
                            {
                                Key = TemplateStore.Text(f, "asx_key"),
                                Parent = f.GetAttributeValue<string>("asx_parentkey"),
                                Name = TemplateStore.Text(f, "asx_expression"),
                                Condition = group == null ? null : ReadGroup(group, 0),
                            };
                        })
                        .ToArray(),
                };
            })
            .ToArray();
        if (
            used.Count != groups.Count
            || conditions.Any(c => !used.Contains(TemplateStore.Text(c, "asx_groupkey")))
            || folders.Count != sections.Sum(s => s.Folders.Length)
        )
            throw new EvaluationBlockedException(
                "Orphaned child rows require configuration repair."
            );
        return new LoadedDraft
        {
            Draft = new DraftDto
            {
                TemplateId = templateRef.Id.ToString("D"),
                Name = template.GetAttributeValue<string>("asx_name"),
                Table = TemplateStore.Text(template, "asx_table"),
                Sources = sources,
                Destinations = sections,
            },
            RevisionId = revisionId.ToString("D"),
            RowVersion = revision.RowVersion,
            Status = TemplateStore.Text(revision, "asx_status"),
        };
    }
}
