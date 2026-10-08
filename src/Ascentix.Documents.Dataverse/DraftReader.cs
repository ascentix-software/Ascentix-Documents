using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
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

    /// <summary>The revision's version number (asx_version), for the version chip.</summary>
    [DataMember]
    public int Version { get; set; }
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
            new ColumnSet("asx_templateid", "asx_status", "asx_version")
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
        var tables = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var source in sources)
            tables[source.Alias] = source.Table;
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
        GroupDto ReadGroup(string key)
        {
            if (!used.Add(key) || !indexed.TryGetValue(key, out var row))
                throw new EvaluationBlockedException("Invalid/shared condition group ownership.");
            return new GroupDto
            {
                All = row.GetAttributeValue<bool>("asx_all"),
                Conditions = conditions
                    .Where(c => TemplateStore.Text(c, "asx_groupkey") == key)
                    .Select(c =>
                        Label(
                            JsonWire.Read<ConditionDto>(TemplateStore.Text(c, "asx_payload")),
                            tables
                        )
                    )
                    .ToArray(),
                Groups = groups
                    .Where(g => g.GetAttributeValue<string>("asx_parentkey") == key)
                    .Select(g => ReadGroup(TemplateStore.Text(g, "asx_key")))
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
                                Condition = group == null ? null : ReadGroup(group),
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
            Version = revision.GetAttributeValue<int>("asx_version"),
        };
    }

    private Guid? caller;

    // Metadata and privileges read once per reader: a draft may compare many conditions with
    // the same lookup column, and each answer is the same for all of them.
    private readonly Dictionary<string, string[]> targets = new Dictionary<string, string[]>(
        StringComparer.Ordinal
    );
    private readonly Dictionary<string, EntityMetadata?> readable = new Dictionary<
        string,
        EntityMetadata?
    >(StringComparer.Ordinal);

    // The record a lookup condition compares with, by name, read as the caller from each target
    // of the lookup in turn; null when it is not found or not readable (spec 6.4). LoadDraft runs
    // in a transaction that a faulting call would end even when caught, so every read here is
    // one that answers instead of faulting: metadata through RetrieveMetadataChanges (a deleted
    // table or column is simply absent), and a table is queried only when the caller holds its
    // Read privilege (a query without it faults).
    private ConditionDto Label(ConditionDto condition, IReadOnlyDictionary<string, string> tables)
    {
        condition.LiteralLabel = null;
        condition.LiteralTable = null;
        if (
            condition.LiteralKind != "Lookup"
            || !Guid.TryParse(condition.Literal, out var id)
            || !tables.TryGetValue(condition.Source, out var table)
        )
            return condition;
        foreach (var target in Targets(table, condition.Column))
        {
            var metadata = Readable(target);
            if (metadata == null)
                continue;
            var query = new QueryExpression(target)
            {
                ColumnSet = new ColumnSet(metadata.PrimaryNameAttribute),
                TopCount = 1,
            };
            query.Criteria.AddCondition(metadata.PrimaryIdAttribute, ConditionOperator.Equal, id);
            var found = service.RetrieveMultiple(query).Entities.FirstOrDefault();
            if (found == null)
                continue;
            condition.LiteralLabel = found.GetAttributeValue<string>(metadata.PrimaryNameAttribute);
            condition.LiteralTable = target;
            break;
        }
        return condition;
    }

    // The tables a lookup column targets.
    private string[] Targets(string table, string column)
    {
        string key = table + "." + column;
        if (!targets.TryGetValue(key, out var found))
            targets[key] = found = TableInfo.LookupTargets(service, table, column);
        return found;
    }

    // The table's metadata when the caller can query it for a name, else null.
    private EntityMetadata? Readable(string table)
    {
        if (!readable.TryGetValue(table, out var metadata))
        {
            metadata = TableInfo.Find(service, table);
            if (
                metadata?.PrimaryIdAttribute == null
                || metadata.PrimaryNameAttribute == null
                || !TableInfo.CanRead(service, metadata, Caller())
            )
                metadata = null;
            readable[table] = metadata;
        }
        return metadata;
    }

    private Guid Caller() =>
        caller ??= (
            (Microsoft.Crm.Sdk.Messages.WhoAmIResponse)
                service.Execute(new Microsoft.Crm.Sdk.Messages.WhoAmIRequest())
        ).UserId;
}
