using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Domain;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public sealed class TemplateStore
{
    private readonly IOrganizationService service;

    public TemplateStore(IOrganizationService service)
    {
        this.service = service;
    }

    /// <summary>
    /// Loads a revision's normalized rows and validates the assembled template and destination catalog.
    /// </summary>
    /// <param name="revisionId">The Dataverse revision row whose configuration is loaded.</param>
    /// <returns>The validated template with its sources, destinations, and folder tree.</returns>
    public DocumentTemplate Read(Guid revisionId)
    {
        var revision = service.Retrieve(
            "asx_revision",
            revisionId,
            new ColumnSet("asx_templateid", "asx_version")
        );
        var templateRef =
            revision.GetAttributeValue<EntityReference>("asx_templateid")
            ?? throw new EvaluationBlockedException("Revision has no template.");
        var templateRow = service.Retrieve(
            "asx_template",
            templateRef.Id,
            new ColumnSet("asx_table")
        );
        var template = new DocumentTemplate
        {
            Id = templateRow.Id,
            Table = Text(templateRow, "asx_table"),
            Revision = revision.GetAttributeValue<int>("asx_version"),
        };
        foreach (var row in Children("asx_source", revisionId, "asx_payload"))
            template.Sources.Add(JsonWire.Read<SourceDto>(Text(row, "asx_payload")).ToModel());
        var folders = Children(
            "asx_folder",
            revisionId,
            "asx_key",
            "asx_sectionkey",
            "asx_parentkey",
            "asx_expression",
            "asx_groupkey",
            "asx_order"
        );
        var groups = Children(
            "asx_documentconditiongroup",
            revisionId,
            "asx_key",
            "asx_parentkey",
            "asx_all"
        );
        var conditions = Children("asx_condition", revisionId, "asx_groupkey", "asx_payload");
        var byGroup = groups.ToDictionary(g => Text(g, "asx_key"), StringComparer.Ordinal);
        var usedGroups = new HashSet<string>(StringComparer.Ordinal);
        Predicate Group(string key, HashSet<string> path)
        {
            if (path.Count >= 10 || !path.Add(key) || !byGroup.TryGetValue(key, out var row))
                throw new EvaluationBlockedException("Invalid condition tree.");
            if (!usedGroups.Add(key))
                throw new EvaluationBlockedException(
                    "Condition groups must belong to one folder tree."
                );
            var children = conditions
                .Where(c => Text(c, "asx_groupkey") == key)
                .Select(c =>
                    (Predicate)JsonWire.Read<ConditionDto>(Text(c, "asx_payload")).ToModel()
                )
                .ToList();
            foreach (
                var child in groups.Where(g => g.GetAttributeValue<string>("asx_parentkey") == key)
            )
                children.Add(
                    Group(Text(child, "asx_key"), new HashSet<string>(path, StringComparer.Ordinal))
                );
            return new ConditionGroup(row.GetAttributeValue<bool>("asx_all"), children.ToArray());
        }
        var sectionKeys = new HashSet<string>();
        foreach (var row in Children("asx_destination", revisionId, "asx_key", "asx_libraryid"))
        {
            var key = Text(row, "asx_key");
            if (!sectionKeys.Add(key))
                throw new EvaluationBlockedException("Duplicate destination section key.");
            var libraryRef =
                row.GetAttributeValue<EntityReference>("asx_libraryid")
                ?? throw new EvaluationBlockedException("Destination has no approved library.");
            var library = service.Retrieve(
                "asx_library",
                libraryRef.Id,
                new ColumnSet(
                    "asx_siteid",
                    "asx_listid",
                    "asx_entryid",
                    "asx_entryurl",
                    "asx_policyrevision",
                    "asx_approved",
                    "asx_policyapplied"
                )
            );
            var siteRef =
                library.GetAttributeValue<EntityReference>("asx_siteid")
                ?? throw new EvaluationBlockedException("Library has no site.");
            var site = service.Retrieve(
                "asx_site",
                siteRef.Id,
                new ColumnSet("asx_webid", "asx_url", "asx_approved")
            );
            var entryUrl = Text(library, "asx_entryurl");
            var siteUrl = new Uri(Text(site, "asx_url")).AbsoluteUri.TrimEnd('/');
            if (!entryUrl.StartsWith(siteUrl + "/", StringComparison.OrdinalIgnoreCase))
                throw new EvaluationBlockedException("Library entry is outside approved web.");
            var section = new DestinationSection
            {
                Key = key,
                Library = new ApprovedLibrary
                {
                    ApprovalId = library.Id,
                    WebId = Guid.Parse(Text(site, "asx_webid")),
                    ListId = Guid.Parse(Text(library, "asx_listid")),
                    EntryId = Guid.Parse(Text(library, "asx_entryid")),
                    // Informational only: a library whose first policy has not been applied has
                    // no revision, and policy state never blocks planning.
                    PolicyRevision = Guid.TryParse(
                        library.GetAttributeValue<string>("asx_policyrevision"),
                        out var policyRevision
                    )
                        ? policyRevision
                        : Guid.Empty,
                    EntryUrl = entryUrl,
                    Approved =
                        site.GetAttributeValue<bool>("asx_approved")
                        && library.GetAttributeValue<bool>("asx_approved"),
                    PolicyApplied = library.GetAttributeValue<bool>("asx_policyapplied"),
                },
            };
            foreach (var folder in folders.Where(f => Text(f, "asx_sectionkey") == key))
            {
                var groupKey = folder.GetAttributeValue<string>("asx_groupkey");
                if (
                    groupKey != null
                    && byGroup.TryGetValue(groupKey, out var rootGroup)
                    && rootGroup.GetAttributeValue<string>("asx_parentkey") != null
                )
                    throw new EvaluationBlockedException(
                        "Folder must reference a root condition group."
                    );
                section.Nodes.Add(
                    new FolderNode
                    {
                        Key = Text(folder, "asx_key"),
                        ParentKey = folder.GetAttributeValue<string>("asx_parentkey"),
                        Name = Text(folder, "asx_expression"),
                        Order = folder.GetAttributeValue<int>("asx_order"),
                        Condition =
                            groupKey == null ? null : Group(groupKey, new HashSet<string>()),
                    }
                );
            }
            template.Destinations.Add(section);
        }
        if (
            folders.Any(f => !sectionKeys.Contains(Text(f, "asx_sectionkey")))
            || usedGroups.Count != groups.Count
            || conditions.Any(c => !usedGroups.Contains(Text(c, "asx_groupkey")))
        )
            throw new EvaluationBlockedException("Orphaned configuration rows are not permitted.");
        TemplateValidator.Validate(template);
        return template;
    }

    public IReadOnlyList<Entity> Children(string table, Guid revision, params string[] columns)
    {
        var query = new QueryExpression(table)
        {
            ColumnSet = new ColumnSet(columns),
            PageInfo = new PagingInfo { Count = 1000, PageNumber = 1 },
        };
        query.Criteria.AddCondition("asx_revisionid", ConditionOperator.Equal, revision);
        var result = service.RetrieveMultiple(query);
        if (result.MoreRecords || result.Entities.Count >= 1000)
            throw new EvaluationBlockedException(
                "Configuration query exceeded the completeness bound."
            );
        return result.Entities.ToArray();
    }

    internal static string Text(Entity entity, string column) =>
        entity.GetAttributeValue<string>(column)
        ?? throw new EvaluationBlockedException("Required configuration value missing: " + column);
}
