using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Domain;
using Microsoft.Xrm.Sdk;

namespace Ascentix.Documents.Dataverse;

/// <summary>
/// Validates a draft and converts it to a template without writing anything (spec 6.3). Save
/// calls it before writing; Preview calls it for unsaved edits. Both refuse with these texts.
/// </summary>
public static class DraftTemplate
{
    public const string BoundsRefusal = "Draft exceeds the supported bounds.";

    public const string GroupRefusal = "Invalid condition group bounds.";

    /// <param name="service">The caller's service; libraries and sites are read as the caller.</param>
    /// <param name="draft">The editor's draft.</param>
    /// <returns>The template the draft describes, validated like a saved revision.</returns>
    public static DocumentTemplate Build(IOrganizationService service, DraftDto draft)
    {
        new FieldReference("root", draft.Table);
        if (
            draft.Sources == null
            || draft.Destinations == null
            || draft.Sources.Length < 1
            || draft.Sources.Length > Bounds.Sources
            || draft.Destinations.Length < 1
            || draft.Destinations.Length > Bounds.Destinations
            || draft.Destinations.Any(d =>
                d.Folders == null
                || d.Folders.Length < 1
                || d.Folders.Length > Bounds.FoldersPerDestination
            )
        )
            throw new EvaluationBlockedException(BoundsRefusal);
        int rows =
            draft.Sources.Length
            + draft.Destinations.Sum(d => 1 + d.Folders.Sum(f => 1 + Count(f.Condition)));
        if (rows > Bounds.ConfigurationRows)
            throw new EvaluationBlockedException(BoundsRefusal);
        var template = new DocumentTemplate
        {
            Id =
                Guid.TryParse(draft.TemplateId, out var id) && id != Guid.Empty
                    ? id
                    : DocumentStore.StableId("draft-preview:" + draft.Table),
            Table = draft.Table,
            Revision = 1,
        };
        foreach (var source in draft.Sources)
            template.Sources.Add(source.ToModel());
        var store = new TemplateStore(service);
        foreach (var destination in draft.Destinations)
        {
            if (destination.Name?.Length > 200)
                throw new EvaluationBlockedException(
                    "Destination names must be 200 characters or fewer."
                );
            var section = store.Destination(destination.Key, Guid.Parse(destination.LibraryId));
            int order = 0;
            foreach (var folder in destination.Folders)
                section.Nodes.Add(
                    new FolderNode
                    {
                        Key = folder.Key,
                        ParentKey = folder.Parent,
                        Name = folder.Name,
                        Order = order++,
                        Condition = folder.Condition == null ? null : Group(folder.Condition),
                    }
                );
            template.Destinations.Add(section);
        }
        TemplateValidator.Validate(template);
        new SnapshotReader(service).ValidateMetadata(template);
        return template;
    }

    // The configuration rows a condition tree writes: one per group and one per condition. A
    // group the serializer left without its lists is refused by Group.
    private static int Count(GroupDto? group) =>
        group == null ? 0 : 1 + (group.Conditions?.Length ?? 0) + (group.Groups?.Sum(Count) ?? 0);

    private static Predicate Group(GroupDto? group)
    {
        if (
            group == null
            || group.Conditions == null
            || group.Groups == null
            || group.Conditions.Length + group.Groups.Length == 0
            || group.Conditions.Any(c => c == null)
        )
            throw new EvaluationBlockedException(GroupRefusal);
        return new ConditionGroup(
            group.All,
            group
                .Conditions.Select(c => (Predicate)c.ToModel())
                .Concat(group.Groups.Select(Group))
                .ToArray()
        );
    }
}
