using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Ascentix.Documents.Conditions;

namespace Ascentix.Documents.Domain;

public sealed class ValueSource
{
    public string Alias { get; set; } = "root";
    public string Table { get; set; } = "";
    public string? RootLookupColumn { get; set; }
    public IDictionary<string, ValueKind> Columns { get; set; } =
        new Dictionary<string, ValueKind>();
}

public sealed class ApprovedLibrary
{
    public Guid ApprovalId { get; set; }
    public Guid WebId { get; set; }
    public Guid ListId { get; set; }
    public Guid EntryId { get; set; }
    public string EntryUrl { get; set; } = "";
    public bool Approved { get; set; }
}

public sealed class FolderNode
{
    public string Key { get; set; } = "";
    public string? ParentKey { get; set; }
    public string Name { get; set; } = "";
    public Predicate? Condition { get; set; }
    public int Order { get; set; }
}

public sealed class DestinationSection
{
    public string Key { get; set; } = "";
    public ApprovedLibrary Library { get; set; } = new ApprovedLibrary();
    public IList<FolderNode> Nodes { get; set; } = new List<FolderNode>();
}

public sealed class DocumentTemplate
{
    public Guid Id { get; set; }
    public int Revision { get; set; }
    public string Table { get; set; } = "";
    public IList<ValueSource> Sources { get; set; } = new List<ValueSource>();
    public IList<DestinationSection> Destinations { get; set; } = new List<DestinationSection>();
}

public sealed class FolderIntent
{
    public string Section { get; }
    public string Node { get; }
    public string? Parent { get; }
    public string Name { get; }
    public string RelativePath { get; }
    public string BindingKey { get; }
    public Guid LibraryApprovalId { get; }
    public bool IsRoot => Parent == null;

    public FolderIntent(
        string section,
        string node,
        string? parent,
        string name,
        string path,
        string bindingKey,
        Guid libraryApprovalId
    )
    {
        Section = section;
        Node = node;
        Parent = parent;
        Name = name;
        RelativePath = path;
        BindingKey = bindingKey;
        LibraryApprovalId = libraryApprovalId;
    }
}

public static class FolderNames
{
    /// <summary>
    /// Rejects empty, oversized, reserved, and path-containing SharePoint folder names.
    /// </summary>
    /// <param name="name">The resolved folder name, excluding its parent path.</param>
    public static void Validate(string name)
    {
        if (
            string.IsNullOrWhiteSpace(name)
            || name.Length > 255
            || name != name.Trim()
            || name.EndsWith(".", StringComparison.Ordinal)
            || name.Any(c => char.IsControl(c) || "\"*:<>?/\\|".IndexOf(c) >= 0)
            || name == "."
            || name == ".."
            || name.Equals("Forms", StringComparison.OrdinalIgnoreCase)
            || name.Equals("desktop.ini", StringComparison.OrdinalIgnoreCase)
            || name.StartsWith("~$", StringComparison.Ordinal)
            || name.IndexOf("_vti_", StringComparison.OrdinalIgnoreCase) >= 0
            || Regex.IsMatch(
                name,
                @"^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])($|\.)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
            )
        )
            throw new EvaluationBlockedException("Resolved folder name is invalid or reserved.");
    }
}

public static class TemplateValidator
{
    internal static void Key(string key)
    {
        if (!Regex.IsMatch(key ?? "", "^[a-z][a-z0-9_-]{0,79}$"))
            throw new EvaluationBlockedException(
                "Stable keys must be lowercase ASCII identifiers."
            );
    }

    public static void Validate(DocumentTemplate template)
    {
        if (
            template.Id == Guid.Empty
            || template.Revision < 1
            || !Regex.IsMatch(template.Table, "^[a-z][a-z0-9_]*$")
        )
            throw new EvaluationBlockedException("Invalid template identity.");
        if (
            template.Destinations.Count < 1
            || template.Destinations.Count > 10
            || template.Sources.Count < 1
            || template.Sources.Count > 6
        )
            throw new EvaluationBlockedException("Template exceeds destination/source limits.");
        var sources = new Dictionary<string, ValueSource>(StringComparer.Ordinal);
        foreach (var source in template.Sources)
        {
            new FieldReference(source.Alias, source.Table);
            if (sources.ContainsKey(source.Alias))
                throw new EvaluationBlockedException("Duplicate source alias.");
            sources.Add(source.Alias, source);
            if (
                source.Alias == "root"
                    ? source.Table != template.Table || source.RootLookupColumn != null
                    : string.IsNullOrWhiteSpace(source.RootLookupColumn)
            )
                throw new EvaluationBlockedException(
                    "Only root and direct lookup sources are supported."
                );
            if (source.RootLookupColumn != null)
                new FieldReference("root", source.RootLookupColumn);
            if (source.Columns.Count > 100)
                throw new EvaluationBlockedException("Too many projected columns.");
            foreach (var column in source.Columns)
            {
                new FieldReference(source.Alias, column.Key);
                if (!Enum.IsDefined(typeof(ValueKind), column.Value))
                    throw new EvaluationBlockedException("Unknown field type.");
            }
        }
        if (!sources.ContainsKey("root"))
            throw new EvaluationBlockedException("Root source is required.");
        var sectionKeys = new HashSet<string>(StringComparer.Ordinal);
        int leaves = 0;
        foreach (var section in template.Destinations)
        {
            Key(section.Key);
            if (!sectionKeys.Add(section.Key))
                throw new EvaluationBlockedException("Invalid destination section.");
            var lib = section.Library;
            // Folders inherit the library's permissions, so the library's policy-update state
            // never gates planning. Only suspension does.
            if (
                !lib.Approved
                || lib.ApprovalId == Guid.Empty
                || lib.WebId == Guid.Empty
                || lib.ListId == Guid.Empty
                || lib.EntryId == Guid.Empty
            )
                throw new EvaluationBlockedException(
                    "Destination must be approved and have a valid identity."
                );
            if (
                !Uri.TryCreate(lib.EntryUrl, UriKind.Absolute, out var url)
                || url.Scheme != "https"
                || url.Query.Length != 0
                || url.Fragment.Length != 0
                || url.UserInfo.Length != 0
            )
                throw new EvaluationBlockedException("Invalid approved entry URL.");
            if (section.Nodes.Count < 1 || section.Nodes.Count > 100)
                throw new EvaluationBlockedException("Folder count exceeds bounds.");
            var nodes = new Dictionary<string, FolderNode>(StringComparer.Ordinal);
            foreach (var node in section.Nodes)
            {
                Key(node.Key);
                if (nodes.ContainsKey(node.Key))
                    throw new EvaluationBlockedException("Duplicate node key.");
                nodes.Add(node.Key, node);
            }
            if (section.Nodes.Count(n => n.ParentKey == null) != 1)
                throw new EvaluationBlockedException("Exactly one root is required.");
            foreach (var node in section.Nodes)
            {
                var seen = new HashSet<string>();
                var cursor = node;
                while (true)
                {
                    if (!seen.Add(cursor.Key) || seen.Count > 10)
                        throw new EvaluationBlockedException("Folder cycle/depth exceeds bounds.");
                    if (cursor.ParentKey == null)
                        break;
                    if (!nodes.TryGetValue(cursor.ParentKey, out cursor))
                        throw new EvaluationBlockedException("Missing parent node.");
                }
                foreach (var field in NameExpression.Fields(node.Name))
                    RequireField(sources, field);
                if (node.Condition != null)
                {
                    node.Condition.Validate();
                    leaves += ValidatePredicate(sources, node.Condition);
                }
            }
        }
        if (leaves > 100)
            throw new EvaluationBlockedException("Template condition leaf limit exceeded.");
    }

    private static ValueKind RequireField(
        IDictionary<string, ValueSource> sources,
        FieldReference field
    )
    {
        if (
            !sources.TryGetValue(field.Source, out var source)
            || !source.Columns.TryGetValue(field.Column, out var kind)
        )
            throw new EvaluationBlockedException(
                "Field is outside the approved source projection: " + field
            );
        return kind;
    }

    private static int ValidatePredicate(
        IDictionary<string, ValueSource> sources,
        Predicate predicate
    )
    {
        if (predicate is ConditionGroup group)
            return group.Children.Sum(child => ValidatePredicate(sources, child));
        if (!(predicate is Condition condition))
            throw new EvaluationBlockedException("Unknown predicate type.");
        var kind = RequireField(sources, condition.Left);
        if (
            !ValueComparer.Supports(kind, condition.Operator)
            || (
                condition.Literal != null
                && (condition.Literal.Kind != kind || condition.Literal.IsNull)
            )
            || (condition.Right != null && RequireField(sources, condition.Right) != kind)
        )
            throw new EvaluationBlockedException("Typed operator/value mismatch.");
        return 1;
    }
}

public static class FolderPlanner
{
    public static IReadOnlyList<FolderIntent> Plan(
        DocumentTemplate template,
        Guid recordId,
        Snapshot snapshot
    )
    {
        TemplateValidator.Validate(template);
        if (recordId == Guid.Empty)
            throw new EvaluationBlockedException("Persisted record identity required.");
        var result = new List<FolderIntent>();
        foreach (var section in template.Destinations)
        {
            var binding =
                template.Id.ToString("N") + ":" + recordId.ToString("N") + ":" + section.Key;
            Action<string?, string> visit = null!;
            visit = (parent, parentPath) =>
            {
                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (
                    var node in section
                        .Nodes.Where(n => n.ParentKey == parent)
                        .OrderBy(n => n.Order)
                        .ThenBy(n => n.Key, StringComparer.Ordinal)
                )
                {
                    if (node.Condition != null && !node.Condition.Evaluate(snapshot))
                        continue;
                    var name = NameExpression.Render(node.Name, snapshot);
                    FolderNames.Validate(name);
                    if (!names.Add(name))
                        throw new EvaluationBlockedException(
                            "Included siblings resolve to the same name."
                        );
                    string path = parent == null ? name : parentPath + "/" + name;
                    if (
                        Uri.UnescapeDataString(new Uri(section.Library.EntryUrl).AbsolutePath)
                            .TrimEnd('/')
                            .Length
                            + 1
                            + path.Length
                            + 100
                        > 400
                    )
                        throw new EvaluationBlockedException(
                            "Folder path does not retain the file-name budget."
                        );
                    result.Add(
                        new FolderIntent(
                            section.Key,
                            node.Key,
                            parent,
                            name,
                            path,
                            binding + ":" + node.Key,
                            section.Library.ApprovalId
                        )
                    );
                    visit(node.Key, path);
                }
            };
            visit(null, "");
        }
        return result.AsReadOnly();
    }
}
