using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

/// <summary>
/// SharePoint folder-name rules, from Microsoft's "Restrictions and limitations in OneDrive and
/// SharePoint", section "Invalid file names and file types"
/// (https://support.microsoft.com/office/64883a5d-228e-48f5-b3d2-eb39e07630fa):
/// the characters " * : &lt; &gt; ? / \ | are not allowed; leading and trailing spaces are not
/// allowed; the names .lock, CON, PRN, AUX, NUL, COM0-COM9, LPT0-LPT9 and desktop.ini are not
/// allowed; _vti_ cannot appear anywhere in a name; a name cannot start with ~$; "forms" cannot
/// be used at the root level of a library; and U+309B and U+1027 cannot be the first character
/// of a folder. Trailing periods, control characters and the 255-character name limit are the
/// Windows and SharePoint limits the product has always enforced.
/// </summary>
public static class FolderNames
{
    private const string Forbidden = "\"*:<>?/\\|";
    private const int MaxLength = 255;

    private static readonly Regex DeviceName = new Regex(
        @"^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])($|\.)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
    );

    private static readonly Regex Vti = new Regex(
        "_(vti)_",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
    );

    /// <summary>
    /// Rejects empty, oversized, reserved, and path-containing SharePoint folder names. This
    /// guards names that already exist or come from outside; names built from record data go
    /// through <see cref="Clean"/> instead.
    /// </summary>
    /// <param name="name">The folder name, excluding its parent path.</param>
    /// <param name="atLibraryRoot">
    /// False only when the caller knows the folder is below the library root, where SharePoint
    /// allows "Forms".
    /// </param>
    public static void Validate(string name, bool atLibraryRoot = true)
    {
        if (
            string.IsNullOrWhiteSpace(name)
            || name.Length > MaxLength
            || name != name.Trim()
            || name.EndsWith(".", StringComparison.Ordinal)
            || name.Any(c => char.IsControl(c) || Forbidden.IndexOf(c) >= 0)
            || name == "."
            || name == ".."
            || (atLibraryRoot && name.Equals("Forms", StringComparison.OrdinalIgnoreCase))
            || name.Equals("desktop.ini", StringComparison.OrdinalIgnoreCase)
            || name.Equals(".lock", StringComparison.OrdinalIgnoreCase)
            || name[0] == (char)0x309B
            || name[0] == (char)0x1027
            || name.StartsWith("~$", StringComparison.Ordinal)
            || name.IndexOf("_vti_", StringComparison.OrdinalIgnoreCase) >= 0
            || DeviceName.IsMatch(name)
        )
            throw new EvaluationBlockedException("Resolved folder name is invalid or reserved.");
    }

    /// <summary>
    /// Turns a name built from record data into a name SharePoint accepts, instead of refusing
    /// it. Forbidden and control characters become "-", the ends are trimmed of spaces and
    /// trailing periods, reserved names get a "_" and the name is cut to 255 characters. The
    /// result always passes <see cref="Validate"/> at the same depth.
    /// </summary>
    /// <param name="raw">The rendered name.</param>
    /// <param name="atLibraryRoot">True when the folder may sit at the root of the library.</param>
    /// <returns>The cleaned name, or null when the value has nothing SharePoint can use.</returns>
    public static string? Clean(string? raw, bool atLibraryRoot)
    {
        // A value made only of characters that are removed or replaced has nothing to name the
        // folder with; "***" would otherwise become "---".
        if (
            raw == null
            || !raw.Any(c =>
                !char.IsControl(c) && !char.IsWhiteSpace(c) && c != '.' && Forbidden.IndexOf(c) < 0
            )
        )
            return null;
        string name = TrimEnds(
            new string(
                raw.Trim()
                    .Select(c => char.IsControl(c) || Forbidden.IndexOf(c) >= 0 ? '-' : c)
                    .ToArray()
            )
        );
        if (name.Length > 0 && (name[0] == (char)0x309B || name[0] == (char)0x1027))
            name = "-" + name.Substring(1);
        if (name.StartsWith("~$", StringComparison.Ordinal))
            name = "-$" + name.Substring(2);
        name = Vti.Replace(name, "-$1-");
        while (true)
        {
            name = Unreserve(name, atLibraryRoot);
            if (name.Length <= MaxLength)
                break;
            int cut = MaxLength;
            if (char.IsHighSurrogate(name[cut - 1]))
                cut--;
            name = TrimEnds(name.Substring(0, cut));
        }
        return name.Length == 0 ? null : name;
    }

    private static string TrimEnds(string name)
    {
        name = name.TrimStart();
        int end = name.Length;
        while (end > 0 && (char.IsWhiteSpace(name[end - 1]) || name[end - 1] == '.'))
            end--;
        return name.Substring(0, end);
    }

    // "." and ".." never reach here: trailing periods are already trimmed away.
    private static string Unreserve(string name, bool atLibraryRoot)
    {
        if (DeviceName.IsMatch(name))
        {
            // "CON" becomes "CON_" and "con.txt" becomes "con_.txt"; a suffix after the
            // extension would still start with the reserved "con.".
            int dot = name.IndexOf('.');
            return dot < 0 ? name + "_" : name.Substring(0, dot) + "_" + name.Substring(dot);
        }
        if (
            name.Equals("desktop.ini", StringComparison.OrdinalIgnoreCase)
            || name.Equals(".lock", StringComparison.OrdinalIgnoreCase)
            || (atLibraryRoot && name.Equals("Forms", StringComparison.OrdinalIgnoreCase))
        )
            return name + "_";
        return name;
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

/// <summary>A record's planned folders, with the notices planning recorded for them.</summary>
public sealed class FolderPlan : ReadOnlyCollection<FolderIntent>
{
    /// <summary>
    /// Things an admin should see about this plan, such as a folder name that was adjusted for
    /// SharePoint. A notice never means the plan failed.
    /// </summary>
    public IReadOnlyList<string> Notices { get; }

    /// <summary>
    /// Folders this plan skipped until the record changes: a blank or unusable name, or a name a
    /// sibling already took. Each one's notice is also in <see cref="Notices"/> unless the
    /// plan reached its notice limit.
    /// </summary>
    public IReadOnlyList<FolderWait> Waits { get; }

    public FolderPlan(IList<FolderIntent> intents, IList<string> notices)
        : this(intents, notices, new List<FolderWait>()) { }

    public FolderPlan(IList<FolderIntent> intents, IList<string> notices, IList<FolderWait> waits)
        : base(intents)
    {
        Notices = new ReadOnlyCollection<string>(notices);
        Waits = new ReadOnlyCollection<FolderWait>(waits);
    }
}

/// <summary>A folder a plan skipped, and everything below it, until the record changes.</summary>
public sealed class FolderWait
{
    public FolderWait(string section, string node, FieldReference? field, string notice)
    {
        Section = section;
        Node = node;
        Field = field;
        Notice = notice;
    }

    public string Section { get; }
    public string Node { get; }

    /// <summary>The blank field the folder waits for; null when it waits for a different name.</summary>
    public FieldReference? Field { get; }

    /// <summary>The notice recorded for it, exactly as in the plan's notices.</summary>
    public string Notice { get; }
}

public static class FolderPlanner
{
    // Bounded like the other stored notices so a plan stays well inside the 500,000-character
    // payload limit of the rows that keep it.
    private const int NoticeLimit = 200;
    private const int NoticeLength = 600;

    /// <summary>Records a notice once and returns its bounded text.</summary>
    private static string Notice(IList<string> notices, string text)
    {
        text = new string(text.Where(c => !char.IsControl(c)).ToArray());
        if (text.Length > NoticeLength)
            text = text.Substring(0, NoticeLength) + "...";
        string recorded = notices.Count >= NoticeLimit ? "More notices were not recorded." : text;
        if (notices.Count <= NoticeLimit && !notices.Contains(recorded))
            notices.Add(recorded);
        return text;
    }

    public static FolderPlan Plan(DocumentTemplate template, Guid recordId, Snapshot snapshot)
    {
        TemplateValidator.Validate(template);
        if (recordId == Guid.Empty)
            throw new EvaluationBlockedException("Persisted record identity required.");
        var result = new List<FolderIntent>();
        var notices = new List<string>();
        var waits = new List<FolderWait>();
        foreach (var section in template.Destinations)
        {
            var binding =
                template.Id.ToString("N") + ":" + recordId.ToString("N") + ":" + section.Key;
            Action<string?, string> visit = null!;
            visit = (parent, parentPath) =>
            {
                // Each cleaned name and the first included sibling that took it.
                var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (
                    var node in section
                        .Nodes.Where(n => n.ParentKey == parent)
                        .OrderBy(n => n.Order)
                        .ThenBy(n => n.Key, StringComparer.Ordinal)
                )
                {
                    if (node.Condition != null && !node.Condition.Evaluate(snapshot))
                        continue;
                    var raw = NameExpression.Render(node.Name, snapshot, out var blank);
                    if (raw == null)
                    {
                        // Skip this folder and everything below it until the field is filled
                        // in; the record's next update or a replan plans it then.
                        waits.Add(
                            new FolderWait(
                                section.Key,
                                node.Key,
                                blank,
                                Notice(
                                    notices,
                                    "Folder '"
                                        + section.Key
                                        + "/"
                                        + node.Key
                                        + "' is waiting for '"
                                        + blank
                                        + "' to have a value."
                                )
                            )
                        );
                        continue;
                    }
                    // A section's root folder may sit at the library root; deeper ones never do.
                    var name = FolderNames.Clean(raw, parent == null);
                    if (name == null)
                    {
                        // Skip this folder and everything below it; the rest of the plan goes on.
                        waits.Add(
                            new FolderWait(
                                section.Key,
                                node.Key,
                                null,
                                Notice(
                                    notices,
                                    "Folder '"
                                        + section.Key
                                        + "/"
                                        + node.Key
                                        + "' is waiting for a usable name: '"
                                        + raw
                                        + "' has no characters SharePoint accepts."
                                )
                            )
                        );
                        continue;
                    }
                    if (name != raw)
                        Notice(
                            notices,
                            "Folder name '"
                                + raw
                                + "' was adjusted to '"
                                + name
                                + "' for SharePoint."
                        );
                    if (names.TryGetValue(name, out var first))
                    {
                        // The first sibling in Order, then Key, keeps the name. This one and
                        // everything below it wait until a change to the record tells them apart.
                        waits.Add(
                            new FolderWait(
                                section.Key,
                                node.Key,
                                null,
                                Notice(
                                    notices,
                                    "Folder '"
                                        + section.Key
                                        + "/"
                                        + node.Key
                                        + "' has the same name '"
                                        + name
                                        + "' as '"
                                        + section.Key
                                        + "/"
                                        + first
                                        + "'; it waits until the names differ."
                                )
                            )
                        );
                        continue;
                    }
                    names.Add(name, node.Key);
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
        return new FolderPlan(result, notices, waits);
    }
}
