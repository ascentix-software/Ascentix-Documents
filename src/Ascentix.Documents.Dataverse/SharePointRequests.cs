using System;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Domain;

namespace Ascentix.Documents.Dataverse;

public sealed class SharePointTarget
{
    public Uri Web { get; }
    public Guid WebId { get; }
    public Guid ListId { get; }
    public string EntryPath { get; }

    public SharePointTarget(
        string approvedWebUrl,
        Guid webId,
        Guid listId,
        string approvedEntryPath
    )
    {
        if (
            !Uri.TryCreate(approvedWebUrl, UriKind.Absolute, out var web)
            || web.Scheme != "https"
            || web.UserInfo.Length != 0
            || web.Query.Length != 0
            || web.Fragment.Length != 0
            || webId == Guid.Empty
            || listId == Guid.Empty
        )
            throw new EvaluationBlockedException(
                "Approved physical site/list identities are required."
            );
        var webPath = Uri.UnescapeDataString(web.AbsolutePath).TrimEnd('/');
        if (
            string.IsNullOrEmpty(approvedEntryPath)
            || !approvedEntryPath.StartsWith(webPath + "/", StringComparison.OrdinalIgnoreCase)
            || approvedEntryPath.Contains("\\")
            || Regex.IsMatch(approvedEntryPath, @"(^|/)\.{1,2}(/|$)")
            || approvedEntryPath.Contains("//")
            || approvedEntryPath.EndsWith("/", StringComparison.Ordinal)
        )
            throw new EvaluationBlockedException(
                "Entry path is outside the approved web or is not canonical."
            );
        Web = web;
        WebId = webId;
        ListId = listId;
        EntryPath = approvedEntryPath;
    }

    internal void ValidateParent(string parent)
    {
        if (parent != EntryPath && !parent.StartsWith(EntryPath + "/", StringComparison.Ordinal))
            throw new EvaluationBlockedException("Parent is outside the approved entry.");
        // The entry may be the library root, so only its first segment may sit at the root;
        // SharePoint reserves "Forms" only there.
        bool first = true;
        foreach (var segment in parent.Substring(EntryPath.Length).Split('/'))
            if (segment.Length > 0)
            {
                FolderNames.Validate(segment, first);
                first = false;
            }
        if (parent.Contains("//") || parent.EndsWith("/", StringComparison.Ordinal))
            throw new EvaluationBlockedException("Parent path is not canonical.");
    }
}

[DataContract]
public sealed class HttpIntent
{
    [DataMember]
    public string Method { get; set; } = "GET";

    [DataMember]
    public string RelativeUri { get; set; } = "";

    [DataMember]
    public string? Body { get; set; }

    [DataMember]
    public string RetryPolicy { get; set; } = "None";
}

public static class SharePointRequests
{
    private static string List(SharePointTarget target) =>
        "_api/web/lists(guid'" + target.ListId.ToString("D") + "')";

    public static HttpIntent FindFolder(SharePointTarget target, string parent, string name)
    {
        target.ValidateParent(parent);
        FolderNames.Validate(name, parent == target.EntryPath);
        string path = parent + "/" + name;
        return new HttpIntent
        {
            RelativeUri =
                "_api/web/GetFolderByServerRelativePath(decodedUrl='"
                + Uri.EscapeDataString(path.Replace("'", "''"))
                + "')?$select=Exists,UniqueId,ServerRelativeUrl,ListItemAllFields/Id,ListItemAllFields/UniqueId,ListItemAllFields/FileLeafRef,ListItemAllFields/FileRef,ListItemAllFields/FSObjType&$expand=ListItemAllFields",
        };
    }

    public static HttpIntent CreateFolder(SharePointTarget target, string parent, string name)
    {
        target.ValidateParent(parent);
        FolderNames.Validate(name, parent == target.EntryPath);
        // Folder path is decoded ResourcePath input; never URL-escape the field value itself.
        var origin = target.Web.GetLeftPart(UriPartial.Authority);
        return new HttpIntent
        {
            Method = "POST",
            RelativeUri = List(target) + "/AddValidateUpdateItemUsingPath",
            Body = JsonWire.Write(
                new FolderCreateBody
                {
                    Info = new FolderCreateInfo
                    {
                        Path = new ResourcePath { Url = origin + parent },
                        Name = new ResourcePath { Url = name },
                        Type = 1,
                    },
                    Values = new[]
                    {
                        new FolderField { Field = "FileLeafRef", Value = name },
                    },
                }
            ),
        };
    }

    public static HttpIntent ReadGroup(int groupId)
    {
        Positive(groupId);
        return new HttpIntent
        {
            RelativeUri =
                "_api/web/sitegroups("
                + groupId
                + ")?$select=Id,Title,Description,Users/Id,Users/LoginName&$expand=Users",
        };
    }

    public static HttpIntent AddMember(int ownedGroupId, string approvedLogin)
    {
        Positive(ownedGroupId);
        if (
            string.IsNullOrWhiteSpace(approvedLogin)
            || approvedLogin.Length > 500
            || approvedLogin.IndexOfAny(new[] { '\r', '\n' }) >= 0
        )
            throw new EvaluationBlockedException("Approved normalized login required.");
        return new HttpIntent
        {
            Method = "POST",
            RelativeUri = "_api/web/sitegroups(" + ownedGroupId + ")/users",
            Body = JsonWire.Write(new AddMemberBody { Login = approvedLogin }),
        };
    }

    public static HttpIntent RemoveMember(int ownedGroupId, int approvedUserId)
    {
        Positive(ownedGroupId);
        Positive(approvedUserId);
        return new HttpIntent
        {
            Method = "POST",
            RelativeUri =
                "_api/web/sitegroups("
                + ownedGroupId
                + ")/users/removebyid("
                + approvedUserId
                + ")",
        };
    }

    public static HttpIntent ChangeOwnedGrant(
        SharePointTarget target,
        int ownedGroupId,
        int resolvedRoleId,
        bool add
    )
    {
        Positive(ownedGroupId);
        Positive(resolvedRoleId);
        return new HttpIntent
        {
            Method = "POST",
            RelativeUri =
                List(target)
                + "/roleassignments/"
                + (add ? "addroleassignment" : "removeroleassignment")
                + "(principalid="
                + ownedGroupId
                + ",roledefid="
                + resolvedRoleId
                + ")",
        };
    }

    private static void Positive(int id)
    {
        if (id <= 0)
            throw new EvaluationBlockedException("Verified positive SharePoint ID required.");
    }
}

[DataContract]
public sealed class FolderCreateBody
{
    [DataMember(Name = "listItemCreateInfo")]
    public FolderCreateInfo Info { get; set; } = new FolderCreateInfo();

    [DataMember(Name = "formValues")]
    public FolderField[] Values { get; set; } = Array.Empty<FolderField>();

    [DataMember(Name = "bNewDocumentUpdate")]
    public bool NewDocument { get; set; }
}

[DataContract]
public sealed class FolderCreateInfo
{
    [DataMember(Name = "FolderPath")]
    public ResourcePath Path { get; set; } = new ResourcePath();

    [DataMember(Name = "LeafName")]
    public ResourcePath Name { get; set; } = new ResourcePath();

    [DataMember(Name = "UnderlyingObjectType")]
    public int Type { get; set; }
}

[DataContract]
public sealed class ResourcePath
{
    [DataMember(Name = "DecodedUrl")]
    public string Url { get; set; } = "";
}

[DataContract]
public sealed class FolderField
{
    [DataMember(Name = "FieldName")]
    public string Field { get; set; } = "";

    [DataMember(Name = "FieldValue")]
    public string Value { get; set; } = "";
}

[DataContract]
public sealed class AddMemberBody
{
    [DataMember(Name = "LoginName")]
    public string Login { get; set; } = "";

    [DataMember(Name = "__metadata")]
    public MemberMetadata Metadata { get; set; } = new MemberMetadata();
}

[DataContract]
public sealed class MemberMetadata
{
    [DataMember(Name = "type")]
    public string Type { get; set; } = "SP.User";
}
