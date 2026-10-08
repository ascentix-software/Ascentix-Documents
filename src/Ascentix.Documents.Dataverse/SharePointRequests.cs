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

    /// <summary>
    /// A read of the folder or file at a server-relative path, with the path passed as an OData
    /// parameter alias in the query string and never in the URL path. The HTTP with Microsoft
    /// Entra ID connector refuses a request whose URL path exceeds its maxUrlLength (401 "The
    /// length of the URL for this request exceeds the configured maxUrlLength value."). That is
    /// the ASP.NET httpRuntime setting, which applies to the URL path only (default 260
    /// characters; the connector's configured value is not published):
    /// https://learn.microsoft.com/dotnet/api/system.web.configuration.httpruntimesection.maxurllength
    /// SharePoint documents parameter aliases for method parameters ("Using parameter aliases in
    /// REST service calls"):
    /// https://learn.microsoft.com/sharepoint/dev/sp-add-ins/determine-sharepoint-rest-service-endpoint-uris
    /// The URL path is then the same short text for every folder, and a 400-character path,
    /// escaped, stays far inside the 16,384-character request URL limit of Power Automate:
    /// https://learn.microsoft.com/power-automate/limits-and-config
    /// </summary>
    /// <param name="function">GetFolderByServerRelativePath or GetFileByServerRelativePath.</param>
    /// <param name="path">The decoded server-relative path.</param>
    /// <param name="query">The rest of the query string, such as "$select=...".</param>
    /// <summary>The notice for a read whose address the connector would refuse.</summary>
    public const string AddressTooLongNotice =
        "This SharePoint address is too long for the HTTP connector: its query string would pass 2,048 characters. Nothing was sent; shorten the names.";

    public static string ByPath(string function, string path, string query)
    {
        string alias = SharePointAddress.Alias(path) + "&" + query;
        // Never sent when the connector would refuse it; planning keeps folders inside this
        // (FolderPlanner), so only an address an admin chose can reach it.
        if (alias.Length > SharePointAddress.MaxQueryString)
            throw new EvaluationBlockedException(AddressTooLongNotice);
        return "_api/web/" + function + "(decodedUrl=@p)?" + alias;
    }

    /// <summary>A read of the folder with this unique ID: the same short address for any path.</summary>
    public static string ById(Guid id, string query)
    {
        if (id == Guid.Empty)
            throw new EvaluationBlockedException("Folder ID required.");
        return "_api/web/GetFolderById('" + id.ToString("D") + "')?" + query;
    }

    public static HttpIntent FindFolder(SharePointTarget target, string parent, string name)
    {
        target.ValidateParent(parent);
        FolderNames.Validate(name, parent == target.EntryPath);
        return new HttpIntent
        {
            RelativeUri = ByPath(
                "GetFolderByServerRelativePath",
                parent + "/" + name,
                SharePointAddress.FolderLookup
            ),
        };
    }

    /// <summary>
    /// Reads a folder Documents already found by its unique ID, with the same fields as a
    /// lookup by path, so a long path never makes the address longer.
    /// </summary>
    public static HttpIntent FindFolder(Guid id) =>
        new HttpIntent { RelativeUri = ById(id, SharePointAddress.FolderLookup) };

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
        Login(approvedLogin);
        return new HttpIntent
        {
            Method = "POST",
            RelativeUri = "_api/web/sitegroups(" + ownedGroupId + ")/users",
            Body = JsonWire.Write(new AddMemberBody { Login = approvedLogin }),
        };
    }

    /// <summary>
    /// Resolves a group claim to a site principal with SPWeb.EnsureUser so it can be added to the
    /// Documents group. POST _api/web/ensureuser with { "logonName": ... } is documented in the
    /// Webs REST API reference:
    /// https://learn.microsoft.com/previous-versions/office/developer/sharepoint-rest-reference/dn499819(v=office.15)
    /// EnsureUser returns the principal already on the site when there is one, so repeating it
    /// is harmless.
    /// </summary>
    public static HttpIntent EnsurePrincipal(string approvedLogin)
    {
        Login(approvedLogin);
        return new HttpIntent
        {
            Method = "POST",
            RelativeUri = "_api/web/ensureuser",
            Body = JsonWire.Write(new EnsureUserBody { Login = approvedLogin }),
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

    /// <summary>
    /// Stops the library inheriting its site's permissions. copyRoleAssignments=true keeps a copy
    /// of the site's current permissions as the library's starting point; clearSubscopes=false
    /// leaves folders and items with their own permissions as they are.
    /// </summary>
    public static HttpIntent BreakInheritance(SharePointTarget target) =>
        new HttpIntent
        {
            Method = "POST",
            RelativeUri =
                List(target)
                + "/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)",
        };

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

    private static void Login(string approvedLogin)
    {
        if (
            string.IsNullOrWhiteSpace(approvedLogin)
            || approvedLogin.Length > 500
            || approvedLogin.IndexOfAny(new[] { '\r', '\n' }) >= 0
        )
            throw new EvaluationBlockedException("Approved normalized login required.");
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
public sealed class EnsureUserBody
{
    [DataMember(Name = "logonName")]
    public string Login { get; set; } = "";
}

[DataContract]
public sealed class MemberMetadata
{
    [DataMember(Name = "type")]
    public string Type { get; set; } = "SP.User";
}
