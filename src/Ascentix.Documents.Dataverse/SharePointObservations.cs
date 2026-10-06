using System;
using System.Linq;
using System.Runtime.Serialization;
using Ascentix.Documents.Conditions;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class ODataEnvelope<T>
{
    [DataMember(Name = "d")]
    public T Data { get; set; } = default!;
}

[DataContract]
public sealed class ODataRows<T>
{
    [DataMember(Name = "results")]
    public T[] Rows { get; set; } = Array.Empty<T>();

    [DataMember(Name = "__next")]
    public string? Next { get; set; }
}

[DataContract]
public sealed class LibraryObservation
{
    [DataMember(Name = "Id")]
    public Guid Id { get; set; }

    [DataMember(Name = "RootFolder")]
    public FolderObservation Root { get; set; } = null!;
}

[DataContract]
public sealed class FolderObservation
{
    [DataMember(Name = "UniqueId")]
    public Guid Id { get; set; }

    [DataMember(Name = "ServerRelativeUrl")]
    public string Path { get; set; } = "";
}

[DataContract]
public sealed class FolderLookupObservation
{
    [DataMember]
    public bool? Exists { get; set; }

    [DataMember(Name = "UniqueId")]
    public Guid Id { get; set; }

    [DataMember(Name = "ServerRelativeUrl")]
    public string Path { get; set; } = "";

    [DataMember]
    public ItemObservation? ListItemAllFields { get; set; }
}

[DataContract]
public sealed class ItemObservation
{
    [DataMember(Name = "Id")]
    public int ItemId { get; set; }

    [DataMember(Name = "UniqueId")]
    public Guid Id { get; set; }

    [DataMember(Name = "FileLeafRef")]
    public string Name { get; set; } = "";

    [DataMember(Name = "FileRef")]
    public string Path { get; set; } = "";

    [DataMember(Name = "FSObjType")]
    public int? Type { get; set; }
}

[DataContract]
public sealed class SharePointErrorBody
{
    [DataMember(Name = "error")]
    public SharePointError? Verbose { get; set; }

    [DataMember(Name = "odata.error")]
    public SharePointError? Light { get; set; }
}

[DataContract]
public sealed class SharePointError
{
    [DataMember(Name = "code")]
    public string? Code { get; set; }

    [DataMember(Name = "message")]
    public SharePointErrorMessage? Message { get; set; }
}

[DataContract]
public sealed class SharePointErrorMessage
{
    [DataMember(Name = "value")]
    public string? Value { get; set; }
}

[DataContract]
public sealed class AclAssignment
{
    [DataMember(Name = "Member")]
    public AclMember Member { get; set; } = null!;

    [DataMember(Name = "RoleDefinitionBindings")]
    public ODataRows<AclRole> Roles { get; set; } = null!;
}

[DataContract]
public sealed class AclMember
{
    [DataMember(Name = "Id")]
    public int Id { get; set; }

    [DataMember(Name = "PrincipalType")]
    public int Type { get; set; }
}

[DataContract]
public sealed class AclRole
{
    [DataMember(Name = "Id")]
    public int Id { get; set; }

    /// <summary>SharePoint's RoleTypeKind; 1 is Limited Access, which SharePoint manages itself.</summary>
    [DataMember(Name = "RoleTypeKind")]
    public int Type { get; set; }

    [DataMember(Name = "BasePermissions")]
    public PermissionMask Permissions { get; set; } = null!;
}

[DataContract]
public sealed class PermissionMask
{
    [DataMember(Name = "High")]
    public string High { get; set; } = "";

    [DataMember(Name = "Low")]
    public string Low { get; set; } = "";
}

[DataContract]
public sealed class CreateObservation
{
    [DataMember(Name = "AddValidateUpdateItemUsingPath")]
    public ODataRows<FieldValidation> Fields { get; set; } = null!;
}

[DataContract]
public sealed class FieldValidation
{
    [DataMember(Name = "FieldName")]
    public string Field { get; set; } = "";

    [DataMember(Name = "HasException")]
    public bool? HasException { get; set; }
}

public static class SharePointObservations
{
    /// <summary>The notice for a request the HTTP connector refused for the length of its URL.</summary>
    public const string UrlTooLongNotice =
        "The HTTP connector refused the request because its URL is too long (maxUrlLength). SharePoint did not receive it.";

    /// <summary>
    /// Whether the HTTP connector itself refused the request for the length of its URL. It
    /// answers 401 with "The length of the URL for this request exceeds the configured
    /// maxUrlLength value." (or the query-string equivalent): a refusal of the request's shape,
    /// not of its sign-in or of what SharePoint holds, so it is never read as either.
    /// </summary>
    public static bool UrlTooLong(WorkerRequest request) =>
        request.HttpStatus >= 400
        && request.HttpStatus < 500
        && request.ResponseBody != null
        && (
            request.ResponseBody.IndexOf("maxUrlLength", StringComparison.OrdinalIgnoreCase) >= 0
            || request.ResponseBody.IndexOf(
                "maxQueryStringLength",
                StringComparison.OrdinalIgnoreCase
            ) >= 0
        );

    public static T Body<T>(WorkerRequest request)
        where T : class
    {
        if (UrlTooLong(request))
            throw new EvaluationBlockedException(UrlTooLongNotice);
        if (request.HttpStatus != 200 || request.ResponseBody == null)
            throw new EvaluationBlockedException(
                "Independent read requires a complete HTTP 200 response."
            );
        return JsonWire.Read<ODataEnvelope<T>>(request.ResponseBody).Data
            ?? throw new EvaluationBlockedException("Missing OData body.");
    }

    /// <summary>
    /// Reads SharePoint's error message from a rejected write, for an admin notice.
    /// </summary>
    /// <param name="request">The worker response carrying SharePoint's status and body.</param>
    /// <returns>SharePoint's message, or the HTTP status when the body has none.</returns>
    public static string ErrorMessage(WorkerRequest request)
    {
        string? message = null;
        try
        {
            var body = JsonWire.Read<SharePointErrorBody>(request.ResponseBody ?? "");
            message = (body.Verbose ?? body.Light)?.Message?.Value;
        }
        catch (Exception error)
            when (error is EvaluationBlockedException || error is SerializationException)
        {
            // Not the verbose or light OData error shape; the status code is reported instead.
        }
        message = new string((message ?? "").Where(c => !char.IsControl(c)).ToArray()).Trim();
        if (message.Length > 400)
            message = message.Substring(0, 400) + "...";
        return message.Length == 0 ? "HTTP " + request.HttpStatus + "." : message;
    }

    /// <summary>Reads the folder at the expected path, or null when SharePoint has none.</summary>
    /// <param name="request">The folder read's response.</param>
    /// <param name="expectedPath">The folder's full server-relative path.</param>
    /// <param name="atLibraryRoot">
    /// False when the folder is known to be below the library root, where "Forms" is allowed.
    /// </param>
    public static ItemObservation? Find(
        WorkerRequest request,
        string expectedPath,
        bool atLibraryRoot = true
    )
    {
        if (request.HttpStatus == 404)
            return null;
        var folder = Body<FolderLookupObservation>(request);
        if (folder.Exists == false && folder.ListItemAllFields == null)
            return null;
        var item = folder.ListItemAllFields;
        if (
            folder.Exists != true
            || item == null
            || folder.Id != item.Id
            || !string.Equals(folder.Path, expectedPath, StringComparison.OrdinalIgnoreCase)
        )
            throw new EvaluationBlockedException(
                "Folder existence and item identity must be explicit."
            );
        if (
            item.Id == Guid.Empty
            || item.ItemId <= 0
            || item.Type != 1
            || !string.Equals(item.Path, expectedPath, StringComparison.OrdinalIgnoreCase)
            || !item.Path.EndsWith("/" + item.Name, StringComparison.Ordinal)
        )
            throw new EvaluationBlockedException("Physical path/type observation differs.");
        Domain.FolderNames.Validate(item.Name, atLibraryRoot);
        return item;
    }

    public static void CreateSucceeded(WorkerRequest request)
    {
        if (request.HttpStatus < 200 || request.HttpStatus >= 300 || request.ResponseBody == null)
            throw new EvaluationBlockedException("Folder create did not return successful HTTP.");
        var result = JsonWire.Read<ODataEnvelope<CreateObservation>>(request.ResponseBody).Data;
        if (
            result?.Fields?.Rows == null
            || result.Fields.Next != null
            || result.Fields.Rows.Length < 1
            || result.Fields.Rows.Any(f => f.HasException != false)
            || !result.Fields.Rows.Any(f => f.Field == "FileLeafRef")
        )
            throw new EvaluationBlockedException(
                "Folder create returned missing or failed field validation."
            );
    }
}
