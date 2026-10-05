using System;
using System.Linq;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;
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
    public static T Body<T>(WorkerRequest request)
        where T : class
    {
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

    public static string AclHash(ODataRows<AclAssignment> assignments)
    {
        if (
            assignments == null
            || assignments.Next != null
            || assignments.Rows == null
            || assignments.Rows.Length > 1000
            || assignments.Rows.Select(a => a.Member?.Id).Distinct().Count()
                != assignments.Rows.Length
        )
            throw new EvaluationBlockedException("Incomplete or duplicate ACL snapshot.");
        var canonical = assignments
            .Rows.SelectMany(assignment =>
            {
                if (
                    assignment.Member == null
                    || assignment.Member.Id <= 0
                    || assignment.Member.Type <= 0
                    || assignment.Roles == null
                    || assignment.Roles.Next != null
                    || assignment.Roles.Rows == null
                    || assignment.Roles.Rows.Length == 0
                    || assignment.Roles.Rows.Length > 100
                    || assignment.Roles.Rows.Select(r => r.Id).Distinct().Count()
                        != assignment.Roles.Rows.Length
                )
                    throw new EvaluationBlockedException("Incomplete principal/role bindings.");
                return assignment.Roles.Rows.Select(role =>
                {
                    if (
                        role.Id <= 0
                        || role.Permissions == null
                        || !Regex.IsMatch(role.Permissions.High ?? "", "^[0-9]{1,20}$")
                        || !Regex.IsMatch(role.Permissions.Low ?? "", "^[0-9]{1,20}$")
                    )
                        throw new EvaluationBlockedException("Role permission mask missing.");
                    return assignment.Member.Id
                        + ":"
                        + assignment.Member.Type
                        + ":"
                        + role.Id
                        + ":"
                        + role.Permissions.High
                        + ":"
                        + role.Permissions.Low;
                });
            })
            .OrderBy(v => v, StringComparer.Ordinal);
        var joined = "acl-v1|" + string.Join("|", canonical);
        // Hashes the complete serialized ACL as UTF-8 bytes.
        using (var sha = System.Security.Cryptography.SHA256.Create())
            return BitConverter
                .ToString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(joined)))
                .Replace("-", "")
                .ToLowerInvariant();
    }

    public static ItemObservation? Find(WorkerRequest request, string expectedPath)
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
        Domain.FolderNames.Validate(item.Name);
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
