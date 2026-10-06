using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Plugins;

public sealed class ApiWriteService : IOrganizationService
{
    private readonly IOrganizationService inner;
    private readonly IPluginExecutionContext owner;
    private static readonly string[] Apis =
    {
        "asx_DocumentWorker",
        "asx_RecoverWorker",
        "asx_SecurityAdmin",
        "asx_RuntimeAdmin",
        "asx_ManageWork",
        "asx_CatalogAdmin",
    };

    public ApiWriteService(IOrganizationService inner, IPluginExecutionContext owner)
    {
        this.inner = inner;
        this.owner = owner;
    }

    /// <summary>
    /// Hashes the API request, caller, correlation, message, and target into a write-context tag.
    /// </summary>
    /// <param name="api">The product API execution context containing the serialized request.</param>
    /// <param name="message">The Create, Update, or Delete operation being tagged.</param>
    /// <param name="target">The entity identity receiving the write.</param>
    /// <returns>The prefixed SHA-256 write-context tag.</returns>
    public static string Tag(IPluginExecutionContext api, string message, Entity target)
    {
        var request =
            api.InputParameters?.Contains("Request") == true
                ? api.InputParameters["Request"] as string
                : null;
        if (string.IsNullOrWhiteSpace(request))
            throw new InvalidPluginExecutionException("Product API request context is missing.");
        string input =
            api.MessageName
            + "\n"
            + api.UserId.ToString("D")
            + "\n"
            + api.CorrelationId.ToString("D")
            + "\n"
            + message
            + "\n"
            + target.LogicalName
            + "\n"
            + target.Id.ToString("D")
            + "\n"
            + request;
        using (var hash = SHA256.Create())
            return "Ascentix.Documents.ApiWrite.v1:"
                + BitConverter
                    .ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(input)))
                    .Replace("-", "");
    }

    public static Entity? SingleTarget(IPluginExecutionContext context)
    {
        if (
            (context.MessageName == "Create" || context.MessageName == "Update")
            && context.InputParameters?.Contains("Target") == true
        )
            return context.InputParameters["Target"] as Entity;
        if (
            context.MessageName == "Delete"
            && context.InputParameters?.Contains("Target") == true
            && context.InputParameters["Target"] is EntityReference reference
        )
            return new Entity(reference.LogicalName, reference.Id);
        return null;
    }

    /// <summary>
    /// Verifies the write tag against the transactional product API ancestry and target.
    /// </summary>
    /// <param name="context">The pre-operation row guard's execution context.</param>
    /// <param name="catalog">Whether to exclude recovery API authorization for a catalog write.</param>
    /// <returns>Whether the target and write tag match an authorized product API ancestor.</returns>
    public static bool Authorizes(IPluginExecutionContext context, bool catalog = false)
    {
        if (
            context.Stage != 20
            || !context.IsInTransaction
            || context.UserId == Guid.Empty
            || context.CorrelationId == Guid.Empty
        )
            return false;
        var target = SingleTarget(context);
        if (
            target == null
            || target.Id == Guid.Empty
            || target.LogicalName != context.PrimaryEntityName
        )
            return false;
        string message =
            context.MessageName == "Delete" ? "Delete"
            : context.MessageName.StartsWith("Create", StringComparison.Ordinal) ? "Create"
            : "Update";
        string? tag =
            context.SharedVariables?.Contains("tag") == true
                ? context.SharedVariables["tag"] as string
                : null;
        var frame = context.ParentContext;
        for (int depth = 0; frame != null && depth < 3; depth++, frame = frame.ParentContext)
        {
            if (
                frame.UserId != context.UserId
                || frame.CorrelationId != context.CorrelationId
                || !frame.IsInTransaction
                || frame.Stage != 30
                || frame.Mode != 0
            )
                return false;
            if (Apis.Contains(frame.MessageName))
            {
                if (catalog && frame.MessageName == "asx_RecoverWorker")
                    return false;
                try
                {
                    return tag != null && tag == Tag(frame, message, target);
                }
                catch
                {
                    return false;
                }
            }
            if (
                (frame.MessageName != message && frame.MessageName != message + "Multiple")
                || frame.PrimaryEntityName != context.PrimaryEntityName
            )
                return false;
            if (tag == null && frame.SharedVariables?.Contains("tag") == true)
                tag = frame.SharedVariables["tag"] as string;
        }
        return false;
    }

    public static bool AuthorizesCatalogIdentityUpgrade(IPluginExecutionContext context)
    {
        if (!Authorizes(context, true))
            return false;
        var frame = context.ParentContext;
        for (int depth = 0; frame != null && depth < 3; depth++, frame = frame.ParentContext)
        {
            if (frame.MessageName != "asx_CatalogAdmin")
                continue;
            try
            {
                return Ascentix
                        .Documents.Dataverse.JsonWire.Read<Ascentix.Documents.Dataverse.CatalogRequest>(
                            (string)frame.InputParameters["Request"]
                        )
                        .Command == "CompleteSiteIdentity";
            }
            catch
            {
                return false;
            }
        }
        return false;
    }

    /// <summary>
    /// A site or library address written by a verified re-point: the worker completing a
    /// re-point probe, which only the catalog API queues. Identity fields stay immutable.
    /// </summary>
    public static bool AuthorizesCatalogRepoint(IPluginExecutionContext context) =>
        AuthorizesCatalogCommand(
            context,
            "asx_DocumentWorker",
            request =>
            {
                var work =
                    Ascentix.Documents.Dataverse.JsonWire.Read<Ascentix.Documents.Dataverse.WorkerRequest>(
                        request
                    );
                return work.Command == "Complete"
                    && work.Key?.StartsWith(
                        Ascentix.Documents.Dataverse.CatalogAdministration.RepointPrefix,
                        StringComparison.Ordinal
                    ) == true;
            }
        );

    /// <summary>A site or library deleted by the catalog API's Remove command.</summary>
    public static bool AuthorizesCatalogRemoval(IPluginExecutionContext context) =>
        AuthorizesCatalogCommand(context, "asx_CatalogAdmin", RemovalRequest);

    private static bool RemovalRequest(string request)
    {
        var command = Ascentix
            .Documents.Dataverse.JsonWire.Read<Ascentix.Documents.Dataverse.CatalogRequest>(request)
            .Command;
        return command == "RemoveLibrary" || command == "RemoveSite";
    }

    /// <summary>Whether the nearest product API frame is the named API with a matching request.</summary>
    private static bool AuthorizesCatalogCommand(
        IPluginExecutionContext context,
        string api,
        Func<string, bool> matches
    )
    {
        if (!Authorizes(context, true))
            return false;
        var frame = context.ParentContext;
        for (int depth = 0; frame != null && depth < 3; depth++, frame = frame.ParentContext)
        {
            if (!Apis.Contains(frame.MessageName))
                continue;
            if (frame.MessageName != api)
                return false;
            try
            {
                return matches((string)frame.InputParameters["Request"]);
            }
            catch
            {
                return false;
            }
        }
        return false;
    }

    private void Mark(OrganizationRequest request, string message, Entity target)
    {
        if (
            !Apis.Contains(owner.MessageName)
            || owner.Stage != 30
            || owner.Mode != 0
            || !owner.IsInTransaction
            || owner.UserId == Guid.Empty
            || owner.CorrelationId == Guid.Empty
            || owner.SharedVariables?.Contains(DocumentWorkerApi.InternalWrite) != true
            || !(owner.SharedVariables[DocumentWorkerApi.InternalWrite] is bool flag)
            || !flag
        )
            throw new InvalidPluginExecutionException(
                "Validated product API write context is required."
            );
        if (target.Id == Guid.Empty)
            throw new InvalidPluginExecutionException("Stable write target identity is required.");
        if (request.Parameters.Contains("tag"))
            throw new InvalidPluginExecutionException(
                "Product write transport does not accept a supplied tag."
            );
        request["tag"] = Tag(owner, message, target);
    }

    public Guid Create(Entity entity)
    {
        if (entity.Id == Guid.Empty)
            entity.Id = Guid.NewGuid();
        var request = new CreateRequest { Target = entity };
        Mark(request, "Create", entity);
        return ((CreateResponse)inner.Execute(request)).id;
    }

    public void Update(Entity entity)
    {
        var request = new UpdateRequest { Target = entity };
        Mark(request, "Update", entity);
        inner.Execute(request);
    }

    public OrganizationResponse Execute(OrganizationRequest request)
    {
        if (request is UpdateRequest update)
            Mark(request, "Update", update.Target);
        else if (request is DeleteRequest delete)
        {
            if (!RetentionRequest(owner) && !CatalogRemovalRequest(owner))
                throw new InvalidPluginExecutionException(
                    "Only history retention and catalog removal may delete through this transport."
                );
            Mark(request, "Delete", new Entity(delete.Target.LogicalName, delete.Target.Id));
        }
        else if (request is CreateRequest create)
        {
            if (create.Target.Id == Guid.Empty)
                create.Target.Id = Guid.NewGuid();
            Mark(request, "Create", create.Target);
        }
        return inner.Execute(request);
    }

    private static bool CatalogRemovalRequest(IPluginExecutionContext context)
    {
        if (
            context.MessageName != "asx_CatalogAdmin"
            || context.InputParameters?.Contains("Request") != true
        )
            return false;
        try
        {
            return RemovalRequest((string)context.InputParameters["Request"]);
        }
        catch
        {
            return false;
        }
    }

    public static bool RetentionRequest(IPluginExecutionContext context)
    {
        if (
            context.MessageName != "asx_DocumentWorker"
            || context.InputParameters?.Contains("Request") != true
        )
            return false;
        try
        {
            return Ascentix
                    .Documents.Dataverse.JsonWire.Read<Ascentix.Documents.Dataverse.WorkerRequest>(
                        (string)context.InputParameters["Request"]
                    )
                    .Command == "PurgeHistory";
        }
        catch
        {
            return false;
        }
    }

    public Entity Retrieve(string name, Guid id, ColumnSet columns) =>
        inner.Retrieve(name, id, columns);

    public EntityCollection RetrieveMultiple(QueryBase query) => inner.RetrieveMultiple(query);

    public void Delete(string name, Guid id) => inner.Delete(name, id);

    public void Associate(
        string name,
        Guid id,
        Relationship relationship,
        EntityReferenceCollection rows
    ) => inner.Associate(name, id, relationship, rows);

    public void Disassociate(
        string name,
        Guid id,
        Relationship relationship,
        EntityReferenceCollection rows
    ) => inner.Disassociate(name, id, relationship, rows);
}
