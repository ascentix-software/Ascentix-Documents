using System;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;

namespace Ascentix.Documents.Plugins;

public sealed class SecurityAdminApi : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        var service = new ApiWriteService(
            (
                (IOrganizationServiceFactory)
                    provider.GetService(typeof(IOrganizationServiceFactory))
            ).CreateOrganizationService(context.UserId),
            context
        );
        context.SharedVariables[DocumentWorkerApi.InternalWrite] = true;
        var request = JsonWire.Read<SecurityRequest>((string)context.InputParameters["Request"]);
        var result = new SecurityAdministration(service).Execute(request, context.IsInTransaction);
        if (request.Command == "RegisterTeam")
            TeamMembershipInvalidationPlugin.Enqueue(
                new DocumentStore(service),
                request.TeamId,
                context.CorrelationId,
                "Registration"
            );
        context.OutputParameters["Result"] = JsonWire.Write(result);
    }
}

public sealed class RuntimeAdminApi : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        var service = new ApiWriteService(
            (
                (IOrganizationServiceFactory)
                    provider.GetService(typeof(IOrganizationServiceFactory))
            ).CreateOrganizationService(context.UserId),
            context
        );
        context.SharedVariables[DocumentWorkerApi.InternalWrite] = true;
        context.OutputParameters["Result"] = JsonWire.Write(
            RuntimeAdministration.Execute(
                service,
                JsonWire.Read<RuntimeRequest>((string)context.InputParameters["Request"]),
                context.IsInTransaction,
                context.UserId
            )
        );
    }
}

public sealed class ManageWorkApi : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        var service = new ApiWriteService(
            (
                (IOrganizationServiceFactory)
                    provider.GetService(typeof(IOrganizationServiceFactory))
            ).CreateOrganizationService(context.UserId),
            context
        );
        var request = JsonWire.Read<WorkerRequest>((string)context.InputParameters["Request"]);
        if (request.Command == "InspectRecord")
        {
            context.OutputParameters["Result"] = JsonWire.Write(
                RecordInspection.Read(service, request, RuntimeProfile.Read(service).Tables)
            );
            return;
        }
        if (request.Command == "Inspect")
        {
            var store = new DocumentStore(service);
            var operation = store.Require<OperationDocument>("asx_operation", request.Key).Value;
            var claim = store
                .Find<DispatcherDocument>(
                    "asx_claim",
                    WorkCoordination.Operation(service, request.Key)
                )
                ?.Value;
            bool owns = claim?.OperationKey == request.Key;
            context.OutputParameters["Result"] = JsonWire.Write(
                new WorkerResult
                {
                    Key = request.Key,
                    Status = operation.Status,
                    Notices =
                        operation.ErrorCode == null
                            ? Array.Empty<string>()
                            : new[] { operation.ErrorCode },
                    RunId = owns ? claim!.RunId : null,
                    Token = owns ? claim!.Token : Guid.Empty,
                    LeaseUntilUtc = owns ? claim!.LeaseUntilUtc : (DateTime?)null,
                }
            );
            return;
        }
        if (request.Command == "PreviewBatch" || request.Command == "QueueBatch")
        {
            context.SharedVariables[DocumentWorkerApi.InternalWrite] = true;
            context.OutputParameters["Result"] = JsonWire.Write(
                new BatchReplan(service, RuntimeProfile.Read(service).Tables).Execute(
                    request,
                    context.IsInTransaction
                )
            );
            return;
        }
        if (
            request.Command != "Retry"
            && request.Command != "Cancel"
            && request.Command != "Replan"
            && request.Command != "Queue"
        )
            throw new InvalidPluginExecutionException("Unsupported operator action.");
        context.SharedVariables[DocumentWorkerApi.InternalWrite] = true;
        context.OutputParameters["Result"] = JsonWire.Write(
            request.Key?.StartsWith("librarycreate:", StringComparison.Ordinal) == true
                ? new LibraryProvisioning(service).Execute(request, context.IsInTransaction)
            : request.Key?.StartsWith("catalogprobe:", StringComparison.Ordinal) == true
                ? new CatalogWorker(service).Execute(request, context.IsInTransaction)
            : request.Key?.StartsWith("policywork:", StringComparison.Ordinal) == true
                ? new SecurityWorker(service).Execute(request, context.IsInTransaction)
            : new WorkerCoordinator(
                service,
                allowedTables: RuntimeProfile.Read(service).Tables
            ).Execute(request, context.IsInTransaction)
        );
    }
}

public sealed class CatalogGuard : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        bool transported = ApiWriteService.Authorizes(context, true);
        string message = context.MessageName;
        if (
            context.Stage != 20
            || !context.IsInTransaction
            || (message != "Create" && message != "Update")
        )
            throw new InvalidPluginExecutionException(
                "Catalog deletion and nontransactional writes are unsupported."
            );
        var parent = context.ParentContext;
        bool marked =
            parent?.SharedVariables.Contains(DocumentWorkerApi.InternalWrite) == true
            && parent.SharedVariables[DocumentWorkerApi.InternalWrite] is bool flag
            && flag
            && parent.UserId == context.UserId;
        bool trusted =
            marked
            && (
                parent?.MessageName == "asx_SecurityAdmin"
                || parent?.MessageName == "asx_RuntimeAdmin"
                || parent?.MessageName == "asx_DocumentWorker"
                || parent?.MessageName == "asx_ManageWork"
                || parent?.MessageName == "asx_CatalogAdmin"
            );
        trusted = trusted || transported;
        var target = ApiWriteService.SingleTarget(context);
        if (context.PrimaryEntityName == "asx_site" || context.PrimaryEntityName == "asx_library")
        {
            if (message == "Update" && target != null)
            {
                var primary = context.PrimaryEntityName + "id";
                if (
                    target.Contains(primary)
                    && target.GetAttributeValue<Guid>(primary) != target.Id
                )
                    throw new InvalidPluginExecutionException("Catalog record identity differs.");
                if (
                    (target.Contains("asx_collectionid") || target.Contains("asx_identity"))
                    && !ApiWriteService.AuthorizesCatalogIdentityUpgrade(context)
                )
                    throw new InvalidPluginExecutionException(
                        "Site identity fields require the verified identity upgrade operation."
                    );
                foreach (
                    var field in context.PrimaryEntityName == "asx_site"
                        ? new[] { "asx_nativeid", "asx_webid", "asx_url" }
                        : new[]
                        {
                            "asx_siteid",
                            "asx_listid",
                            "asx_entryid",
                            "asx_entryurl",
                            "asx_nativeparentid",
                        }
                )
                    if (target.Contains(field))
                        throw new InvalidPluginExecutionException(
                            "Approved catalog physical identities are immutable; register a new destination for changes."
                        );
                foreach (
                    var field in new[] { "asx_policyrevision", "asx_policyapplied", "asx_aclhash" }
                )
                    if (target.Contains(field) && !trusted)
                        throw new InvalidPluginExecutionException(
                            "Policy observations are server-owned."
                        );
            }
            if (
                message == "Create"
                && target?.GetAttributeValue<bool>("asx_policyapplied") == true
                && !trusted
            )
                throw new InvalidPluginExecutionException(
                    "New libraries require worker-verified policy application."
                );
            if (!trusted)
                throw new InvalidPluginExecutionException(
                    "Catalog changes require fresh observation-based approval through the catalog API."
                );
            return;
        }
        if (
            context.PrimaryEntityName == "asx_teamregistration"
            && !trusted
            && message == "Update"
            && context.Mode == 0
            && context.UserId != Guid.Empty
            && target?.GetAttributeValue<string>("asx_status") == "Revoking"
        )
        {
            var service = (
                (IOrganizationServiceFactory)
                    provider.GetService(typeof(IOrganizationServiceFactory))
            ).CreateOrganizationService(context.UserId);
            if (RuntimeProfile.Read(service).WorkerId != context.UserId)
                throw new InvalidPluginExecutionException(
                    "Team retirement requires the configured worker."
                );
            var registration = JsonWire.Read<TeamRegistration>(
                target.GetAttributeValue<string>("asx_payload")
            );
            if (
                registration == null
                || registration.SchemaVersion != 1
                || registration.Enabled
                || registration.Status != "Revoking"
                || registration.TeamId == Guid.Empty
                || registration.Key != "team:" + registration.TeamId.ToString("N")
                || target.Id != DocumentStore.StableId("asx_teamregistration:" + registration.Key)
            )
                throw new InvalidPluginExecutionException(
                    "Team retirement identity or state is inconsistent."
                );
            return;
        }
        if (
            context.PrimaryEntityName == "asx_runtime"
            && message == "Create"
            && target != null
            && !target.GetAttributeValue<bool>("asx_enabled")
            && !target.GetAttributeValue<bool>("asx_processrecordupdates")
        )
            return;
        if (!trusted)
            throw new InvalidPluginExecutionException(
                "Security/runtime configuration is writable only through guarded product APIs."
            );
    }
}

public sealed class TeamRetirementPlugin : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        if (
            context.Stage != 40
            || context.Mode != 0
            || !context.IsInTransaction
            || context.MessageName != "Delete"
            || context.PrimaryEntityName != "team"
            || context.PrimaryEntityId == Guid.Empty
        )
            throw new InvalidPluginExecutionException("Invalid team retirement registration.");
        var service = (
            (IOrganizationServiceFactory)provider.GetService(typeof(IOrganizationServiceFactory))
        ).CreateOrganizationService(context.UserId);
        if (RuntimeProfile.Read(service).WorkerId != context.UserId)
            throw new InvalidPluginExecutionException(
                "Team lifecycle identity differs from runtime."
            );
        var store = new DocumentStore(service);
        var registration = store.Find<TeamRegistration>(
            "asx_teamregistration",
            "team:" + context.PrimaryEntityId.ToString("N")
        );
        if (registration == null)
            return;
        registration.Value.Enabled = false;
        registration.Value.Status = "Revoking";
        store.Save(registration);
        TeamMembershipInvalidationPlugin.Enqueue(
            store,
            context.PrimaryEntityId,
            context.CorrelationId,
            "Delete"
        );
    }
}

public sealed class CatalogAdminApi : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        // Marks the API context before requesting its organization service.
        context.SharedVariables[DocumentWorkerApi.InternalWrite] = true;
        var service = (
            (IOrganizationServiceFactory)provider.GetService(typeof(IOrganizationServiceFactory))
        ).CreateOrganizationService(context.UserId);
        context.OutputParameters["Result"] = JsonWire.Write(
            new CatalogAdministration(new ApiWriteService(service, context)).Execute(
                JsonWire.Read<CatalogRequest>((string)context.InputParameters["Request"]),
                context.IsInTransaction
            )
        );
    }
}
