using System;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;

namespace Ascentix.Documents.Plugins;

public sealed class DocumentWorkerApi : IPlugin
{
    public const string InternalWrite = "Ascentix.Documents.InternalStateWrite";

    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        if (!context.IsInTransaction)
            throw new InvalidPluginExecutionException("Worker requires an ambient transaction.");
        var service = new ApiWriteService(
            (
                (IOrganizationServiceFactory)
                    provider.GetService(typeof(IOrganizationServiceFactory))
            ).CreateOrganizationService(context.UserId),
            context
        );
        var profile = RuntimeProfile.Read(service);
        if (profile.WorkerId != context.UserId)
            throw new InvalidPluginExecutionException(
                "Worker caller differs from the approved runtime identity."
            );
        if (!profile.Enabled)
        {
            context.OutputParameters["Result"] = JsonWire.Write(
                new WorkerResult { Status = "Disabled" }
            );
            return;
        }
        context.SharedVariables[InternalWrite] = true;
        var request = JsonWire.Read<WorkerRequest>((string)context.InputParameters["Request"]);
        if (request.Command == "PurgeHistory")
        {
            context.OutputParameters["Result"] = JsonWire.Write(
                WorkRetention.Purge(service, DateTime.UtcNow)
            );
            return;
        }
        if (request.Command == "BeginHttp")
        {
            context.OutputParameters["Result"] = JsonWire.Write(
                WorkCoordination.BeginHttp(service, request, DateTime.UtcNow)
            );
            return;
        }
        if (request.Command == "Observe" || request.Command == "CreateResponse")
        {
            var held = WorkCoordination.Response(service, request, DateTime.UtcNow);
            if (held != null)
            {
                context.OutputParameters["Result"] = JsonWire.Write(held);
                return;
            }
        }
        var result =
            request.Command == "RefreshSecurity" ? new SecurityRefresh(service).Scan()
            : request.Key?.StartsWith("librarycreate:", StringComparison.Ordinal) == true
                ? new LibraryProvisioning(service).Execute(request, context.IsInTransaction)
            : request.Key?.StartsWith("catalogprobe:", StringComparison.Ordinal) == true
                ? new CatalogWorker(service).Execute(request, context.IsInTransaction)
            : request.Key?.StartsWith("policywork:", StringComparison.Ordinal) == true
                ? new SecurityWorker(service).Execute(request, context.IsInTransaction)
            : new WorkerCoordinator(service, allowedTables: profile.Tables).Execute(
                request,
                context.IsInTransaction
            );
        profile.ValidateTransport(result);
        context.OutputParameters["Result"] = JsonWire.Write(result);
    }
}

public sealed class RecoverWorkerApi : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        if (!context.IsInTransaction)
            throw new InvalidPluginExecutionException("Recovery requires an ambient transaction.");
        var service = new ApiWriteService(
            (
                (IOrganizationServiceFactory)
                    provider.GetService(typeof(IOrganizationServiceFactory))
            ).CreateOrganizationService(context.UserId),
            context
        );
        context.SharedVariables[DocumentWorkerApi.InternalWrite] = true;
        var request = JsonWire.Read<WorkerRequest>((string)context.InputParameters["Request"]);
        context.OutputParameters["Result"] = JsonWire.Write(
            new WorkerCoordinator(service).PermitRecovery(request, context.IsInTransaction)
        );
    }
}

public sealed class StateGuard : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        var parent = context.ParentContext;
        if (context.Stage != 20 || !context.IsInTransaction)
            throw new InvalidPluginExecutionException(
                "State deletion or unguarded mutation is not supported. " + DescribeContext(context)
            );
        if (context.MessageName == "Delete" && ApiWriteService.Authorizes(context))
        {
            bool retention = false;
            for (var frame = parent; frame != null; frame = frame.ParentContext)
                retention |= ApiWriteService.RetentionRequest(frame);
            if (!retention)
                throw new InvalidPluginExecutionException(
                    "Deletion is restricted to expired work history."
                );
            var target = (EntityReference)context.InputParameters["Target"];
            var service = (
                (IOrganizationServiceFactory)
                    provider.GetService(typeof(IOrganizationServiceFactory))
            ).CreateOrganizationService(context.UserId);
            WorkRetention.ValidateDelete(
                service,
                service.Retrieve(
                    target.LogicalName,
                    target.Id,
                    new Microsoft.Xrm.Sdk.Query.ColumnSet(
                        "asx_key",
                        "asx_status",
                        "asx_payload",
                        "asx_retireafter"
                    )
                ),
                DateTime.UtcNow
            );
            return;
        }
        if (context.MessageName != "Create" && context.MessageName != "Update")
            throw new InvalidPluginExecutionException(
                "State deletion or unguarded mutation is not supported. " + DescribeContext(context)
            );
        bool marked =
            parent?.SharedVariables.Contains(DocumentWorkerApi.InternalWrite) == true
            && parent.SharedVariables[DocumentWorkerApi.InternalWrite] is bool flag
            && flag;
        bool api =
            parent?.MessageName == "asx_DocumentWorker"
            || parent?.MessageName == "asx_RecoverWorker"
            || parent?.MessageName == "asx_SecurityAdmin"
            || parent?.MessageName == "asx_RuntimeAdmin"
            || parent?.MessageName == "asx_ManageWork"
            || parent?.MessageName == "asx_CatalogAdmin";
        if (
            (!marked || parent?.UserId != context.UserId || !api)
            && !ApiWriteService.Authorizes(context)
            && !WorkerOutboxAppend.Authorizes(
                context,
                () =>
                    RuntimeProfile
                        .ReadCapture(
                            (
                                (IOrganizationServiceFactory)
                                    provider.GetService(typeof(IOrganizationServiceFactory))
                            ).CreateOrganizationService(context.UserId)
                        )
                        .WorkerId
            )
        )
            throw new InvalidPluginExecutionException(
                "Durable state is writable only through guarded product server operations. "
                    + DescribeContext(context)
            );
        if (context.PrimaryEntityName == "asx_attempt" && context.MessageName != "Create")
            throw new InvalidPluginExecutionException("Attempt receipts are append-only.");
    }

    private static string DescribeContext(IPluginExecutionContext context)
    {
        // Structural diagnostics only: no user IDs, entity IDs, payloads, or variable values.
        string result = "Guard context:";
        var current = context;
        for (int depth = 0; current != null && depth < 4; depth++, current = current.ParentContext)
        {
            bool marked =
                current.SharedVariables?.Contains(DocumentWorkerApi.InternalWrite) == true
                && current.SharedVariables[DocumentWorkerApi.InternalWrite] is bool value
                && value;
            result +=
                " ["
                + depth
                + ":"
                + current.MessageName
                + ",stage="
                + current.Stage
                + ",marked="
                + marked
                + ",sameCaller="
                + (current.UserId == context.UserId)
                + "]";
        }
        return result;
    }
}
