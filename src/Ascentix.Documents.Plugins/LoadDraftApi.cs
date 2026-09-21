using System;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;

namespace Ascentix.Documents.Plugins;

public sealed class LoadDraftApi : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        var service = (
            (IOrganizationServiceFactory)provider.GetService(typeof(IOrganizationServiceFactory))
        ).CreateOrganizationService(context.InitiatingUserId);
        context.OutputParameters["Result"] = JsonWire.Write(
            new DraftReader(service).Read(Guid.Parse((string)context.InputParameters["RevisionId"]))
        );
    }
}
