using System;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;

namespace Ascentix.Documents.Plugins;

public sealed class TeamMembershipInvalidationPlugin : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        if (
            context.Stage != 40
            || context.Mode != 0
            || !context.IsInTransaction
            || (context.MessageName != "Associate" && context.MessageName != "Disassociate")
        )
            throw new InvalidPluginExecutionException("Invalid membership event registration.");
        if (
            !context.InputParameters.Contains("Relationship")
            || !(context.InputParameters["Relationship"] is Relationship relationship)
            || relationship.SchemaName != "teammembership_association"
        )
            return;
        var service = (
            (IOrganizationServiceFactory)provider.GetService(typeof(IOrganizationServiceFactory))
        ).CreateOrganizationService(context.UserId);
        if (RuntimeProfile.Read(service).WorkerId != context.UserId)
            throw new InvalidPluginExecutionException(
                "Membership event identity differs from approved runtime."
            );
        var target = context.InputParameters.Contains("Target")
            ? context.InputParameters["Target"] as EntityReference
            : null;
        var related = context.InputParameters.Contains("RelatedEntities")
            ? context.InputParameters["RelatedEntities"] as EntityReferenceCollection
            : null;
        if (
            target == null
            || target.Id == Guid.Empty
            || related == null
            || related.Count > 100
            || related.Any(r => r.Id == Guid.Empty)
        )
            throw new InvalidPluginExecutionException(
                "Bounded exact membership relationship identities required."
            );
        Guid[] teams;
        if (target.LogicalName == "team" && related.All(r => r.LogicalName == "systemuser"))
            teams = new[] { target.Id };
        else if (target.LogicalName == "systemuser" && related.All(r => r.LogicalName == "team"))
            teams = related.Select(r => r.Id).Distinct().ToArray();
        else
            throw new InvalidPluginExecutionException(
                "Unsupported membership relationship direction."
            );
        var store = new DocumentStore(service);
        foreach (var team in teams)
            if (
                store.Find<TeamRegistration>("asx_teamregistration", "team:" + team.ToString("N"))
                != null
            )
                Enqueue(store, team, context.CorrelationId, context.MessageName);
    }

    public static void Enqueue(DocumentStore store, Guid team, Guid correlation, string message)
    {
        string key =
            "team-event:" + correlation.ToString("N") + ":" + team.ToString("N") + ":" + message;
        if (store.Find<OutboxDocument>("asx_outbox", key) == null)
            store.Create(
                "asx_outbox",
                new OutboxDocument
                {
                    Key = key,
                    SecurityTeamId = team,
                    SecurityPage = 1,
                }
            );
    }
}
