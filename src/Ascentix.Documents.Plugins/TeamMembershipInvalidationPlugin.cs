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
            || (context.MessageName != "Associate" && context.MessageName != "Disassociate")
        )
            throw new InvalidPluginExecutionException("Invalid membership event registration.");
        if (
            !context.InputParameters.Contains("Relationship")
            || !(context.InputParameters["Relationship"] is Relationship relationship)
            || relationship.SchemaName != "teammembership_association"
        )
            return;
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
            || related.Any(r => r.Id == Guid.Empty)
        )
            throw new InvalidPluginExecutionException(
                "Exact membership relationship identities required."
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
        var service = (
            (IOrganizationServiceFactory)provider.GetService(typeof(IOrganizationServiceFactory))
        ).CreateOrganizationService(context.UserId);
        var store = new DocumentStore(service);
        var registered = teams
            .Where(team =>
                store.Find<TeamRegistration>("asx_teamregistration", "team:" + team.ToString("N"))
                != null
            )
            .ToArray();
        if (registered.Length == 0)
            return;
        if (RuntimeProfile.ReadCapture(service).WorkerId != context.UserId)
            throw new InvalidPluginExecutionException(
                "Membership event identity differs from approved runtime."
            );
        foreach (var team in registered)
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
