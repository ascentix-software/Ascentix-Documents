using System;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;

namespace Ascentix.Documents.Plugins;

public static class WorkerOutboxAppend
{
    /// <summary>
    /// Allows the configured worker to create outbox rows with consistent stored identities.
    /// </summary>
    public static bool Authorizes(IPluginExecutionContext context, Func<Guid> workerId)
    {
        if (
            context.Stage != 20
            || context.Mode != 0
            || !context.IsInTransaction
            || context.MessageName != "Create"
            || context.PrimaryEntityName != "asx_outbox"
            || context.UserId == Guid.Empty
        )
            return false;
        var target = ApiWriteService.SingleTarget(context);
        if (target == null || target.LogicalName != "asx_outbox" || target.Id == Guid.Empty)
            return false;
        if (context.UserId != workerId())
            return false;

        var document = JsonWire.Read<StoredDocument>(
            target.GetAttributeValue<string>("asx_payload")
        );
        if (
            document == null
            || document.SchemaVersion != 1
            || string.IsNullOrWhiteSpace(document.Key)
            || string.IsNullOrWhiteSpace(document.Status)
            || document.Status != target.GetAttributeValue<string>("asx_status")
            || target.GetAttributeValue<string>("asx_key") != DocumentStore.Hash(document.Key)
            || target.Id != DocumentStore.StableId("asx_outbox:" + document.Key)
        )
            throw new InvalidPluginExecutionException(
                "Outbox identity, schema or status is inconsistent."
            );
        return true;
    }
}
