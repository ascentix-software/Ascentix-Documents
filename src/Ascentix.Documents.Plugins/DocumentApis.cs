using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Domain;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Plugins;

public sealed class PreviewTemplateApi : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        var service = (
            (IOrganizationServiceFactory)provider.GetService(typeof(IOrganizationServiceFactory))
        ).CreateOrganizationService(context.InitiatingUserId);
        try
        {
            var request = JsonWire.Read<PreviewRequest>((string)context.InputParameters["Request"]);
            var template = new TemplateStore(service).Read(Guid.Parse(request.RevisionId));
            var recordId = Guid.Parse(request.RecordId);
            var snapshot = new SnapshotReader(service).Read(template, recordId);
            context.OutputParameters["Result"] = JsonWire.Write(
                PlanDto.From(FolderPlanner.Plan(template, recordId, snapshot.Values))
            );
        }
        catch (Exception error) when (!(error is InvalidPluginExecutionException))
        {
            throw new InvalidPluginExecutionException(
                "Preview blocked. Check caller access, approved catalog and configuration. "
                    + (
                        error is EvaluationBlockedException
                            ? error.Message
                            : "No plan was returned."
                    )
            );
        }
    }
}

public sealed class PublishTemplateApi : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        if (!context.IsInTransaction)
            throw new InvalidPluginExecutionException("Publish requires a transaction.");
        var service = (
            (IOrganizationServiceFactory)provider.GetService(typeof(IOrganizationServiceFactory))
        ).CreateOrganizationService(context.InitiatingUserId);
        var revisionId = Guid.Parse((string)context.InputParameters["RevisionId"]);
        var expectedVersion = (string)context.InputParameters["RowVersion"];
        var row = service.Retrieve(
            "asx_revision",
            revisionId,
            new ColumnSet("asx_status", "asx_templateid")
        );
        if (
            row.GetAttributeValue<string>("asx_status") != "Draft"
            || row.RowVersion != expectedVersion
        )
            throw new InvalidPluginExecutionException("Draft changed; reload before publishing.");
        // Validate current metadata and the structural/catalog contract before freezing a revision.
        var publishedTemplate = new TemplateStore(service).Read(revisionId);
        var notices = RuntimeProfile.Read(service).ValidateSources(publishedTemplate);
        new SnapshotReader(service).ValidateMetadata(publishedTemplate);
        service.Execute(
            new UpdateRequest
            {
                Target = new Entity("asx_revision", revisionId)
                {
                    RowVersion = expectedVersion,
                    ["asx_status"] = "Published",
                },
                ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
            }
        );
        var templateRef = row.GetAttributeValue<EntityReference>("asx_templateid");
        var template = service.Retrieve(
            "asx_template",
            templateRef.Id,
            new ColumnSet("asx_publishedrevisionid")
        );
        service.Execute(
            new UpdateRequest
            {
                Target = new Entity("asx_template", template.Id)
                {
                    RowVersion = template.RowVersion,
                    ["asx_publishedrevisionid"] = row.ToEntityReference(),
                },
                ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
            }
        );
        context.OutputParameters["Result"] = JsonWire.Write(
            new PublishResult { Status = "Published", Notices = notices }
        );
    }
}

public sealed class RevisionGuard : IPlugin
{
    private static bool Publishing(IPluginExecutionContext context)
    {
        // Dataverse inserts main-operation frames between a custom API and row guards.
        // Each intermediate frame must match the caller, correlation, entity, and transaction.
        for (var parent = context.ParentContext; parent != null; parent = parent.ParentContext)
        {
            if (
                parent.UserId != context.UserId
                || parent.CorrelationId != context.CorrelationId
                || !parent.IsInTransaction
                || parent.Stage != 30
                || parent.Mode != 0
            )
                return false;
            if (parent.MessageName == "asx_PublishTemplate")
                return true;
            if (
                (
                    parent.MessageName != "Create"
                    && parent.MessageName != "Update"
                    && parent.MessageName != "CreateMultiple"
                    && parent.MessageName != "UpdateMultiple"
                )
                || parent.PrimaryEntityName != context.PrimaryEntityName
            )
                return false;
        }
        return false;
    }

    private static readonly string[] Children =
    {
        "asx_source",
        "asx_destination",
        "asx_folder",
        "asx_documentconditiongroup",
        "asx_condition",
    };

    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        if (
            context.Stage != 20
            || !context.IsInTransaction
            || !new[] { "Create", "Update", "Delete" }.Contains(context.MessageName)
        )
            throw new InvalidPluginExecutionException(
                "Guard requires transactional pre-operation registration."
            );
        var service = (
            (IOrganizationServiceFactory)provider.GetService(typeof(IOrganizationServiceFactory))
        ).CreateOrganizationService(context.UserId);
        bool publish = Publishing(context);
        var entity = context.InputParameters["Target"] as Entity;
        var reference =
            entity?.ToEntityReference()
            ?? context.InputParameters["Target"] as EntityReference
            ?? throw new InvalidPluginExecutionException("Unsupported mutation target.");
        // Template deletion cascades to the template's owned configuration rows.
        bool removingTemplate =
            reference.LogicalName == "asx_template" && context.MessageName == "Delete";
        for (
            var ancestor = context.ParentContext;
            ancestor != null;
            ancestor = ancestor.ParentContext
        )
            removingTemplate |=
                ancestor.MessageName == "Delete"
                && ancestor.PrimaryEntityName == "asx_template"
                && ancestor.UserId == context.UserId;
        if (removingTemplate)
            return;
        Entity? existing =
            context.MessageName == "Create"
                ? null
                : service.Retrieve(
                    reference.LogicalName,
                    reference.Id,
                    new ColumnSet(
                        Children.Contains(reference.LogicalName) ? new[] { "asx_revisionid" }
                        : reference.LogicalName == "asx_revision" ? new[] { "asx_status" }
                        : new[] { "asx_publishedrevisionid", "asx_startsutc", "asx_endsutc" }
                    )
                );
        if (reference.LogicalName == "asx_revision")
        {
            if (
                context.MessageName == "Create"
                && entity!.GetAttributeValue<string>("asx_status") != "Draft"
            )
                throw new InvalidPluginExecutionException("New revisions must be Draft.");
            if (
                existing?.GetAttributeValue<string>("asx_status") != null
                && existing.GetAttributeValue<string>("asx_status") != "Draft"
            )
                throw new InvalidPluginExecutionException("Published revision is immutable.");
            if (
                context.MessageName != "Create"
                && entity?.Contains("asx_status") == true
                && !publish
            )
                throw new InvalidPluginExecutionException(
                    "Publication requires the dedicated API."
                );
        }
        else if (reference.LogicalName == "asx_template")
        {
            if (entity != null)
                TemplateLifecycle.Validate(entity, existing);
            if (entity?.Contains("asx_publishedrevisionid") == true && !publish)
                throw new InvalidPluginExecutionException("Published pointer is API-owned.");
            if (
                existing?.GetAttributeValue<EntityReference>("asx_publishedrevisionid") != null
                && entity?.Contains("asx_table") == true
            )
                throw new InvalidPluginExecutionException(
                    "Published template table/identity cannot change."
                );
        }
        else if (Children.Contains(reference.LogicalName))
        {
            var oldRevision = existing?.GetAttributeValue<EntityReference>("asx_revisionid");
            var nextRevision =
                entity?.Contains("asx_revisionid") == true
                    ? entity.GetAttributeValue<EntityReference>("asx_revisionid")
                    : oldRevision;
            if (nextRevision == null || (oldRevision != null && oldRevision.Id != nextRevision.Id))
                throw new InvalidPluginExecutionException(
                    "Revision ownership is required and immutable."
                );
            var revision = service.Retrieve(
                "asx_revision",
                nextRevision.Id,
                new ColumnSet("asx_status")
            );
            if (revision.GetAttributeValue<string>("asx_status") != "Draft")
                throw new InvalidPluginExecutionException("Published child rows are immutable.");
            // Updates the revision edit stamp using the row version read for this child edit.
            service.Execute(
                new UpdateRequest
                {
                    Target = new Entity("asx_revision", revision.Id)
                    {
                        RowVersion = revision.RowVersion,
                        ["asx_editstamp"] = Guid.NewGuid().ToString("N"),
                    },
                    ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                }
            );
        }
    }
}

public sealed class RecordInvalidationPlugin : IPlugin
{
    public void Execute(IServiceProvider provider)
    {
        var context = (IPluginExecutionContext)provider.GetService(typeof(IPluginExecutionContext));
        if (
            context.Stage != 40
            || !new[] { "Create", "Update", "CreateMultiple", "UpdateMultiple", "Delete" }.Contains(
                context.MessageName
            )
        )
            throw new InvalidPluginExecutionException("Invalid outbox event registration.");
        var service = (
            (IOrganizationServiceFactory)provider.GetService(typeof(IOrganizationServiceFactory))
        ).CreateOrganizationService(context.UserId);
        var profile = RuntimeProfile.ReadCapture(service);
        if (!profile.Tables.Contains(context.PrimaryEntityName))
            return;
        if (
            context.MessageName.StartsWith("Update", StringComparison.Ordinal)
            && !profile.ProcessRecordUpdates
        )
            return;
        if (context.UserId != profile.WorkerId)
            throw new InvalidPluginExecutionException(
                "Event registration differs from the approved planner identity."
            );
        Guid[] ids;
        if (context.MessageName.EndsWith("Multiple", StringComparison.Ordinal))
        {
            var created =
                context.OutputParameters != null
                && context.OutputParameters.Contains("Ids")
                && context.OutputParameters["Ids"] is Guid[] output
                    ? output
                    : null;
            var targets = context.InputParameters.Contains("Targets")
                ? context.InputParameters["Targets"] as EntityCollection
                : null;
            if (
                targets == null
                || targets.Entities.Any(e => e.LogicalName != context.PrimaryEntityName)
            )
                throw new InvalidPluginExecutionException("Bulk event targets are inconsistent.");
            ids = (created ?? targets.Entities.Select(e => e.Id).ToArray())
                .Where(id => id != Guid.Empty)
                .Distinct()
                .ToArray();
        }
        else
        {
            if (context.PrimaryEntityId == Guid.Empty)
                throw new InvalidPluginExecutionException("Record identity missing.");
            ids = new[] { context.PrimaryEntityId };
        }
        var relatedStore = new DocumentStore(service);
        if (!context.MessageName.StartsWith("Create", StringComparison.Ordinal))
            foreach (var id in ids)
                TargetedReplan.Append(
                    service,
                    relatedStore,
                    context.PrimaryEntityName,
                    id,
                    context.CorrelationId
                );
        var query = new QueryExpression("asx_template")
        {
            ColumnSet = TemplateLifecycle.Columns(),
            PageInfo = new PagingInfo { Count = 100, PageNumber = 1 },
        };
        query.Criteria.AddCondition(
            "asx_table",
            ConditionOperator.Equal,
            context.PrimaryEntityName
        );
        query.AddOrder("asx_templateid", OrderType.Ascending);
        var templates = new System.Collections.Generic.List<Entity>();
        while (true)
        {
            var page = service.RetrieveMultiple(query);
            templates.AddRange(page.Entities);
            if (!page.MoreRecords)
                break;
            if (string.IsNullOrEmpty(page.PagingCookie))
                throw new InvalidPluginExecutionException("Template catalog continuation missing.");
            query.PageInfo.PageNumber++;
            query.PageInfo.PagingCookie = page.PagingCookie;
        }
        var allIds = ids;
        foreach (var row in templates)
        {
            ids = allIds;
            var revision = row.GetAttributeValue<EntityReference>("asx_publishedrevisionid");
            if (revision == null || !TemplateLifecycle.Active(row, DateTime.UtcNow))
                continue;
            if (context.MessageName.StartsWith("Update", StringComparison.Ordinal))
            {
                var sources = new TemplateStore(service)
                    .Children("asx_source", revision.Id, "asx_payload")
                    .Select(e => JsonWire.Read<SourceDto>(TemplateStore.Text(e, "asx_payload")))
                    .ToArray();
                var dependencies = sources
                    .Where(s => s.Alias == "root")
                    .SelectMany(s => s.Columns.Select(c => c.Name))
                    .Concat(sources.Where(s => s.Alias != "root").Select(s => s.Lookup))
                    .Where(c => c != null)
                    .ToArray();
                if (
                    context.InputParameters.Contains("Targets")
                    && context.InputParameters["Targets"] is EntityCollection batch
                )
                    ids = ids.Where(id =>
                            batch.Entities.Any(e =>
                                e.Id == id && e.Attributes.Keys.Any(k => dependencies.Contains(k))
                            )
                        )
                        .ToArray();
                else if (
                    context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity target
                    && !target.Attributes.Keys.Any(k => dependencies.Contains(k))
                )
                    continue;
            }
            var store = new DocumentStore(service);
            foreach (var id in ids)
            {
                if (context.MessageName == "Delete")
                {
                    string retired = WorkerCoordinator.RetirementKey(context.PrimaryEntityName, id);
                    if (store.Find<OutboxDocument>("asx_outbox", retired) == null)
                        store.Create(
                            "asx_outbox",
                            new OutboxDocument
                            {
                                Key = retired,
                                TemplateId = row.Id,
                                RevisionId = revision.Id,
                                RecordId = id,
                                Table = context.PrimaryEntityName,
                                Status = "DecommissionReview",
                                Notices = new[]
                                {
                                    "Business record deleted. Existing documents, bindings and receipts are retained; review decommission separately.",
                                },
                            }
                        );
                    continue;
                }
                string key =
                    "event:"
                    + row.Id.ToString("N")
                    + ":"
                    + context.CorrelationId.ToString("N")
                    + ":"
                    + context.PrimaryEntityName
                    + ":"
                    + id.ToString("N")
                    + ":"
                    + (
                        context.MessageName.StartsWith("Create", StringComparison.Ordinal)
                            ? "Create"
                            : "Update"
                    );
                if (store.Find<OutboxDocument>("asx_outbox", key) == null)
                    store.Create(
                        "asx_outbox",
                        new OutboxDocument
                        {
                            Key = key,
                            TemplateId = row.Id,
                            RevisionId = revision.Id,
                            RecordId = id,
                            Table = context.PrimaryEntityName,
                        }
                    );
            }
        }
    }
}
