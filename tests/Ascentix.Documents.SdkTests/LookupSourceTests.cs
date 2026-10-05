using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Plugins;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

/// <summary>Templates whose lookup (related) source tables are not enabled for Documents.</summary>
public sealed class LookupSourceTests
{
    private const string ContactNotice =
        "Changes to contact records don't update folders until the account record changes or is replanned. Enable contact in the Tables panel to react to its changes.";

    [Fact]
    public void PublishSucceedsWithANoticeWhenALookupTableIsNotEnabled()
    {
        var f = LookupTemplate();
        RuntimeSeed.Seed(f.Service, Guid.NewGuid(), "account");
        f.Service.Rows[f.RevisionId]["asx_status"] = "Draft";

        var result = Publish(f);

        Assert.Equal("Published", result.Status);
        Assert.Equal(new[] { ContactNotice }, result.Notices);
        Assert.Equal(
            "Published",
            f.Service.Rows[f.RevisionId].GetAttributeValue<string>("asx_status")
        );
        Assert.DoesNotContain(
            f.Service.Rows.Values,
            r =>
                r.LogicalName == "asx_runtimetable"
                && r.GetAttributeValue<string>("asx_logicalname") == "contact"
        );
        Assert.DoesNotContain(
            f.Service.Rows.Values,
            r => r.LogicalName == "sdkmessageprocessingstep"
        );
    }

    [Fact]
    public void PublishHasNoNoticeWhenEveryLookupTableIsEnabled()
    {
        var f = LookupTemplate();
        RuntimeSeed.Seed(f.Service, Guid.NewGuid(), "account", "contact");
        f.Service.Rows[f.RevisionId]["asx_status"] = "Draft";

        var result = Publish(f);

        Assert.Equal("Published", result.Status);
        Assert.Empty(result.Notices);
    }

    [Fact]
    public void PublishIsRefusedWhenTheRootTableIsNotEnabled()
    {
        var f = LookupTemplate();
        RuntimeSeed.Seed(f.Service, Guid.NewGuid(), "contact");
        f.Service.Rows[f.RevisionId]["asx_status"] = "Draft";

        Assert.Throws<EvaluationBlockedException>(() => Publish(f));
        Assert.Equal("Draft", f.Service.Rows[f.RevisionId].GetAttributeValue<string>("asx_status"));
    }

    [Fact]
    public void PlanningReadsALookupValueFromATableThatIsNotEnabled()
    {
        var f = LookupTemplate();
        var coordinator = new WorkerCoordinator(f.Service, () => f.Now, new[] { "account" });

        var queued = f.Service.Transaction(() => coordinator.Execute(Queue(f), true));
        var planned = f.Service.Transaction(() =>
            coordinator.Execute(new WorkerRequest { Command = "Plan", Key = queued.Key }, true)
        );

        var operation = f.Store.Require<OperationDocument>("asx_operation", planned.Keys.Single());
        Assert.Equal("Example Pat", operation.Value.Folder.Candidate);
    }

    [Fact]
    public void BatchReviewQueuesAndPlansWithALookupTableThatIsNotEnabled()
    {
        var f = LookupTemplate();
        var api = new BatchReplan(f.Service, new[] { "account" });
        var review = api.Execute(
            new WorkerRequest
            {
                Command = "PreviewBatch",
                TemplateId = f.TemplateId,
                RecordIds = new[] { f.RecordId },
                RequestId = Guid.NewGuid(),
            },
            true
        );
        var queued = f.Service.Transaction(() =>
            api.Execute(
                new WorkerRequest
                {
                    Command = "QueueBatch",
                    Key = review.Key,
                    RowVersion = review.RowVersion,
                },
                true
            )
        );
        var coordinator = new WorkerCoordinator(f.Service, () => f.Now, new[] { "account" });

        var planned = f.Service.Transaction(() =>
            coordinator.Execute(
                new WorkerRequest { Command = "Plan", Key = queued.Keys.Single() },
                true
            )
        );

        Assert.Equal("Planned", planned.Status);
    }

    [Fact]
    public void PlanningIsRefusedWhenTheRootTableIsNotEnabled()
    {
        var f = LookupTemplate();
        var queued = f.Service.Transaction(() => f.Coordinator.Execute(Queue(f), true));
        var scoped = new WorkerCoordinator(f.Service, () => f.Now, new[] { "contact" });

        Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() =>
                scoped.Execute(new WorkerRequest { Command = "Plan", Key = queued.Key }, true)
            )
        );
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_operation");
    }

    /// <summary>
    /// An account template whose root folder name also uses the primary contact's full name.
    /// </summary>
    private static DurableWorkerTests.Fixture LookupTemplate()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var contact = Guid.NewGuid();
        f.Service.Seed(new Entity("contact", contact) { ["fullname"] = "Pat" });
        f.Service.Rows[f.RecordId]["primarycontactid"] = new EntityReference("contact", contact);
        f.Service.Lookups["account.primarycontactid"] = "contact";
        f.Service.Seed(
            new Entity("asx_source", Guid.NewGuid())
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", f.RevisionId),
                ["asx_payload"] = JsonWire.Write(
                    new SourceDto
                    {
                        Alias = "lookup_1",
                        Table = "contact",
                        Lookup = "primarycontactid",
                        Columns = new[]
                        {
                            new ColumnDto { Name = "fullname", Kind = "Text" },
                        },
                    }
                ),
            }
        );
        f.Service.Rows.Values.Single(r => r.LogicalName == "asx_folder")["asx_expression"] =
            "{root.name} {lookup_1.fullname}";
        return f;
    }

    private static WorkerRequest Queue(DurableWorkerTests.Fixture f) =>
        new WorkerRequest
        {
            Command = "Queue",
            RequestId = Guid.NewGuid(),
            TemplateId = f.TemplateId,
            RecordId = f.RecordId,
        };

    private static PublishResult Publish(DurableWorkerTests.Fixture f)
    {
        var context = GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["IsInTransaction"] = true,
                ["InitiatingUserId"] = Guid.NewGuid(),
                ["InputParameters"] = new ParameterCollection
                {
                    ["RevisionId"] = f.RevisionId.ToString(),
                    ["RowVersion"] = f.Service.Rows[f.RevisionId].RowVersion,
                },
                ["OutputParameters"] = new ParameterCollection(),
            }
        );
        new PublishTemplateApi().Execute(new Provider(context, f.Service));
        return JsonWire.Read<PublishResult>((string)context.OutputParameters["Result"]);
    }

    private sealed class Provider : IServiceProvider, IOrganizationServiceFactory
    {
        private readonly IPluginExecutionContext context;
        private readonly IOrganizationService service;

        public Provider(IPluginExecutionContext context, IOrganizationService service)
        {
            this.context = context;
            this.service = service;
        }

        public object GetService(Type type) =>
            type == typeof(IPluginExecutionContext) ? context : this;

        public IOrganizationService CreateOrganizationService(Guid? userId) => service;
    }
}
