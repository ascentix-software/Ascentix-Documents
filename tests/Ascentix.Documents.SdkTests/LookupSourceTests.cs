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
    public void PlanningStopsCleanlyWhenTheRootTableIsNotEnabled()
    {
        var f = LookupTemplate();
        var queued = f.Service.Transaction(() => f.Coordinator.Execute(Queue(f), true));
        var scoped = new WorkerCoordinator(f.Service, () => f.Now, new[] { "contact" });

        var planned = f.Service.Transaction(() =>
            scoped.Execute(new WorkerRequest { Command = "Plan", Key = queued.Key }, true)
        );
        Assert.Equal("Cancelled", planned.Status);
        Assert.Contains(planned.Notices, n => n.Contains("account is no longer enabled"));
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_operation");
    }

    [Theory]
    [InlineData("contact")]
    [InlineData("account,contact")]
    public void LoadDraftNamesTheRecordALookupConditionComparesWith(string targets)
    {
        var f = LookupTemplate();
        var contact = Guid.NewGuid();
        f.Service.Seed(new Entity("contact", contact) { ["fullname"] = "Jane Smith" });
        f.Service.PrimaryNames["contact"] = "fullname";
        f.Service.Lookups["account.primarycontactid"] = targets;
        var draft = new DraftReader(f.Service).Read(f.RevisionId);
        Assert.Equal(1, draft.Version);
        var condition = new GroupDto
        {
            Conditions = new[]
            {
                new ConditionDto
                {
                    Column = "primarycontactid",
                    Operator = "Equal",
                    LiteralKind = "Lookup",
                    Literal = contact.ToString("D"),
                },
            },
        };
        var root = draft.Draft.Sources.Single(s => s.Alias == "root");
        root.Columns = root
            .Columns.Concat(
                new[]
                {
                    new ColumnDto { Name = "primarycontactid", Kind = "Lookup" },
                }
            )
            .ToArray();
        draft.Draft.Destinations[0].Folders[0].Condition = condition;
        Save(f, draft.Draft);
        var loaded = new DraftReader(f.Service)
            .Read(LatestRevision(f))
            .Draft.Destinations[0]
            .Folders[0]
            .Condition!.Conditions[0];
        Assert.Equal("Jane Smith", loaded.LiteralLabel);
        Assert.Equal("contact", loaded.LiteralTable);
        Assert.Equal(contact.ToString("D"), loaded.Literal);
    }

    [Fact]
    public void ARecordTheCallerCannotFindHasNoLabelAndSaveIgnoresTheLabels()
    {
        var f = LookupTemplate();
        f.Service.Lookups["account.primarycontactid"] = "contact";
        var draft = new DraftReader(f.Service).Read(f.RevisionId).Draft;
        var root = draft.Sources.Single(s => s.Alias == "root");
        root.Columns = root
            .Columns.Concat(
                new[]
                {
                    new ColumnDto { Name = "primarycontactid", Kind = "Lookup" },
                }
            )
            .ToArray();
        draft.Destinations[0].Folders[0].Condition = new GroupDto
        {
            Conditions = new[]
            {
                new ConditionDto
                {
                    Column = "primarycontactid",
                    Operator = "Equal",
                    LiteralKind = "Lookup",
                    Literal = Guid.NewGuid().ToString("D"),
                    LiteralLabel = "Forged",
                    LiteralTable = "systemuser",
                },
            },
        };
        Save(f, draft);
        var loaded = new DraftReader(f.Service)
            .Read(LatestRevision(f))
            .Draft.Destinations[0]
            .Folders[0]
            .Condition!.Conditions[0];
        Assert.Null(loaded.LiteralLabel);
        Assert.Null(loaded.LiteralTable);
        Assert.DoesNotContain(
            f.Service.Rows.Values,
            r =>
                r.LogicalName == "asx_condition"
                && r.GetAttributeValue<string>("asx_payload").Contains("Forged")
        );
    }

    [Fact]
    public void ARecordInATableTheCallerCannotReadHasNoLabelAndTheLoadStillSucceeds()
    {
        var f = LookupTemplate();
        var contact = Guid.NewGuid();
        f.Service.Seed(new Entity("contact", contact) { ["fullname"] = "Jane Smith" });
        f.Service.PrimaryNames["contact"] = "fullname";
        Save(f, WithLookupCondition(f, contact));
        // Querying a table without its Read privilege faults, and inside LoadDraft's transaction
        // even a caught fault would end it, so the label is left out without that call.
        f.Service.CallerCannotRead.Add("contact");
        var loaded = f
            .Service.Transaction(() => new DraftReader(f.Service).Read(LatestRevision(f)))
            .Draft.Destinations[0]
            .Folders[0]
            .Condition!.Conditions[0];
        Assert.Null(loaded.LiteralLabel);
        Assert.Null(loaded.LiteralTable);
        Assert.Equal(contact.ToString("D"), loaded.Literal);
    }

    [Fact]
    public void ALookupColumnDeletedFromTheTableHasNoLabelAndTheLoadStillSucceeds()
    {
        var f = LookupTemplate();
        var contact = Guid.NewGuid();
        f.Service.Seed(new Entity("contact", contact) { ["fullname"] = "Jane Smith" });
        Save(f, WithLookupCondition(f, contact));
        // RetrieveAttribute faults for a deleted column; the load reads metadata without faulting.
        f.Service.MissingColumns.Add("account.primarycontactid");
        var loaded = f
            .Service.Transaction(() => new DraftReader(f.Service).Read(LatestRevision(f)))
            .Draft.Destinations[0]
            .Folders[0]
            .Condition!.Conditions[0];
        Assert.Null(loaded.LiteralLabel);
        Assert.Equal(contact.ToString("D"), loaded.Literal);
    }

    [Fact]
    public void ManyLookupConditionsOnOneColumnReadEachTargetsMetadataAndPrivilegeOnce()
    {
        var f = LookupTemplate();
        f.Service.PrimaryNames["contact"] = "fullname";
        f.Service.Lookups["account.primarycontactid"] = "account,contact";
        var contacts = Enumerable
            .Range(0, 4)
            .Select(i =>
            {
                var id = Guid.NewGuid();
                f.Service.Seed(new Entity("contact", id) { ["fullname"] = "Contact " + i });
                return id;
            })
            .ToArray();
        var draft = WithLookupCondition(f, contacts[0]);
        var group = draft.Destinations[0].Folders[0].Condition!;
        group.All = false;
        group.Conditions = contacts
            .Select(id => new ConditionDto
            {
                Column = "primarycontactid",
                Operator = "Equal",
                LiteralKind = "Lookup",
                Literal = id.ToString("D"),
            })
            .ToArray();
        Save(f, draft);
        f.Service.Executed.Clear();

        var loaded = new DraftReader(f.Service).Read(LatestRevision(f));

        Assert.Equal(
            new[] { "Contact 0", "Contact 1", "Contact 2", "Contact 3" },
            loaded
                .Draft.Destinations[0]
                .Folders[0]
                .Condition!.Conditions.Select(c => c.LiteralLabel)
        );
        var metadata = f
            .Service.Executed.OfType<Microsoft.Xrm.Sdk.Messages.RetrieveMetadataChangesRequest>()
            .ToArray();
        Assert.Single(metadata, r => r.Query.AttributeQuery != null);
        Assert.Equal(
            new[] { "account", "contact" },
            metadata
                .Where(r => r.Query.AttributeQuery == null)
                .Select(r => (string)r.Query.Criteria.Conditions.Single().Value)
                .OrderBy(t => t)
        );
        Assert.Equal(
            2,
            f.Service.Executed.OfType<Microsoft.Crm.Sdk.Messages.RetrieveUserPrivilegeByPrivilegeIdRequest>()
                .Count()
        );
        Assert.Single(f.Service.Executed.OfType<Microsoft.Crm.Sdk.Messages.WhoAmIRequest>());
    }

    /// <summary>The template's draft with its root folder kept only when the primary contact is the given one.</summary>
    private static DraftDto WithLookupCondition(DurableWorkerTests.Fixture f, Guid contact)
    {
        var draft = new DraftReader(f.Service).Read(f.RevisionId).Draft;
        var root = draft.Sources.Single(s => s.Alias == "root");
        root.Columns = root
            .Columns.Concat(
                new[]
                {
                    new ColumnDto { Name = "primarycontactid", Kind = "Lookup" },
                }
            )
            .ToArray();
        draft.Destinations[0].Folders[0].Condition = new GroupDto
        {
            Conditions = new[]
            {
                new ConditionDto
                {
                    Column = "primarycontactid",
                    Operator = "Equal",
                    LiteralKind = "Lookup",
                    Literal = contact.ToString("D"),
                },
            },
        };
        return draft;
    }

    private static DraftResult Save(DurableWorkerTests.Fixture f, DraftDto draft)
    {
        var context = GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["IsInTransaction"] = true,
                ["InitiatingUserId"] = Guid.NewGuid(),
                ["InputParameters"] = new ParameterCollection
                {
                    ["Request"] = JsonWire.Write(draft),
                },
                ["OutputParameters"] = new ParameterCollection(),
            }
        );
        return f.Service.Transaction(() =>
        {
            new CreateDraftApi().Execute(new Provider(context, f.Service));
            return JsonWire.Read<DraftResult>((string)context.OutputParameters["Result"]);
        });
    }

    private static Guid LatestRevision(DurableWorkerTests.Fixture f) =>
        f
            .Service.Rows.Values.Where(r => r.LogicalName == "asx_revision")
            .OrderByDescending(r => r.GetAttributeValue<int>("asx_version"))
            .First()
            .Id;

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
