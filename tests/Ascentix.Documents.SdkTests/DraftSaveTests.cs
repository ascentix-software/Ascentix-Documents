using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Plugins;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class DraftSaveTests
{
    private sealed class Provider : IServiceProvider, IOrganizationServiceFactory
    {
        private readonly DurableWorkerTests.Fixture f;
        public IPluginExecutionContext Context;

        public Provider(DurableWorkerTests.Fixture f, DraftDto draft)
        {
            this.f = f;
            Context = GuardTests.ContextProxy.Create(
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
        }

        public object GetService(Type type) =>
            type == typeof(IPluginExecutionContext) ? Context : this;

        public IOrganizationService CreateOrganizationService(Guid? id) => f.Service;
    }

    private static DraftResult Save(DurableWorkerTests.Fixture f, DraftDto draft)
    {
        var p = new Provider(f, draft);
        return f.Service.Transaction(() =>
        {
            new CreateDraftApi().Execute(p);
            return JsonWire.Read<DraftResult>((string)p.Context.OutputParameters["Result"]);
        });
    }

    [Fact]
    public void RecordCreateQueuesEveryActiveTemplateIndependentlyAndDeduplicatesReplay()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var worker = Guid.NewGuid();
        f.Service.Seed(
            new Entity("asx_runtime", Guid.NewGuid())
            {
                ["asx_name"] = "Default",
                ["asx_workeruserid"] = worker.ToString(),
                ["asx_enabled"] = true,
                ["asx_allowedtables"] = "[\"account\"]",
                ["asx_sharepointhosts"] = "[\"example.sharepoint.com\"]",
            }
        );
        var second = Guid.NewGuid();
        f.Service.Seed(
            new Entity("asx_template", second)
            {
                ["asx_table"] = "account",
                ["asx_publishedrevisionid"] = new EntityReference("asx_revision", Guid.NewGuid()),
            }
        );
        var disabled = Guid.NewGuid();
        f.Service.Seed(
            new Entity("asx_template", disabled)
            {
                ["asx_table"] = "account",
                ["asx_disabled"] = true,
                ["asx_publishedrevisionid"] = new EntityReference("asx_revision", Guid.NewGuid()),
            }
        );
        var p = new Provider(f, new DraftDto());
        p.Context = GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["Stage"] = 40,
                ["Mode"] = 0,
                ["IsInTransaction"] = true,
                ["MessageName"] = "Create",
                ["UserId"] = worker,
                ["PrimaryEntityName"] = "account",
                ["PrimaryEntityId"] = Guid.NewGuid(),
                ["CorrelationId"] = Guid.NewGuid(),
                ["OwningExtension"] = new EntityReference(
                    "sdkmessageprocessingstep",
                    Guid.NewGuid()
                ),
                ["SharedVariables"] = new ParameterCollection(),
                ["InputParameters"] = new ParameterCollection(),
            }
        );
        f.Service.Transaction(() =>
        {
            new RecordInvalidationPlugin().Execute(p);
            return true;
        });
        f.Service.Transaction(() =>
        {
            new RecordInvalidationPlugin().Execute(p);
            return true;
        });
        var jobs = f
            .Service.Rows.Values.Where(r => r.LogicalName == "asx_outbox")
            .Select(r => JsonWire.Read<OutboxDocument>(r.GetAttributeValue<string>("asx_payload")))
            .ToArray();
        Assert.Equal(2, jobs.Length);
        Assert.Contains(jobs, j => j.TemplateId == f.TemplateId);
        Assert.Contains(jobs, j => j.TemplateId == second);
    }

    [Fact]
    public void MultipleTemplatesOnSameTableHaveIndependentDraftsAndRejectCrossTemplateRevision()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var draft = new DraftReader(f.Service).Read(f.RevisionId).Draft;
        draft.TemplateId = null;
        draft.Name = "Second template";
        var second = Save(f, draft);
        Assert.NotEqual(f.TemplateId.ToString("D"), second.TemplateId);
        Assert.Equal(
            1,
            f.Service.Rows[Guid.Parse(second.RevisionId)].GetAttributeValue<int>("asx_version")
        );
        Assert.Equal(
            f.RevisionId,
            f.Service.Rows[f.TemplateId]
                .GetAttributeValue<EntityReference>("asx_publishedrevisionid")
                .Id
        );
        draft.TemplateId = second.TemplateId;
        draft.RevisionId = f.RevisionId.ToString();
        draft.RowVersion = f.Service.Rows[f.RevisionId].RowVersion;
        Assert.Throws<InvalidPluginExecutionException>(() => Save(f, draft));
        draft.RevisionId = second.RevisionId;
        draft.RowVersion = second.RowVersion;
        Assert.Equal(second.RevisionId, Save(f, draft).RevisionId);
        draft.Table = "contact";
        Assert.Throws<InvalidPluginExecutionException>(() => Save(f, draft));
    }

    [Fact]
    public void SavesReuseDraftWhilePublishedRevisionAndPointerRemainUnchanged()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var draft = new DraftReader(f.Service).Read(f.RevisionId).Draft;
        draft.RevisionId = f.RevisionId.ToString();
        draft.RowVersion = f.Service.Rows[f.RevisionId].RowVersion;
        var first = Save(f, draft);
        var id = Guid.Parse(first.RevisionId);
        Assert.NotEqual(f.RevisionId, id);
        Assert.Equal(2, f.Service.Rows[id].GetAttributeValue<int>("asx_version"));
        draft.RevisionId = first.RevisionId;
        draft.RowVersion = first.RowVersion;
        draft.Destinations[0].Name = "Updated draft";
        var second = Save(f, draft);
        Assert.Equal(first.RevisionId, second.RevisionId);
        Assert.NotEqual(first.RowVersion, second.RowVersion);
        Assert.Equal(2, f.Service.Rows.Values.Count(r => r.LogicalName == "asx_revision"));
        Assert.Equal(
            "Updated draft",
            new DraftReader(f.Service).Read(id).Draft.Destinations[0].Name
        );
        Assert.Equal(
            f.RevisionId,
            f.Service.Rows[f.TemplateId]
                .GetAttributeValue<EntityReference>("asx_publishedrevisionid")
                .Id
        );
        Assert.Equal(
            "Published",
            f.Service.Rows[f.RevisionId].GetAttributeValue<string>("asx_status")
        );
        Assert.Throws<InvalidPluginExecutionException>(() => Save(f, draft)); // stale row version
        draft.RevisionId = null;
        draft.RowVersion = null;
        Assert.Throws<InvalidPluginExecutionException>(() => Save(f, draft)); // don't overwrite someone else's draft
        f.Service.Rows[id]["asx_status"] = "Published";
        draft.RevisionId = id.ToString();
        draft.RowVersion = second.RowVersion;
        var third = Save(f, draft);
        Assert.NotEqual(id, Guid.Parse(third.RevisionId));
        Assert.Equal(
            3,
            f.Service.Rows[Guid.Parse(third.RevisionId)].GetAttributeValue<int>("asx_version")
        );
    }
}
