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
        var runtimeRow = RuntimeSeed.Seed(f.Service, worker, "account");
        f.Service.Rows[runtimeRow]["asx_enabled"] = true;
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

    private static PlanDto Preview(DurableWorkerTests.Fixture f, PreviewRequest request)
    {
        var context = GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["InitiatingUserId"] = Guid.NewGuid(),
                ["InputParameters"] = new ParameterCollection
                {
                    ["Request"] = JsonWire.Write(request),
                },
                ["OutputParameters"] = new ParameterCollection(),
            }
        );
        new PreviewTemplateApi().Execute(new Provider(f, new DraftDto()) { Context = context });
        return JsonWire.Read<PlanDto>((string)context.OutputParameters["Result"]);
    }

    [Fact]
    public void PreviewOfUnsavedEditsMatchesTheSavedRevisionAndWritesNothing()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var draft = new DraftReader(f.Service).Read(f.RevisionId).Draft;
        int rows = f.Service.Rows.Count;
        long updates = f.Service.Updates.Count;
        var saved = Preview(
            f,
            new PreviewRequest
            {
                RevisionId = f.RevisionId.ToString("D"),
                RecordId = f.RecordId.ToString("D"),
            }
        );
        var unsaved = Preview(
            f,
            new PreviewRequest { Draft = draft, RecordId = f.RecordId.ToString("D") }
        );
        Assert.Equal(JsonWire.Write(saved), JsonWire.Write(unsaved));
        Assert.Equal(rows, f.Service.Rows.Count);
        Assert.Equal(updates, f.Service.Updates.Count);
        draft.Destinations[0].Folders[0].Name = "Edited {root.name}";
        Assert.Equal(
            "Edited Example",
            Preview(f, new PreviewRequest { Draft = draft, RecordId = f.RecordId.ToString("D") })
                .Folders[0]
                .Name
        );
    }

    [Fact]
    public void PreviewNeedsExactlyOneOfRevisionAndDraft()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var draft = new DraftReader(f.Service).Read(f.RevisionId).Draft;
        Assert.Throws<InvalidPluginExecutionException>(() =>
            Preview(f, new PreviewRequest { RecordId = f.RecordId.ToString("D") })
        );
        Assert.Throws<InvalidPluginExecutionException>(() =>
            Preview(
                f,
                new PreviewRequest
                {
                    RevisionId = f.RevisionId.ToString("D"),
                    Draft = draft,
                    RecordId = f.RecordId.ToString("D"),
                }
            )
        );
    }

    [Theory]
    [InlineData("sources")]
    [InlineData("destinations")]
    [InlineData("folders")]
    [InlineData("rows")]
    [InlineData("name")]
    public void BuildRefusesWhatSaveRefusesWithTheSameText(string breach)
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var draft = new DraftReader(f.Service).Read(f.RevisionId).Draft;
        var root = draft.Destinations[0];
        switch (breach)
        {
            case "sources":
                draft.Sources = Enumerable
                    .Repeat(draft.Sources[0], Ascentix.Documents.Domain.Bounds.Sources + 1)
                    .ToArray();
                break;
            case "destinations":
                draft.Destinations = Enumerable
                    .Range(0, Ascentix.Documents.Domain.Bounds.Destinations + 1)
                    .Select(i => new DestinationDto
                    {
                        Key = "d" + i,
                        LibraryId = root.LibraryId,
                        Folders = root.Folders,
                    })
                    .ToArray();
                break;
            case "folders":
                root.Folders = root
                    .Folders.Concat(
                        Enumerable
                            .Range(0, Ascentix.Documents.Domain.Bounds.FoldersPerDestination)
                            .Select(i => new FolderDto
                            {
                                Key = "f" + i,
                                Parent = root.Folders[0].Key,
                                Name = "F" + i,
                            })
                    )
                    .ToArray();
                break;
            case "rows":
                root.Folders[0].Condition = new GroupDto
                {
                    Conditions = Enumerable
                        .Range(0, Ascentix.Documents.Domain.Bounds.ConfigurationRows)
                        .Select(_ => new ConditionDto { Column = "name", Operator = "IsNotNull" })
                        .ToArray(),
                };
                break;
            case "name":
                root.Name = new string('n', 201);
                break;
        }
        int revisions = f.Service.Rows.Values.Count(r => r.LogicalName == "asx_revision");
        var built = Assert.ThrowsAny<Exception>(() => DraftTemplate.Build(f.Service, draft));
        var saved = Assert.Throws<InvalidPluginExecutionException>(() => Save(f, draft));
        Assert.Equal(built.Message, saved.Message);
        Assert.Equal(revisions, f.Service.Rows.Values.Count(r => r.LogicalName == "asx_revision"));
    }

    [Fact]
    public void FoldersNestUpToTheFolderDepthBound()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var draft = new DraftReader(f.Service).Read(f.RevisionId).Draft;
        var destination = draft.Destinations[0];
        var top = destination.Folders.Single(x => x.Parent == null);
        FolderDto[] Chain(int levels) =>
            new[] { top }
                .Concat(
                    Enumerable
                        .Range(1, levels - 1)
                        .Select(i => new FolderDto
                        {
                            Key = "level" + i,
                            Parent = i == 1 ? top.Key : "level" + (i - 1),
                            Name = "L" + i,
                        })
                )
                .ToArray();
        destination.Folders = Chain(Ascentix.Documents.Domain.Bounds.FolderDepth);
        DraftTemplate.Build(f.Service, draft);
        destination.Folders = Chain(Ascentix.Documents.Domain.Bounds.FolderDepth + 1);
        var refused = Assert.ThrowsAny<Exception>(() => DraftTemplate.Build(f.Service, draft));
        Assert.Equal("Folder cycle/depth exceeds bounds.", refused.Message);
    }

    [Fact]
    public void ConditionsAreBoundedByTheRowsOneSaveWritesNotByAHundred()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        var draft = new DraftReader(f.Service).Read(f.RevisionId).Draft;
        draft.Destinations[0].Folders[0].Condition = new GroupDto
        {
            Conditions = Enumerable
                .Range(0, 150)
                .Select(_ => new ConditionDto { Column = "name", Operator = "IsNotNull" })
                .ToArray(),
        };
        var saved = Save(f, draft);
        var loaded = new DraftReader(f.Service).Read(Guid.Parse(saved.RevisionId));
        Assert.Equal(150, loaded.Draft.Destinations[0].Folders[0].Condition!.Conditions.Length);
    }
}
