using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Plugins;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class AsyncCaptureTests
{
    private static DurableWorkerTests.Fixture Setup(out Guid runtime)
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        runtime = RuntimeSeed.Seed(f.Service, f.RecordId, "account");
        return f;
    }

    private static IPluginExecutionContext Context(
        DurableWorkerTests.Fixture f,
        string message,
        string entity,
        ParameterCollection inputs,
        Guid? user = null,
        Guid? primaryId = null
    ) =>
        GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["OwningExtension"] = new EntityReference(
                    "sdkmessageprocessingstep",
                    Guid.NewGuid()
                ),
                ["Stage"] = 40,
                ["Mode"] = 1,
                ["IsInTransaction"] = false,
                ["MessageName"] = message,
                ["PrimaryEntityName"] = entity,
                ["PrimaryEntityId"] = primaryId ?? f.RecordId,
                ["UserId"] = user ?? f.RecordId,
                ["InputParameters"] = inputs,
                ["OutputParameters"] = new ParameterCollection(),
                ["SharedVariables"] = new ParameterCollection(),
                ["CorrelationId"] = Guid.NewGuid(),
            }
        );

    private sealed class Provider : IServiceProvider, IOrganizationServiceFactory
    {
        private readonly DurableWorkerTests.Fixture f;
        private readonly IPluginExecutionContext c;

        public Provider(DurableWorkerTests.Fixture f, IPluginExecutionContext c)
        {
            this.f = f;
            this.c = c;
        }

        public object GetService(Type type) => type == typeof(IPluginExecutionContext) ? c : this;

        public IOrganizationService CreateOrganizationService(Guid? userId) => f.Service;
    }

    [Fact]
    public void AsyncCreateIsAcceptedAndCorruptHostsDoNotStopCapture()
    {
        var f = Setup(out var runtime);
        f.Service.Rows[runtime]["asx_sharepointhosts"] = "not json";
        new RecordInvalidationPlugin().Execute(
            new Provider(f, Context(f, "Create", "account", new ParameterCollection()))
        );
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void TablesOutsideTheAllowlistAreSkippedSilently()
    {
        var f = Setup(out _);
        new RecordInvalidationPlugin().Execute(
            new Provider(f, Context(f, "Create", "contact", new ParameterCollection()))
        );
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void BulkTargetsAboveOneHundredAreProcessed()
    {
        var f = Setup(out _);
        var targets = new EntityCollection(
            Enumerable.Range(0, 150).Select(_ => new Entity("account", Guid.NewGuid())).ToList()
        );
        new RecordInvalidationPlugin().Execute(
            new Provider(
                f,
                Context(
                    f,
                    "CreateMultiple",
                    "account",
                    new ParameterCollection { ["Targets"] = targets }
                )
            )
        );
        Assert.Equal(150, f.Service.Rows.Values.Count(r => r.LogicalName == "asx_outbox"));
    }

    [Fact]
    public void MembershipForUnregisteredTeamsSkipsBeforeReadingRuntime()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        var reads = new List<string>();
        f.Service.QueryHook = q =>
        {
            reads.Add(q.EntityName);
            return null;
        };
        var related = new EntityReferenceCollection(
            Enumerable
                .Range(0, 150)
                .Select(_ => new EntityReference("systemuser", Guid.NewGuid()))
                .ToList()
        );
        new TeamMembershipInvalidationPlugin().Execute(
            new Provider(
                f,
                Context(
                    f,
                    "Associate",
                    "team",
                    new ParameterCollection
                    {
                        ["Relationship"] = new Relationship("teammembership_association"),
                        ["Target"] = new EntityReference("team", Guid.NewGuid()),
                        ["RelatedEntities"] = related,
                    }
                )
            )
        );
        Assert.DoesNotContain("asx_runtime", reads);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void RetirementOfUnregisteredTeamSkipsBeforeReadingRuntime()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        var reads = new List<string>();
        f.Service.QueryHook = q =>
        {
            reads.Add(q.EntityName);
            return null;
        };
        new TeamRetirementPlugin().Execute(
            new Provider(
                f,
                Context(f, "Delete", "team", new ParameterCollection(), primaryId: Guid.NewGuid())
            )
        );
        Assert.DoesNotContain("asx_runtime", reads);
    }

    private static WorkerResult Rerun(DurableWorkerTests.Fixture f, WorkerRequest request) =>
        f.Service.Transaction(() =>
            MissedChanges.Rerun(f.Service, request, new[] { "account" }, () => f.Now)
        );

    [Fact]
    public void RerunRecordQueuesOnlyPublishedActiveTemplatesAndIsIdempotent()
    {
        var f = Setup(out _);
        var off = Guid.NewGuid();
        f.Service.Seed(
            new Entity("asx_template", off)
            {
                ["asx_name"] = "Contract documents",
                ["asx_table"] = "account",
                ["asx_disabled"] = true,
                ["asx_publishedrevisionid"] = new EntityReference("asx_revision", Guid.NewGuid()),
            }
        );
        var draftOnly = Guid.NewGuid();
        f.Service.Seed(
            new Entity("asx_template", draftOnly)
            {
                ["asx_name"] = "Draft only",
                ["asx_table"] = "account",
            }
        );
        var request = new WorkerRequest
        {
            Table = "account",
            RecordId = f.RecordId,
            RequestId = Guid.NewGuid(),
        };
        var first = Rerun(f, request);
        Assert.Equal("Queued", first.Status);
        Assert.Single(first.Keys);
        Assert.Equal(new[] { "Template 'Contract documents' is off; skipped." }, first.Notices);
        var again = Rerun(f, request);
        Assert.Equal(first.Keys, again.Keys);
        Assert.Single(
            f.Service.Rows.Values,
            r =>
                r.LogicalName == "asx_outbox"
                && r.GetAttributeValue<string>("asx_payload").Contains("request:")
        );
        var queued = f.Store.Require<OutboxDocument>("asx_outbox", first.Keys[0]).Value;
        Assert.Equal(f.TemplateId, queued.TemplateId);
        Assert.Equal(1, queued.Priority);
    }

    [Fact]
    public void RerunRecordWithNoActiveTemplateIsInactive()
    {
        var f = Setup(out _);
        f.Service.Rows[f.TemplateId]["asx_disabled"] = true;
        var result = Rerun(
            f,
            new WorkerRequest
            {
                Table = "account",
                RecordId = f.RecordId,
                RequestId = Guid.NewGuid(),
            }
        );
        Assert.Equal("Inactive", result.Status);
        Assert.Empty(result.Keys);
    }

    [Fact]
    public void RerunRecordRefusesATableThatIsNotEnabled()
    {
        var f = Setup(out _);
        var refused = Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            f.Service.Transaction(() =>
                MissedChanges.Rerun(
                    f.Service,
                    new WorkerRequest
                    {
                        Table = "account",
                        RecordId = f.RecordId,
                        RequestId = Guid.NewGuid(),
                    },
                    new[] { "contact" },
                    () => f.Now
                )
            )
        );
        Assert.Equal(WorkerCoordinator.TableNotEnabled("account"), refused.Message);
        Assert.DoesNotContain(f.Service.Rows.Values, r => r.LogicalName == "asx_outbox");
    }

    [Fact]
    public void DismissOfAJobThatIsAlreadyGoneReturnsDismissedWithoutAFault()
    {
        var f = Setup(out _);
        Assert.Equal("Dismissed", MissedChanges.Dismiss(f.Service, Guid.NewGuid()).Status);
    }

    [Fact]
    public void DismissDeletesOnlyAFailedDocumentsCaptureJob()
    {
        var f = Setup(out _);
        Guid Job(string handler, int status)
        {
            var step = Guid.NewGuid();
            var type = Guid.NewGuid();
            f.Service.Seed(new Entity("plugintype", type) { ["typename"] = handler });
            f.Service.Seed(
                new Entity("sdkmessageprocessingstep", step)
                {
                    ["eventhandler"] = new EntityReference("plugintype", type),
                }
            );
            var job = Guid.NewGuid();
            f.Service.Seed(
                new Entity("asyncoperation", job)
                {
                    ["statuscode"] = new OptionSetValue(status),
                    ["owningextensionid"] = new EntityReference("sdkmessageprocessingstep", step),
                }
            );
            return job;
        }
        var ours = Job(EventRegistrations.RecordHandler, 31);
        Assert.Equal("Dismissed", MissedChanges.Dismiss(f.Service, ours).Status);
        Assert.False(f.Service.Rows.ContainsKey(ours));
        var other = Job("Contoso.Plugins.Something", 31);
        Assert.Equal(
            "Only a failed Documents capture job can be dismissed.",
            Assert
                .Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
                    MissedChanges.Dismiss(f.Service, other)
                )
                .Message
        );
        Assert.True(f.Service.Rows.ContainsKey(other));
        var running = Job(EventRegistrations.RecordHandler, 20);
        Assert.Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
            MissedChanges.Dismiss(f.Service, running)
        );
        Assert.True(f.Service.Rows.ContainsKey(running));
        // Dataverse refuses the delete for a caller without the privilege: a plain sentence,
        // and nothing runs after the refusal.
        var refused = Job(EventRegistrations.MembershipHandler, 31);
        Assert.Equal(
            "You don't have permission to remove system jobs.",
            Assert
                .Throws<Ascentix.Documents.Conditions.EvaluationBlockedException>(() =>
                    f.Service.Transaction(() =>
                        MissedChanges.Dismiss(new DeleteDenied(f.Service), refused)
                    )
                )
                .Message
        );
        Assert.True(f.Service.Rows.ContainsKey(refused));
    }

    /// <summary>A caller whose deletes Dataverse refuses for want of the privilege.</summary>
    private sealed class DeleteDenied : IOrganizationService
    {
        private readonly DurableWorkerTests.MemoryService inner;

        public DeleteDenied(DurableWorkerTests.MemoryService inner) => this.inner = inner;

        public void Delete(string entityName, Guid id) =>
            throw inner.Refuse(
                new System.ServiceModel.FaultException<OrganizationServiceFault>(
                    new OrganizationServiceFault
                    {
                        ErrorCode = unchecked((int)0x80040220),
                        Message = "Principal user is missing prvDeleteAsyncOperation privilege.",
                    }
                )
            );

        public Entity Retrieve(
            string entityName,
            Guid id,
            Microsoft.Xrm.Sdk.Query.ColumnSet columnSet
        ) => inner.Retrieve(entityName, id, columnSet);

        public EntityCollection RetrieveMultiple(Microsoft.Xrm.Sdk.Query.QueryBase query) =>
            inner.RetrieveMultiple(query);

        public OrganizationResponse Execute(OrganizationRequest request) => inner.Execute(request);

        public Guid Create(Entity entity) => inner.Create(entity);

        public void Update(Entity entity) => inner.Update(entity);

        public void Associate(
            string entityName,
            Guid entityId,
            Relationship relationship,
            EntityReferenceCollection relatedEntities
        ) => inner.Associate(entityName, entityId, relationship, relatedEntities);

        public void Disassociate(
            string entityName,
            Guid entityId,
            Relationship relationship,
            EntityReferenceCollection relatedEntities
        ) => inner.Disassociate(entityName, entityId, relationship, relatedEntities);
    }
}
