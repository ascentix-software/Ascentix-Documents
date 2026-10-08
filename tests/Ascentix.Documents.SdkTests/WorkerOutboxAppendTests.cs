using System;
using System.Collections.Generic;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Plugins;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class WorkerOutboxAppendTests
{
    [Theory]
    [InlineData("none")]
    [InlineData("Create")]
    [InlineData("Update")]
    [InlineData("Upsert")]
    public void WorkerCanAppendWithoutAncestryEvidence(string parentMessage)
    {
        var service = new GuardedService(parentMessage);
        new DocumentStore(service).Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = "event:fixture",
                Table = "account",
                RecordId = Guid.NewGuid(),
                TemplateId = Guid.NewGuid(),
            }
        );
        Assert.Equal(1, service.Creates);
        Assert.Equal(1, service.RuntimeReads);
    }

    [Theory]
    [InlineData("caller")]
    [InlineData("missing-worker")]
    [InlineData("update")]
    [InlineData("delete")]
    [InlineData("other-table")]
    [InlineData("async")]
    [InlineData("outside-transaction")]
    [InlineData("key")]
    [InlineData("id")]
    [InlineData("schema")]
    [InlineData("status")]
    [InlineData("payload")]
    [InlineData("null-payload")]
    public void WorkerTrustDoesNotAuthorizeOtherWritersOrInvalidState(string variant)
    {
        var service = new GuardedService("Upsert");
        var document = new OutboxDocument { Key = "event:fixture" };
        if (variant == "schema")
            document.SchemaVersion = 2;
        var row = new Entity("asx_outbox", DocumentStore.StableId("asx_outbox:" + document.Key))
        {
            ["asx_key"] = DocumentStore.Hash(document.Key),
            ["asx_status"] = document.Status,
            ["asx_payload"] = JsonWire.Write(document),
        };
        var values = service.Values(row);
        if (variant == "caller")
        {
            values["UserId"] = Guid.NewGuid();
            values["SharedVariables"] = new ParameterCollection
            {
                [DocumentWorkerApi.InternalWrite] = true,
                ["tag"] = "Ascentix.Documents.RecordEvent.v1:forged",
            };
        }
        if (variant == "missing-worker")
            values["UserId"] = Guid.Empty;
        if (variant == "update" || variant == "delete")
            values["MessageName"] = variant == "update" ? "Update" : "Delete";
        if (variant == "other-table")
            values["PrimaryEntityName"] = row.LogicalName = "asx_operation";
        if (variant == "async")
            values["Mode"] = 1;
        if (variant == "outside-transaction")
            values["IsInTransaction"] = false;
        if (variant == "key")
            row["asx_key"] = "wrong";
        if (variant == "id")
            row.Id = Guid.NewGuid();
        if (variant == "status")
            row["asx_status"] = "Applied";
        if (variant == "payload")
            row["asx_payload"] = "not json";
        if (variant == "null-payload")
            row["asx_payload"] = "null";
        service.Context = GuardTests.ContextProxy.Create(values);
        Assert.ThrowsAny<Exception>(() => new StateGuard().Execute(service));
    }

    private sealed class GuardedService
        : IOrganizationService,
            IServiceProvider,
            IOrganizationServiceFactory
    {
        private readonly Guid worker = Guid.NewGuid();
        private readonly string parentMessage;
        public IPluginExecutionContext Context = null!;
        public int Creates;
        public int RuntimeReads;

        public GuardedService(string parentMessage) => this.parentMessage = parentMessage;

        public Dictionary<string, object> Values(Entity row)
        {
            var values = new Dictionary<string, object>
            {
                ["Stage"] = 20,
                ["Mode"] = 0,
                ["IsInTransaction"] = true,
                ["MessageName"] = "Create",
                ["PrimaryEntityName"] = "asx_outbox",
                ["UserId"] = worker,
                ["CorrelationId"] = Guid.NewGuid(),
                ["InputParameters"] = new ParameterCollection { ["Target"] = row },
                ["SharedVariables"] = new ParameterCollection(),
            };
            if (parentMessage != "none")
            {
                var ancestor = GuardTests.ContextProxy.Create(
                    new Dictionary<string, object>
                    {
                        ["Stage"] = 30,
                        ["MessageName"] = parentMessage,
                        ["PrimaryEntityName"] = "account",
                        ["UserId"] = Guid.NewGuid(),
                        ["SharedVariables"] = new ParameterCollection(),
                    }
                );
                values["ParentContext"] = GuardTests.ContextProxy.Create(
                    new Dictionary<string, object>
                    {
                        ["Stage"] = 30,
                        ["Mode"] = 0,
                        ["IsInTransaction"] = true,
                        ["MessageName"] = "Create",
                        ["PrimaryEntityName"] = "asx_outbox",
                        ["UserId"] = worker,
                        ["CorrelationId"] = values["CorrelationId"],
                        ["ParentContext"] = ancestor,
                        ["SharedVariables"] = new ParameterCollection(),
                    }
                );
            }
            return values;
        }

        public object GetService(Type type) =>
            type == typeof(IPluginExecutionContext) ? Context : this;

        public IOrganizationService CreateOrganizationService(Guid? userId) => this;

        public Guid Create(Entity row)
        {
            Context = GuardTests.ContextProxy.Create(Values(row));
            new StateGuard().Execute(this);
            Creates++;
            return row.Id;
        }

        public EntityCollection RetrieveMultiple(QueryBase query)
        {
            var entity = Assert.IsType<QueryExpression>(query).EntityName;
            if (entity == "asx_runtimetable")
                return new EntityCollection();
            Assert.Equal("asx_runtime", entity);
            RuntimeReads++;
            return new EntityCollection(
                new[]
                {
                    new Entity("asx_runtime", Guid.NewGuid())
                    {
                        ["asx_workeruserid"] = worker.ToString(),
                        ["asx_allowedtables"] = "[\"account\"]",
                        ["asx_sharepointhosts"] = "[\"ascentixsoftware.sharepoint.com\"]",
                    },
                }
            );
        }

        public Entity Retrieve(string table, Guid id, ColumnSet columns) =>
            throw new NotSupportedException();

        public OrganizationResponse Execute(OrganizationRequest request) =>
            throw new NotSupportedException();

        public void Update(Entity row) => throw new NotSupportedException();

        public void Delete(string table, Guid id) => throw new NotSupportedException();

        public void Associate(
            string table,
            Guid id,
            Relationship relationship,
            EntityReferenceCollection rows
        ) => throw new NotSupportedException();

        public void Disassociate(
            string table,
            Guid id,
            Relationship relationship,
            EntityReferenceCollection rows
        ) => throw new NotSupportedException();
    }
}
