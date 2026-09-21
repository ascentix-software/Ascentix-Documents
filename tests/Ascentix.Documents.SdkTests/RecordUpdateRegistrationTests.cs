using System;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class RecordUpdateRegistrationTests
{
    [Fact]
    public void RuntimeSaveTogglesOnlyOwnedUpdateStepsAcrossSourceTables()
    {
        var s = new Service();
        var account = s.Step("account", "Update");
        var bulk = s.Step("account", "UpdateMultiple");
        var contact = s.Step("contact", "Update");
        var create = s.Step("account", "Create");
        var delete = s.Step("account", "Delete");
        var foreign = s.Step("account", "Update", owned: false);
        var otherHandler = s.Step("account", "Update", correctHandler: false);
        var request = RuntimeAdministration.Execute(s, new RuntimeRequest(), true);
        Assert.False(request.ProcessRecordUpdates);
        request.Command = "Save";
        var saved = RuntimeAdministration.Execute(s, request, true);
        foreach (var id in new[] { account, bulk, contact })
            Assert.Equal(1, s.State(id));
        foreach (var id in new[] { create, delete, foreign, otherHandler })
            Assert.Equal(0, s.State(id));
        int writes = s.StepWrites;
        saved.Command = "Save";
        saved = RuntimeAdministration.Execute(s, saved, true);
        Assert.Equal(writes, s.StepWrites);
        saved.Command = "Save";
        saved.ProcessRecordUpdates = true;
        saved = RuntimeAdministration.Execute(s, saved, true);
        Assert.True(saved.ProcessRecordUpdates);
        foreach (var id in new[] { account, bulk, contact })
            Assert.Equal(0, s.State(id));
        saved.Command = "Save";
        saved.Tables = new[] { "account" };
        RuntimeAdministration.Execute(s, saved, true);
        Assert.Equal(1, s.State(contact));
        Assert.Equal(0, s.State(account));
    }

    [Theory]
    [InlineData("worker")]
    [InlineData("stage")]
    [InlineData("filter")]
    [InlineData("renamed")]
    [InlineData("duplicate")]
    public void InvalidRegistrationPreventsAllActivation(string variant)
    {
        var s = new Service();
        var first = s.Step("account", "Update", active: false);
        var bad = s.Step("contact", "Update", active: false);
        var row = s.Memory.Rows[bad];
        if (variant == "worker")
            row["impersonatinguserid"] = new EntityReference("systemuser", Guid.NewGuid());
        if (variant == "stage")
            row["stage"] = new OptionSetValue(20);
        if (variant == "filter")
            row["sdkmessagefilterid"] = null;
        if (variant == "renamed")
            row["name"] = "Renamed update step";
        if (variant == "duplicate")
            s.Step("account", "Update", active: false);
        Assert.Throws<InvalidPluginExecutionException>(() =>
            RecordUpdateRegistrations.Apply(s, s.Worker, new[] { "account", "contact" }, true)
        );
        Assert.Equal(0, s.StepWrites);
        Assert.Equal(1, s.State(first));
    }

    [Fact]
    public void MissingStepsCannotBeReportedAsEnabledAndReadbackFailuresSurface()
    {
        var s = new Service();
        Assert.Throws<InvalidPluginExecutionException>(() =>
            RecordUpdateRegistrations.Apply(s, s.Worker, new[] { "account" }, true)
        );
        s.Step("account", "Update");
        s.IgnoreWrites = true;
        Assert.Throws<InvalidPluginExecutionException>(() =>
            RecordUpdateRegistrations.Apply(s, s.Worker, new[] { "account" }, false)
        );
    }

    [Fact]
    public void StaleRuntimeSaveDoesNotChangeRegistrations()
    {
        var s = new Service();
        var step = s.Step("account", "Update");
        var request = RuntimeAdministration.Execute(s, new RuntimeRequest(), true);
        request.Command = "Save";
        request.RowVersion = "stale";
        Assert.ThrowsAny<Exception>(() => RuntimeAdministration.Execute(s, request, true));
        Assert.Equal(0, s.StepWrites);
        Assert.Equal(0, s.State(step));
    }

    private sealed class Service : IOrganizationService
    {
        internal readonly DurableWorkerTests.MemoryService Memory =
            new DurableWorkerTests.MemoryService();
        internal readonly Guid Worker = Guid.NewGuid();
        private readonly Guid handler = Guid.NewGuid();
        internal int StepWrites;
        internal bool IgnoreWrites;

        internal Service()
        {
            Memory.Seed(
                new Entity("plugintype", handler)
                {
                    ["typename"] = "Ascentix.Documents.Plugins.RecordInvalidationPlugin",
                }
            );
            Memory.Seed(new Entity("systemuser", Worker) { ["isdisabled"] = false });
            Memory.Seed(
                new Entity("asx_runtime", Guid.NewGuid())
                {
                    ["asx_name"] = "Default",
                    ["asx_workeruserid"] = Worker.ToString(),
                    ["asx_allowedtables"] = "[\"account\",\"contact\"]",
                    ["asx_sharepointhosts"] = "[\"example.sharepoint.com\"]",
                }
            );
        }

        internal Guid Step(
            string table,
            string message,
            bool owned = true,
            bool active = true,
            bool correctHandler = true
        )
        {
            var sdk = new Entity("sdkmessage", Guid.NewGuid()) { ["name"] = message };
            Memory.Seed(sdk);
            var filter = new Entity("sdkmessagefilter", Guid.NewGuid())
            {
                ["primaryobjecttypecode"] = table,
                ["sdkmessageid"] = sdk.ToEntityReference(),
            };
            Memory.Seed(filter);
            var row = new Entity("sdkmessageprocessingstep", Guid.NewGuid())
            {
                ["name"] = "Ascentix Documents: event " + message + " " + table,
                ["description"] = owned
                    ? "Owned by standalone Ascentix Documents development setup v1"
                    : "Another owner",
                ["eventhandler"] = new EntityReference(
                    "plugintype",
                    correctHandler ? handler : Guid.NewGuid()
                ),
                ["sdkmessageid"] = sdk.ToEntityReference(),
                ["sdkmessagefilterid"] = filter.ToEntityReference(),
                ["impersonatinguserid"] = new EntityReference("systemuser", Worker),
                ["stage"] = new OptionSetValue(40),
                ["mode"] = new OptionSetValue(0),
                ["statecode"] = new OptionSetValue(active ? 0 : 1),
            };
            Memory.Seed(row);
            return row.Id;
        }

        internal int State(Guid id) =>
            Memory.Rows[id].GetAttributeValue<OptionSetValue>("statecode").Value;

        public void Update(Entity entity)
        {
            Assert.Equal("sdkmessageprocessingstep", entity.LogicalName);
            StepWrites++;
            if (!IgnoreWrites)
                foreach (var pair in entity.Attributes)
                    Memory.Rows[entity.Id][pair.Key] = pair.Value;
        }

        public OrganizationResponse Execute(OrganizationRequest request) => Memory.Execute(request);

        public Guid Create(Entity entity) => Memory.Create(entity);

        public Entity Retrieve(string name, Guid id, ColumnSet columns) =>
            Memory.Retrieve(name, id, columns);

        public EntityCollection RetrieveMultiple(QueryBase query) => Memory.RetrieveMultiple(query);

        public void Delete(string name, Guid id) => throw new NotSupportedException();

        public void Associate(
            string name,
            Guid id,
            Relationship relationship,
            EntityReferenceCollection rows
        ) => throw new NotSupportedException();

        public void Disassociate(
            string name,
            Guid id,
            Relationship relationship,
            EntityReferenceCollection rows
        ) => throw new NotSupportedException();
    }
}
