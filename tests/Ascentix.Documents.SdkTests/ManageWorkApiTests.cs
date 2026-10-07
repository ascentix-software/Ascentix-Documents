using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Plugins;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

/// <summary>Read commands write nothing; write commands reach the guarded transport (spec 7.2 ApiWriteGuardTests row).</summary>
public sealed class ManageWorkApiTests
{
    private sealed class Counting : IOrganizationService
    {
        internal readonly IOrganizationService Inner;
        internal int Writes;
        internal bool Deny;

        internal Counting(IOrganizationService inner) => Inner = inner;

        private void Write()
        {
            Writes++;
            if (Deny)
                throw new InvalidOperationException("This command must not write.");
        }

        public Guid Create(Entity entity)
        {
            Write();
            return Inner.Create(entity);
        }

        public Entity Retrieve(string name, Guid id, ColumnSet columns) =>
            Inner.Retrieve(name, id, columns);

        public EntityCollection RetrieveMultiple(QueryBase query) => Inner.RetrieveMultiple(query);

        public void Update(Entity entity)
        {
            Write();
            Inner.Update(entity);
        }

        public void Delete(string name, Guid id)
        {
            Write();
            Inner.Delete(name, id);
        }

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            if (request is CreateRequest || request is UpdateRequest || request is DeleteRequest)
                Write();
            return Inner.Execute(request);
        }

        public void Associate(string a, Guid b, Relationship c, EntityReferenceCollection d) =>
            throw new NotSupportedException();

        public void Disassociate(string a, Guid b, Relationship c, EntityReferenceCollection d) =>
            throw new NotSupportedException();
    }

    private static WorkerResult Call(
        DurableWorkerTests.Fixture f,
        IOrganizationService service,
        WorkerRequest request
    ) =>
        f.Service.Transaction(() =>
            JsonWire.Read<WorkerResult>(
                ApiHarness.Invoke(
                    new ManageWorkApi(),
                    "asx_ManageWork",
                    service,
                    request,
                    Guid.NewGuid()
                )
            )
        );

    private static DurableWorkerTests.Fixture Setup()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        RuntimeSeed.Seed(f.Service, Guid.NewGuid(), "account");
        // Counts come from returntotalrecordcount; an aggregate would fail the test.
        f.Service.FetchHook = fetch =>
            fetch.Query.Contains("aggregate=")
                ? throw new InvalidOperationException("A plug-in must not run an aggregate.")
                : new EntityCollection { TotalRecordCount = 3 };
        return f;
    }

    [Theory]
    [InlineData("Summary")]
    [InlineData("CountRecords")]
    [InlineData("ListProblems")]
    public void ReadCommandsWriteNothing(string command)
    {
        var f = Setup();
        var service = new Counting(f.Service) { Deny = true };
        var result = Call(
            f,
            service,
            new WorkerRequest
            {
                Command = command,
                TemplateId = f.TemplateId,
                List = "BlockedJobs",
            }
        );
        Assert.NotEqual("", result.Status);
        Assert.Equal(0, service.Writes);
    }

    [Fact]
    public void SummaryAndListProblemsAnswerThroughTheApi()
    {
        var f = Setup();
        var summary = Call(f, f.Service, new WorkerRequest { Command = "Summary" });
        Assert.Equal("Summary", summary.Status);
        Assert.Equal(3, summary.Summary!.BlockedJobs);
        var job = f.Operation;
        job.Status = "Blocked";
        job.ErrorCode = "WorkerFailed";
        f.Store.Create("asx_operation", job);
        var page = Call(
            f,
            f.Service,
            new WorkerRequest { Command = "ListProblems", List = "BlockedJobs" }
        );
        Assert.Equal("Page", page.Status);
        Assert.Equal(job.Key, Assert.Single(page.Problems).Key);
        Assert.Null(page.Next);
    }

    [Fact]
    public void ListProblemsWritesOnlyToPickUpA0103LostCreate()
    {
        var f = Setup();
        f.Store.Create(
            "asx_operation",
            new LibrarySetup
            {
                Key = "librarycreate:old",
                Status = "RecoveryRequired",
                SiteId = f.SiteId,
                Name = "Old",
                Mutation = "CreateLibrary",
                ExternalSubmitted = true,
            }
        );
        var service = new Counting(f.Service);
        Call(f, service, new WorkerRequest { Command = "ListProblems", List = "BlockedJobs" });
        Assert.Equal(1, service.Writes);
        Call(f, service, new WorkerRequest { Command = "ListProblems", List = "BlockedJobs" });
        Assert.Equal(1, service.Writes);
    }

    [Theory]
    [InlineData("StartTemplateRun")]
    [InlineData("RerunRecord")]
    public void WriteCommandsPassTheGuardedTransport(string command)
    {
        var f = Setup();
        var result = Call(
            f,
            f.Service,
            new WorkerRequest
            {
                Command = command,
                TemplateId = f.TemplateId,
                Table = "account",
                RecordId = f.RecordId,
                RequestId = Guid.NewGuid(),
            }
        );
        Assert.Contains(result.Status, new[] { "Pending", "Queued" });
        Assert.Contains(
            f.Service.Rows.Values,
            r =>
                r.LogicalName == "asx_outbox"
                && r.GetAttributeValue<string>("asx_status") == "Pending"
        );
    }

    [Fact]
    public void RunControlsPassTheGuardedTransport()
    {
        var f = Setup();
        var started = Call(
            f,
            f.Service,
            new WorkerRequest
            {
                Command = "StartTemplateRun",
                TemplateId = f.TemplateId,
                RequestId = Guid.NewGuid(),
            }
        );
        Assert.Equal(
            "Paused",
            Call(
                f,
                f.Service,
                new WorkerRequest { Command = "PauseTemplateRun", Key = started.Key }
            ).Status
        );
        Assert.Equal(
            "Pending",
            Call(
                f,
                f.Service,
                new WorkerRequest { Command = "ResumeTemplateRun", Key = started.Key }
            ).Status
        );
        Assert.Equal(
            "Cancelled",
            Call(
                f,
                f.Service,
                new WorkerRequest { Command = "CancelTemplateRun", Key = started.Key }
            ).Status
        );
    }

    [Fact]
    public void DismissCaptureJobPassesTheGuardedTransport()
    {
        var f = Setup();
        var type = Guid.NewGuid();
        var step = Guid.NewGuid();
        var job = Guid.NewGuid();
        f.Service.Seed(
            new Entity("plugintype", type) { ["typename"] = EventRegistrations.RecordHandler }
        );
        f.Service.Seed(
            new Entity("sdkmessageprocessingstep", step)
            {
                ["eventhandler"] = new EntityReference("plugintype", type),
            }
        );
        f.Service.Seed(
            new Entity("asyncoperation", job)
            {
                ["statuscode"] = new OptionSetValue(31),
                ["owningextensionid"] = new EntityReference("sdkmessageprocessingstep", step),
            }
        );
        Assert.Equal(
            "Dismissed",
            Call(
                f,
                f.Service,
                new WorkerRequest { Command = "DismissCaptureJob", JobId = job }
            ).Status
        );
        Assert.False(f.Service.Rows.ContainsKey(job));
    }

    [Fact]
    public void RecheckSetupAndResolveSetupPassTheGuardedTransport()
    {
        var f = new LibraryProvisioningTests.Fixture();
        var key = f.Lost();
        f.Run(key);
        CatalogResult Catalog(CatalogRequest request) =>
            f.Service.Transaction(() =>
                JsonWire.Read<CatalogResult>(
                    ApiHarness.Invoke(
                        new CatalogAdminApi(),
                        "asx_CatalogAdmin",
                        f.Service,
                        request,
                        Guid.NewGuid()
                    )
                )
            );
        Assert.Equal(
            "Reconciling",
            Catalog(new CatalogRequest { Command = "RecheckSetup", Key = key }).Status
        );
        f.Run(key);
        var resolved = Catalog(
            new CatalogRequest
            {
                Command = "ResolveSetup",
                Key = key,
                Choice = "UseLibrary",
                ListId = f.List,
                RowVersion = f.Worker.Inspect(key).RowVersion,
            }
        );
        Assert.Equal("Pending", resolved.Status);
    }
}

/// <summary>
/// Runs a custom API plug-in the way Dataverse does: a context carrying the JSON Request and the
/// caller, and a service factory. The plug-in wraps that service in ApiWriteService itself, so a
/// write command that forgets InternalWrite fails here (spec 7.2 ApiWriteGuardTests row).
/// </summary>
internal static class ApiHarness
{
    private sealed class Provider : IServiceProvider, IOrganizationServiceFactory
    {
        private readonly IPluginExecutionContext context;
        private readonly IOrganizationService service;

        internal Provider(IPluginExecutionContext context, IOrganizationService service)
        {
            this.context = context;
            this.service = service;
        }

        public object GetService(Type type) =>
            type == typeof(IPluginExecutionContext) ? context : this;

        public IOrganizationService CreateOrganizationService(Guid? userId) => service;
    }

    /// <returns>The plug-in's Result string.</returns>
    internal static string Invoke<TRequest>(
        IPlugin api,
        string message,
        IOrganizationService service,
        TRequest request,
        Guid user
    )
    {
        var context = GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["Stage"] = 30,
                ["Mode"] = 0,
                ["IsInTransaction"] = true,
                ["UserId"] = user,
                ["CorrelationId"] = Guid.NewGuid(),
                ["MessageName"] = message,
                ["InputParameters"] = new ParameterCollection
                {
                    ["Request"] = JsonWire.Write(request),
                },
                ["OutputParameters"] = new ParameterCollection(),
                ["SharedVariables"] = new ParameterCollection(),
            }
        );
        api.Execute(new Provider(context, service));
        return (string)context.OutputParameters["Result"];
    }
}
