using System;
using System.Collections.Generic;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Plugins;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class GuardTests
{
    [Fact]
    public void AuthorsCanDeletePublishedTemplateAndChangeItsAvailability()
    {
        var deletion = new Setup(
            "asx_template",
            "Delete",
            new EntityReference("asx_template", Guid.NewGuid())
        );
        deletion.Service.Existing["asx_publishedrevisionid"] = new EntityReference(
            "asx_revision",
            Guid.NewGuid()
        );
        new RevisionGuard().Execute(deletion);
        var update = new Setup(
            "asx_template",
            "Update",
            new Entity("asx_template")
            {
                ["asx_disabled"] = true,
                ["asx_endsutc"] = DateTime.UtcNow,
            }
        );
        update.Service.Existing["asx_publishedrevisionid"] = new EntityReference(
            "asx_revision",
            Guid.NewGuid()
        );
        new RevisionGuard().Execute(update);
    }

    [Fact]
    public void TemplateCascadeCanDeletePublishedChildConfiguration()
    {
        var user = Guid.NewGuid();
        var parent = ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["MessageName"] = "Delete",
                ["PrimaryEntityName"] = "asx_template",
                ["UserId"] = user,
            }
        );
        var setup = new Setup(
            "asx_folder",
            "Delete",
            new EntityReference("asx_folder", Guid.NewGuid()),
            parent,
            user
        );
        setup.Service.RevisionStatus = "Published";
        new RevisionGuard().Execute(setup);
        Assert.Empty(setup.Service.Requests);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PublicationRecognizesPlatformFramesButNotADifferentCaller(bool differentCaller)
    {
        var user = Guid.NewGuid();
        var api = ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["MessageName"] = "asx_PublishTemplate",
                ["Stage"] = 30,
                ["IsInTransaction"] = true,
                ["UserId"] = differentCaller ? Guid.NewGuid() : user,
            }
        );
        var parent = ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["MessageName"] = "Update",
                ["PrimaryEntityName"] = "asx_revision",
                ["Stage"] = 30,
                ["IsInTransaction"] = true,
                ["UserId"] = user,
                ["ParentContext"] = api,
            }
        );
        var setup = new Setup(
            "asx_revision",
            "Update",
            new Entity("asx_revision") { ["asx_status"] = "Published" },
            parent,
            user
        );
        setup.Service.Existing["asx_status"] = "Draft";
        if (differentCaller)
            Assert.Throws<InvalidPluginExecutionException>(() =>
                new RevisionGuard().Execute(setup)
            );
        else
            new RevisionGuard().Execute(setup);
    }

    [Fact]
    public void DirectPublicationIsRejected()
    {
        var setup = new Setup(
            "asx_revision",
            "Update",
            new Entity("asx_revision") { ["asx_status"] = "Published" }
        );
        setup.Service.Existing["asx_status"] = "Draft";
        Assert.Throws<InvalidPluginExecutionException>(() => new RevisionGuard().Execute(setup));
    }

    [Fact]
    public void NewRevisionMustBeDraft()
    {
        var setup = new Setup(
            "asx_revision",
            "Create",
            new Entity("asx_revision") { ["asx_status"] = "Published" }
        );
        Assert.Throws<InvalidPluginExecutionException>(() => new RevisionGuard().Execute(setup));
    }

    [Fact]
    public void PublishedRevisionCannotBeDeleted()
    {
        var setup = new Setup(
            "asx_revision",
            "Delete",
            new EntityReference("asx_revision", Guid.NewGuid())
        );
        setup.Service.Existing["asx_status"] = "Published";
        Assert.Throws<InvalidPluginExecutionException>(() => new RevisionGuard().Execute(setup));
    }

    [Fact]
    public void DirectPublishedPointerIsRejected()
    {
        var setup = new Setup(
            "asx_template",
            "Update",
            new Entity("asx_template")
            {
                ["asx_publishedrevisionid"] = new EntityReference("asx_revision", Guid.NewGuid()),
            }
        );
        Assert.Throws<InvalidPluginExecutionException>(() => new RevisionGuard().Execute(setup));
    }

    [Fact]
    public void PublishedChildIsImmutable()
    {
        var setup = new Setup(
            "asx_folder",
            "Update",
            new Entity("asx_folder") { ["asx_expression"] = "Changed" }
        );
        setup.Service.Existing["asx_revisionid"] = new EntityReference(
            "asx_revision",
            setup.Service.RevisionId
        );
        setup.Service.RevisionStatus = "Published";
        Assert.Throws<InvalidPluginExecutionException>(() => new RevisionGuard().Execute(setup));
    }

    [Fact]
    public void DraftChildTouchesAggregateWithExpectedVersion()
    {
        var setup = new Setup(
            "asx_folder",
            "Update",
            new Entity("asx_folder") { ["asx_expression"] = "Changed" }
        );
        setup.Service.Existing["asx_revisionid"] = new EntityReference(
            "asx_revision",
            setup.Service.RevisionId
        );
        new RevisionGuard().Execute(setup);
        var update = Assert.IsType<UpdateRequest>(Assert.Single(setup.Service.Requests));
        Assert.Equal(ConcurrencyBehavior.IfRowVersionMatches, update.ConcurrencyBehavior);
        Assert.Equal("73", update.Target.RowVersion);
        Assert.Equal(setup.Service.RevisionId, update.Target.Id);
        Assert.True(update.Target.Contains("asx_editstamp"));
    }

    [Fact]
    public void ChildCannotBeReparentedAcrossRevisions()
    {
        var setup = new Setup(
            "asx_folder",
            "Update",
            new Entity("asx_folder")
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", Guid.NewGuid()),
            }
        );
        setup.Service.Existing["asx_revisionid"] = new EntityReference(
            "asx_revision",
            setup.Service.RevisionId
        );
        Assert.Throws<InvalidPluginExecutionException>(() => new RevisionGuard().Execute(setup));
    }

    [Fact]
    public void JsonReaderEnforcesDepthQuota()
    {
        var nested = "{\"All\":true,\"Conditions\":[],\"Groups\":[]}";
        for (int i = 0; i < 40; i++)
            nested = "{\"All\":true,\"Conditions\":[],\"Groups\":[" + nested + "]}";
        Assert.ThrowsAny<Exception>(() => JsonWire.Read<GroupDto>(nested));
    }

    [Fact]
    public void StateRowsRejectDirectEditsEvenWithTableRights()
    {
        var setup = new Setup("asx_operation", "Update", new Entity("asx_operation"));
        Assert.Throws<InvalidPluginExecutionException>(() => new StateGuard().Execute(setup));
    }

    [Fact]
    public void StateGuardRequiresMarkedSameIdentityWorkerApi()
    {
        var user = Guid.NewGuid();
        var parent = ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["MessageName"] = "asx_DocumentWorker",
                ["UserId"] = user,
                ["SharedVariables"] = new ParameterCollection
                {
                    [DocumentWorkerApi.InternalWrite] = true,
                },
            }
        );
        new StateGuard().Execute(
            new Setup("asx_operation", "Update", new Entity("asx_operation"), parent, user)
        );
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new StateGuard().Execute(
                new Setup(
                    "asx_operation",
                    "Update",
                    new Entity("asx_operation"),
                    parent,
                    Guid.NewGuid()
                )
            )
        );
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new StateGuard().Execute(
                new Setup("asx_attempt", "Update", new Entity("asx_attempt"), parent, user)
            )
        );
    }

    [Fact]
    public void CatalogCannotForgeAppliedPolicyOrChangePhysicalDestination()
    {
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new CatalogGuard().Execute(
                new Setup(
                    "asx_library",
                    "Update",
                    new Entity("asx_library") { ["asx_policyapplied"] = true }
                )
            )
        );
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new CatalogGuard().Execute(
                new Setup(
                    "asx_library",
                    "Update",
                    new Entity("asx_library") { ["asx_listid"] = Guid.NewGuid().ToString() }
                )
            )
        );
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new CatalogGuard().Execute(
                new Setup(
                    "asx_library",
                    "Create",
                    new Entity("asx_library") { ["asx_policyapplied"] = true }
                )
            )
        );
    }

    [Fact]
    public void SecurityAndRuntimeConfigurationRequireServerApiParent()
    {
        foreach (var table in new[] { "asx_policy", "asx_teamregistration", "asx_runtime" })
            Assert.Throws<InvalidPluginExecutionException>(() =>
                new CatalogGuard().Execute(new Setup(table, "Update", new Entity(table)))
            );
        var user = Guid.NewGuid();
        var parent = ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["MessageName"] = "asx_SecurityAdmin",
                ["UserId"] = user,
                ["SharedVariables"] = new ParameterCollection
                {
                    [DocumentWorkerApi.InternalWrite] = true,
                },
            }
        );
        new CatalogGuard().Execute(
            new Setup("asx_policy", "Update", new Entity("asx_policy"), parent, user)
        );
        new CatalogGuard().Execute(
            new Setup(
                "asx_teamregistration",
                "Update",
                new Entity("asx_teamregistration") { ["asx_status"] = "Revoking" },
                parent,
                user
            )
        );
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new CatalogGuard().Execute(
                new Setup("asx_policy", "Update", new Entity("asx_policy"), parent, Guid.NewGuid())
            )
        );
    }

    [Fact]
    public void BulkDirectStateAndCatalogWritesFailClosed()
    {
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new StateGuard().Execute(
                new Setup("asx_operation", "CreateMultiple", new Entity("asx_operation"))
            )
        );
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new CatalogGuard().Execute(
                new Setup("asx_library", "UpdateMultiple", new Entity("asx_library"))
            )
        );
    }

    [Fact]
    public void CatalogApiMarksContextBeforeServiceFactoryCapturesIt()
    {
        var setup = new CatalogSetup();
        Assert.Throws<NotSupportedException>(() => new CatalogAdminApi().Execute(setup));
        Assert.True(setup.MarkedAtServiceCreation);
    }

    private sealed class CatalogSetup : IServiceProvider, IOrganizationServiceFactory
    {
        private readonly ParameterCollection shared = new ParameterCollection();
        private readonly Guid user = Guid.NewGuid();
        public bool MarkedAtServiceCreation;

        public object GetService(Type type) =>
            type == typeof(IPluginExecutionContext)
                ? ContextProxy.Create(
                    new Dictionary<string, object>
                    {
                        ["UserId"] = user,
                        ["IsInTransaction"] = true,
                        ["SharedVariables"] = shared,
                        ["InputParameters"] = new ParameterCollection
                        {
                            ["Request"] = "{\"Command\":\"Inspect\",\"Key\":\"catalogprobe:test\"}",
                        },
                    }
                )
                : this;

        public IOrganizationService CreateOrganizationService(Guid? id)
        {
            Assert.Equal(user, id);
            MarkedAtServiceCreation =
                shared.Contains(DocumentWorkerApi.InternalWrite)
                && shared[DocumentWorkerApi.InternalWrite] is bool marker
                && marker;
            return new FakeService();
        }
    }

    private sealed class Setup : IServiceProvider, IOrganizationServiceFactory
    {
        public FakeService Service { get; } = new FakeService();
        private readonly IPluginExecutionContext context;

        public Setup(
            string table,
            string message,
            object target,
            IPluginExecutionContext? parent = null,
            Guid? user = null
        )
        {
            if (target is Entity entity)
                entity.Id = Guid.NewGuid();
            var properties = new Dictionary<string, object>
            {
                ["Stage"] = 20,
                ["IsInTransaction"] = true,
                ["UserId"] = user ?? Guid.NewGuid(),
                ["PrimaryEntityName"] = table,
                ["MessageName"] = message,
                ["InputParameters"] = new ParameterCollection { ["Target"] = target },
            };
            if (parent != null)
                properties["ParentContext"] = parent;
            context = ContextProxy.Create(properties);
            Service.Existing = new Entity(
                table,
                target is Entity e ? e.Id : ((EntityReference)target).Id
            );
        }

        public object GetService(Type type) =>
            type == typeof(IPluginExecutionContext) ? context : this;

        public IOrganizationService CreateOrganizationService(Guid? userId)
        {
            Assert.NotNull(userId);
            Assert.NotEqual(Guid.Empty, userId);
            return Service;
        }
    }

    private static IPluginExecutionContext RuntimeFrame(
        string message,
        Guid user,
        Guid correlation,
        bool marked,
        IPluginExecutionContext? parent = null
    ) =>
        ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["MessageName"] = message,
                ["UserId"] = user,
                ["CorrelationId"] = correlation,
                ["IsInTransaction"] = true,
                ["Stage"] = 30,
                ["Mode"] = 0,
                ["ParentContext"] = parent!,
                ["SharedVariables"] = marked
                    ? new ParameterCollection { [DocumentWorkerApi.InternalWrite] = true }
                    : new ParameterCollection(),
            }
        );

    private static void RunRuntimeTableGuard(
        string message,
        Guid user,
        Guid correlation,
        IPluginExecutionContext? parent
    ) =>
        new CatalogGuard().Execute(
            new SingleContextProvider(
                ContextProxy.Create(
                    new Dictionary<string, object>
                    {
                        ["Stage"] = 20,
                        ["Mode"] = 0,
                        ["IsInTransaction"] = true,
                        ["MessageName"] = message,
                        ["PrimaryEntityName"] = "asx_runtimetable",
                        ["PrimaryEntityId"] = Guid.NewGuid(),
                        ["UserId"] = user,
                        ["CorrelationId"] = correlation,
                        ["ParentContext"] = parent!,
                        ["InputParameters"] = new ParameterCollection(),
                        ["SharedVariables"] = new ParameterCollection(),
                    }
                )
            )
        );

    [Theory]
    [InlineData("Create", true)]
    [InlineData("Delete", true)]
    [InlineData("Create", false)]
    [InlineData("Delete", false)]
    public void RuntimeTableRowsAreWritableOnlyFromRuntimeAdministration(
        string message,
        bool fromApi
    )
    {
        var user = Guid.NewGuid();
        var correlation = Guid.NewGuid();
        var parent = RuntimeFrame(
            fromApi ? "asx_RuntimeAdmin" : "Create",
            user,
            correlation,
            fromApi
        );
        if (fromApi)
            RunRuntimeTableGuard(message, user, correlation, parent);
        else
            Assert.Throws<InvalidPluginExecutionException>(() =>
                RunRuntimeTableGuard(message, user, correlation, parent)
            );
    }

    [Theory]
    [InlineData("Create", "allowed")]
    [InlineData("Delete", "allowed")]
    [InlineData("Create", "unmarked")]
    [InlineData("Delete", "unmarked")]
    [InlineData("Create", "otheruser")]
    [InlineData("Delete", "otheruser")]
    [InlineData("Create", "othercorrelation")]
    [InlineData("Delete", "othercorrelation")]
    public void RuntimeTableRowsAreRecognizedThroughNestedPlatformFrames(
        string message,
        string scenario
    )
    {
        var user = Guid.NewGuid();
        var correlation = Guid.NewGuid();
        var api = RuntimeFrame(
            "asx_RuntimeAdmin",
            scenario == "otheruser" ? Guid.NewGuid() : user,
            scenario == "othercorrelation" ? Guid.NewGuid() : correlation,
            scenario != "unmarked"
        );
        var intermediate = RuntimeFrame("Create", user, correlation, false, api);
        if (scenario == "allowed")
            RunRuntimeTableGuard(message, user, correlation, intermediate);
        else
            Assert.Throws<InvalidPluginExecutionException>(() =>
                RunRuntimeTableGuard(message, user, correlation, intermediate)
            );
    }

    private sealed class SingleContextProvider : IServiceProvider
    {
        private readonly IPluginExecutionContext context;

        public SingleContextProvider(IPluginExecutionContext context) => this.context = context;

        public object GetService(Type type) =>
            type == typeof(IPluginExecutionContext)
                ? context
                : throw new InvalidOperationException("No service calls expected.");
    }

    internal sealed class ContextProxy : RealProxy
    {
        private readonly IDictionary<string, object> properties;

        private ContextProxy(IDictionary<string, object> properties)
            : base(typeof(IPluginExecutionContext))
        {
            this.properties = properties;
        }

        public static IPluginExecutionContext Create(IDictionary<string, object> properties) =>
            (IPluginExecutionContext)new ContextProxy(properties).GetTransparentProxy();

        public override IMessage Invoke(IMessage message)
        {
            var call = (IMethodCallMessage)message;
            properties.TryGetValue(call.MethodName.Substring(4), out var value);
            if (
                value == null
                && call.MethodBase is System.Reflection.MethodInfo method
                && method.ReturnType.IsValueType
            )
                value = Activator.CreateInstance(method.ReturnType);
            return new ReturnMessage(value, null, 0, call.LogicalCallContext, call);
        }
    }

    private sealed class FakeService : IOrganizationService
    {
        public Entity Existing = null!;
        public Guid RevisionId { get; } = Guid.NewGuid();
        public string RevisionStatus = "Draft";
        public List<OrganizationRequest> Requests { get; } = new List<OrganizationRequest>();

        public Entity Retrieve(string name, Guid id, ColumnSet columns) =>
            id == RevisionId
                ? new Entity("asx_revision", id)
                {
                    RowVersion = "73",
                    ["asx_status"] = RevisionStatus,
                }
                : Existing;

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            Requests.Add(request);
            return new UpdateResponse();
        }

        public Guid Create(Entity entity) => throw new NotSupportedException();

        public void Update(Entity entity) => throw new NotSupportedException();

        public void Delete(string name, Guid id) => throw new NotSupportedException();

        public EntityCollection RetrieveMultiple(QueryBase query) =>
            throw new NotSupportedException();

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
