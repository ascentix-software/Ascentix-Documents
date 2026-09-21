using System;
using System.Collections.Generic;
using Ascentix.Documents.Plugins;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class ApiWriteGuardTests
{
    private sealed class Provider : IServiceProvider
    {
        private readonly IPluginExecutionContext context;

        public Provider(IPluginExecutionContext context)
        {
            this.context = context;
        }

        public object GetService(Type type) => context;
    }

    private static IPluginExecutionContext Context(Dictionary<string, object> p) =>
        GuardTests.ContextProxy.Create(p);

    [Theory]
    [InlineData("Create", "asx_CatalogAdmin")]
    [InlineData("Update", "asx_ManageWork")]
    [InlineData("Create", "asx_DocumentWorker")]
    [InlineData("Update", "asx_DocumentWorker")]
    [InlineData("Update", "asx_RuntimeAdmin")]
    [InlineData("Create", "asx_SecurityAdmin")]
    [InlineData("Update", "asx_RecoverWorker")]
    public void ExactApiTransportAllowsSingleTargetWrites(string message, string api)
    {
        new StateGuard().Execute(new Provider(Setup("valid", message, api)));
    }

    [Theory]
    [InlineData("direct")]
    [InlineData("no-tag")]
    [InlineData("wrong-caller")]
    [InlineData("wrong-correlation")]
    [InlineData("wrong-api")]
    [InlineData("wrong-api-stage")]
    [InlineData("async-api")]
    [InlineData("outside-transaction")]
    [InlineData("changed-request")]
    [InlineData("changed-target")]
    [InlineData("changed-table")]
    [InlineData("wrong-intermediate")]
    [InlineData("wrong-guard-stage")]
    [InlineData("bulk")]
    [InlineData("delete")]
    public void ForgedOrMismatchedTransportFailsClosed(string variant)
    {
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new StateGuard().Execute(new Provider(Setup(variant, "Update", "asx_ManageWork")))
        );
    }

    [Fact]
    public void ReceiptsRemainAppendOnly()
    {
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new StateGuard().Execute(
                new Provider(Setup("valid", "Update", "asx_DocumentWorker", "asx_attempt"))
            )
        );
    }

    [Fact]
    public void CatalogAllowsTrustedSingleWrite()
    {
        new CatalogGuard().Execute(
            new Provider(Setup("valid", "Create", "asx_CatalogAdmin", "asx_site"))
        );
    }

    [Fact]
    public void CatalogPhysicalIdentitiesRemainImmutable()
    {
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new CatalogGuard().Execute(
                new Provider(Setup("immutable", "Update", "asx_CatalogAdmin", "asx_library"))
            )
        );
    }

    [Fact]
    public void RecoveryCannotWriteCatalog()
    {
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new CatalogGuard().Execute(
                new Provider(Setup("valid", "Update", "asx_RecoverWorker", "asx_library"))
            )
        );
    }

    [Fact]
    public void SitePrimaryKeyIsNotALibraryParentReference()
    {
        new CatalogGuard().Execute(
            new Provider(Setup("primary", "Update", "asx_CatalogAdmin", "asx_site"))
        );
    }

    [Fact]
    public void LibraryParentRemainsImmutable()
    {
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new CatalogGuard().Execute(
                new Provider(Setup("site-parent", "Update", "asx_CatalogAdmin", "asx_library"))
            )
        );
    }

    [Fact]
    public void IdentityBackfillRequiresItsExactAuthorizedCommand()
    {
        new CatalogGuard().Execute(
            new Provider(Setup("identity-upgrade", "Update", "asx_CatalogAdmin", "asx_site"))
        );
        Assert.Throws<InvalidPluginExecutionException>(() =>
            new CatalogGuard().Execute(
                new Provider(Setup("identity-edit", "Update", "asx_CatalogAdmin", "asx_site"))
            )
        );
    }

    private static IPluginExecutionContext Setup(
        string variant,
        string message,
        string apiName,
        string table = "asx_operation"
    )
    {
        var user = Guid.NewGuid();
        var correlation = Guid.NewGuid();
        var target = new Entity(table, Guid.NewGuid());
        if (variant == "primary")
            target[table + "id"] = target.Id;
        if (variant == "identity-upgrade" || variant == "identity-edit")
            target["asx_collectionid"] = Guid.NewGuid().ToString();
        if (variant == "site-parent")
            target["asx_siteid"] = Guid.NewGuid();
        if (variant == "immutable")
            target["asx_listid"] = Guid.NewGuid().ToString();
        var input = new ParameterCollection
        {
            ["Request"] = "{\"Command\":\"Cancel\",\"Key\":\"catalogprobe:test\"}",
        };
        if (variant == "identity-upgrade")
            input["Request"] = "{\"Command\":\"CompleteSiteIdentity\"}";
        var apiValues = new Dictionary<string, object>
        {
            ["Stage"] = 30,
            ["Mode"] = 0,
            ["IsInTransaction"] = true,
            ["UserId"] = user,
            ["CorrelationId"] = correlation,
            ["MessageName"] = apiName,
            ["InputParameters"] = input,
            ["SharedVariables"] = new ParameterCollection(),
        };
        var api = Context(apiValues);
        string operation = message.StartsWith("Create", StringComparison.Ordinal)
            ? "Create"
            : "Update";
        string tag = ApiWriteService.Tag(api, operation, target);
        if (variant == "wrong-caller")
            apiValues["UserId"] = Guid.NewGuid();
        if (variant == "wrong-correlation")
            apiValues["CorrelationId"] = Guid.NewGuid();
        if (variant == "wrong-api")
            apiValues["MessageName"] = "other_Api";
        if (variant == "wrong-api-stage")
            apiValues["Stage"] = 20;
        if (variant == "async-api")
            apiValues["Mode"] = 1;
        if (variant == "outside-transaction")
            apiValues["IsInTransaction"] = false;
        if (variant == "changed-request")
            input["Request"] = "different";
        if (variant == "changed-target")
            target.Id = Guid.NewGuid();
        var platform = Context(
            new Dictionary<string, object>
            {
                ["Stage"] = 30,
                ["Mode"] = 0,
                ["IsInTransaction"] = true,
                ["UserId"] = user,
                ["CorrelationId"] = correlation,
                ["MessageName"] = variant == "wrong-intermediate" ? "Execute" : operation,
                ["PrimaryEntityName"] = table,
                ["SharedVariables"] = new ParameterCollection(),
                ["ParentContext"] = api,
            }
        );
        var inputs = new ParameterCollection { ["Target"] = target };
        if (message.EndsWith("Multiple", StringComparison.Ordinal) || variant == "bulk")
        {
            var rows = new EntityCollection();
            rows.Entities.Add(target);
            if (variant == "bulk")
            {
                rows.Entities.Add(new Entity(table, Guid.NewGuid()));
                message = "UpdateMultiple";
            }
            inputs = new ParameterCollection { ["Targets"] = rows };
        }
        var shared = new ParameterCollection();
        if (variant != "no-tag")
            shared["tag"] = tag;
        var values = new Dictionary<string, object>
        {
            ["Stage"] = variant == "wrong-guard-stage" ? 10 : 20,
            ["IsInTransaction"] = true,
            ["UserId"] = user,
            ["CorrelationId"] = correlation,
            ["MessageName"] = variant == "delete" ? "Delete" : message,
            ["PrimaryEntityName"] = variant == "changed-table" ? "asx_claim" : table,
            ["SharedVariables"] = shared,
            ["InputParameters"] = inputs,
        };
        if (variant != "direct")
            values["ParentContext"] = platform;
        return Context(values);
    }

    [Fact]
    public void TransportPreservesConcurrencyAndRequiresLocallyValidatedContext()
    {
        var shared = new ParameterCollection();
        var api = Context(
            new Dictionary<string, object>
            {
                ["Stage"] = 30,
                ["Mode"] = 0,
                ["IsInTransaction"] = true,
                ["UserId"] = Guid.NewGuid(),
                ["CorrelationId"] = Guid.NewGuid(),
                ["MessageName"] = "asx_ManageWork",
                ["InputParameters"] = new ParameterCollection { ["Request"] = "{}" },
                ["SharedVariables"] = shared,
            }
        );
        var inner = new Capture();
        var service = new ApiWriteService(inner, api);
        var target = new Entity("asx_operation", Guid.NewGuid()) { RowVersion = "123" };
        var request = new UpdateRequest
        {
            Target = target,
            ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
        };
        Assert.Throws<InvalidPluginExecutionException>(() => service.Execute(request));
        Assert.Null(inner.Last);
        shared[DocumentWorkerApi.InternalWrite] = true;
        service.Execute(request);
        Assert.Same(request, inner.Last);
        Assert.Equal(ConcurrencyBehavior.IfRowVersionMatches, request.ConcurrencyBehavior);
        Assert.Equal("123", request.Target.RowVersion);
        Assert.Equal(ApiWriteService.Tag(api, "Update", target), request["tag"]);
    }

    private sealed class Capture : IOrganizationService
    {
        public OrganizationRequest? Last;

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            Last = request;
            return new UpdateResponse();
        }

        public Guid Create(Entity entity) => throw new NotSupportedException();

        public void Update(Entity entity) => throw new NotSupportedException();

        public void Delete(string name, Guid id) => throw new NotSupportedException();

        public Entity Retrieve(string name, Guid id, ColumnSet columns) =>
            throw new NotSupportedException();

        public EntityCollection RetrieveMultiple(QueryBase query) =>
            throw new NotSupportedException();

        public void Associate(
            string name,
            Guid id,
            Relationship r,
            EntityReferenceCollection rows
        ) => throw new NotSupportedException();

        public void Disassociate(
            string name,
            Guid id,
            Relationship r,
            EntityReferenceCollection rows
        ) => throw new NotSupportedException();
    }
}
