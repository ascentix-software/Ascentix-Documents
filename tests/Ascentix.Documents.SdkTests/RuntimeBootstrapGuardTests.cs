using System;
using System.Collections.Generic;
using Ascentix.Documents.Plugins;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class RuntimeBootstrapGuardTests
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

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void OnlyDisabledRuntimeCanBeCreatedDirectly(bool enabled, bool updates)
    {
        var target = new Entity("asx_runtime", Guid.NewGuid())
        {
            ["asx_enabled"] = enabled,
            ["asx_processrecordupdates"] = updates,
        };
        var context = GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["MessageName"] = "Create",
                ["PrimaryEntityName"] = "asx_runtime",
                ["Stage"] = 20,
                ["IsInTransaction"] = true,
                ["UserId"] = Guid.NewGuid(),
                ["InputParameters"] = new ParameterCollection { ["Target"] = target },
            }
        );
        if (enabled || updates)
            Assert.Throws<InvalidPluginExecutionException>(() =>
                new CatalogGuard().Execute(new Provider(context))
            );
        else
            new CatalogGuard().Execute(new Provider(context));
    }
}
