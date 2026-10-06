using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Plugins;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

/// <summary>
/// Pausing the runtime stops new work, but never refuses to record the outcome of a write the
/// product already permitted: the response of a sent write is recorded while paused, and the
/// job then finishes or waits for resume with its outcome known.
/// </summary>
public sealed class PauseTests
{
    /// <summary>The worker custom API (asx_DocumentWorker) over a test Dataverse.</summary>
    internal sealed class WorkerApi
    {
        private readonly DurableWorkerTests.MemoryService service;
        private readonly Guid worker = Guid.NewGuid();
        private readonly Guid runtime;

        public WorkerApi(
            DurableWorkerTests.MemoryService service,
            string siteUrl,
            params string[] tables
        )
        {
            this.service = service;
            runtime = RuntimeSeed.Seed(service, worker, tables);
            service.Rows[runtime]["asx_enabled"] = true;
            if (
                !service.Rows.Values.Any(r =>
                    r.LogicalName == "sharepointsite"
                    && r.GetAttributeValue<string>("absoluteurl") == siteUrl
                )
            )
                service.Seed(
                    new Entity("sharepointsite", Guid.NewGuid())
                    {
                        ["absoluteurl"] = siteUrl,
                        ["statecode"] = new OptionSetValue(0),
                    }
                );
        }

        /// <summary>Pauses (false) or resumes (true) the runtime, as Runtime admin does.</summary>
        public void Enabled(bool enabled) => service.Rows[runtime]["asx_enabled"] = enabled;

        public WorkerResult Execute(WorkerRequest request)
        {
            var context = GuardTests.ContextProxy.Create(
                new Dictionary<string, object>
                {
                    ["Stage"] = 30,
                    ["Mode"] = 0,
                    ["IsInTransaction"] = true,
                    ["UserId"] = worker,
                    ["CorrelationId"] = Guid.NewGuid(),
                    ["MessageName"] = "asx_DocumentWorker",
                    ["InputParameters"] = new ParameterCollection
                    {
                        ["Request"] = JsonWire.Write(request),
                    },
                    ["OutputParameters"] = new ParameterCollection(),
                    ["SharedVariables"] = new ParameterCollection(),
                }
            );
            return service.Transaction(() =>
            {
                new DocumentWorkerApi().Execute(new Provider(context, service));
                return JsonWire.Read<WorkerResult>((string)context.OutputParameters["Result"]);
            });
        }

        public WorkerResult Step(
            string command,
            WorkerResult work,
            string run,
            int status = 0,
            string? body = null
        ) =>
            Execute(
                new WorkerRequest
                {
                    Command = command,
                    Key = work.Key,
                    RunId = run,
                    Token = work.Token,
                    ProbeId = work.ProbeId,
                    ProbeKind = work.ProbeKind,
                    HttpStatus = status,
                    ResponseBody = body,
                }
            );

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

    private static DispatcherDocument Writer(
        DurableWorkerTests.MemoryService service,
        string key
    ) =>
        new DocumentStore(service)
            .Require<DispatcherDocument>("asx_claim", WorkCoordination.Operation(service, key))
            .Value;

    private static void Expire(DurableWorkerTests.MemoryService service, string key)
    {
        var store = new DocumentStore(service);
        var claim = store.Require<DispatcherDocument>(
            "asx_claim",
            WorkCoordination.Operation(service, key)
        );
        claim.Value.LeaseUntilUtc = DateTime.UtcNow.AddMinutes(-1);
        store.Save(claim);
    }

    [Fact]
    public void PauseRefusesOnlyCommandsThatStartNewWork()
    {
        foreach (
            var command in new[]
            {
                "Claim",
                "BeginHttp",
                "PrepareCreate",
                "Plan",
                "ListOutbox",
                "ListOperations",
                "RefreshSecurity",
                "PurgeHistory",
                "Queue",
            }
        )
            Assert.True(WorkerPause.Refuses(command), command);
        foreach (
            var command in new[]
            {
                "CreateResponse",
                "Observe",
                "Complete",
                "Fail",
                "FailUnclaimed",
                "FailOutbox",
                "Renew",
            }
        )
            Assert.False(WorkerPause.Refuses(command), command);
    }

    [Fact]
    public void FolderCreateAnsweredDuringAPauseIsRecordedAndFinishesAfterResume()
    {
        var f = new DurableWorkerTests.Fixture();
        var api = new WorkerApi(f.Service, "https://example.sharepoint.com/sites/proto", "account");
        const string run = "dispatch/run-1";
        var work = api.Execute(
            new WorkerRequest
            {
                Command = "Claim",
                Key = f.Operation.Key,
                RunId = run,
            }
        );
        Assert.Equal("Library", work.ProbeKind);
        work = api.Step("Observe", work, run, 200, f.LibraryBody());
        work = api.Step("Observe", work, run, 200, f.ParentBody());
        Assert.Equal("Folder", work.ProbeKind);
        work = api.Step("Observe", work, run, 404, "{}");
        Assert.Equal("ReadyToCreate", work.Status);
        var prepared = api.Step("PrepareCreate", work, run);
        Assert.Equal("Create", prepared.Status);
        Assert.Equal("Permit", api.Step("BeginHttp", prepared, run).Status);

        // The POST is on its way when the admin pauses; SharePoint answers 200.
        api.Enabled(false);
        var recorded = api.Step(
            "CreateResponse",
            prepared,
            run,
            200,
            DurableWorkerTests.CreateBody()
        );
        Assert.Equal("Read", recorded.Status);
        Assert.Equal("Folder", recorded.ProbeKind);
        var op = f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value;
        Assert.True(op.ExternalSubmitted && op.ExternalResponseKnown, "The answer is recorded");
        Assert.False(Writer(f.Service, f.Operation.Key).HttpOutstanding);
        // No new SharePoint call while paused: the run ends at its next step.
        Assert.Equal("Disabled", api.Step("BeginHttp", recorded, run).Status);
        Assert.Equal(
            "Disabled",
            api.Execute(
                new WorkerRequest
                {
                    Command = "Claim",
                    Key = f.Operation.Key,
                    RunId = "dispatch/run-2",
                }
            ).Status
        );

        // After resume the next run takes over, finds the folder and completes without a POST.
        api.Enabled(true);
        Expire(f.Service, f.Operation.Key);
        const string next = "dispatch/run-2";
        work = api.Execute(
            new WorkerRequest
            {
                Command = "Claim",
                Key = f.Operation.Key,
                RunId = next,
            }
        );
        work = api.Step("Observe", work, next, 200, f.LibraryBody());
        work = api.Step("Observe", work, next, 200, f.ParentBody());
        var item = f.Item(Guid.NewGuid(), null);
        work = api.Step(
            "Observe",
            work,
            next,
            200,
            JsonWire.Write(
                new ODataEnvelope<FolderLookupObservation>
                {
                    Data = new FolderLookupObservation
                    {
                        Exists = true,
                        Id = item.Id,
                        Path = item.Path,
                        ListItemAllFields = item,
                    },
                }
            )
        );
        Assert.Equal("FinalParent", work.ProbeKind);
        work = api.Step("Observe", work, next, 200, f.ParentBody());
        Assert.Equal("Verified", work.Status);
        Assert.Equal("Applied", api.Step("Complete", work, next).Status);
    }

    [Fact]
    public void FolderJobVerifiedBeforeAPauseCompletesWhilePaused()
    {
        var f = new DurableWorkerTests.Fixture();
        var api = new WorkerApi(f.Service, "https://example.sharepoint.com/sites/proto", "account");
        // The worker API runs on the real clock.
        f.Now = DateTime.UtcNow;
        var verified = f.ObserveAndFinalize(f.Preflight(f.Claim()), f.Item(Guid.NewGuid(), null));
        api.Enabled(false);
        // Completing records what SharePoint already showed; it calls nothing.
        Assert.Equal("Applied", api.Step("Complete", verified, "run-1").Status);
    }

    [Fact]
    public void LibraryCreateAnsweredDuringAPauseIsRecordedAndNeedsNoRecovery()
    {
        var f = new LibraryProvisioningTests.Fixture { StopBeforePost = true };
        var api = new WorkerApi(f.Service, "https://example.sharepoint.com/sites/test");
        var key = f.Queue().Key;
        var prepared = f.Run(key);
        Assert.Equal("Create", prepared.Status);
        f.StopBeforePost = false;
        Assert.Equal("Permit", api.Step("BeginHttp", prepared, "library-test").Status);

        // The library create is on its way when the admin pauses; SharePoint answers 201.
        f.Exists = true;
        api.Enabled(false);
        var recorded = api.Step(
            "CreateResponse",
            prepared,
            "library-test",
            201,
            JsonWire.Write(
                new ODataEnvelope<CreatedLibrary>
                {
                    Data = new CreatedLibrary { Id = f.List, Title = "Documents" },
                }
            )
        );
        Assert.Equal("Read", recorded.Status);
        Assert.Equal("Library", recorded.ProbeKind);
        var setup = f.Store.Require<LibrarySetup>("asx_operation", key).Value;
        Assert.Equal(f.List, setup.ListId);
        Assert.False(Writer(f.Service, key).HttpOutstanding);
        Assert.Equal("Disabled", api.Step("BeginHttp", recorded, "library-test").Status);

        // After resume the next run takes over and finishes; no create is sent again and no
        // recovery is needed.
        api.Enabled(true);
        f.Expire(key);
        Assert.Equal("AccessPending", f.Run(key).Status);
        Assert.NotEqual("RecoveryRequired", f.Status(key));
        Assert.DoesNotContain(f.Posts, p => p == "_api/web/lists");
        Assert.Single(f.Service.Rows.Values, r => r.LogicalName == "asx_library");
    }

    [Fact]
    public void AccessWriteAnsweredDuringAPauseIsRecordedAndFinishesAfterResume()
    {
        var f = new SecurityWorkerTests.Fixture();
        var api = new WorkerApi(f.Service, "https://example.sharepoint.com/sites/proto");
        f.Queue("Read");
        var work = f.Start();
        while (work.Status == "Read")
            work = f.Observe(work);
        Assert.Equal("ReadyToCreate", work.Status);
        var prepared = f.Call("PrepareCreate", work);
        Assert.Equal("Create", prepared.Status);
        Assert.Equal("Permit", api.Step("BeginHttp", prepared, "security/run-1").Status);

        api.Enabled(false);
        f.Apply();
        string sent = f.Writes.Single();
        var recorded = api.Step("CreateResponse", prepared, "security/run-1", 200, "{}");
        Assert.Equal("Read", recorded.Status);
        Assert.True(f.Operation().ExternalResponseKnown, "The answer is recorded");
        Assert.False(Writer(f.Service, f.Key).HttpOutstanding);
        Assert.Equal("Disabled", api.Step("BeginHttp", recorded, "security/run-1").Status);

        api.Enabled(true);
        Expire(f.Service, f.Key);
        Assert.Equal("Applied", f.Drive().Status);
        // The write answered during the pause is not sent again.
        Assert.Single(f.Writes, w => w == sent);
    }
}
