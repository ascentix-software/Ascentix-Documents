using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

/// <summary>Re-running a template for every record of its table in the background (spec 6.5).</summary>
public sealed class TemplateRunTests
{
    private static readonly Guid Admin = Guid.NewGuid();
    private static readonly Guid Worker = Guid.NewGuid();

    /// <summary>
    /// A published template on account, account enabled, and `records` more account rows. The
    /// platform's daily snapshot counts the rows unless `snapshot` says otherwise; an aggregate
    /// fails the test, because a plug-in never runs one.
    /// </summary>
    private static DurableWorkerTests.Fixture Setup(int records, long snapshot = -1)
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false)
        {
            AllowedTables = new[] { "account" },
        };
        f.SeedTemplate();
        RuntimeSeed.Seed(f.Service, Worker, "account");
        f.Service.Seed(new Entity("systemuser", Admin) { ["fullname"] = "Matt LaCasse" });
        for (int i = 0; i < records; i++)
            f.Service.Seed(new Entity("account", Guid.NewGuid()) { ["name"] = "Account " + i });
        if (snapshot >= 0)
            f.Service.SnapshotCounts["account"] = snapshot;
        f.Service.FetchHook = _ =>
            throw new InvalidOperationException("A plug-in must not run an aggregate.");
        return f;
    }

    private static TemplateRun Runs(DurableWorkerTests.Fixture f) =>
        new TemplateRun(f.Service, new[] { "account" }, () => f.Now);

    private static WorkerResult Start(DurableWorkerTests.Fixture f, Guid? request = null) =>
        f.Service.Transaction(() =>
            Runs(f)
                .Start(
                    new WorkerRequest
                    {
                        TemplateId = f.TemplateId,
                        RequestId = request ?? Guid.NewGuid(),
                    },
                    Admin
                )
        );

    private static WorkerResult Consume(DurableWorkerTests.Fixture f) =>
        f.Execute(new WorkerRequest { Command = "Plan", Key = TemplateRun.Key(f.TemplateId) });

    private static OutboxDocument Run(DurableWorkerTests.Fixture f) =>
        f.Store.Require<OutboxDocument>("asx_outbox", TemplateRun.Key(f.TemplateId)).Value;

    private static OutboxDocument[] Requests(DurableWorkerTests.Fixture f) =>
        f
            .Service.Rows.Values.Where(r => r.LogicalName == "asx_outbox")
            .Select(r => JsonWire.Read<OutboxDocument>(r.GetAttributeValue<string>("asx_payload")))
            .Where(o => o.Key.StartsWith("request:", StringComparison.Ordinal))
            .ToArray();

    /// <summary>Plans every request row of the run's current page, as the dispatcher would.</summary>
    private static void PlanPage(DurableWorkerTests.Fixture f)
    {
        foreach (var key in Run(f).PageKeys)
            f.Execute(new WorkerRequest { Command = "Plan", Key = key });
    }

    /// <summary>One dispatcher pass: the outbox page the flow lists, each key planned in order, failures recorded.</summary>
    private static string[] Dispatch(DurableWorkerTests.Fixture f)
    {
        var page = f.Execute(new WorkerRequest { Command = "ListOutbox" }).Keys;
        foreach (var key in page)
        {
            try
            {
                f.Execute(new WorkerRequest { Command = "Plan", Key = key });
            }
            catch (EvaluationBlockedException)
            {
                f.Execute(new WorkerRequest { Command = "FailOutbox", Key = key });
            }
        }
        return page;
    }

    /// <summary>A record event, as the record's Create or Update queues it.</summary>
    private static string Event(DurableWorkerTests.Fixture f)
    {
        var id = Guid.NewGuid();
        f.Service.Seed(new Entity("account", id) { ["name"] = "Changed" });
        return f.Execute(
            new WorkerRequest
            {
                Command = "Queue",
                TemplateId = f.TemplateId,
                RecordId = id,
                RequestId = Guid.NewGuid(),
            }
        ).Key;
    }

    private static int PendingRerunRows(DurableWorkerTests.Fixture f) =>
        Requests(f).Count(r => r.Priority == TemplateRun.Priority && r.Status == "Pending");

    [Fact]
    public void StartIsIdempotentAndKeepsOneActiveRunPerTemplate()
    {
        var f = Setup(3);
        var request = Guid.NewGuid();
        var started = Start(f, request);
        Assert.Equal("Pending", started.Status);
        Assert.Equal(TemplateRun.Key(f.TemplateId), started.Key);
        Assert.Equal("Running", started.Run!.State);
        Assert.Equal(4, started.Run.Total);
        Assert.True(started.Run.TotalEstimated);
        Assert.Equal("Matt LaCasse", started.Run.StartedBy);
        Assert.Equal(started.Key, Start(f, request).Key);
        var again = Start(f);
        Assert.Equal(started.Key, again.Key);
        Assert.Equal(new[] { "A re-run of this template is already in progress." }, again.Notices);
        Assert.Single(
            f.Service.Rows.Values,
            r =>
                r.LogicalName == "asx_outbox"
                && r.GetAttributeValue<string>("asx_table") == TemplateRun.IndexMarker
        );
    }

    [Fact]
    public void StartIsRefusedWithoutAPublishedVersion()
    {
        var f = Setup(1);
        f.Service.Rows[f.TemplateId]["asx_publishedrevisionid"] = null;
        Assert.Equal(
            "Publish the template first.",
            Assert.Throws<EvaluationBlockedException>(() => Start(f)).Message
        );
    }

    [Fact]
    public void ConsumeQueuesOnePageWithStableRequestIdsAndPriorityTwo()
    {
        var f = Setup(60);
        var request = Guid.NewGuid();
        Start(f, request);
        Consume(f);
        var run = Run(f);
        Assert.Equal(TemplateRun.PageSize, run.PageKeys.Length);
        Assert.Equal(2, run.SourcePage);
        var queued = Requests(f);
        Assert.Equal(TemplateRun.PageSize, queued.Length);
        Assert.All(queued, q => Assert.Equal(TemplateRun.Priority, q.Priority));
        Assert.All(
            queued,
            q =>
                Assert.Equal(
                    "request:"
                        + DocumentStore
                            .StableId(
                                run.Key
                                    + ":"
                                    + request.ToString("N")
                                    + ":"
                                    + q.RecordId.ToString("N")
                            )
                            .ToString("N"),
                    q.Key
                )
        );
    }

    [Fact]
    public void TheNextPageWaitsWhileAnyRowOfThisPageIsPending()
    {
        var f = Setup(60);
        Start(f);
        Consume(f);
        var first = Run(f).PageKeys;
        Assert.Equal("Waiting", Consume(f).Run!.State);
        Assert.Equal(first, Run(f).PageKeys);
        Assert.Equal(TemplateRun.PageSize, Requests(f).Length);
        PlanPage(f);
        Consume(f);
        Assert.Equal(TemplateRun.PageSize, Run(f).Planned);
        Assert.Equal(2 * TemplateRun.PageSize, Requests(f).Length);
    }

    [Fact]
    public void ARepeatedConsumeAfterARollbackQueuesNothingTwice()
    {
        var f = Setup(10);
        Start(f);
        Assert.ThrowsAny<InvalidOperationException>(() =>
            f.Service.Transaction<int>(() =>
            {
                Runs(f)
                    .Consume(
                        f.Store.Require<OutboxDocument>("asx_outbox", TemplateRun.Key(f.TemplateId))
                    );
                throw new InvalidOperationException("The flow's Plan call failed after queuing.");
            })
        );
        Assert.Empty(Requests(f));
        Assert.Equal(1, Run(f).SourcePage);
        Consume(f);
        var keys = Requests(f).Select(r => r.Key).OrderBy(k => k).ToArray();
        // The same page queued again by a second run of the same request finds its rows.
        var run = f.Store.Require<OutboxDocument>("asx_outbox", TemplateRun.Key(f.TemplateId));
        run.Value.SourcePage = 1;
        run.Value.SourceCookie = null;
        run.Value.PageKeys = Array.Empty<string>();
        f.Store.Save(run);
        Consume(f);
        Assert.Equal(keys, Requests(f).Select(r => r.Key).OrderBy(k => k).ToArray());
    }

    [Fact]
    public void RetiredRecordsAreSkippedAndCounted()
    {
        var f = Setup(4);
        var retired = f
            .Service.Rows.Values.Where(r => r.LogicalName == "account")
            .Take(2)
            .Select(r => r.Id)
            .ToArray();
        foreach (var id in retired)
            f.Store.Create(
                "asx_outbox",
                new OutboxDocument
                {
                    Key = WorkerCoordinator.RetirementKey("account", id),
                    Status = "Retired",
                }
            );
        Start(f);
        Consume(f);
        Assert.DoesNotContain(Requests(f), r => retired.Contains(r.RecordId));
        Assert.Equal(2, Run(f).Planned);
        Assert.Equal(3, Run(f).PageKeys.Length);
    }

    [Fact]
    public void TheRunEndsDoneWithItsCounts()
    {
        var f = Setup(30);
        Start(f);
        for (int pass = 0; pass < 5 && Run(f).Status == "Pending"; pass++)
        {
            Consume(f);
            PlanPage(f);
        }
        Consume(f);
        var run = Run(f);
        Assert.Equal("Planned", run.Status);
        Assert.Equal(31, run.Planned);
        Assert.Equal(31, run.Total);
        Assert.Equal(f.Now, run.EndedUtc);
        var done = Runs(f).Describe(run);
        Assert.Equal("Done", done.State);
        Assert.Equal(31, done.Planned);
        Assert.Equal(31, done.Total);
        Assert.False(done.TotalEstimated);
    }

    [Fact]
    public void AStaleSnapshotTotalGrowsWithTheRunAndIsExactWhenDone()
    {
        var f = Setup(30, snapshot: 5);
        Assert.Equal(5, Start(f).Run!.Total);
        Consume(f);
        Assert.Equal(TemplateRun.PageSize, Run(f).Total);
        PlanPage(f);
        Consume(f);
        // Planned plus the page in flight.
        Assert.Equal(31, Runs(f).Describe(Run(f)).Total);
        Assert.True(Runs(f).Describe(Run(f)).TotalEstimated);
        PlanPage(f);
        Consume(f);
        Assert.Equal("Done", Runs(f).Describe(Run(f)).State);
        Assert.Equal(31, Run(f).Total);
    }

    [Fact]
    public void TemplateSwitchedOffCancelsTheRun()
    {
        var f = Setup(5);
        Start(f);
        f.Service.Rows[f.TemplateId]["asx_disabled"] = true;
        var result = Consume(f);
        Assert.Equal("Cancelled", result.Status);
        Assert.Equal(new[] { "The template is off, so the re-run stopped." }, Run(f).Notices);
        Assert.DoesNotContain(
            TemplateRun.Key(f.TemplateId),
            f.Execute(new WorkerRequest { Command = "ListOutbox" }).Keys
        );
    }

    [Fact]
    public void RemovedTableCancelsTheRun()
    {
        var f = Setup(5);
        Start(f);
        var runs = new TemplateRun(f.Service, Array.Empty<string>(), () => f.Now);
        var result = f.Service.Transaction(() =>
            runs.Consume(
                f.Store.Require<OutboxDocument>("asx_outbox", TemplateRun.Key(f.TemplateId))
            )
        );
        Assert.Equal("Cancelled", result.Status);
        Assert.Equal(WorkerCoordinator.TableNotEnabled("account"), Assert.Single(Run(f).Notices));
    }

    [Fact]
    public void ATableDeletedFromTheOrganizationCancelsTheRun()
    {
        var f = Setup(30);
        Start(f);
        Consume(f);
        f.Service.MissingTables.Add("account");
        var result = Consume(f);
        Assert.Equal("Cancelled", result.Status);
        Assert.Equal(
            new[] { "The table no longer exists, so the re-run stopped." },
            Run(f).Notices
        );
        Assert.Equal("account", result.Run!.TableLabel);
        Assert.Equal(
            "The table no longer exists.",
            Assert
                .Throws<EvaluationBlockedException>(() =>
                    f.Service.Transaction(() => Runs(f).Count(f.TemplateId))
                )
                .Message
        );
    }

    [Fact]
    public void BlockedRequestRowsCountAsPlannedAndTheRunContinues()
    {
        var f = Setup(30);
        Start(f);
        Consume(f);
        var keys = Run(f).PageKeys;
        // One record's planning fails for a reason that is not temporary.
        f.Execute(new WorkerRequest { Command = "FailOutbox", Key = keys[0] });
        foreach (var key in keys.Skip(1))
            f.Execute(new WorkerRequest { Command = "Plan", Key = key });
        Consume(f);
        Assert.Equal(TemplateRun.PageSize, Run(f).Planned);
        Assert.Equal("Pending", Run(f).Status);
        Assert.Equal(31 - TemplateRun.PageSize, Run(f).PageKeys.Length);
    }

    [Fact]
    public void APausedRunIsNotListedAndResumeContinues()
    {
        var f = Setup(5);
        Start(f);
        var key = TemplateRun.Key(f.TemplateId);
        Assert.Equal("Paused", f.Service.Transaction(() => Runs(f).Pause(key)).Run!.State);
        Assert.DoesNotContain(key, f.Execute(new WorkerRequest { Command = "ListOutbox" }).Keys);
        Assert.Equal("Paused", Consume(f).Status);
        Assert.Empty(Requests(f));
        var resumed = f.Service.Transaction(() => Runs(f).Resume(key));
        Assert.Equal("Running", resumed.Run!.State);
        Assert.Contains(key, f.Execute(new WorkerRequest { Command = "ListOutbox" }).Keys);
    }

    [Fact]
    public void CancelAlsoCancelsThePagesPendingRows()
    {
        var f = Setup(5);
        Start(f);
        Consume(f);
        var key = TemplateRun.Key(f.TemplateId);
        var cancelled = f.Service.Transaction(() => Runs(f).Cancel(key));
        Assert.Equal("Cancelled", cancelled.Status);
        Assert.Equal(new[] { "Re-run cancelled." }, cancelled.Notices);
        Assert.All(Requests(f), r => Assert.Equal("Cancelled", r.Status));
        Assert.Empty(f.Execute(new WorkerRequest { Command = "ListOutbox" }).Keys);
    }

    [Fact]
    public void ATemporaryFailureWaitsAndAnyOtherBlocksAndRetryResumesAtTheCursor()
    {
        var f = Setup(60);
        Start(f);
        Consume(f);
        PlanPage(f);
        var key = TemplateRun.Key(f.TemplateId);
        Assert.Equal(
            "Pending",
            f.Execute(
                new WorkerRequest
                {
                    Command = "FailOutbox",
                    Key = key,
                    StatusCode = 429,
                }
            ).Status
        );
        Assert.NotNull(Run(f).NextAttemptUtc);
        Assert.Equal("Retrying", Runs(f).Describe(Run(f)).State);
        Assert.Equal(
            "Blocked",
            f.Execute(
                new WorkerRequest
                {
                    Command = "FailOutbox",
                    Key = key,
                    StatusCode = 400,
                }
            ).Status
        );
        Assert.Equal("Blocked", Runs(f).Describe(Run(f)).State);
        Assert.Equal(
            "Pending",
            f.Execute(new WorkerRequest { Command = "RetryOutbox", Key = key }).Status
        );
        Assert.Equal(2, Run(f).SourcePage);
    }

    [Fact]
    public void ADueWaitEndsAtTheNextDispatchEvenWhileThePageIsInFlight()
    {
        var f = Setup(60);
        Start(f);
        Consume(f);
        var key = TemplateRun.Key(f.TemplateId);
        f.Execute(
            new WorkerRequest
            {
                Command = "FailOutbox",
                Key = key,
                StatusCode = 429,
            }
        );
        f.Now = Run(f).NextAttemptUtc!.Value;
        Assert.Equal("Waiting", Consume(f).Run!.State);
        Assert.Null(Run(f).NextAttemptUtc);
        Assert.Equal(0, Run(f).Attempts);
        Assert.Empty(Run(f).Notices);
    }

    [Fact]
    public void FolderJobsFromARerunComeAfterNormalWork()
    {
        var f = Setup(1);
        Start(f);
        Consume(f);
        PlanPage(f);
        var normal = f.PlanRecord();
        var rerun = f
            .Service.Rows.Values.Where(r => r.LogicalName == "asx_operation")
            .Select(r =>
                (
                    Priority: r.GetAttributeValue<int>("asx_priority"),
                    Key: JsonWire
                        .Read<OperationDocument>(r.GetAttributeValue<string>("asx_payload"))
                        .Key
                )
            )
            .ToArray();
        Assert.Contains(rerun, o => o.Priority == TemplateRun.Priority);
        Assert.Contains(rerun, o => o.Priority == 1);
        var pending = f.Store.Pending("asx_operation", now: f.Now);
        var firstRerun = Array.FindIndex(
            pending,
            k => rerun.Single(o => o.Key == k).Priority == TemplateRun.Priority
        );
        Assert.All(
            pending.Take(firstRerun),
            k => Assert.Equal(1, rerun.Single(o => o.Key == k).Priority)
        );
        Assert.Contains(normal.Keys.Single(), pending.Take(firstRerun));
    }

    [Fact]
    public void RowsQueuedBeforePrioritiesExistedStillMakeNormalFolderJobs()
    {
        var f = Setup(0);
        var queued = f.Execute(
            new WorkerRequest
            {
                Command = "Queue",
                TemplateId = f.TemplateId,
                RecordId = f.RecordId,
                RequestId = Guid.NewGuid(),
            }
        );
        // A row an earlier release stored has no Priority member; it reads as 0.
        var row = f.Service.Rows.Values.Single(r =>
            r.LogicalName == "asx_outbox"
            && r.GetAttributeValue<string>("asx_key") == DocumentStore.Hash(queued.Key)
        );
        string payload = row.GetAttributeValue<string>("asx_payload");
        Assert.Contains("\"Priority\":1,", payload);
        row["asx_payload"] = payload.Replace("\"Priority\":1,", "");
        Assert.Equal(0, f.Store.Require<OutboxDocument>("asx_outbox", queued.Key).Value.Priority);
        var planned = f.Execute(new WorkerRequest { Command = "Plan", Key = queued.Key });
        var operation = f.Service.Rows.Values.Single(r =>
            r.LogicalName == "asx_operation"
            && r.GetAttributeValue<string>("asx_key") == DocumentStore.Hash(planned.Keys.Single())
        );
        Assert.Equal(1, operation.GetAttributeValue<int>("asx_priority"));
    }

    [Fact]
    public void ListOutboxPutsRunsLastSoARunSeesItsPagePlanned()
    {
        var f = Setup(60);
        Start(f);
        Consume(f);
        var page = f.Execute(new WorkerRequest { Command = "ListOutbox" }).Keys;
        Assert.Equal(DocumentStore.DispatchPage, page.Length);
        Assert.Equal(TemplateRun.Key(f.TemplateId), page.Last());
        Assert.Equal(Run(f).PageKeys, page.Take(TemplateRun.PageSize));
    }

    [Fact]
    public void WithNoOtherWorkARunPlansAPagePerDispatch()
    {
        var f = Setup(200);
        Start(f);
        Dispatch(f);
        for (int dispatch = 1; dispatch <= 5; dispatch++)
        {
            var page = Dispatch(f);
            Assert.Equal(TemplateRun.Key(f.TemplateId), page.Last());
            Assert.Equal(dispatch * TemplateRun.PageSize, Run(f).Planned);
            Assert.Equal(TemplateRun.PageSize, PendingRerunRows(f));
        }
    }

    [Fact]
    public void WithOtherWorkTheRunInterleavesFairly()
    {
        var f = Setup(200);
        Start(f);
        Dispatch(f);
        Dispatch(f);
        Assert.Equal(TemplateRun.PageSize, Run(f).Planned);
        // Thirty record events arrive behind the run's page in flight.
        var events = Enumerable.Range(0, 30).Select(_ => Event(f)).ToArray();
        int planned = Run(f).Planned;
        for (int dispatch = 0; dispatch < 3; dispatch++)
        {
            Dispatch(f);
            Assert.True(PendingRerunRows(f) <= TemplateRun.PageSize);
        }
        // Each event waited behind at most one page of re-run work, and the run kept going.
        Assert.All(
            events,
            e =>
                Assert.Equal(
                    "Planned",
                    f.Store.Require<OutboxDocument>("asx_outbox", e).Value.Status
                )
        );
        Assert.True(Run(f).Planned >= planned + TemplateRun.PageSize);
    }

    [Fact]
    public void NewRecordEventsArePlannedWithinTwoDispatchPages()
    {
        var f = Setup(40000);
        Start(f);
        Assert.Equal(new[] { TemplateRun.Key(f.TemplateId) }, Dispatch(f));
        Dispatch(f);
        // Ten records change while the re-run has a page in flight.
        var events = Enumerable.Range(0, 10).Select(_ => Event(f)).ToArray();
        var next = Dispatch(f).Concat(Dispatch(f)).ToArray();
        Assert.All(events, e => Assert.Contains(e, next));
        Assert.All(
            events,
            e =>
                Assert.Equal(
                    "Planned",
                    f.Store.Require<OutboxDocument>("asx_outbox", e).Value.Status
                )
        );
        Assert.True(PendingRerunRows(f) <= TemplateRun.PageSize);
    }

    [Fact]
    public void ProgressShowsAnEstimatedFinishOnceAPageIsDone()
    {
        var f = Setup(100);
        Start(f);
        Assert.Null(Runs(f).Describe(Run(f)).EstimatedFinishUtc);
        Consume(f);
        PlanPage(f);
        f.Now = f.Now.AddMinutes(10);
        Consume(f);
        var state = Runs(f).Describe(Run(f));
        Assert.Equal(TemplateRun.PageSize, state.Planned);
        Assert.Equal(101, state.Total);
        // One page in 10 minutes: the records left take as long each.
        double perRecord = 600.0 / TemplateRun.PageSize;
        Assert.Equal(
            f.Now.AddSeconds(perRecord * (101 - TemplateRun.PageSize)),
            state.EstimatedFinishUtc
        );
        f.Service.Transaction(() => Runs(f).Pause(TemplateRun.Key(f.TemplateId)));
        Assert.Null(Runs(f).Describe(Run(f)).EstimatedFinishUtc);
    }

    [Fact]
    public void CountsComeFromTheDailySnapshotAtAnySize()
    {
        var f = Setup(1, snapshot: 120000);
        var counted = f.Service.Transaction(() => Runs(f).Count(f.TemplateId));
        Assert.Equal("Counted", counted.Status);
        Assert.Equal(120000, counted.Run!.Total);
        Assert.True(counted.Run.TotalEstimated);
        var started = Start(f);
        Assert.Equal(120000, started.Run!.Total);
        Assert.Null(started.Run.EstimatedFinishUtc);
    }

    [Fact]
    public void CountRefusesATableThatIsNotEnabledAsStartDoes()
    {
        var f = Setup(5);
        var runs = new TemplateRun(f.Service, Array.Empty<string>(), () => f.Now);
        var counted = Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() => runs.Count(f.TemplateId))
        );
        var started = Assert.Throws<EvaluationBlockedException>(() =>
            f.Service.Transaction(() =>
                runs.Start(
                    new WorkerRequest { TemplateId = f.TemplateId, RequestId = Guid.NewGuid() },
                    Admin
                )
            )
        );
        Assert.Equal(WorkerCoordinator.TableNotEnabled("account"), counted.Message);
        Assert.Equal(started.Message, counted.Message);
    }

    [Fact]
    public void AnOperatorWhoCannotReadTheTableOrUsersStartsARunThroughTheWorker()
    {
        var f = Setup(3);
        var caller = new ReadDenied(f.Service, "account", "systemuser");
        IOrganizationService Factory(Guid? user) =>
            user == Worker ? f.Service
            : user == Admin ? caller
            : throw new InvalidOperationException("Unexpected identity " + user);
        var count = ManageWork(
            f,
            Factory,
            new WorkerRequest { Command = "CountRecords", TemplateId = f.TemplateId }
        );
        Assert.Equal(4, count.Run!.Total);
        var started = ManageWork(
            f,
            Factory,
            new WorkerRequest
            {
                Command = "StartTemplateRun",
                TemplateId = f.TemplateId,
                RequestId = Guid.NewGuid(),
            }
        );
        Assert.Equal("Pending", started.Status);
        Assert.Equal(4, started.Run!.Total);
        Assert.Equal("Matt LaCasse", started.Run.StartedBy);
        Assert.Equal("Matt LaCasse", Run(f).StartedByName);
        var paused = ManageWork(
            f,
            Factory,
            new WorkerRequest { Command = "PauseTemplateRun", Key = started.Key }
        );
        Assert.Equal("Paused", paused.Run!.State);
        Assert.Equal("Matt LaCasse", paused.Run.StartedBy);
    }

    private static WorkerResult ManageWork(
        DurableWorkerTests.Fixture f,
        Func<Guid?, IOrganizationService> factory,
        WorkerRequest request
    )
    {
        var context = GuardTests.ContextProxy.Create(
            new Dictionary<string, object>
            {
                ["Stage"] = 30,
                ["Mode"] = 0,
                ["IsInTransaction"] = true,
                ["UserId"] = Admin,
                ["InitiatingUserId"] = Admin,
                ["CorrelationId"] = Guid.NewGuid(),
                ["MessageName"] = "asx_ManageWork",
                ["InputParameters"] = new ParameterCollection
                {
                    ["Request"] = JsonWire.Write(request),
                },
                ["OutputParameters"] = new ParameterCollection(),
                ["SharedVariables"] = new ParameterCollection(),
            }
        );
        return f.Service.Transaction(() =>
        {
            new Ascentix.Documents.Plugins.ManageWorkApi().Execute(new Provider(context, factory));
            return JsonWire.Read<WorkerResult>((string)context.OutputParameters["Result"]);
        });
    }

    private sealed class Provider : IServiceProvider, IOrganizationServiceFactory
    {
        private readonly IPluginExecutionContext context;
        private readonly Func<Guid?, IOrganizationService> factory;

        public Provider(IPluginExecutionContext context, Func<Guid?, IOrganizationService> factory)
        {
            this.context = context;
            this.factory = factory;
        }

        public object GetService(Type type) =>
            type == typeof(IPluginExecutionContext) ? context : this;

        public IOrganizationService CreateOrganizationService(Guid? userId) => factory(userId);
    }

    /// <summary>
    /// A caller with the Documents Operator role only: no Read on the listed tables. Dataverse
    /// refuses those reads with a fault, which ends the caller's transaction.
    /// </summary>
    private sealed class ReadDenied : IOrganizationService
    {
        private readonly DurableWorkerTests.MemoryService inner;
        private readonly string[] denied;

        public ReadDenied(DurableWorkerTests.MemoryService inner, params string[] denied)
        {
            this.inner = inner;
            this.denied = denied;
        }

        private T Read<T>(string table, Func<T> read) =>
            denied.Contains(table)
                ? throw inner.Refuse(
                    new FaultException<OrganizationServiceFault>(
                        new OrganizationServiceFault
                        {
                            ErrorCode = unchecked((int)0x80040220),
                            Message = "Principal user is missing prvRead privilege on " + table,
                        }
                    )
                )
                : read();

        public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet) =>
            Read(entityName, () => inner.Retrieve(entityName, id, columnSet));

        public EntityCollection RetrieveMultiple(QueryBase query) =>
            query is QueryExpression expression
                ? Read(expression.EntityName, () => inner.RetrieveMultiple(query))
                : throw inner.Refuse(new InvalidOperationException("Aggregates are not allowed."));

        public OrganizationResponse Execute(OrganizationRequest request) =>
            request is Microsoft.Crm.Sdk.Messages.RetrieveTotalRecordCountRequest totals
                ? Read(totals.EntityNames.Single(), () => inner.Execute(request))
                : inner.Execute(request);

        public Guid Create(Entity entity) => inner.Create(entity);

        public void Update(Entity entity) => inner.Update(entity);

        public void Delete(string entityName, Guid id) => inner.Delete(entityName, id);

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
