using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

/// <summary>Monitor's counts and readable lists (spec 6.1, 6.2).</summary>
public sealed class ProblemListTests
{
    private static DurableWorkerTests.Fixture Setup()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false)
        {
            AllowedTables = new[] { "account" },
        };
        f.SeedTemplate();
        RuntimeSeed.Seed(f.Service, Guid.NewGuid(), "account");
        f.Service.DisplayNames["account"] = "Account";
        f.Service.Rows[f.LibraryId]["asx_name"] = "General";
        f.Service.Rows[f.SiteId]["asx_name"] = "Proto";
        return f;
    }

    private static WorkerResult List(
        DurableWorkerTests.Fixture f,
        string list,
        string? page = null
    ) =>
        f.Service.Transaction(() =>
            ProblemList.List(
                f.Service,
                new WorkerRequest { List = list, Page = page },
                f.Now,
                new[] { "account" }
            )
        );

    private static void Blocked(DurableWorkerTests.Fixture f, string code)
    {
        var job = f.Operation;
        job.Status = "Blocked";
        job.ErrorCode = code;
        f.Store.Create("asx_operation", job);
    }

    [Fact]
    public void SummaryCountsEachListWithoutAggregatesAndReportsThePlatformLimit()
    {
        var f = Setup();
        var asked = new List<string>();
        f.Service.FetchHook = fetch =>
        {
            asked.Add(fetch.Query);
            if (fetch.Query.Contains("asyncoperation"))
                return new EntityCollection
                {
                    TotalRecordCount = ProblemList.CountLimit,
                    TotalRecordCountLimitExceeded = true,
                };
            int n =
                fetch.Query.Contains("RetryWait") ? 4
                : fetch.Query.Contains("operator=\"eq\" value=\"templaterun\"") ? 1
                : fetch.Query.Contains("value=\"Waiting\"") ? 3
                : 2;
            return new EntityCollection { TotalRecordCount = n };
        };
        var summary = f.Service.Transaction(() => ProblemList.Summary(f.Service, f.Now));
        Assert.Equal(6, asked.Count);
        Assert.All(asked, q => Assert.DoesNotContain("aggregate=", q));
        Assert.All(asked, q => Assert.Contains("returntotalrecordcount=\"true\"", q));
        Assert.Equal(1, summary.TemplateRuns);
        Assert.Equal(2, summary.BlockedRecords);
        Assert.Equal(3, summary.WaitingRecords);
        Assert.Equal(2, summary.BlockedJobs);
        Assert.Equal(4, summary.RetryingJobs);
        Assert.Equal(5000, summary.NotCaptured);
        Assert.Equal(new[] { "NotCaptured" }, summary.Capped);
        Assert.Equal(f.Now, summary.CountedUtc);
        var blockedRecords = asked.Single(q =>
            q.Contains("asx_outbox") && q.Contains("value=\"Blocked\"")
        );
        Assert.Contains(
            "<condition attribute=\"asx_table\" operator=\"ne\" value=\"templaterun\" />",
            blockedRecords
        );
        var blockedJobs = asked.Single(q =>
            q.Contains("asx_operation") && q.Contains("RecoveryRequired")
        );
        Assert.Contains("<value>Reconciling</value>", blockedJobs);
        Assert.Contains(
            EventRegistrations.RecordHandler,
            asked.Single(q => q.Contains("asyncoperation"))
        );
    }

    [Fact]
    public void BlockedJobsResolveNamesAndOfferActions()
    {
        var f = Setup();
        Blocked(f, "RequestUrlTooLong");
        var row = Assert.Single(List(f, "BlockedJobs").Problems);
        Assert.Equal("FolderJob", row.Kind);
        Assert.Equal("Folder job", row.KindLabel);
        Assert.Equal("Example · Accounts", row.Title);
        Assert.Equal("Example", row.Record!.Name);
        Assert.Equal("Account", row.Record.TableLabel);
        Assert.Equal("Accounts", row.TemplateName);
        Assert.Equal("General", row.Library!.Name);
        Assert.Equal("Proto", row.Site!.Name);
        Assert.Equal("The folder path is too long for SharePoint's address limit.", row.Problem);
        Assert.Equal(
            "Shorten the record's value or the template's folder names, then re-run the record.",
            row.Fix
        );
        Assert.Equal(new[] { "Retry", "Cancel", "OpenRecord", "Check" }, row.Actions);
    }

    [Fact]
    public void ARecordTheCallerCannotReadHasNoName()
    {
        var f = Setup();
        Blocked(f, "WorkerFailed");
        f.Service.Rows.Remove(f.RecordId);
        var row = Assert.Single(List(f, "BlockedJobs").Problems);
        Assert.Null(row.Record!.Name);
        Assert.Equal("Record not available · Accounts", row.Title);
        Assert.DoesNotContain("OpenRecord", row.Actions);
    }

    [Fact]
    public void ARecordOfATableDeletedFromTheOrganizationIsListedWithoutAFault()
    {
        var f = Setup();
        Blocked(f, "WorkerFailed");
        f.Service.MissingTables.Add("account");
        var row = Assert.Single(List(f, "BlockedJobs").Problems);
        Assert.Null(row.Record!.Name);
        Assert.Equal("account", row.Record.TableLabel);
    }

    [Fact]
    public void EveryKindHasATitle()
    {
        var f = Setup();
        f.Store.Create(
            "asx_operation",
            new SecurityOperation
            {
                Key = "policywork:a",
                Status = "Blocked",
                LibraryId = f.LibraryId,
                ErrorCode = "SecurityWriteRejected",
            }
        );
        f.Store.Create(
            "asx_operation",
            new LibrarySetup
            {
                Key = "librarycreate:a",
                Status = "Blocked",
                SiteId = f.SiteId,
                Name = "Project documents",
            }
        );
        f.Store.Create(
            "asx_operation",
            new CatalogProbe
            {
                Key = "catalogprobe:repoint:a",
                Status = "Blocked",
                Repoint = true,
                ListId = f.ListId,
                CatalogId = f.LibraryId,
            }
        );
        f.Store.Create(
            "asx_operation",
            new CatalogProbe
            {
                Key = "catalogprobe:site",
                Status = "Blocked",
                SiteId = f.SiteId,
                DisplayName = "",
            }
        );
        var titles = List(f, "BlockedJobs")
            .Problems.ToDictionary(p => p.Key, p => (p.Kind, p.Title));
        Assert.Equal(("AccessRun", "General"), titles["policywork:a"]);
        Assert.Equal(("LibrarySetup", "Project documents"), titles["librarycreate:a"]);
        Assert.Equal(("Repoint", "General"), titles["catalogprobe:repoint:a"]);
        Assert.Equal(("SiteCheck", "Proto"), titles["catalogprobe:site"]);
    }

    [Fact]
    public void ALibrarySetupRowCarriesItsFindingAndChoices()
    {
        var f = Setup();
        var finding = LibraryReconcile.Decide(
            f.SiteId,
            "Docs",
            "/sites/proto/Docs",
            null,
            null,
            f.Now,
            Array.Empty<CatalogEntryRow>(),
            f.Now
        );
        f.Store.Create(
            "asx_operation",
            new LibrarySetup
            {
                Key = "librarycreate:lost",
                Status = "RecoveryRequired",
                SiteId = f.SiteId,
                Name = "Docs",
                Mutation = "CreateLibrary",
                ExternalSubmitted = true,
                WritePermitted = true,
                Recovery = finding,
            }
        );
        var row = Assert.Single(List(f, "BlockedJobs").Problems);
        Assert.Equal("NotFound", row.Recovery!.State);
        Assert.Equal(new[] { "CreateAgain", "CheckAgain", "Cancel" }, row.Actions);
        Assert.Equal(
            "SharePoint has no library named Docs. The creation didn't happen.",
            row.Problem
        );
        Assert.Equal(
            f.Store.Require<LibrarySetup>("asx_operation", "librarycreate:lost").Row.RowVersion,
            row.RowVersion
        );
    }

    [Fact]
    public void A0103LostCreateIsPickedUpWhenListed()
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
        var row = Assert.Single(List(f, "BlockedJobs").Problems);
        Assert.Equal("Reconciling", row.Status);
        Assert.Equal("Checking", row.Recovery!.State);
        Assert.Equal(new[] { "Cancel" }, row.Actions);
        Assert.Equal("Checking SharePoint for Old.", row.Problem);
    }

    [Fact]
    public void RunRowsAreListedAsRunsAndNeverAsBlockedRecords()
    {
        var f = Setup();
        f.Service.Transaction(() =>
            new TemplateRun(f.Service, new[] { "account" }, () => f.Now).Start(
                new WorkerRequest { TemplateId = f.TemplateId, RequestId = Guid.NewGuid() },
                Guid.NewGuid()
            )
        );
        var run = f.Store.Require<OutboxDocument>("asx_outbox", TemplateRun.Key(f.TemplateId));
        run.Value.Status = "Blocked";
        run.Value.Notices = new[] { WorkerCoordinator.PlanningFailedNotice };
        f.Store.Save(run);
        f.BlockRecord();
        var blocked = List(f, "BlockedRecords").Problems;
        Assert.DoesNotContain(blocked, p => p.Key == run.Value.Key);
        var record = Assert.Single(blocked);
        Assert.Equal("RecordPlan", record.Kind);
        // The run's wording is for runs only: a record keeps the notice it was given.
        Assert.Equal(WorkerCoordinator.PlanningFailedNotice, record.Problem);
        Assert.Equal(new[] { "Retry", "OpenRecord", "Check" }, record.Actions);
        var runs = Assert.Single(List(f, "TemplateRuns").Problems);
        Assert.Equal("TemplateRun", runs.Kind);
        Assert.Equal("Template re-run", runs.KindLabel);
        Assert.Equal("Accounts · Account", runs.Title);
        Assert.Equal("Blocked", runs.Run!.State);
        Assert.Equal("The re-run stopped while planning a page of records.", runs.Problem);
        Assert.Equal("Retry. Records already planned are kept.", runs.Fix);
        Assert.Equal(new[] { "Retry", "CancelRun" }, runs.Actions);
    }

    [Fact]
    public void RunningRunsArePausedAndPausedRunsResumed()
    {
        var f = Setup();
        var runs = new TemplateRun(f.Service, new[] { "account" }, () => f.Now);
        var started = f.Service.Transaction(() =>
            runs.Start(
                new WorkerRequest { TemplateId = f.TemplateId, RequestId = Guid.NewGuid() },
                Guid.NewGuid()
            )
        );
        Assert.Equal(
            new[] { "Pause", "CancelRun" },
            Assert.Single(List(f, "TemplateRuns").Problems).Actions
        );
        f.Service.Transaction(() => runs.Pause(started.Key));
        Assert.Equal(
            new[] { "Resume", "CancelRun" },
            Assert.Single(List(f, "TemplateRuns").Problems).Actions
        );
    }

    [Fact]
    public void WaitingRecordsShowTheFirstNoticeShortAndTheRest()
    {
        var f = Setup();
        f.Store.Create(
            "asx_outbox",
            new RecordPlanDocument
            {
                Key = "recordplan:" + f.TemplateId.ToString("N") + ":" + f.RecordId.ToString("N"),
                Status = WorkerCoordinator.WaitingStatus,
                Table = "account",
                TemplateId = f.TemplateId,
                RecordId = f.RecordId,
                Waiting = new[]
                {
                    "Folder 'general/deep' is waiting for a shorter path: its path would be 404 characters, and SharePoint allows 400 including file names. Shorten the record's value or the template's folder names, then replan the record.",
                    "Folder 'general/a' is waiting for 'root.code' to have a value. Fill in 'root.code', then replan the record.",
                    "Folder 'general/b' is waiting for 'root.code' to have a value. Fill in 'root.code', then replan the record.",
                },
            }
        );
        var row = Assert.Single(List(f, "WaitingRecords").Problems);
        Assert.Equal(
            "Folder 'general/deep' needs a shorter path: 404 characters; the limit is 400. Shorten the record's value or the template's folder names, then re-run the record.",
            row.Problem
        );
        Assert.Equal(2, row.More.Length);
        Assert.Equal(
            "Folder 'general/a' is waiting for 'root.code' to have a value. Fill in 'root.code', then re-run the record.",
            row.More[0]
        );
        Assert.Equal("Example · Accounts", row.Title);
        Assert.Equal(new[] { "Rerun", "OpenRecord", "Check" }, row.Actions);
    }

    [Fact]
    public void MissedChangesListTheCaptureJobsOfTheDocumentsHandlers()
    {
        var f = Setup();
        QueryExpression? asked = null;
        f.Service.QueryHook = q =>
        {
            if (q.EntityName == "asyncoperation")
                asked = q;
            return null;
        };
        var job = Guid.NewGuid();
        f.Service.Seed(
            new Entity("asyncoperation", job)
            {
                ["statuscode"] = new OptionSetValue(31),
                ["message"] = "Runtime identity/source scope is incomplete.\r\nat …",
                ["createdon"] = f.Now,
                ["regardingobjectid"] = new EntityReference("account", f.RecordId),
            }
        );
        var row = Assert.Single(List(f, "NotCaptured").Problems);
        Assert.Equal(job.ToString("D"), row.Key);
        Assert.Equal("CaptureJob", row.Kind);
        Assert.Equal("Missed change", row.KindLabel);
        Assert.Equal("Example · Account", row.Title);
        Assert.Equal("Runtime identity/source scope is incomplete.", row.Problem);
        Assert.Equal(new[] { "Rerun", "Dismiss", "OpenRecord" }, row.Actions);
        var plugin = asked!.LinkEntities.Single().LinkEntities.Single();
        Assert.Equal("plugintype", plugin.LinkToEntityName);
        Assert.Equal(
            new[]
            {
                EventRegistrations.RecordHandler,
                EventRegistrations.RetirementHandler,
                EventRegistrations.MembershipHandler,
            },
            plugin.LinkCriteria.Conditions.Single().Values.Cast<string>().ToArray()
        );
    }

    [Fact]
    public void PagesHoldFiftyRowsAndNextLeadsToTheRest()
    {
        var f = Setup();
        for (int i = 0; i < 60; i++)
            f.Store.Create(
                "asx_outbox",
                new OutboxDocument
                {
                    Key = "request:" + i,
                    Status = "Blocked",
                    Table = "account",
                    TemplateId = f.TemplateId,
                    RecordId = f.RecordId,
                }
            );
        var first = List(f, "BlockedRecords");
        Assert.Equal(ProblemList.PageSize, first.Problems.Length);
        Assert.NotNull(first.Next);
        var second = List(f, "BlockedRecords", first.Next);
        Assert.Equal(10, second.Problems.Length);
        Assert.Null(second.Next);
        Assert.Empty(
            first.Problems.Select(p => p.Key).Intersect(second.Problems.Select(p => p.Key))
        );
    }

    [Theory]
    [InlineData("not base64!")]
    [InlineData("bm8gbmV3bGluZQ==")]
    [InlineData("eAo=")]
    public void APageTheServerDidNotWriteIsRefused(string page)
    {
        var f = Setup();
        Assert.Equal(
            "The list changed. Refresh it.",
            Assert.Throws<EvaluationBlockedException>(() => List(f, "BlockedRecords", page)).Message
        );
    }

    [Fact]
    public void AnUnknownListIsRefused()
    {
        var f = Setup();
        Assert.Throws<EvaluationBlockedException>(() => List(f, "Everything"));
    }

    [Fact]
    public void RecentOperationsListEveryStatus()
    {
        var f = Setup();
        var applied = f.Operation;
        applied.Status = "Applied";
        f.Store.Create("asx_operation", applied);
        var row = Assert.Single(List(f, "RecentOperations").Problems);
        Assert.Equal("Applied", row.Status);
        Assert.Equal(new[] { "OpenRecord", "Check" }, row.Actions);
        Assert.Empty(List(f, "BlockedJobs").Problems);
    }

    [Fact]
    public void ABlockedTeamRefreshIsAnOutboxRowRetriedAsOne()
    {
        // A team-membership refresh is an outbox row: Retry sends RetryOutbox (Kind RecordPlan),
        // and there is no Cancel, which only operations have.
        var f = Setup();
        var team = Guid.NewGuid();
        f.Service.Seed(new Entity("team", team) { ["name"] = "Sales" });
        f.Store.Create(
            "asx_outbox",
            new OutboxDocument
            {
                Key = "team-event:" + Guid.NewGuid().ToString("N") + ":" + team.ToString("N"),
                Status = "Blocked",
                SecurityTeamId = team,
                Notices = new[] { WorkerCoordinator.PlanningFailedNotice },
            }
        );
        var row = Assert.Single(List(f, "BlockedRecords").Problems);
        Assert.Equal("RecordPlan", row.Kind);
        Assert.Equal("Library access", row.KindLabel);
        Assert.Equal("Sales · Library access", row.Title);
        Assert.Equal(new[] { "Retry", "OpenRecord" }, row.Actions);
    }

    [Fact]
    public void AMissedChangeOfADeletedRecordDoesNotOfferRerun()
    {
        var f = Setup();
        f.Service.Seed(
            new Entity("asyncoperation", Guid.NewGuid())
            {
                ["statuscode"] = new OptionSetValue(31),
                ["message"] = "Record was deleted.",
                ["createdon"] = f.Now,
                ["regardingobjectid"] = new EntityReference("account", Guid.NewGuid()),
            }
        );
        var row = Assert.Single(List(f, "NotCaptured").Problems);
        Assert.NotNull(row.Record);
        Assert.Null(row.Record!.Name);
        Assert.DoesNotContain("Rerun", row.Actions);
        Assert.Contains("Dismiss", row.Actions);
    }

    [Fact]
    public void AMissedChangeOffersRerunOnlyForARecordOfATable()
    {
        // RerunRecord refuses a team (the membership and retirement handlers' jobs) and a job
        // with no record, so neither offers it.
        var f = Setup();
        var team = Guid.NewGuid();
        f.Service.Seed(new Entity("team", team) { ["name"] = "Sales" });
        f.Service.Seed(
            new Entity("asyncoperation", Guid.NewGuid())
            {
                ["statuscode"] = new OptionSetValue(31),
                ["message"] = "Team membership failed.",
                ["createdon"] = f.Now,
                ["regardingobjectid"] = new EntityReference("team", team),
            }
        );
        f.Service.Seed(
            new Entity("asyncoperation", Guid.NewGuid())
            {
                ["statuscode"] = new OptionSetValue(31),
                ["message"] = "No record.",
                ["createdon"] = f.Now.AddMinutes(-1),
            }
        );
        var rows = List(f, "NotCaptured").Problems;
        Assert.Equal(2, rows.Length);
        Assert.Equal(new[] { "Dismiss", "OpenRecord" }, rows.Single(r => r.Record != null).Actions);
        Assert.Equal(new[] { "Dismiss" }, rows.Single(r => r.Record == null).Actions);
    }

    [Fact]
    public void RetryingJobsShowTheCauseAndTheNextAttempt()
    {
        var f = Setup();
        var job = f.Operation;
        DocumentStore.Wait(job, 429, null, null, f.Now);
        f.Store.Create("asx_operation", job);
        var row = Assert.Single(List(f, "RetryingJobs").Problems);
        Assert.Equal("SharePoint is busy or limiting requests (HTTP 429).", row.Problem);
        Assert.Null(row.Fix);
        Assert.Equal(1, row.Attempt);
        Assert.Equal(job.NextAttemptUtc, row.NextAttemptUtc);
        Assert.Equal(new[] { "Retry", "Cancel", "OpenRecord", "Check" }, row.Actions);
    }

    [Theory]
    [InlineData(
        "its path would be 2291 characters, and SharePoint allows 400 including file names.",
        "needs a shorter path: 2,291 characters; the limit is 400."
    )]
    [InlineData(
        "written into a SharePoint address it would be 2,291 characters, and the HTTP connector accepts 2,048.",
        "needs a shorter path: 2,291 characters; the limit is 2,048."
    )]
    public void LongWaitingNoticesAreShortenedOnRead(string tail, string shortTail)
    {
        Assert.Equal(
            "Folder 'general/deep' " + shortTail + " Fill in 'root.code', then re-run the record.",
            ProblemText.Shorten(
                "Folder 'general/deep' is waiting for a shorter path: "
                    + tail
                    + " Fill in 'root.code', then replan the record."
            )
        );
    }

    [Fact]
    public void TheShortFileNameNoticeIsShortenedOnRead()
    {
        Assert.Equal(
            "Folder 'general/root': files inside need short names; its path is 350 of 400 characters.",
            ProblemText.Shorten(
                "Folder 'general/root': This folder's path is 350 characters; SharePoint allows 400 including file names, so files inside need short names."
            )
        );
    }

    [Fact]
    public void AnUnknownCodeFallsBackToAPlainSentence()
    {
        var (problem, fix) = ProblemText.Describe("SomethingNew", "Blocked", null);
        Assert.Equal("Stopped with code SomethingNew.", problem);
        Assert.Equal("Retry, or cancel the job.", fix);
        Assert.Equal(
            ("SharePoint is busy or limiting requests (HTTP 429).", (string?)null),
            ProblemText.Describe(
                null,
                "RetryWait",
                "Waiting to retry after a temporary error (HTTP 429, 0x80072321); attempt 3."
            )
        );
        Assert.Equal(
            ("A temporary error stopped the last attempt (HTTP 503).", (string?)null),
            ProblemText.Describe(
                "Waiting to retry after a temporary error (HTTP 503); attempt 1.",
                "RetryWait",
                null
            )
        );
        Assert.Equal(
            ("A sentence the server wrote.", (string?)null),
            ProblemText.Describe("A sentence the server wrote.", "Blocked", null)
        );
    }

    [Fact]
    public void EveryBlockCodeHasReadableText()
    {
        var codes = new SortedSet<string>(StringComparer.Ordinal);
        foreach (
            var file in Directory.GetFiles(
                Path.Combine(SourceRoot(), "src", "Ascentix.Documents.Dataverse"),
                "*.cs"
            )
        )
        {
            var text = File.ReadAllText(file);
            foreach (
                Match call in Regex.Matches(
                    text,
                    @"\b(?:return\s+Block|Stop)\((?<args>[^;]*?)\);",
                    RegexOptions.Singleline
                )
            )
            foreach (
                Match literal in Regex.Matches(
                    call.Groups["args"].Value,
                    "\"(?<code>[A-Z][a-z]+(?:[A-Z][a-z]+)+)\""
                )
            )
                codes.Add(literal.Groups["code"].Value);
            foreach (
                Match set in Regex.Matches(
                    text,
                    "ErrorCode = \"(?<code>[A-Z][a-z]+(?:[A-Z][a-z]+)+)\""
                )
            )
                codes.Add(set.Groups["code"].Value);
        }
        Assert.Contains("RequestUrlTooLong", codes);
        Assert.Contains("TableNotEnabled", codes);
        Assert.DoesNotContain(codes, code => !ProblemText.Codes.Contains(code));
    }

    private static string SourceRoot()
    {
        for (
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            dir != null;
            dir = dir.Parent
        )
            if (File.Exists(Path.Combine(dir.FullName, "Ascentix.Documents.sln")))
                return dir.FullName;
        throw new InvalidOperationException(
            "The repository root was not found above the test output."
        );
    }
}
