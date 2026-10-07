using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Security;
using System.Text;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

/// <summary>Monitor's counts (spec 6.1).</summary>
[DataContract]
public sealed class ProblemSummary
{
    [DataMember]
    public int TemplateRuns { get; set; }

    [DataMember]
    public int NotCaptured { get; set; }

    [DataMember]
    public int BlockedRecords { get; set; }

    [DataMember]
    public int WaitingRecords { get; set; }

    [DataMember]
    public int BlockedJobs { get; set; }

    [DataMember]
    public int RetryingJobs { get; set; }

    /// <summary>
    /// Lists with more rows than Dataverse counts (ProblemList.CountLimit); their count is
    /// CountLimit and the page shows it as "5,000+".
    /// </summary>
    [DataMember]
    public string[] Capped { get; set; } = Array.Empty<string>();

    [DataMember]
    public DateTime CountedUtc { get; set; }
}

[DataContract]
public sealed class ProblemRecord
{
    [DataMember]
    public string Table { get; set; } = "";

    [DataMember]
    public string TableLabel { get; set; } = "";

    [DataMember]
    public Guid Id { get; set; }

    /// <summary>Null when the caller cannot read the record.</summary>
    [DataMember]
    public string? Name { get; set; }
}

[DataContract]
public sealed class NamedRef
{
    [DataMember]
    public Guid Id { get; set; }

    [DataMember]
    public string? Name { get; set; }
}

/// <summary>One row of a Monitor list (spec 6.2).</summary>
[DataContract]
public sealed class ProblemRow
{
    [DataMember]
    public string Key { get; set; } = "";

    /// <summary>LibrarySetup rows: the setup row's version, which ResolveSetup checks.</summary>
    [DataMember]
    public string? RowVersion { get; set; }

    [DataMember]
    public string Kind { get; set; } = "";

    [DataMember]
    public string KindLabel { get; set; } = "";

    [DataMember]
    public string Title { get; set; } = "";

    [DataMember]
    public ProblemRecord? Record { get; set; }

    [DataMember]
    public Guid? TemplateId { get; set; }

    [DataMember]
    public string? TemplateName { get; set; }

    [DataMember]
    public NamedRef? Library { get; set; }

    [DataMember]
    public NamedRef? Site { get; set; }

    [DataMember]
    public string Status { get; set; } = "";

    [DataMember]
    public string Problem { get; set; } = "";

    [DataMember]
    public string? Fix { get; set; }

    [DataMember]
    public string? Code { get; set; }

    [DataMember]
    public string[] More { get; set; } = Array.Empty<string>();

    [DataMember]
    public DateTime? SinceUtc { get; set; }

    [DataMember]
    public DateTime? NextAttemptUtc { get; set; }

    [DataMember]
    public int Attempt { get; set; }

    [DataMember]
    public TemplateRunState? Run { get; set; }

    [DataMember]
    public LibraryRecovery? Recovery { get; set; }

    [DataMember]
    public string[] Actions { get; set; } = Array.Empty<string>();
}

/// <summary>
/// Monitor's counts and lists (spec 6.1, 6.2). Every read runs as the caller. Nothing here runs
/// a FetchXML aggregate or catches a service fault: inside a plug-in either would end the
/// transaction (see TableInfo).
/// </summary>
public static class ProblemList
{
    /// <summary>Rows per list page: the page of every Monitor list (spec 3, Paging).</summary>
    public const int PageSize = 50;

    /// <summary>
    /// The most a count reports. Summary counts with returntotalrecordcount, which Dataverse
    /// counts up to 5,000 rows; past that it reports 5,000 and sets
    /// TotalRecordCountLimitExceeded. A FetchXML aggregate would count further, but Dataverse
    /// refuses one over 50,000 rows (AggregateQueryRecordLimit), and inside a plug-in that
    /// refusal ends the transaction even when caught.
    /// </summary>
    public const int CountLimit = 5000;

    private static readonly string[] CaptureHandlers =
    {
        EventRegistrations.RecordHandler,
        EventRegistrations.RetirementHandler,
        EventRegistrations.MembershipHandler,
    };

    /// <summary>Six counts; a count Dataverse stopped at CountLimit is reported capped.</summary>
    public static ProblemSummary Summary(IOrganizationService service, DateTime now)
    {
        var capped = new List<string>();
        int Count(string list, string fetch)
        {
            var counted = service.RetrieveMultiple(new FetchExpression(fetch));
            if (counted.TotalRecordCountLimitExceeded)
            {
                capped.Add(list);
                return CountLimit;
            }
            return Math.Max(0, counted.TotalRecordCount);
        }
        return new ProblemSummary
        {
            TemplateRuns = Count(
                "TemplateRuns",
                Fetch(
                    "asx_outbox",
                    "<condition attribute=\"asx_table\" operator=\"eq\" value=\""
                        + TemplateRun.IndexMarker
                        + "\" />"
                        + In("asx_status", "Pending", "Paused", "Blocked")
                )
            ),
            BlockedRecords = Count(
                "BlockedRecords",
                Fetch(
                    "asx_outbox",
                    "<condition attribute=\"asx_status\" operator=\"eq\" value=\"Blocked\" />"
                        + "<filter type=\"or\"><condition attribute=\"asx_table\" operator=\"null\" />"
                        + "<condition attribute=\"asx_table\" operator=\"ne\" value=\""
                        + TemplateRun.IndexMarker
                        + "\" /></filter>"
                )
            ),
            WaitingRecords = Count(
                "WaitingRecords",
                Fetch(
                    "asx_outbox",
                    "<condition attribute=\"asx_status\" operator=\"eq\" value=\""
                        + WorkerCoordinator.WaitingStatus
                        + "\" />"
                )
            ),
            BlockedJobs = Count(
                "BlockedJobs",
                Fetch("asx_operation", In("asx_status", BlockedJobStatuses))
            ),
            RetryingJobs = Count(
                "RetryingJobs",
                Fetch(
                    "asx_operation",
                    "<condition attribute=\"asx_status\" operator=\"eq\" value=\"RetryWait\" />"
                )
            ),
            NotCaptured = Count(
                "NotCaptured",
                "<fetch count=\"1\" returntotalrecordcount=\"true\"><entity name=\"asyncoperation\">"
                    + "<attribute name=\"asyncoperationid\" />"
                    + "<filter><condition attribute=\"statuscode\" operator=\"eq\" value=\"31\" /></filter>"
                    + "<link-entity name=\"sdkmessageprocessingstep\" from=\"sdkmessageprocessingstepid\" to=\"owningextensionid\" link-type=\"inner\">"
                    + "<link-entity name=\"plugintype\" from=\"plugintypeid\" to=\"eventhandler\" link-type=\"inner\"><filter>"
                    + In("typename", CaptureHandlers)
                    + "</filter></link-entity></link-entity></entity></fetch>"
            ),
            Capped = capped.ToArray(),
            CountedUtc = now,
        };
    }

    // Library setups whose create is being looked up (Reconciling) wait for the admin too.
    private static readonly string[] BlockedJobStatuses =
    {
        "Blocked",
        "RecoveryRequired",
        "Reconciling",
    };

    // One row per page is enough: only TotalRecordCount is read.
    private static string Fetch(string table, string conditions) =>
        "<fetch count=\"1\" returntotalrecordcount=\"true\"><entity name=\""
        + table
        + "\"><attribute name=\""
        + table
        + "id\" /><filter>"
        + conditions
        + "</filter></entity></fetch>";

    private static string In(string attribute, params string[] values) =>
        "<condition attribute=\""
        + attribute
        + "\" operator=\"in\">"
        + string.Concat(values.Select(v => "<value>" + SecurityElement.Escape(v) + "</value>"))
        + "</condition>";

    /// <summary>One page of one list.</summary>
    /// <param name="service">The caller's service (writes only for the legacy pickup, decision D6).</param>
    /// <param name="request">List and Page.</param>
    /// <param name="now">Now, for times and the 24-hour window of ended runs.</param>
    /// <param name="allowed">The enabled tables, for template-run descriptions.</param>
    public static WorkerResult List(
        IOrganizationService service,
        WorkerRequest request,
        DateTime now,
        string[] allowed
    )
    {
        var (page, cookie) = ReadPage(request.Page);
        var names = new Names(service);
        var rows = new List<ProblemRow>();
        EntityCollection found;
        switch (request.List)
        {
            case "NotCaptured":
                found = service.RetrieveMultiple(CaptureQuery(page, cookie));
                rows.AddRange(found.Entities.Select(row => CaptureRow(row, names)));
                break;
            case "BlockedRecords":
            case "WaitingRecords":
            case "TemplateRuns":
                found = service.RetrieveMultiple(OutboxQuery(request.List, page, cookie, now));
                var runs = new TemplateRun(service, allowed, () => now);
                rows.AddRange(
                    found.Entities.Select(row => OutboxRow(request.List, row, names, runs))
                );
                break;
            case "BlockedJobs":
            case "RetryingJobs":
            case "RecentOperations":
                found = service.RetrieveMultiple(OperationQuery(request.List, page, cookie));
                var setups = new LibraryProvisioning(service, () => now);
                rows.AddRange(found.Entities.Select(row => OperationRow(row, names, setups)));
                break;
            default:
                throw new EvaluationBlockedException(
                    "Choose a list: TemplateRuns, NotCaptured, BlockedRecords, WaitingRecords, BlockedJobs, RetryingJobs or RecentOperations."
                );
        }
        names.Resolve(rows);
        foreach (var row in rows)
            row.Actions = Actions(row);
        return new WorkerResult
        {
            Status = "Page",
            Problems = rows.ToArray(),
            Next = found.MoreRecords ? WritePage(page + 1, found.PagingCookie) : null,
        };
    }

    private static QueryExpression Paged(
        string table,
        int page,
        string? cookie,
        string order,
        params string[] columns
    )
    {
        var query = new QueryExpression(table)
        {
            ColumnSet = new ColumnSet(columns),
            PageInfo = new PagingInfo
            {
                Count = PageSize,
                PageNumber = page,
                PagingCookie = cookie,
            },
        };
        // Newest first; the row ID breaks ties so pages never overlap.
        query.AddOrder(order, OrderType.Descending);
        query.AddOrder(table + "id", OrderType.Descending);
        return query;
    }

    private static QueryExpression CaptureQuery(int page, string? cookie)
    {
        var query = Paged(
            "asyncoperation",
            page,
            cookie,
            "createdon",
            "asyncoperationid",
            "message",
            "createdon",
            "regardingobjectid"
        );
        query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 31);
        var step = query.AddLink(
            "sdkmessageprocessingstep",
            "owningextensionid",
            "sdkmessageprocessingstepid"
        );
        var plugin = step.AddLink("plugintype", "eventhandler", "plugintypeid");
        plugin.LinkCriteria.AddCondition(
            "typename",
            ConditionOperator.In,
            CaptureHandlers.Cast<object>().ToArray()
        );
        return query;
    }

    private static QueryExpression OutboxQuery(string list, int page, string? cookie, DateTime now)
    {
        var query = Paged(
            "asx_outbox",
            page,
            cookie,
            list == "TemplateRuns" ? "createdon" : "modifiedon",
            "asx_payload",
            "modifiedon"
        );
        if (list == "TemplateRuns")
        {
            // Active runs, and runs that ended in the last 24 hours (spec 6.2 Paging).
            query.Criteria.AddCondition(
                "asx_table",
                ConditionOperator.Equal,
                TemplateRun.IndexMarker
            );
            var shown = new FilterExpression(LogicalOperator.Or);
            shown.AddCondition("asx_status", ConditionOperator.In, "Pending", "Paused", "Blocked");
            shown.AddCondition("modifiedon", ConditionOperator.GreaterEqual, now.AddHours(-24));
            query.Criteria.AddFilter(shown);
            return query;
        }
        query.Criteria.AddCondition(
            "asx_status",
            ConditionOperator.Equal,
            list == "WaitingRecords" ? WorkerCoordinator.WaitingStatus : "Blocked"
        );
        if (list == "BlockedRecords")
        {
            var notRuns = new FilterExpression(LogicalOperator.Or);
            notRuns.AddCondition("asx_table", ConditionOperator.Null);
            notRuns.AddCondition("asx_table", ConditionOperator.NotEqual, TemplateRun.IndexMarker);
            query.Criteria.AddFilter(notRuns);
        }
        return query;
    }

    private static QueryExpression OperationQuery(string list, int page, string? cookie)
    {
        var query = Paged(
            "asx_operation",
            page,
            cookie,
            list == "RecentOperations" ? "createdon" : "modifiedon",
            "asx_payload",
            "modifiedon"
        );
        if (list == "BlockedJobs")
            query.Criteria.AddCondition(
                "asx_status",
                ConditionOperator.In,
                BlockedJobStatuses.Cast<object>().ToArray()
            );
        if (list == "RetryingJobs")
            query.Criteria.AddCondition("asx_status", ConditionOperator.Equal, "RetryWait");
        return query;
    }

    private static ProblemRow CaptureRow(Entity row, Names names)
    {
        var regarding = row.GetAttributeValue<EntityReference>("regardingobjectid");
        string message =
            (row.GetAttributeValue<string>("message") ?? "")
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(line => line.Trim())
                .FirstOrDefault(line => line.Length > 0)
            ?? "The background job failed without a message.";
        return new ProblemRow
        {
            Key = row.Id.ToString("D"),
            Kind = "CaptureJob",
            KindLabel = "Missed change",
            Status = "Failed",
            Problem = message,
            SinceUtc = row.GetAttributeValue<DateTime?>("createdon"),
            Record = regarding == null ? null : names.Record(regarding.LogicalName, regarding.Id),
        };
    }

    private static ProblemRow OutboxRow(string list, Entity row, Names names, TemplateRun runs)
    {
        string payload = row.GetAttributeValue<string>("asx_payload");
        var since = row.GetAttributeValue<DateTime?>("modifiedon");
        if (list == "WaitingRecords")
        {
            var plan = JsonWire.Read<RecordPlanDocument>(payload);
            var waits = (plan.Waiting.Length > 0 ? plan.Waiting : plan.Notices)
                .Select(ProblemText.Shorten)
                .ToArray();
            return new ProblemRow
            {
                Key = plan.Key,
                Kind = "RecordPlan",
                KindLabel = "Record",
                Status = plan.Status,
                Problem = waits.FirstOrDefault() ?? "Waiting for record data.",
                More = waits.Skip(1).ToArray(),
                SinceUtc = since,
                Record = names.Record(plan.Table, plan.RecordId),
                TemplateId = names.Template(plan.TemplateId),
            };
        }
        var work = JsonWire.Read<OutboxDocument>(payload);
        if (work.Key.StartsWith(TemplateRun.Prefix, StringComparison.Ordinal))
        {
            var state = runs.Describe(work);
            var (runProblem, runFix) =
                state.State == "Blocked"
                    ? ProblemText.DescribeRun(work.Notices.FirstOrDefault())
                    : (state.Problem ?? "", (string?)null);
            return new ProblemRow
            {
                Key = work.Key,
                Kind = "TemplateRun",
                KindLabel = "Template re-run",
                Title = state.TemplateName + " · " + state.TableLabel,
                Status = work.Status,
                Problem = runProblem,
                Fix = runFix,
                SinceUtc = work.StartedUtc,
                NextAttemptUtc = state.NextAttemptUtc,
                Run = state,
                TemplateId = names.Template(work.TemplateId),
            };
        }
        var (problem, fix) = ProblemText.Describe(null, work.Status, work.Notices.FirstOrDefault());
        var outbox = new ProblemRow
        {
            Key = work.Key,
            Kind = "RecordPlan",
            KindLabel = "Record",
            Status = work.Status,
            Problem = problem,
            Fix = fix,
            More = work.Notices.Skip(1).Select(ProblemText.Shorten).ToArray(),
            SinceUtc = since,
        };
        if (work.RecordId != Guid.Empty)
        {
            outbox.Record = names.Record(work.Table, work.RecordId);
            outbox.TemplateId = names.Template(work.TemplateId);
        }
        else if (work.RelatedRecordId != Guid.Empty)
        {
            outbox.Record = names.Record(work.RelatedTable, work.RelatedRecordId);
            outbox.Title = "Records that use this " + work.RelatedTable;
        }
        else if (work.SecurityTeamId != Guid.Empty)
        {
            // A team-membership refresh stays an outbox row (Kind RecordPlan): Retry sends
            // RetryOutbox and there is no Cancel. Only its label says what it is.
            outbox.KindLabel = "Library access";
            outbox.Record = names.Record("team", work.SecurityTeamId);
        }
        return outbox;
    }

    private static ProblemRow OperationRow(Entity row, Names names, LibraryProvisioning setups)
    {
        string payload = row.GetAttributeValue<string>("asx_payload");
        var basic = JsonWire.Read<OperationDocument>(payload);
        var result = new ProblemRow
        {
            Key = basic.Key,
            Status = basic.Status,
            SinceUtc = row.GetAttributeValue<DateTime?>("modifiedon"),
            NextAttemptUtc = basic.NextAttemptUtc,
            Attempt = basic.RetryCount,
            Code = basic.ErrorCode,
        };
        if (basic.Key.StartsWith("librarycreate:", StringComparison.Ordinal))
        {
            // Starts the lookup of a 0.1.0.3-era lost create the first time it is listed (D6).
            var stored = setups.PickUp(
                new DocumentStore(names.Service).Require<LibrarySetup>("asx_operation", basic.Key)
            );
            var setup = stored.Value;
            result.RowVersion = stored.Row.RowVersion;
            result.Kind = "LibrarySetup";
            result.KindLabel = "Library setup";
            result.Title = setup.Name;
            result.Status = setup.Status;
            result.Site = names.Site(setup.SiteId);
            result.Recovery = setup.Recovery;
            if (setup.Recovery != null && LibraryProvisioning.LostCreate(setup))
            {
                result.Problem = LibraryReconcile.Sentence(setup.Name, setup.Recovery);
                return result;
            }
        }
        else if (basic.Key.StartsWith("policywork:", StringComparison.Ordinal))
        {
            var access = JsonWire.Read<SecurityOperation>(payload);
            result.Kind = "AccessRun";
            result.KindLabel = "Library access";
            result.Library = names.Library(access.LibraryId);
        }
        else if (basic.Key.StartsWith("catalogprobe:", StringComparison.Ordinal))
        {
            var probe = JsonWire.Read<CatalogProbe>(payload);
            (result.Kind, result.KindLabel) =
                probe.Repoint ? ("Repoint", "Re-point")
                : probe.DiscoverLibraries ? ("LibraryDiscovery", "Finding libraries")
                : probe.ListId == Guid.Empty ? ("SiteCheck", "Site check")
                : ("LibraryCheck", "Library check");
            result.Title = probe.DisplayName;
            if (probe.CatalogId != Guid.Empty && probe.ListId != Guid.Empty)
                result.Library = names.Library(probe.CatalogId);
            if (probe.SiteId != Guid.Empty)
                result.Site = names.Site(probe.SiteId);
        }
        else
        {
            result.Kind = "FolderJob";
            result.KindLabel = "Folder job";
            if (basic.Folders.Length > 0)
            {
                // A compacted history keeps the first folder; otherwise the cursor's folder.
                var folder =
                    basic.HistoryCompacted
                    || basic.Cursor < 0
                    || basic.Cursor >= basic.Folders.Length
                        ? basic.Folders[0]
                        : basic.Folder;
                result.Record = names.Record(folder.Table, folder.RecordId);
                result.TemplateId = names.Template(folder.TemplateId);
                result.Library = names.Library(folder.LibraryId);
            }
        }
        (result.Problem, result.Fix) = ProblemText.Describe(basic.ErrorCode, result.Status, null);
        return result;
    }

    /// <summary>The actions a row offers, from its kind and status (spec 6.2 Actions table).</summary>
    internal static string[] Actions(ProblemRow row)
    {
        var actions = new List<string>();
        bool terminal =
            row.Status == "Applied"
            || row.Status == "Cancelled"
            || row.Status == "Superseded"
            || row.Status == "Planned";
        switch (row.Kind)
        {
            case "TemplateRun":
                string state = row.Run?.State ?? "";
                if (state == "Running" || state == "Waiting" || state == "Retrying")
                    actions.Add("Pause");
                if (state == "Paused")
                    actions.Add("Resume");
                if (state == "Blocked")
                    actions.Add("Retry");
                if (state != "Done" && state != "Cancelled")
                    actions.Add("CancelRun");
                return actions.ToArray();
            case "CaptureJob":
                // RerunRecord queues a record of an enabled table; a team (the membership and
                // retirement handlers) or a job with no record would always be refused.
                if (row.Record != null && row.Record.Table != "team")
                    actions.Add("Rerun");
                actions.Add("Dismiss");
                break;
            case "RecordPlan":
                if (row.Status == "Blocked")
                    actions.Add("Retry");
                if (row.Status == WorkerCoordinator.WaitingStatus)
                    actions.Add("Rerun");
                break;
            case "LibrarySetup"
                when row.Recovery != null
                    && (row.Status == "RecoveryRequired" || row.Status == "Reconciling"):
                actions.AddRange(row.Recovery.Choices);
                return actions.ToArray();
            default:
                if (
                    row.Status == "Blocked"
                    || row.Status == "RecoveryRequired"
                    || row.Status == "RetryWait"
                )
                    actions.Add("Retry");
                if (!terminal)
                    actions.Add("Cancel");
                break;
        }
        if (row.Record?.Name != null)
            actions.Add("OpenRecord");
        if (row.Record != null && row.TemplateId != null && row.Kind != "CaptureJob")
            actions.Add("Check");
        return actions.ToArray();
    }

    // The opaque Next: the page number and Dataverse's paging cookie.
    private static string WritePage(int page, string? cookie) =>
        Convert.ToBase64String(
            Encoding.UTF8.GetBytes(
                page.ToString(CultureInfo.InvariantCulture) + "\n" + (cookie ?? "")
            )
        );

    private static (int Page, string? Cookie) ReadPage(string? next)
    {
        if (string.IsNullOrEmpty(next))
            return (1, null);
        const string Changed = "The list changed. Refresh it.";
        string text;
        try
        {
            text = Encoding.UTF8.GetString(Convert.FromBase64String(next));
        }
        catch (FormatException)
        {
            throw new EvaluationBlockedException(Changed);
        }
        var parts = text.Split(new[] { '\n' }, 2);
        if (
            parts.Length != 2
            || !int.TryParse(
                parts[0],
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out int page
            )
            || page < 2
        )
            throw new EvaluationBlockedException(Changed);
        return (page, parts[1].Length == 0 ? null : parts[1]);
    }

    /// <summary>
    /// Collects the IDs a page refers to, reads each table's names once as the caller, and fills
    /// the rows. Table metadata comes from TableInfo, which answers nothing for a table deleted
    /// from the organization instead of faulting.
    /// </summary>
    private sealed class Names
    {
        internal readonly IOrganizationService Service;
        private readonly Dictionary<(string Table, Guid Id), ProblemRecord> records = new();
        private readonly HashSet<Guid> templates = new();
        private readonly HashSet<Guid> libraries = new();
        private readonly HashSet<Guid> sites = new();

        internal Names(IOrganizationService service) => Service = service;

        internal ProblemRecord Record(string table, Guid id)
        {
            if (!records.TryGetValue((table, id), out var record))
                records[(table, id)] = record = new ProblemRecord { Table = table, Id = id };
            return record;
        }

        internal Guid? Template(Guid id)
        {
            if (id == Guid.Empty)
                return null;
            templates.Add(id);
            return id;
        }

        internal NamedRef? Library(Guid id) => Add(libraries, id);

        internal NamedRef? Site(Guid id) => Add(sites, id);

        private static NamedRef? Add(HashSet<Guid> set, Guid id)
        {
            if (id == Guid.Empty)
                return null;
            set.Add(id);
            return new NamedRef { Id = id };
        }

        internal void Resolve(List<ProblemRow> rows)
        {
            foreach (var group in records.Values.GroupBy(r => r.Table))
            {
                var metadata = string.IsNullOrEmpty(group.Key)
                    ? null
                    : TableInfo.Find(Service, group.Key);
                string label = TableInfo.Label(metadata, group.Key);
                var found =
                    metadata?.PrimaryIdAttribute is string primary
                    && metadata.PrimaryNameAttribute is string name
                        ? Read(group.Key, primary, name, group.Select(r => r.Id))
                        : new Dictionary<Guid, string?>();
                foreach (var record in group)
                {
                    record.TableLabel = label;
                    record.Name = found.TryGetValue(record.Id, out var text) ? text : null;
                }
            }
            var templateNames = Read("asx_template", "asx_templateid", "asx_name", templates);
            var libraryRows = ReadRows(
                "asx_library",
                "asx_libraryid",
                libraries,
                "asx_name",
                "asx_siteid"
            );
            foreach (var library in libraryRows.Values)
                if (
                    library.GetAttributeValue<EntityReference>("asx_siteid") is EntityReference site
                )
                    sites.Add(site.Id);
            var siteNames = Read("asx_site", "asx_siteid", "asx_name", sites);
            foreach (var row in rows)
            {
                if (row.TemplateId is Guid template)
                    row.TemplateName = templateNames.TryGetValue(template, out var t) ? t : null;
                if (row.Library != null && libraryRows.TryGetValue(row.Library.Id, out var library))
                {
                    row.Library.Name = library.GetAttributeValue<string>("asx_name");
                    if (
                        row.Site == null
                        && library.GetAttributeValue<EntityReference>("asx_siteid")
                            is EntityReference site
                    )
                        row.Site = new NamedRef { Id = site.Id };
                }
                if (row.Site != null)
                    row.Site.Name = siteNames.TryGetValue(row.Site.Id, out var s) ? s : null;
                row.Title = Title(row);
            }
        }

        /// <summary>Spec 6.2 Title: renames show, because names are read when the list is.</summary>
        private static string Title(ProblemRow row)
        {
            string record = row.Record?.Name ?? "Record not available";
            return row.Kind switch
            {
                "FolderJob"
                or "RecordPlan" when row.Record != null && string.IsNullOrEmpty(row.Title) => record
                    + " · "
                    + (row.TemplateName ?? row.Library?.Name ?? row.KindLabel),
                "RecordPlan" when !string.IsNullOrEmpty(row.Title) => row.Title + " · " + record,
                "AccessRun" => row.Library?.Name
                    ?? (row.Record?.Name != null ? "Team " + row.Record.Name : row.KindLabel),
                "CaptureJob" => record + " · " + (row.Record?.TableLabel ?? row.KindLabel),
                "LibrarySetup" or "TemplateRun" => row.Title,
                _ => !string.IsNullOrEmpty(row.Title)
                    ? row.Title
                    : row.Library?.Name ?? row.Site?.Name ?? row.KindLabel,
            };
        }

        private Dictionary<Guid, string?> Read(
            string table,
            string primary,
            string name,
            IEnumerable<Guid> ids
        ) =>
            ReadRows(table, primary, ids, name)
                .ToDictionary(p => p.Key, p => (string?)p.Value.GetAttributeValue<string>(name));

        private Dictionary<Guid, Entity> ReadRows(
            string table,
            string primary,
            IEnumerable<Guid> ids,
            params string[] columns
        )
        {
            var list = ids.Distinct().ToArray();
            if (list.Length == 0)
                return new Dictionary<Guid, Entity>();
            var query = new QueryExpression(table) { ColumnSet = new ColumnSet(columns) };
            query.Criteria.AddCondition(
                primary,
                ConditionOperator.In,
                list.Cast<object>().ToArray()
            );
            // A page refers to at most 50 rows of each kind, so one read per table is complete;
            // rows the caller cannot read are simply absent (their names stay null).
            return Service.RetrieveMultiple(query).Entities.ToDictionary(e => e.Id);
        }
    }
}
