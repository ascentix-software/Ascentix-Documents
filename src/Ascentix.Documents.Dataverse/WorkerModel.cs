using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public class StoredDocument
{
    [DataMember]
    public DateTime? RetireAfterUtc { get; set; }

    [DataMember]
    public int SchemaVersion { get; set; } = 1;

    [DataMember]
    public string Key { get; set; } = "";

    [DataMember]
    public string Status { get; set; } = "Pending";

    /// <summary>
    /// Rows saved by an earlier release lack the lists added since, and the serializer runs no
    /// constructor or initializer, so those lists would read as null. Every stored list reads
    /// as empty instead, whichever release wrote the row.
    /// </summary>
    [OnDeserialized]
    private void FillMissingLists(StreamingContext context) => StoredLists.Fill(this);
}

/// <summary>Gives a deserialized document an empty list for each list member its row lacked.</summary>
internal static class StoredLists
{
    private static readonly ConcurrentDictionary<Type, PropertyInfo[]> Lists =
        new ConcurrentDictionary<Type, PropertyInfo[]>();

    public static void Fill(object document)
    {
        foreach (var list in Lists.GetOrAdd(document.GetType(), Find))
            if (list.GetValue(document) == null)
                list.SetValue(document, Empty(list.PropertyType));
    }

    private static PropertyInfo[] Find(Type type) =>
        type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p =>
                p.CanRead
                && p.CanWrite
                && p.GetIndexParameters().Length == 0
                && p.GetCustomAttributes(typeof(DataMemberAttribute), true).Length > 0
                && p.PropertyType != typeof(string)
                && typeof(IEnumerable).IsAssignableFrom(p.PropertyType)
            )
            .ToArray();

    private static object Empty(Type type) =>
        type.IsArray
            ? Array.CreateInstance(type.GetElementType()!, 0)
            : Activator.CreateInstance(type)!;
}

[DataContract]
public sealed class OutboxDocument : StoredDocument
{
    [DataMember]
    public string RelatedTable { get; set; } = "";

    [DataMember]
    public Guid RelatedRecordId { get; set; }

    [DataMember]
    public int SourcePage { get; set; } = 1;

    [DataMember]
    public string? SourceCookie { get; set; }

    [DataMember]
    public Guid TemplateId { get; set; }

    [DataMember]
    public Guid RevisionId { get; set; }

    [DataMember]
    public Guid RecordId { get; set; }

    [DataMember]
    public string Table { get; set; } = "";

    [DataMember]
    public Guid SecurityTeamId { get; set; }

    [DataMember]
    public int SecurityPage { get; set; } = 1;

    [DataMember]
    public string? SecurityCookie { get; set; }

    [DataMember]
    public bool PinnedRevision { get; set; }

    [DataMember]
    public SourceVersion[] ReviewedSources { get; set; } = Array.Empty<SourceVersion>();

    [DataMember]
    public bool Replan { get; set; }

    [DataMember]
    public string[] Operations { get; set; } = Array.Empty<string>();

    [DataMember]
    public string[] Notices { get; set; } = Array.Empty<string>();

    // A Pending row that failed temporarily waits until this UTC time (indexed as
    // asx_nextattempt) and counts its waits in Attempts. A successful Plan clears both.
    [DataMember]
    public DateTime? NextAttemptUtc { get; set; }

    [DataMember]
    public int Attempts { get; set; }

    // Template re-run rows (TemplateRun): the folder-job priority of the records it queues, the
    // request rows of the page in flight, progress, who started it and when (spec 6.5).

    /// <summary>
    /// asx_priority of the folder jobs this row's plan creates. A row stored before this member
    /// existed reads as 0 and makes normal folder jobs (1); see WorkerCoordinator.Plan.
    /// </summary>
    [DataMember]
    public int Priority { get; set; } = 1;

    [DataMember]
    public string[] PageKeys { get; set; } = Array.Empty<string>();

    [DataMember]
    public int Planned { get; set; }

    [DataMember]
    public int Total { get; set; }

    [DataMember]
    public Guid StartedBy { get; set; }

    /// <summary>The starter's name, read as the worker at Start: an operator may not read users.</summary>
    [DataMember]
    public string? StartedByName { get; set; }

    [DataMember]
    public DateTime? StartedUtc { get; set; }

    /// <summary>The Start request that began this run; it seeds every record's request ID.</summary>
    [DataMember]
    public Guid RunRequestId { get; set; }

    /// <summary>Start or the last Resume, and Planned then: the base of the estimated finish.</summary>
    [DataMember]
    public DateTime? ResumedUtc { get; set; }

    [DataMember]
    public int PlannedAtResume { get; set; }

    [DataMember]
    public DateTime? EndedUtc { get; set; }
}

[DataContract]
public sealed class SourceVersion
{
    [DataMember]
    public string Table { get; set; } = "";

    [DataMember]
    public Guid Id { get; set; }

    [DataMember]
    public string Version { get; set; } = "";
}

[DataContract]
public sealed class RecordPlanDocument : StoredDocument
{
    [DataMember]
    public string Table { get; set; } = "";

    [DataMember]
    public Guid TemplateId { get; set; }

    [DataMember]
    public Guid RecordId { get; set; }

    [DataMember]
    public Guid RevisionId { get; set; }

    [DataMember]
    public string[] Operations { get; set; } = Array.Empty<string>();

    [DataMember]
    public SourceVersion[] Sources { get; set; } = Array.Empty<SourceVersion>();

    [DataMember]
    public string[] IncludedSections { get; set; } = Array.Empty<string>();

    private string[]? notices;

    /// <summary>
    /// What planning recorded for the admin, such as an adjusted folder name or a folder that
    /// waits for a value. Rows stored before this field read as empty.
    /// </summary>
    [DataMember]
    public string[] Notices
    {
        get => notices ?? Array.Empty<string>();
        set => notices = value;
    }

    /// <summary>
    /// The notices of folders the plan skipped until the record changes, each with what it
    /// needs. While any wait, the row's status is Waiting, which Blocked records lists.
    /// </summary>
    [DataMember]
    public string[] Waiting { get; set; } = Array.Empty<string>();
}

[DataContract]
public sealed class FolderStep : StoredDocument
{
    [DataMember]
    public Guid TemplateId { get; set; }

    [DataMember]
    public Guid RecordId { get; set; }

    [DataMember]
    public string Table { get; set; } = "";

    [DataMember]
    public string Section { get; set; } = "";

    [DataMember]
    public string Node { get; set; } = "";

    [DataMember]
    public string? ParentBinding { get; set; }

    [DataMember]
    public Guid LibraryId { get; set; }

    [DataMember]
    public Guid EntryId { get; set; }

    [DataMember]
    public string OriginalName { get; set; } = "";

    [DataMember]
    public string Candidate { get; set; } = "";

    [DataMember]
    public Guid PhysicalId { get; set; }

    [DataMember]
    public string? PhysicalPath { get; set; }

    [DataMember]
    public Guid LocationId { get; set; }
}

[DataContract]
public class OperationDocument : StoredDocument
{
    [DataMember]
    public bool HistoryCompacted { get; set; }

    [DataMember]
    public Guid ResultLibraryId { get; set; }

    [DataMember]
    public Guid ResultLocationId { get; set; }

    [DataMember]
    public Guid ResultPhysicalId { get; set; }

    [DataMember]
    public FolderStep[] Folders { get; set; } = Array.Empty<FolderStep>();

    [DataMember]
    public int Cursor { get; set; }

    [DataMember]
    public DateTime? CompletedUtc { get; set; }
    public FolderStep Folder =>
        Folders.Length > 0 && Cursor >= 0 && Cursor < Folders.Length
            ? Folders[Cursor]
            : throw new Ascentix.Documents.Conditions.EvaluationBlockedException(
                "Folder job cursor is invalid."
            );

    [DataMember]
    public int Priority { get; set; } = 1;

    [DataMember]
    public Guid RevisionId { get; set; }

    // The policy generation a SecurityOperation applies. Folder jobs no longer set or read it;
    // it stays on the shared model so stored folder-job payloads still deserialize.
    [DataMember]
    public Guid PolicyRevision { get; set; }

    [DataMember]
    public string? ParentPath { get; set; }

    // The library entry path this job's stored paths were written against. When a re-point
    // moves the entry, the job's paths follow it at its next step (WorkerCoordinator.Follow).
    // Null in jobs stored before re-point existed; set at their next step.
    [DataMember]
    public string? EntryPath { get; set; }

    [DataMember]
    public Guid ProbeId { get; set; }

    [DataMember]
    public string ProbeKind { get; set; } = "";

    [DataMember]
    public bool AbsenceVerified { get; set; }

    [DataMember]
    public string? AncestorPath { get; set; }

    [DataMember]
    public string? LibraryRootPath { get; set; }

    [DataMember]
    public Guid LibraryRootId { get; set; }

    [DataMember]
    public bool ExternalSubmitted { get; set; }

    [DataMember]
    public bool ExternalResponseKnown { get; set; }

    [DataMember]
    public string? ErrorCode { get; set; }

    [DataMember]
    public int RetryCount { get; set; }

    [DataMember]
    public DateTime? NextAttemptUtc { get; set; }

    // Set when a new claim takes over a write whose outcome is unknown (an expired lease or an
    // operator retry). The next reads establish what SharePoint did; only then is the write
    // either adopted or cleared so it can be prepared again.
    [DataMember]
    public bool Reprobe { get; set; }

    // SharePoint answered the create with 409: something already uses the folder path. Kept
    // apart from ErrorCode, which a later wait notice overwrites, so the conflict is still
    // recognised when the folder read is retried.
    [DataMember]
    public bool NameConflict { get; set; }
}

[DataContract]
public sealed class DispatcherDocument : StoredDocument
{
    [DataMember]
    public bool HttpOutstanding { get; set; }

    [DataMember]
    public string? OperationKey { get; set; }

    [DataMember]
    public string? RunId { get; set; }

    [DataMember]
    public Guid Token { get; set; }

    [DataMember]
    public DateTime LeaseUntilUtc { get; set; }
}

[DataContract]
public sealed class AttemptDocument : StoredDocument
{
    [DataMember]
    public string OperationKey { get; set; } = "";

    [DataMember]
    public string RunId { get; set; } = "";

    [DataMember]
    public string Event { get; set; } = "";

    [DataMember]
    public DateTime AtUtc { get; set; }
}

[DataContract]
public sealed class WorkerRequest
{
    [DataMember]
    public Guid[] RecordIds { get; set; } = Array.Empty<Guid>();

    [DataMember]
    public string RowVersion { get; set; } = "";

    [DataMember]
    public string Command { get; set; } = "";

    [DataMember]
    public string Key { get; set; } = "";

    [DataMember]
    public Guid TemplateId { get; set; }

    [DataMember]
    public Guid RecordId { get; set; }

    [DataMember]
    public Guid RequestId { get; set; }

    [DataMember]
    public string RunId { get; set; } = "";

    [DataMember]
    public Guid Token { get; set; }

    [DataMember]
    public Guid ProbeId { get; set; }

    [DataMember]
    public string ProbeKind { get; set; } = "";

    [DataMember]
    public string? RetryAfter { get; set; }

    [DataMember]
    public string? ResponseBody { get; set; }

    [DataMember]
    public int HttpStatus { get; set; }

    // The failed flow action's HTTP status, error code and message, sent by the flow's failure
    // branches to FailOutbox, FailUnclaimed and Fail. A flow that sends none of them keeps
    // the earlier behavior and blocks.
    [DataMember]
    public int? StatusCode { get; set; }

    [DataMember]
    public string? ErrorCode { get; set; }

    [DataMember]
    public string? Error { get; set; }

    /// <summary>ListProblems: TemplateRuns, NotCaptured, BlockedRecords, WaitingRecords, BlockedJobs, RetryingJobs or RecentOperations.</summary>
    [DataMember]
    public string List { get; set; } = "";

    /// <summary>ListProblems: the Next of the previous page; null for the first.</summary>
    [DataMember]
    public string? Page { get; set; }

    /// <summary>RerunRecord: the record's table.</summary>
    [DataMember]
    public string Table { get; set; } = "";

    /// <summary>DismissCaptureJob: the failed system job.</summary>
    [DataMember]
    public Guid JobId { get; set; }
}

[DataContract]
public sealed class WorkerResult
{
    [DataMember]
    public int WaitSeconds { get; set; }

    [DataMember]
    public string RowVersion { get; set; } = "";

    [DataMember]
    public BatchDocument? Batch { get; set; }

    [DataMember]
    public RecordInspection? Record { get; set; }

    [DataMember]
    public string? RunId { get; set; }

    [DataMember]
    public DateTime? LeaseUntilUtc { get; set; }

    [DataMember]
    public string Status { get; set; } = "";

    [DataMember]
    public string Key { get; set; } = "";

    [DataMember]
    public Guid Token { get; set; }

    [DataMember]
    public Guid ProbeId { get; set; }

    [DataMember]
    public string ProbeKind { get; set; } = "";

    [DataMember]
    public string? SiteUrl { get; set; }

    [DataMember]
    public HttpIntent? Http { get; set; }

    [DataMember]
    public string[] Keys { get; set; } = Array.Empty<string>();

    [DataMember]
    public string[] Notices { get; set; } = Array.Empty<string>();

    [DataMember]
    public Guid PhysicalId { get; set; }

    [DataMember]
    public Guid LocationId { get; set; }

    /// <summary>A library setup's lookup finding (spec 6.8); null for everything else.</summary>
    [DataMember]
    public LibraryRecovery? Recovery { get; set; }

    /// <summary>A template re-run's state (spec 6.5); null for everything else.</summary>
    [DataMember]
    public TemplateRunState? Run { get; set; }

    /// <summary>Summary: Monitor's counts (spec 6.1); null for everything else.</summary>
    [DataMember]
    public ProblemSummary? Summary { get; set; }

    /// <summary>ListProblems: one page of a Monitor list (spec 6.2).</summary>
    [DataMember]
    public ProblemRow[] Problems { get; set; } = Array.Empty<ProblemRow>();

    /// <summary>ListProblems: the Page that reads the rest of the list; null on the last page.</summary>
    [DataMember]
    public string? Next { get; set; }
}
