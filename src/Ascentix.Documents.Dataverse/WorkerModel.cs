using System;
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

    // Unused: folder jobs no longer compare the library ACL. Kept so stored payloads deserialize.
    [DataMember]
    public string? ApprovedAclHash { get; set; }

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

    [DataMember]
    public bool RecoveryPermitted { get; set; }

    [DataMember]
    public string? TerminationEvidence { get; set; }
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

    [DataMember]
    public string? Evidence { get; set; }

    // The failed flow action's HTTP status, error code and message, sent by the flow's failure
    // branches to FailOutbox, FailUnclaimed and Fail. A flow that sends none of them keeps
    // the earlier behavior and blocks.
    [DataMember]
    public int? StatusCode { get; set; }

    [DataMember]
    public string? ErrorCode { get; set; }

    [DataMember]
    public string? Error { get; set; }
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
}
