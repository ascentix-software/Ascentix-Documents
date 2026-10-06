using System;
using System.Linq;

namespace Ascentix.Documents.Dataverse;

/// <summary>
/// What a paused runtime refuses. Pause stops new work; it never refuses to record the outcome
/// of work already under way, above all the answer to a SharePoint write the product already
/// permitted. A refused answer is lost, and a library create whose answer is lost needs manual
/// recovery.
/// </summary>
public static class WorkerPause
{
    // Each records an outcome: CreateResponse and Observe a SharePoint answer, Complete a
    // verified result, Fail, FailUnclaimed and FailOutbox a failed step. Renew only keeps the
    // run's claim. None calls SharePoint or starts a job.
    private static readonly string[] Records =
    {
        "CreateResponse",
        "Observe",
        "Complete",
        "Fail",
        "FailUnclaimed",
        "FailOutbox",
        "Renew",
    };

    /// <summary>
    /// Whether a paused runtime refuses the worker command. Every other command starts new work
    /// and is refused: a scan or plan that makes jobs (PurgeHistory, RefreshSecurity, ListOutbox,
    /// Plan, ListOperations, Queue), a claim (Claim), a prepared write (PrepareCreate) or a
    /// SharePoint request (BeginHttp, read or write: a paused runtime calls SharePoint no more).
    /// </summary>
    public static bool Refuses(string? command) =>
        command == null || !Records.Contains(command, StringComparer.Ordinal);
}
