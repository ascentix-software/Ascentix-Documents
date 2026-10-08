using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Ascentix.Documents.Dataverse;

/// <summary>A plain Problem and Fix for each stop code and notice (spec 6.2, "Readable problems").</summary>
public static class ProblemText
{
    private static readonly Dictionary<string, (string Problem, string? Fix)> Known = new(
        StringComparer.Ordinal
    )
    {
        ["RequestUrlTooLong"] = (
            "The folder path is too long for SharePoint's address limit.",
            "Shorten the record's value or the template's folder names, then re-run the record."
        ),
        ["CreateConflictNoFolder"] = (
            "A file already uses this folder's name in SharePoint.",
            "Rename or move that file in SharePoint, then retry."
        ),
        ["FileAtExpectedFolderPath"] = (
            "A file already uses this folder's name in SharePoint.",
            "Rename or move that file in SharePoint, then retry."
        ),
        ["CreateNameConflict"] = (
            "Another folder already has this name.",
            "Change the record or the folder name, then re-run the record."
        ),
        ["FolderChangedDuringRequest"] = (
            "The folder changed in SharePoint while Documents was working on it.",
            "Retry."
        ),
        ["ObservationMismatch"] = (
            "The folder changed in SharePoint while Documents was working on it.",
            "Retry."
        ),
        ["EstablishedFolderMissing"] = (
            "A folder Documents created earlier is gone from SharePoint.",
            "Retry to create it again, or cancel the job."
        ),
        ["CreateRejected"] = (
            "SharePoint refused to create the folder.",
            "Check the library's permissions for the connection, then retry."
        ),
        ["ReadAfterUnknownWrite"] = (
            "Documents couldn't confirm whether its last change reached SharePoint.",
            "Retry; Documents checks SharePoint first."
        ),
        ["WorkerFailed"] = (
            "The flow run stopped with an error.",
            "Check the flow run in Power Automate, then retry."
        ),
        ["SecurityWorkerFailed"] = (
            "The flow run stopped with an error.",
            "Check the flow run in Power Automate, then retry."
        ),
        ["AmbiguousSecurityWrite"] = (
            "Documents couldn't confirm a permission change in SharePoint.",
            "Retry; Documents reads SharePoint first."
        ),
        ["UnknownSecurityWrite"] = (
            "Documents couldn't confirm a permission change in SharePoint.",
            "Retry; Documents reads SharePoint first."
        ),
        ["UnknownGroupWrite"] = (
            "Documents couldn't confirm a permission change in SharePoint.",
            "Retry; Documents reads SharePoint first."
        ),
        ["SecurityWriteRejected"] = (
            "SharePoint refused a permission change.",
            "Check that the connection can manage permissions on the library, then retry."
        ),
        ["DestinationRemoved"] = (
            "The library was removed from Documents.",
            "Cancel the job, or add the library again."
        ),
        ["TemplateUnavailable"] = ("The template is off or outside its active dates.", null),
        ["TableNotEnabled"] = (
            "The table is no longer enabled in Documents.",
            "Enable the table, then re-run the record."
        ),
    };

    /// <summary>Every code with its own text; a test checks each Block and Stop code literal is here.</summary>
    public static IReadOnlyCollection<string> Codes => Known.Keys;

    private static readonly Regex Transient = new(
        @"^Waiting to retry after a temporary error \((?<cause>[^)]*)\); attempt \d+\.$"
    );
    private static readonly Regex PathTooLong = new(
        @"^Folder '(?<folder>[^']+)' is waiting for a shorter path: its path would be (?<length>[\d,]+) characters, and SharePoint allows (?<limit>[\d,]+) including file names\.(?<rest>.*)$",
        RegexOptions.Singleline
    );
    private static readonly Regex AddressTooLong = new(
        @"^Folder '(?<folder>[^']+)' is waiting for a shorter path: written into a SharePoint address it would be (?<length>[\d,]+) characters, and the HTTP connector accepts (?<limit>[\d,]+)\.(?<rest>.*)$",
        RegexOptions.Singleline
    );
    private static readonly Regex ShortNames = new(
        @"^Folder '(?<folder>[^']+)': This folder's path is (?<length>[\d,]+) characters; SharePoint allows 400 including file names, so files inside need short names\.$"
    );

    /// <summary>The Problem and Fix for a stopped item.</summary>
    /// <param name="code">The operation's ErrorCode: a code, a sentence, or null.</param>
    /// <param name="status">The row's status.</param>
    /// <param name="notice">The row's first notice, for outbox rows and waits.</param>
    public static (string Problem, string? Fix) Describe(
        string? code,
        string status,
        string? notice
    )
    {
        if (code != null && Known.TryGetValue(code, out var known))
            return known;
        string? text = notice ?? code;
        if (text == null)
            return status == "Blocked"
                ? ("Stopped without a reason recorded.", "Retry, or cancel the job.")
                : (status, null);
        var transient = Transient.Match(text);
        if (transient.Success)
        {
            string cause = transient.Groups["cause"].Value;
            return (
                cause.StartsWith("HTTP 429", StringComparison.Ordinal)
                    ? "SharePoint is busy or limiting requests (HTTP 429)."
                    : "A temporary error stopped the last attempt (" + cause + ").",
                null
            );
        }
        if (Regex.IsMatch(text, "^[A-Z][a-z]+(?:[A-Z][a-z]+)+$"))
            return ("Stopped with code " + text + ".", "Retry, or cancel the job.");
        return (Shorten(text), null);
    }

    /// <summary>
    /// The Problem and Fix for a stopped template re-run. Its planning-failed notice is the one
    /// every outbox row gets, so only a run row reads it as the run's (spec 6.2 table).
    /// </summary>
    /// <param name="notice">The run row's first notice.</param>
    public static (string Problem, string? Fix) DescribeRun(string? notice) =>
        notice == WorkerCoordinator.PlanningFailedNotice
            ? (
                "The re-run stopped while planning a page of records.",
                "Retry. Records already planned are kept."
            )
            : Describe(null, "Blocked", notice);

    /// <summary>
    /// Shortens the long waiting notices stored before 0.1.0.4 changed them (spec 6.2, "Shorter
    /// waiting notices"). Why the limits exist is in operations.md › Folder names.
    /// </summary>
    public static string Shorten(string notice)
    {
        string text = notice;
        foreach (var pattern in new[] { PathTooLong, AddressTooLong })
        {
            var match = pattern.Match(text);
            if (match.Success)
                text =
                    "Folder '"
                    + match.Groups["folder"].Value
                    + "' needs a shorter path: "
                    + Number(match.Groups["length"].Value)
                    + " characters; the limit is "
                    + Number(match.Groups["limit"].Value)
                    + "."
                    + match.Groups["rest"].Value;
        }
        var names = ShortNames.Match(text);
        if (names.Success)
            text =
                "Folder '"
                + names.Groups["folder"].Value
                + "': files inside need short names; its path is "
                + Number(names.Groups["length"].Value)
                + " of 400 characters.";
        // Also turns "replan the records" into "re-run the records".
        return text.Replace("replan the record", "re-run the record");
    }

    private static string Number(string digits) =>
        long.Parse(digits.Replace(",", ""), CultureInfo.InvariantCulture)
            .ToString("N0", CultureInfo.InvariantCulture);
}
