using System;
using System.Globalization;
using System.Linq;

namespace Ascentix.Documents.Dataverse;

/// <summary>
/// Decides whether a failed worker call reported by a flow is temporary. A temporary failure
/// waits and retries with no attempt cap; anything else is a genuine failure and blocks.
/// </summary>
public static class TransientFailure
{
    // Dataverse error codes that mean "try again later". Sources:
    // - Service protection limits (Web API 429), learn.microsoft.com/power-apps/developer/data-platform/api-limits:
    //   0x80072322 (-2147015902) number of requests, 0x80072321 (-2147015903) combined execution
    //   time, 0x80072326 (-2147015898) concurrent requests.
    // - Optimistic concurrency, learn.microsoft.com/power-apps/developer/data-platform/optimistic-concurrency:
    //   0x80060882 (-2147088254) ConcurrencyVersionMismatch, raised by DocumentStore.Save's
    //   IfRowVersionMatches update when another run saved the row first.
    // - SQL, learn.microsoft.com/troubleshoot/power-platform/dataverse/bulk-deleting-data/bulk-operation-errors:
    //   0x80044151 (-2147204783) SQL timeout expired, 0x80040255 (-2147220907) transaction already
    //   rolled back by a swallowed SQL deadlock.
    private static readonly int[] TransientCodes =
    {
        unchecked((int)0x80072322),
        unchecked((int)0x80072321),
        unchecked((int)0x80072326),
        unchecked((int)0x80060882),
        unchecked((int)0x80044151),
        unchecked((int)0x80040255),
    };

    // 0x80044150 (-2147204784) is the generic SQL error. The same bulk-operation-errors page shows
    // it carrying "Sql Number: 1205" for a deadlock; other SQL numbers under it are not temporary.
    private const int GenericSqlError = unchecked((int)0x80044150);

    // Codes raised by plug-in code itself, which the platform can report with HTTP 500. They are
    // genuine failures (a refused configuration or a product check), never a reason to wait.
    // learn.microsoft.com/troubleshoot/power-platform/dataverse/plug-in-execution/dataverse-plug-ins-errors:
    // 0x80040265 (-2147220891) ISV code aborted the operation, 0x80040224 (-2147220956) unexpected
    // error from ISV code (an unhandled product exception), 0x8004418d (-2147204723) sandbox worker
    // crashed (reported with ExceptionRetriable false).
    private static readonly int[] GenuineCodes =
    {
        unchecked((int)0x80040265),
        unchecked((int)0x80040224),
        unchecked((int)0x8004418d),
    };

    /// <summary>
    /// Classifies one failed call.
    /// </summary>
    /// <param name="statusCode">The failed action's HTTP status; 0 or null means no response arrived.</param>
    /// <param name="errorCode">The Dataverse error code, hexadecimal ("0x80060882") or decimal.</param>
    /// <param name="error">The error message, used only to recognise a SQL deadlock.</param>
    /// <returns>True when the call should wait and retry; false when it should block.</returns>
    public static bool Is(int? statusCode, string? errorCode, string? error)
    {
        // The flow's drive branch reports DriveFailed when no Dataverse action failed (an
        // expression error, a failed variable step or a loop limit): not temporary.
        if (string.Equals(errorCode?.Trim(), DriveFailed, StringComparison.Ordinal))
            return false;
        var code = Parse(errorCode);
        if (code.HasValue)
        {
            if (TransientCodes.Contains(code.Value))
                return true;
            if (code.Value == GenericSqlError)
                return error != null
                    && (
                        error.IndexOf("1205", StringComparison.Ordinal) >= 0
                        || error.IndexOf("deadlock", StringComparison.OrdinalIgnoreCase) >= 0
                    );
            if (GenuineCodes.Contains(code.Value))
                return false;
        }
        // Some permanent faults also look temporary here: a 5xx without a known code, or a
        // ConcurrencyVersionMismatch that a product bug repeats on every attempt. They keep
        // backing off rather than blocking, but never silently: the notice names the cause and
        // counts the attempts, and an operator can Cancel the work at any time.
        switch (statusCode ?? 0)
        {
            case 0: // No response: connector timeout or cancellation.
            case 408:
            case 429:
            case 500:
            case 502:
            case 503:
            case 504:
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// True when the flow reported a failure cause and it is temporary. A request without any
    /// cause comes from a flow version that predates classification and keeps blocking.
    /// </summary>
    public static bool Is(WorkerRequest request) =>
        Reported(request) && Is(request.StatusCode, request.ErrorCode, request.Error);

    public static bool Reported(WorkerRequest request) =>
        request.StatusCode != null || request.ErrorCode != null || request.Error != null;

    /// <summary>The visible notice for a row that waits after a temporary failure.</summary>
    public static string Notice(int? statusCode, string? errorCode, int attempt) =>
        "Waiting to retry after a temporary error ("
        + Describe(statusCode, errorCode)
        + "); attempt "
        + attempt.ToString(CultureInfo.InvariantCulture)
        + ".";

    public static string Notice(WorkerRequest request, int attempt) =>
        Notice(request.StatusCode, request.ErrorCode, attempt);

    public const string DriveFailed = "DriveFailed";

    /// <summary>The failed call's status, code and a bounded message, for a visible notice.</summary>
    public static string Cause(int? statusCode, string? errorCode, string? error)
    {
        var cause = Describe(statusCode, errorCode);
        if (string.IsNullOrWhiteSpace(error))
            return cause;
        var clean = new string(error!.Trim().Where(c => !char.IsControl(c)).Take(200).ToArray());
        return cause + ": " + clean;
    }

    private static string Describe(int? statusCode, string? errorCode)
    {
        string status =
            (statusCode ?? 0) > 0
                ? "HTTP " + statusCode!.Value.ToString(CultureInfo.InvariantCulture)
                : "no response";
        if (string.IsNullOrWhiteSpace(errorCode))
            return status;
        // The code comes from a flow expression; keep the notice short and printable.
        var clean = new string(errorCode!.Trim().Where(c => !char.IsControl(c)).Take(40).ToArray());
        return status + ", " + clean;
    }

    private static int? Parse(string? errorCode)
    {
        if (string.IsNullOrWhiteSpace(errorCode))
            return null;
        var text = errorCode!.Trim();
        if (text.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            return uint.TryParse(
                text.Substring(2),
                NumberStyles.AllowHexSpecifier,
                CultureInfo.InvariantCulture,
                out var hex
            )
                ? unchecked((int)hex)
                : (int?)null;
        return int.TryParse(
            text,
            NumberStyles.AllowLeadingSign,
            CultureInfo.InvariantCulture,
            out var value
        )
            ? value
            : (int?)null;
    }
}
