using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;

namespace Ascentix.Documents.Dataverse;

/// <summary>What Documents found in SharePoint for a library create whose answer was lost.</summary>
[DataContract]
public sealed class LibraryRecovery
{
    /// <summary>Checking, Found, NotFound or Ambiguous.</summary>
    [DataMember]
    public string State { get; set; } = "Checking";

    [DataMember]
    public string? Reason { get; set; }

    [DataMember]
    public DateTime? CheckedUtc { get; set; }

    [DataMember]
    public RecoveryCandidate[] Candidates { get; set; } = Array.Empty<RecoveryCandidate>();

    /// <summary>What the admin may do with this finding: UseLibrary, UseCandidate, CreateAgain, CheckAgain, Cancel.</summary>
    [DataMember]
    public string[] Choices { get; set; } = Array.Empty<string>();
}

/// <summary>A list the lookup found, with each check it passed or failed.</summary>
[DataContract]
public sealed class RecoveryCandidate
{
    [DataMember]
    public Guid ListId { get; set; }

    [DataMember]
    public string Title { get; set; } = "";

    [DataMember]
    public string Url { get; set; } = "";

    [DataMember]
    public DateTime? CreatedUtc { get; set; }

    [DataMember]
    public bool IsLibrary { get; set; }

    [DataMember]
    public bool TitleMatches { get; set; }

    [DataMember]
    public bool UrlMatches { get; set; }

    [DataMember]
    public bool CreatedAfterRequest { get; set; }

    /// <summary>None; ThisSite (adopted on use); Conflict (another site or removed: use is refused).</summary>
    [DataMember]
    public string CatalogEntry { get; set; } = "None";
}

/// <summary>A list as the lookup reads it, by title or by address.</summary>
[DataContract]
public sealed class ReconcileObservation
{
    [DataMember(Name = "Id")]
    public Guid Id { get; set; }

    [DataMember(Name = "Title")]
    public string Title { get; set; } = "";

    [DataMember(Name = "BaseTemplate")]
    public int BaseTemplate { get; set; }

    /// <summary>SharePoint's ISO 8601 creation time, kept as text: the serializer only reads its own date form.</summary>
    [DataMember(Name = "Created")]
    public string? Created { get; set; }

    [DataMember(Name = "RootFolder")]
    public FolderObservation? Root { get; set; }
}

/// <summary>A Documents catalog row the lookup compares with what SharePoint has.</summary>
public sealed class CatalogEntryRow
{
    public Guid Id { get; set; }

    public Guid SiteId { get; set; }

    public string Name { get; set; } = "";

    public Guid ListId { get; set; }

    public bool Removed { get; set; }
}

/// <summary>
/// The decision of the library lookup (spec 6.8): one candidate that passes every check is Found;
/// no list by title or by address, and no catalog entry by name, is NotFound; anything else is
/// Ambiguous and the admin decides.
/// </summary>
public static class LibraryReconcile
{
    /// <summary>The fields both lookup reads select.</summary>
    public const string Select =
        "$select=Id,Title,BaseTemplate,Created,HasUniqueRoleAssignments,RootFolder/Name,RootFolder/ServerRelativeUrl,RootFolder/UniqueId&$expand=RootFolder";

    public const string CatalogConflict =
        "The Documents catalog already has a different entry for this library.";

    // Characters SharePoint leaves out of the folder name it derives from a list title when the
    // create body has only a Title (NewLibraryBody). Confirmed in TEST (Task 2, Step 1).
    private const string Stripped = "~\"#%&*:<>?/\\{|}";

    /// <summary>The server-relative address SharePoint gives a library created with only this title.</summary>
    /// <param name="webUrl">The site's absolute URL.</param>
    /// <param name="title">The requested library name.</param>
    /// <returns>The site's server-relative path, a slash, then the derived folder name.</returns>
    public static string ExpectedUrl(string webUrl, string title)
    {
        string site = Uri.UnescapeDataString(new Uri(webUrl).AbsolutePath).TrimEnd('/');
        string folder = new string(title.Where(c => Stripped.IndexOf(c) < 0).ToArray())
            .Trim()
            .Trim('.');
        return site + "/" + folder;
    }

    public static LibraryRecovery Checking() =>
        new LibraryRecovery { State = "Checking", Choices = new[] { "Cancel" } };

    /// <summary>A lookup read that could not be made or was refused: the admin may check again or cancel.</summary>
    public static LibraryRecovery ReadFailed(string cause, DateTime now) =>
        new LibraryRecovery
        {
            State = "Ambiguous",
            Reason = "Documents couldn't read SharePoint: " + cause,
            CheckedUtc = now,
            Choices = new[] { "CheckAgain", "Cancel" },
        };

    /// <summary>Decides what the two reads and the catalog show.</summary>
    /// <param name="siteId">The setup's site.</param>
    /// <param name="name">The requested library name.</param>
    /// <param name="expectedUrl">ExpectedUrl for the setup.</param>
    /// <param name="titleHit">The list read by title, or null for 404.</param>
    /// <param name="urlHit">The list read by address, or null for 404.</param>
    /// <param name="requestedUtc">When SharePoint last had no list with this title (AbsentUtc), else when the setup was created.</param>
    /// <param name="catalog">Catalog rows on this site with the name, and rows for either list ID.</param>
    /// <param name="now">The time of the check.</param>
    /// <returns>The finding with its candidates and the choices it allows.</returns>
    public static LibraryRecovery Decide(
        Guid siteId,
        string name,
        string expectedUrl,
        ReconcileObservation? titleHit,
        ReconcileObservation? urlHit,
        DateTime requestedUtc,
        IReadOnlyCollection<CatalogEntryRow> catalog,
        DateTime now
    )
    {
        // SharePoint reports Created to the second; a create in the same second still counts.
        var since = new DateTime(
            requestedUtc.Ticks - requestedUtc.Ticks % TimeSpan.TicksPerSecond,
            DateTimeKind.Utc
        );
        var candidates = new[] { titleHit, urlHit }
            .Where(hit => hit != null)
            .Select(hit => hit!)
            .GroupBy(hit => hit.Id)
            .Select(group => Candidate(group.First(), siteId, name, expectedUrl, since, catalog))
            .ToArray();
        bool otherEntry = catalog.Any(row =>
            !row.Removed
            && row.SiteId == siteId
            && string.Equals(row.Name, name, StringComparison.OrdinalIgnoreCase)
            && candidates.All(c => c.ListId != row.ListId)
        );
        var finding = new LibraryRecovery { CheckedUtc = now, Candidates = candidates };
        if (candidates.Length == 0)
        {
            finding.State = otherEntry ? "Ambiguous" : "NotFound";
            finding.Reason = otherEntry ? CatalogConflict : null;
        }
        else if (candidates.Length == 1 && Matches(candidates[0]) && !otherEntry)
            finding.State = "Found";
        else
        {
            finding.State = "Ambiguous";
            finding.Reason = Why(candidates);
        }
        finding.Choices = Choices(finding);
        return finding;
    }

    /// <summary>The choices a finding allows (spec 3.3 recovery table).</summary>
    public static string[] Choices(LibraryRecovery finding)
    {
        switch (finding.State)
        {
            case "Found":
                return new[] { "UseLibrary", "CheckAgain", "Cancel" };
            case "NotFound":
                return new[] { "CreateAgain", "CheckAgain", "Cancel" };
            case "Ambiguous":
                var choices = new List<string>();
                if (finding.Candidates.Any(Usable))
                    choices.Add("UseCandidate");
                if (finding.Candidates.All(c => !c.TitleMatches))
                    choices.Add("CreateAgain");
                choices.Add("CheckAgain");
                choices.Add("Cancel");
                return choices.ToArray();
            default:
                return new[] { "Cancel" };
        }
    }

    /// <summary>The finding as one sentence, for notices and Problem text.</summary>
    public static string Sentence(string name, LibraryRecovery finding) =>
        finding.State switch
        {
            "Checking" => "Checking SharePoint for " + name + ".",
            "Found" => "SharePoint has a library named " + name + " that matches this request.",
            "NotFound" => "SharePoint has no library named "
                + name
                + ". The creation didn't happen.",
            _ => finding.Reason ?? "Documents found more than one possible library.",
        };

    internal static bool Usable(RecoveryCandidate candidate) =>
        candidate.IsLibrary && candidate.CatalogEntry != "Conflict";

    private static bool Matches(RecoveryCandidate c) =>
        c.IsLibrary
        && c.TitleMatches
        && c.UrlMatches
        && c.CreatedAfterRequest
        && c.CatalogEntry != "Conflict";

    private static string Why(RecoveryCandidate[] candidates)
    {
        if (candidates.Length > 1)
            return candidates.Length.ToString(CultureInfo.InvariantCulture)
                + " libraries could be this one.";
        var c = candidates[0];
        return !c.IsLibrary ? "A list has this name but is not a document library."
            : !c.TitleMatches ? "A library has this address but a different name."
            : !c.UrlMatches ? "A library has this name but a different address."
            : !c.CreatedAfterRequest ? "A library has this name but was there before the request."
            : CatalogConflict;
    }

    private static RecoveryCandidate Candidate(
        ReconcileObservation hit,
        Guid siteId,
        string name,
        string expectedUrl,
        DateTime since,
        IReadOnlyCollection<CatalogEntryRow> catalog
    )
    {
        DateTime? created = DateTime.TryParse(
            hit.Created,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
            out var parsed
        )
            ? parsed
            : (DateTime?)null;
        var rows = catalog.Where(row => row.ListId == hit.Id).ToArray();
        // The same condition Register applies (LibraryProvisioning.Register): only this site's
        // active entry for this list, with the identity approval gives it, is adopted.
        string entry =
            rows.Length == 0 ? "None"
            : rows.Length == 1
            && rows[0].SiteId == siteId
            && !rows[0].Removed
            && rows[0].Id == SiteIdentity.LibraryId(siteId, hit.Id)
                ? "ThisSite"
            : "Conflict";
        return new RecoveryCandidate
        {
            ListId = hit.Id,
            Title = hit.Title,
            Url = hit.Root?.Path ?? "",
            CreatedUtc = created,
            IsLibrary = hit.BaseTemplate == 101,
            TitleMatches = string.Equals(hit.Title, name, StringComparison.Ordinal),
            UrlMatches = string.Equals(
                hit.Root?.Path,
                expectedUrl,
                StringComparison.OrdinalIgnoreCase
            ),
            CreatedAfterRequest = created >= since,
            CatalogEntry = entry,
        };
    }
}
