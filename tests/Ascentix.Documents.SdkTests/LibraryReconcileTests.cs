using System;
using Ascentix.Documents.Dataverse;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class LibraryReconcileTests
{
    private static readonly Guid Site = Guid.NewGuid();
    private static readonly DateTime Requested = new DateTime(
        2026,
        10,
        6,
        17,
        40,
        0,
        400,
        DateTimeKind.Utc
    );
    private const string Expected = "/sites/test/Project documents";

    private static ReconcileObservation List(
        Guid? id = null,
        string title = "Project documents",
        string path = Expected,
        int template = 101,
        string created = "2026-10-06T17:41:02Z"
    ) =>
        new ReconcileObservation
        {
            Id = id ?? Guid.NewGuid(),
            Title = title,
            BaseTemplate = template,
            Created = created,
            Root = new FolderObservation { Id = Guid.NewGuid(), Path = path },
        };

    private static LibraryRecovery Decide(
        ReconcileObservation? title,
        ReconcileObservation? url,
        params CatalogEntryRow[] catalog
    ) =>
        LibraryReconcile.Decide(
            Site,
            "Project documents",
            Expected,
            title,
            url,
            Requested,
            catalog,
            Requested.AddMinutes(6)
        );

    [Theory]
    [InlineData("https://example.sharepoint.com/sites/test", "Documents", "/sites/test/Documents")]
    [InlineData(
        "https://example.sharepoint.com/sites/test/",
        "Project documents",
        "/sites/test/Project documents"
    )]
    [InlineData(
        "https://example.sharepoint.com/sites/my%20site",
        "R&D: plans #1",
        "/sites/my site/RD plans 1"
    )]
    [InlineData(
        "https://example.sharepoint.com/sites/test",
        "Bob's files",
        "/sites/test/Bob's files"
    )]
    [InlineData(
        "https://example.sharepoint.com/sites/test",
        "Ünïcode Åbc",
        "/sites/test/Ünïcode Åbc"
    )]
    [InlineData("https://example.sharepoint.com/sites/test", "50% done", "/sites/test/50 done")]
    public void ExpectedUrlFollowsSharePoint(string web, string title, string expected) =>
        Assert.Equal(expected, LibraryReconcile.ExpectedUrl(web, title));

    [Fact]
    public void OneListThatPassesEveryCheckIsFound()
    {
        var made = List();
        var finding = Decide(made, made);
        Assert.Equal("Found", finding.State);
        Assert.Equal(new[] { "UseLibrary", "CheckAgain", "Cancel" }, finding.Choices);
        Assert.Equal("None", Assert.Single(finding.Candidates).CatalogEntry);
    }

    [Fact]
    public void ACreateInTheSameSecondAsTheMissingReadStillMatches()
    {
        var made = List(created: "2026-10-06T17:40:00Z");
        Assert.Equal("Found", Decide(made, made).State);
    }

    [Fact]
    public void NoListAndNoCatalogEntryIsNotFound()
    {
        var finding = Decide(null, null);
        Assert.Equal("NotFound", finding.State);
        Assert.Equal(new[] { "CreateAgain", "CheckAgain", "Cancel" }, finding.Choices);
    }

    [Fact]
    public void ACatalogEntryWithTheNameAndNoListIsAmbiguousButMayCreate()
    {
        var finding = Decide(
            null,
            null,
            new CatalogEntryRow
            {
                Id = Guid.NewGuid(),
                SiteId = Site,
                Name = "project documents",
                ListId = Guid.NewGuid(),
            }
        );
        Assert.Equal("Ambiguous", finding.State);
        Assert.Equal(LibraryReconcile.CatalogConflict, finding.Reason);
        Assert.Equal(new[] { "CreateAgain", "CheckAgain", "Cancel" }, finding.Choices);
    }

    [Fact]
    public void TwoDifferentListsAreAmbiguous()
    {
        var finding = Decide(List(), List(title: "Other", path: Expected));
        Assert.Equal("2 libraries could be this one.", finding.Reason);
        Assert.Equal(new[] { "UseCandidate", "CheckAgain", "Cancel" }, finding.Choices);
    }

    [Theory]
    [InlineData("path", "A library has this name but a different address.")]
    [InlineData("template", "A list has this name but is not a document library.")]
    [InlineData("before", "A library has this name but was there before the request.")]
    [InlineData("title", "A library has this address but a different name.")]
    public void ASingleListThatFailsACheckIsAmbiguousWithItsReason(string failure, string reason)
    {
        var list = failure switch
        {
            "path" => List(path: "/sites/test/Project documents1"),
            "template" => List(template: 100),
            "before" => List(created: "2026-10-06T17:39:59Z"),
            _ => List(title: "Renamed"),
        };
        var finding = failure == "title" ? Decide(null, list) : Decide(list, null);
        Assert.Equal("Ambiguous", finding.State);
        Assert.Equal(reason, finding.Reason);
        // Create it again only while no list has the requested title.
        Assert.Equal(failure == "title", Array.IndexOf(finding.Choices, "CreateAgain") >= 0);
        // A list that is not a document library cannot be used.
        Assert.Equal(failure != "template", Array.IndexOf(finding.Choices, "UseCandidate") >= 0);
    }

    [Theory]
    [InlineData(false, true, "Conflict")]
    [InlineData(true, false, "Conflict")]
    [InlineData(true, true, "ThisSite")]
    public void TheCatalogDecidesWhetherAListIsAdoptedOrInConflict(
        bool sameSite,
        bool active,
        string entry
    )
    {
        var made = List();
        var owner = sameSite ? Site : Guid.NewGuid();
        var finding = Decide(
            made,
            made,
            new CatalogEntryRow
            {
                Id = SiteIdentity.LibraryId(owner, made.Id),
                SiteId = owner,
                Name = "Project documents",
                ListId = made.Id,
                Removed = !active,
            }
        );
        Assert.Equal(entry, Assert.Single(finding.Candidates).CatalogEntry);
        Assert.Equal(entry == "ThisSite" ? "Found" : "Ambiguous", finding.State);
        if (entry == "Conflict")
        {
            Assert.Equal(LibraryReconcile.CatalogConflict, finding.Reason);
            Assert.DoesNotContain("UseCandidate", finding.Choices);
        }
    }
}
