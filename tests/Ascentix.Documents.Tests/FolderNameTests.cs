using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Domain;
using Xunit;

namespace Ascentix.Documents.Tests;

public class FolderNameTests
{
    [Theory]
    [InlineData("Smith & Sons: Holdings", "Smith & Sons- Holdings")]
    [InlineData("A/B Testing Co", "A-B Testing Co")]
    [InlineData("Acme Inc.", "Acme Inc")]
    [InlineData("Acme ", "Acme")]
    [InlineData("  Acme", "Acme")]
    [InlineData("Acme . . ", "Acme")]
    [InlineData("a\u0001b", "a-b")]
    [InlineData("Acme\t\r\n", "Acme")]
    [InlineData("a\"b*c:d<e>f?g/h\\i|j", "a-b-c-d-e-f-g-h-i-j")]
    [InlineData("~$x", "-$x")]
    [InlineData("a~$x", "a~$x")]
    [InlineData("a_vti_b", "a-vti-b")]
    [InlineData("A_VTI_B", "A-VTI-B")]
    [InlineData("CON", "CON_")]
    [InlineData("con.txt", "con_.txt")]
    [InlineData("COM1", "COM1_")]
    [InlineData("LPT9.log", "LPT9_.log")]
    [InlineData("CONSOLE", "CONSOLE")]
    [InlineData("desktop.ini", "desktop.ini_")]
    [InlineData(".lock", ".lock_")]
    [InlineData("゛abc", "-abc")]
    [InlineData("ဧabc", "-abc")]
    [InlineData("Café Ñandú 東京 Zürich", "Café Ñandú 東京 Zürich")]
    [InlineData("-", "-")]
    [InlineData(".hidden", ".hidden")]
    public void CleanMakesRecordDataAValidFolderName(string raw, string expected)
    {
        Assert.Equal(expected, FolderNames.Clean(raw, true));
        FolderNames.Validate(expected);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(".")]
    [InlineData("..")]
    [InlineData("...")]
    [InlineData("\"*:<>?/\\|")]
    [InlineData(" : / ")]
    [InlineData("\u0001\u0002")]
    public void CleanReturnsNullWhenNothingUsableRemains(string raw) =>
        Assert.Null(FolderNames.Clean(raw, true));

    [Fact]
    public void CleanTruncatesTo255AndRetrimsTheEnd()
    {
        Assert.Equal(new string('a', 255), FolderNames.Clean(new string('a', 300), true));
        Assert.Equal(
            new string('a', 253),
            FolderNames.Clean(new string('a', 253) + " ." + new string('b', 40), true)
        );
        Assert.Equal("Forms_", FolderNames.Clean("Forms" + new string(' ', 300) + "x", true));
        var device = FolderNames.Clean("con." + new string('a', 251), true)!;
        Assert.Equal(255, device.Length);
        Assert.StartsWith("con_.", device, StringComparison.Ordinal);
        // A surrogate pair is never split.
        var emoji = FolderNames.Clean(new string('a', 254) + "\U0001F600", true)!;
        Assert.Equal(new string('a', 254), emoji);
    }

    [Fact]
    public void FormsIsReservedOnlyAtTheLibraryRoot()
    {
        Assert.Equal("Forms_", FolderNames.Clean("Forms", true));
        Assert.Equal("forms_", FolderNames.Clean("forms", true));
        Assert.Equal("Forms", FolderNames.Clean("Forms", false));
        Assert.Throws<EvaluationBlockedException>(() => FolderNames.Validate("Forms"));
        Assert.Throws<EvaluationBlockedException>(() => FolderNames.Validate("Forms", true));
        FolderNames.Validate("Forms", false);
    }

    [Fact]
    public void CleanedNamesAlwaysPassValidateAndAreStable()
    {
        var pieces = new[]
        {
            "a",
            "Z",
            "é",
            "東",
            "\U0001F600",
            " ",
            "  ",
            ".",
            "..",
            "\"",
            "*",
            ":",
            "<",
            ">",
            "?",
            "/",
            "\\",
            "|",
            "\t",
            "\u0000",
            "\u001F",
            "\u007F",
            ((char)0x00A0).ToString(),
            ((char)0x3000).ToString(),
            "~$",
            "~",
            "$",
            "_vti_",
            "_VtI_",
            "_",
            "CON",
            "nul",
            "COM0",
            "lpt5",
            "AUX.",
            "desktop.ini",
            "Forms",
            ".lock",
            "゛",
            "ဧ",
            "&",
            "#",
            "%",
            "-",
        };
        var random = new Random(20261005);
        for (int i = 0; i < 20000; i++)
        {
            var builder = new StringBuilder();
            int count = random.Next(0, 8);
            for (int j = 0; j < count; j++)
                builder.Append(pieces[random.Next(pieces.Length)]);
            if (random.Next(50) == 0)
                builder.Append(new string('x', random.Next(240, 320)));
            string raw = builder.ToString();
            foreach (var atRoot in new[] { true, false })
            {
                var clean = FolderNames.Clean(raw, atRoot);
                if (clean == null)
                    continue;
                var error = Record.Exception(() => FolderNames.Validate(clean, atRoot));
                Assert.True(
                    error == null,
                    "Validate rejected '" + clean + "' cleaned from '" + raw + "'."
                );
                Assert.Equal(clean, FolderNames.Clean(clean, atRoot));
            }
        }
    }

    private static Snapshot Names(string? root, string? customer) =>
        new(
            new Dictionary<string, Value>
            {
                ["root.flag"] = Value.Boolean(true),
                ["root.name"] = Value.Text(root),
                ["customer.name"] = Value.Text(customer),
            }
        );

    [Fact]
    public void PlannerCleansRecordDataAndRecordsANotice()
    {
        var template = AcceptanceTests.Template();
        var plan = FolderPlanner.Plan(
            template,
            Guid.NewGuid(),
            Names("Holdings.", "Smith & Sons: A/B")
        );
        Assert.Equal(4, plan.Count);
        Assert.All(
            plan.Where(n => n.IsRoot),
            n => Assert.Equal("Smith & Sons- A-B-Holdings", n.Name)
        );
        Assert.Contains(
            plan,
            n => n.Node == "child" && n.RelativePath == "Smith & Sons- A-B-Holdings/Included"
        );
        Assert.Equal(
            new[]
            {
                "Folder name 'Smith & Sons: A/B-Holdings.' was adjusted to 'Smith & Sons- A-B-Holdings' for SharePoint.",
            },
            plan.Notices
        );
        Assert.Empty(FolderPlanner.Plan(template, Guid.NewGuid(), Names("Example", "C")).Notices);
    }

    [Fact]
    public void PlannerReservesFormsOnlyAtTheSectionRoot()
    {
        var template = AcceptanceTests.Template();
        foreach (var section in template.Destinations)
        {
            section.Nodes[0].Name = "Forms";
            section.Nodes[1].Name = "Forms";
        }
        var plan = FolderPlanner.Plan(template, Guid.NewGuid(), Names("x", "y"));
        Assert.All(plan.Where(n => n.IsRoot), n => Assert.Equal("Forms_", n.Name));
        Assert.All(
            plan.Where(n => n.Node == "child"),
            n => Assert.Equal("Forms_/Forms", n.RelativePath)
        );
        Assert.Equal(
            new[] { "Folder name 'Forms' was adjusted to 'Forms_' for SharePoint." },
            plan.Notices
        );
    }

    [Fact]
    public void BlankNamingValueAtARootWaitsWhileOtherSectionsArePlanned()
    {
        var template = AcceptanceTests.Template();
        template.Destinations[1].Nodes[0].Name = "Static";
        var plan = FolderPlanner.Plan(template, Guid.NewGuid(), Names("Example", null));
        Assert.DoesNotContain(plan, n => n.Section == "general");
        Assert.Equal(2, plan.Count(n => n.Section == "sensitive"));
        Assert.Equal(
            new[] { "Folder 'general/root' is waiting for 'customer.name' to have a value." },
            plan.Notices
        );
    }

    [Fact]
    public void EachSkippedFolderIsReportedAsAWaitWithItsFieldAndNotice()
    {
        var template = AcceptanceTests.Template();
        template.Destinations[1].Nodes[0].Name = "Static";
        var blank = FolderPlanner.Plan(template, Guid.NewGuid(), Names("Example", null));
        var wait = Assert.Single(blank.Waits);
        Assert.Equal(("general", "root"), (wait.Section, wait.Node));
        Assert.Equal("customer.name", wait.Field?.ToString());
        Assert.Equal(Assert.Single(blank.Notices), wait.Notice);
        // A name SharePoint cannot take, or one a sibling took, waits for a different name.
        template.Destinations[0].Nodes[0].Name = "{customer.name}";
        var unusable = FolderPlanner.Plan(template, Guid.NewGuid(), Names("Example", "..."));
        Assert.Null(Assert.Single(unusable.Waits).Field);
        Assert.Contains(Assert.Single(unusable.Waits).Notice, unusable.Notices);
        // Nothing waits once every name has a value of its own.
        Assert.Empty(FolderPlanner.Plan(template, Guid.NewGuid(), Names("Example", "C")).Waits);
    }

    [Fact]
    public void PartlyBlankNamesArePlannedAndWhollyBlankNamesWait()
    {
        var template = AcceptanceTests.Template();
        template.Destinations[1].Nodes[0].Name = "Static";
        // A blank token inside a composite name is kept, as before.
        foreach (var customer in new[] { "", " " })
        {
            var plan = FolderPlanner.Plan(template, Guid.NewGuid(), Names("Example", customer));
            Assert.All(
                plan.Where(n => n.Section == "general" && n.IsRoot),
                n => Assert.Equal("-Example", n.Name)
            );
            Assert.Equal(4, plan.Count);
        }
        template.Destinations[0].Nodes[0].Name = "X ({customer.name})";
        Assert.Contains(
            FolderPlanner.Plan(template, Guid.NewGuid(), Names("Example", " ")),
            n => n.IsRoot && n.Name == "X ( )"
        );
        // A name that is blank as a whole waits for its field.
        template.Destinations[0].Nodes[0].Name = "{customer.name}";
        foreach (var customer in new[] { "", "   " })
        {
            var plan = FolderPlanner.Plan(template, Guid.NewGuid(), Names("Example", customer));
            Assert.DoesNotContain(plan, n => n.Section == "general");
            Assert.Equal(
                new[] { "Folder 'general/root' is waiting for 'customer.name' to have a value." },
                plan.Notices
            );
        }
    }

    [Fact]
    public void ADuplicateSiblingIsPlannedOnceTheNamesDiffer()
    {
        var template = AcceptanceTests.Template();
        foreach (var section in template.Destinations)
        {
            section.Nodes[1].Name = "{root.name}";
            section.Nodes.Add(
                new FolderNode
                {
                    Key = "other",
                    ParentKey = "root",
                    Name = "{customer.name}",
                    Order = 1,
                }
            );
            section.Nodes.Add(
                new FolderNode
                {
                    Key = "deep",
                    ParentKey = "other",
                    Name = "Deep",
                }
            );
        }
        var record = Guid.NewGuid();
        var waiting = FolderPlanner.Plan(template, record, Names("A/B", "A:B"));
        Assert.Equal(4, waiting.Count);
        Assert.DoesNotContain(waiting, n => n.Node == "other" || n.Node == "deep");
        Assert.Contains(
            "Folder 'general/other' has the same name 'A-B' as 'general/child'; it waits until the names differ.",
            waiting.Notices
        );
        Assert.Contains(
            "Folder 'sensitive/other' has the same name 'A-B' as 'sensitive/child'; it waits until the names differ.",
            waiting.Notices
        );
        var replanned = FolderPlanner.Plan(template, record, Names("A/B", "C"));
        Assert.Equal(8, replanned.Count);
        Assert.Contains(replanned, n => n.RelativePath == "C-A-B/C/Deep");
        Assert.DoesNotContain(replanned.Notices, n => n.Contains("same name"));
    }

    [Theory]
    [InlineData(".lock")]
    [InlineData(".LOCK")]
    [InlineData("゛abc")]
    [InlineData("ဧabc")]
    public void ValidateRejectsTheOtherMicrosoftReservedNames(string name)
    {
        Assert.Throws<EvaluationBlockedException>(() => FolderNames.Validate(name));
        Assert.Throws<EvaluationBlockedException>(() => FolderNames.Validate(name, false));
        FolderNames.Validate("a" + name);
    }

    [Fact]
    public void BlankNamingValueSkipsOnlyThatFolderAndItsChildrenUntilFilledIn()
    {
        var template = AcceptanceTests.Template();
        foreach (var section in template.Destinations)
        {
            section.Nodes[0].Name = "Fixed {root.name}";
            section.Nodes[1].Name = "{customer.name}";
            section.Nodes.Add(
                new FolderNode
                {
                    Key = "deep",
                    ParentKey = "child",
                    Name = "Deep",
                }
            );
        }
        var record = Guid.NewGuid();
        var waiting = FolderPlanner.Plan(template, record, Names("Example", null));
        Assert.Equal(2, waiting.Count);
        Assert.All(waiting, n => Assert.Equal("root", n.Node));
        Assert.Equal(
            new[]
            {
                "Folder 'general/child' is waiting for 'customer.name' to have a value.",
                "Folder 'sensitive/child' is waiting for 'customer.name' to have a value.",
            },
            waiting.Notices
        );
        var filled = FolderPlanner.Plan(template, record, Names("Example", "Acme"));
        Assert.Equal(6, filled.Count);
        Assert.Contains(filled, n => n.RelativePath == "Fixed Example/Acme/Deep");
        Assert.Empty(filled.Notices);
        Assert.Equal(
            waiting.Select(n => n.BindingKey),
            filled.Where(n => n.Node == "root").Select(n => n.BindingKey)
        );
    }

    // A template whose entry folder path has the given length, so a record name of a chosen
    // length gives its root folder an exact path length.
    private static DocumentTemplate LongEntryTemplate(int entry = 121)
    {
        var template = AcceptanceTests.Template();
        foreach (var section in template.Destinations)
        {
            section.Library.EntryUrl =
                "https://example.sharepoint.com/sites/proto/"
                + section.Key.Substring(0, 7)
                + "/"
                + new string('e', entry - 21);
            section.Nodes.Add(
                new FolderNode
                {
                    Key = "deep",
                    ParentKey = "child",
                    Name = "Deep",
                }
            );
        }
        return template;
    }

    // The root folder's name is "C-" plus the record name, under the entry path.
    private static string RootNamed(int pathLength, int entry = 121) =>
        new string('r', pathLength - entry - 1 - 2);

    [Fact]
    public void PathsUpToSharePointsLimitAreCreatedWithANoticeAbove300Characters()
    {
        var plan = FolderPlanner.Plan(
            LongEntryTemplate(),
            Guid.NewGuid(),
            Names(RootNamed(350), "C")
        );
        // root 350, child 359 and deep 364 characters: all are created, each with a notice.
        Assert.Equal(6, plan.Count);
        Assert.Empty(plan.Waits);
        foreach (var section in new[] { "general", "sensitive" })
        {
            Assert.Contains(
                "Folder '"
                    + section
                    + "/root': files inside need short names; its path is 350 of 400 characters.",
                plan.Notices
            );
            Assert.Contains(
                "Folder '"
                    + section
                    + "/deep': files inside need short names; its path is 364 of 400 characters.",
                plan.Notices
            );
        }
        // At 300 characters or fewer there is no notice.
        Assert.Empty(
            FolderPlanner
                .Plan(LongEntryTemplate(), Guid.NewGuid(), Names(RootNamed(286), "C"))
                .Notices
        );
    }

    [Fact]
    public void AFolderWhoseEscapedAddressPassesTheConnectorLimitWaitsWithANotice()
    {
        // 399 characters, mostly CJK: each escapes to 9 characters, so the folder's address in
        // a SharePoint read would pass the connector's 2,048-character query string.
        var plan = FolderPlanner.Plan(
            LongEntryTemplate(200),
            Guid.NewGuid(),
            Names(new string('文', 196), "C")
        );
        Assert.Empty(plan);
        var wait = plan.Waits.First(w => w.Section == "general");
        Assert.Equal("root", wait.Node);
        Assert.Equal(FolderWaitReason.PathTooLong, wait.Reason);
        Assert.StartsWith("Folder 'general/root' needs a shorter path: ", wait.Notice);
        Assert.EndsWith(
            " characters; the limit is "
                + SharePointAddress.MaxQueryString.ToString(
                    "N0",
                    System.Globalization.CultureInfo.InvariantCulture
                )
                + ".",
            wait.Notice
        );
        Assert.Contains(wait.Notice, plan.Notices);
        // A 399-character path whose address fits is created.
        var fits = FolderPlanner.Plan(
            LongEntryTemplate(200),
            Guid.NewGuid(),
            Names(new string('文', 60) + new string('r', 136), "C")
        );
        Assert.Contains(fits, n => n.Node == "root");
        Assert.All(
            fits.Where(n => n.Section == "general"),
            n =>
                Assert.True(
                    SharePointAddress.Fits(
                        "/sites/proto/general/" + new string('e', 179) + "/" + n.RelativePath
                    )
                )
        );
    }

    [Fact]
    public void AFolderPastSharePointsPathLimitWaitsWithItsSubfoldersAndTheRestIsPlanned()
    {
        // root 395 characters is created; child would be 404 and waits with deep below it.
        var plan = FolderPlanner.Plan(
            LongEntryTemplate(200),
            Guid.NewGuid(),
            Names(RootNamed(395, 200), "C")
        );
        Assert.Equal(new[] { "root", "root" }, plan.Select(n => n.Node));
        Assert.Equal(2, plan.Waits.Count);
        var wait = plan.Waits.First(w => w.Section == "general");
        Assert.Equal("child", wait.Node);
        Assert.Null(wait.Field);
        Assert.Equal(FolderWaitReason.PathTooLong, wait.Reason);
        Assert.Equal(
            "Folder 'general/child' needs a shorter path: 404 characters; the limit is 400.",
            wait.Notice
        );
        Assert.Contains(wait.Notice, plan.Notices);
        // A root folder past the limit waits too; the plan itself never fails.
        var root = FolderPlanner.Plan(
            LongEntryTemplate(321),
            Guid.NewGuid(),
            Names(RootNamed(404, 321), "C")
        );
        Assert.Empty(root);
        Assert.All(root.Waits, w => Assert.Equal("root", w.Node));
    }

    [Fact]
    public void ANullValueFormatsAsNullForNaming()
    {
        Assert.Null(Value.Null(ValueKind.Text).Format());
        Assert.Null(Value.Null(ValueKind.DateOnly).Format());
        Assert.Equal("", Value.Text("").Format());
    }

    [Fact]
    public void PlanNoticesAreBounded()
    {
        var template = AcceptanceTests.Template();
        var plan = FolderPlanner.Plan(
            template,
            Guid.NewGuid(),
            Names(new string('a', 3000) + ":", "b")
        );
        Assert.Single(plan.Notices);
        Assert.True(plan.Notices[0].Length <= 603);
        Assert.EndsWith("...", plan.Notices[0], StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("../escape")]
    [InlineData("bad\\path")]
    [InlineData("trailing.")]
    [InlineData("CON.txt")]
    [InlineData(" lead")]
    public void ValidateStillRejectsUnsafeObservedNames(string name)
    {
        Assert.Throws<EvaluationBlockedException>(() => FolderNames.Validate(name));
        Assert.Throws<EvaluationBlockedException>(() => FolderNames.Validate(name, false));
    }
}
