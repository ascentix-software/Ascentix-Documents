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
