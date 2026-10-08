using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Domain;
using Xunit;

namespace Ascentix.Documents.Tests;

public class AcceptanceTests
{
    private static readonly FieldReference Flag = new("root", "flag");
    private static readonly Snapshot Snapshot = new(
        new Dictionary<string, Value>
        {
            ["root.flag"] = Value.Boolean(true),
            ["root.name"] = Value.Text("Example"),
            ["customer.name"] = Value.Text("Customer"),
        }
    );

    private static Condition True() => new(Flag, Comparison.Equal, Value.Boolean(true));

    private static Condition False() => new(Flag, Comparison.Equal, Value.Boolean(false));

    public static DocumentTemplate Template()
    {
        var template = new DocumentTemplate
        {
            Id = Guid.NewGuid(),
            Revision = 1,
            Table = "account",
        };
        template.Sources.Add(
            new ValueSource
            {
                Table = "account",
                Columns = new Dictionary<string, ValueKind>
                {
                    ["name"] = ValueKind.Text,
                    ["flag"] = ValueKind.Boolean,
                },
            }
        );
        template.Sources.Add(
            new ValueSource
            {
                Alias = "customer",
                Table = "account",
                RootLookupColumn = "parentaccountid",
                Columns = new Dictionary<string, ValueKind> { ["name"] = ValueKind.Text },
            }
        );
        foreach (var key in new[] { "general", "sensitive" })
            template.Destinations.Add(
                new DestinationSection
                {
                    Key = key,
                    Library = new ApprovedLibrary
                    {
                        ApprovalId = Guid.NewGuid(),
                        WebId = Guid.NewGuid(),
                        ListId = Guid.NewGuid(),
                        EntryId = Guid.NewGuid(),
                        EntryUrl = "https://example.sharepoint.com/sites/proto/" + key,
                        Approved = true,
                    },
                    Nodes = new List<FolderNode>
                    {
                        new() { Key = "root", Name = "{customer.name}-{root.name}" },
                        new()
                        {
                            Key = "child",
                            ParentKey = "root",
                            Name = "Included",
                            Condition = True(),
                        },
                        new()
                        {
                            Key = "hidden",
                            ParentKey = "root",
                            Name = "Hidden",
                            Condition = False(),
                        },
                        new()
                        {
                            Key = "grandchild",
                            ParentKey = "hidden",
                            Name = "Never",
                        },
                    },
                }
            );
        return template;
    }

    [Theory]
    [InlineData(true, 4)]
    [InlineData(false, 2)]
    public void RootConditionControlsWholeDestination(bool include, int count)
    {
        var template = Template();
        template.Destinations[0].Nodes[0].Condition = include ? True() : False();
        var plan = FolderPlanner.Plan(template, Guid.NewGuid(), Snapshot);
        Assert.Equal(count, plan.Count);
        Assert.Equal(include ? 2 : 0, plan.Count(n => n.Section == "general"));
        Assert.Equal(2, plan.Count(n => n.Section == "sensitive"));
    }

    [Fact]
    public void NestedGroupsUseChildVerdicts()
    {
        Assert.True(new ConditionGroup(false, new ConditionGroup(true, True())).Evaluate(Snapshot));
        Assert.True(
            new ConditionGroup(true, new ConditionGroup(false, False(), True())).Evaluate(Snapshot)
        );
        Assert.False(new ConditionGroup(true, True(), False()).Evaluate(Snapshot));
        Assert.Throws<EvaluationBlockedException>(() => new ConditionGroup(true));
    }

    [Fact]
    public void MissingSnapshotIsBlockedEvenBehindTrueOr()
    {
        var missing = new Condition(new FieldReference("customer", "secret"), Comparison.IsNull);
        Assert.Throws<EvaluationBlockedException>(() =>
            new ConditionGroup(false, True(), missing).Evaluate(Snapshot)
        );
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("fr-CA")]
    public void TypedValuesIgnoreAmbientCulture(string culture)
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(culture);
            Assert.True(
                ValueComparer.Compare(Value.Number(1.2m), Comparison.Greater, Value.Number(1.1m))
            );
            Assert.Equal("1.2", Value.Number(1.2m).Format());
            Assert.False(
                ValueComparer.Compare(Value.Text("01"), Comparison.Equal, Value.Text("1"))
            );
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }

    [Fact]
    public void TypedNullAndOperatorsAreExplicit()
    {
        Assert.True(ValueComparer.Compare(Value.Null(ValueKind.Text), Comparison.IsNull));
        Assert.False(ValueComparer.Compare(Value.Text(""), Comparison.IsNull));
        Assert.False(
            ValueComparer.Compare(Value.Null(ValueKind.Text), Comparison.NotEqual, Value.Text("a"))
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            ValueComparer.Compare(Value.Boolean(true), Comparison.Greater, Value.Boolean(false))
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            ValueComparer.Compare(Value.Choice(1), Comparison.Equal, Value.Number(1))
        );
    }

    [Fact]
    public void DatesAndChoiceSetsAreNormalized()
    {
        Assert.True(
            ValueComparer.Compare(
                Value.Instant(
                    DateTimeOffset.Parse("2026-09-08T01:00:00-04:00", CultureInfo.InvariantCulture)
                ),
                Comparison.Equal,
                Value.Instant(
                    DateTimeOffset.Parse("2026-09-08T05:00:00Z", CultureInfo.InvariantCulture)
                )
            )
        );
        Assert.True(
            ValueComparer.Compare(
                Value.MultiChoice(new[] { 3, 1, 1 }),
                Comparison.Equal,
                Value.MultiChoice(new[] { 1, 3 })
            )
        );
        Assert.True(
            ValueComparer.Compare(
                Value.MultiChoice(new[] { 1, 3 }),
                Comparison.Contains,
                Value.MultiChoice(new[] { 3 })
            )
        );
    }

    [Fact]
    public void MultipleDestinationsFanOutWithParentSuppressionAndStableKeys()
    {
        var template = Template();
        var record = Guid.NewGuid();
        var first = FolderPlanner.Plan(template, record, Snapshot);
        Assert.Equal(4, first.Count);
        Assert.Equal(2, first.Count(n => n.IsRoot));
        Assert.DoesNotContain(first, n => n.Node == "grandchild");
        template.Revision++;
        Assert.Equal(
            first.Select(n => n.BindingKey),
            FolderPlanner.Plan(template, record, Snapshot).Select(n => n.BindingKey)
        );
        Assert.All(first.Where(n => n.IsRoot), n => Assert.Equal("Customer-Example", n.Name));
    }

    [Fact]
    public void UnsupportedSourceAndTypedPayloadFailBeforePlanning()
    {
        var template = Template();
        template.Destinations[0].Nodes[0].Name = "{root.customer.name}";
        Assert.Throws<EvaluationBlockedException>(() => TemplateValidator.Validate(template));
        template = Template();
        template.Destinations[0].Nodes[1].Condition = new Condition(
            Flag,
            Comparison.Greater,
            Value.Boolean(false)
        );
        Assert.Throws<EvaluationBlockedException>(() => TemplateValidator.Validate(template));
    }

    [Fact]
    public void InvalidTreeAndSuspendedDestinationBlock()
    {
        var template = Template();
        template.Destinations[0].Nodes[1].ParentKey = "child";
        Assert.Throws<EvaluationBlockedException>(() => TemplateValidator.Validate(template));
        template = Template();
        template.Destinations[0].Library.Approved = false;
        var suspended = Assert.Throws<EvaluationBlockedException>(() =>
            TemplateValidator.Validate(template)
        );
        Assert.Equal("Destination must be approved and have a valid identity.", suspended.Message);
        template = Template();
        template.Destinations[0].Library.EntryId = Guid.Empty;
        Assert.Throws<EvaluationBlockedException>(() => TemplateValidator.Validate(template));
    }

    [Theory]
    [InlineData("../escape")]
    [InlineData("Forms")]
    [InlineData("bad\\path")]
    [InlineData("trailing.")]
    [InlineData("CON.txt")]
    public void UnsafeNamesFail(string name) =>
        Assert.Throws<EvaluationBlockedException>(() => FolderNames.Validate(name));

    [Fact]
    public void IncludedSiblingDuplicatesWaitWhileTheFirstIsPlanned()
    {
        var template = Template();
        template
            .Destinations[0]
            .Nodes.Add(
                new FolderNode
                {
                    Key = "duplicate",
                    ParentKey = "root",
                    Name = "included",
                }
            );
        template
            .Destinations[0]
            .Nodes.Add(
                new FolderNode
                {
                    Key = "below",
                    ParentKey = "duplicate",
                    Name = "Below",
                }
            );
        var plan = FolderPlanner.Plan(template, Guid.NewGuid(), Snapshot);
        Assert.Equal(4, plan.Count);
        Assert.Contains(plan, n => n.Section == "general" && n.Node == "child");
        Assert.DoesNotContain(plan, n => n.Node == "duplicate" || n.Node == "below");
        Assert.Equal(
            new[]
            {
                "Folder 'general/duplicate' has the same name 'included' as 'general/child'; it waits until the names differ.",
            },
            plan.Notices
        );
    }

    [Fact]
    public void PureAssembliesHaveNoSdkOrEngineReferences()
    {
        foreach (var assembly in new[] { typeof(Value).Assembly, typeof(FolderPlanner).Assembly })
            Assert.DoesNotContain(
                assembly.GetReferencedAssemblies(),
                a => a.Name!.Contains("Xrm") || a.Name.Contains("RulesEngine")
            );
    }
}
