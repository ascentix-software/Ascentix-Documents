using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Domain;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;
using Xunit.Abstractions;

namespace Ascentix.Documents.SdkTests;

/// <summary>
/// The measurements behind the Bounds constants (spec 6.9): worst-case payload sizes, and the
/// Dataverse calls each unit costs, which Step 7 multiplies by the per-call time measured in DEV.
/// </summary>
public sealed class CapTraceTests
{
    private const int PayloadLimit = 500000;
    private readonly ITestOutputHelper output;

    public CapTraceTests(ITestOutputHelper output) => this.output = output;

    private static FolderStep WorstFolder(int i) =>
        new FolderStep
        {
            Key =
                Guid.NewGuid().ToString("N")
                + ":"
                + Guid.NewGuid().ToString("N")
                + ":"
                + new string('s', 80)
                + ":"
                + new string('n', 79)
                + i,
            TemplateId = Guid.NewGuid(),
            RecordId = Guid.NewGuid(),
            Table = new string('t', 100),
            Section = new string('s', 80),
            Node = new string('n', 80),
            ParentBinding = new string('p', 80),
            LibraryId = Guid.NewGuid(),
            EntryId = Guid.NewGuid(),
            OriginalName = new string('o', 255),
            Candidate = new string('c', 255),
            PhysicalId = Guid.NewGuid(),
            PhysicalPath = new string('x', 400),
            LocationId = Guid.NewGuid(),
            Status = "Applied",
        };

    [Fact]
    public void AFolderJobStaysInsideItsPayloadAtTheFolderBound()
    {
        int one = JsonWire
            .Write(
                new OperationDocument
                {
                    Key = "folderjob:" + new string('k', 43),
                    Folders = new[] { WorstFolder(0) },
                }
            )
            .Length;
        int two = JsonWire
            .Write(
                new OperationDocument
                {
                    Key = "folderjob:" + new string('k', 43),
                    Folders = new[] { WorstFolder(0), WorstFolder(1) },
                }
            )
            .Length;
        int perFolder = two - one,
            fixedPart = one - perFolder;
        output.WriteLine(
            "Folder job payload: "
                + fixedPart
                + " + "
                + perFolder
                + " per folder; "
                + (PayloadLimit - fixedPart) / perFolder
                + " folders fit."
        );
        var full = new OperationDocument
        {
            Key = "folderjob:" + new string('k', 43),
            Folders = Enumerable
                .Range(0, Bounds.FoldersPerDestination)
                .Select(WorstFolder)
                .ToArray(),
        };
        Assert.True(JsonWire.Write(full).Length <= PayloadLimit);
    }

    [Fact]
    public void EveryBoundIsUsedByTheServerCheckAndSharedWithThePage()
    {
        string root = SourceRoot();
        string Source(string path) => File.ReadAllText(Path.Combine(root, path));
        Assert.Contains(
            "Bounds.FoldersPerDestination",
            Source("src/Ascentix.Documents.Dataverse/WorkerCoordinator.cs")
        );
        Assert.Contains(
            "Bounds.FoldersPerDestination",
            Source("src/Ascentix.Documents.Dataverse/DraftTemplate.cs")
        );
        Assert.Contains("Bounds.Destinations", Source("src/Ascentix.Documents.Domain/Planning.cs"));
        Assert.Contains(
            "Bounds.TeamEntries",
            Source("src/Ascentix.Documents.Dataverse/SecurityAdministration.cs")
        );
        Assert.Contains(
            "Bounds.TeamEntries",
            Source("src/Ascentix.Documents.Dataverse/LibraryProvisioning.cs")
        );
        Assert.Contains(
            "Bounds.PreviewRecords",
            Source("src/Ascentix.Documents.Dataverse/BatchReplan.cs")
        );
        Assert.Contains(
            "Bounds.Sources",
            Source("src/Ascentix.Documents.Dataverse/DocumentStore.cs")
        );
        foreach (var old in new[] { "Sources.Length > 6", "i < 6" })
            Assert.DoesNotContain(old, Source("src/Ascentix.Documents.Dataverse/DocumentStore.cs"));
        Assert.Contains(
            "Bounds.ConditionDepth",
            Source("src/Ascentix.Documents.Domain/Planning.cs")
        );
        Assert.DoesNotContain(
            "depth > 10",
            Source("src/Ascentix.Documents.Conditions/Expressions.cs")
        );
        Assert.Contains(
            "Bounds.ConfigurationRows",
            Source("src/Ascentix.Documents.Dataverse/DraftTemplate.cs")
        );
        // The literal caps are gone from the checks.
        foreach (
            var old in new[]
            {
                "Sources.Length > 6",
                "Destinations.Length > 10",
                "Folders.Length > 100",
                "++groupIndex > 100",
                "++conditionIndex > 100",
                "depth > 10",
            }
        )
            Assert.DoesNotContain(old, Source("src/Ascentix.Documents.Plugins/CreateDraftApi.cs"));
        Assert.DoesNotContain(
            "Destinations.Count > 10",
            Source("src/Ascentix.Documents.Domain/Planning.cs")
        );
    }

    [Fact]
    public void PlanningCostsAFixedNumberOfCallsPlusSomePerDestination()
    {
        int Calls(int destinations)
        {
            var f = new DurableWorkerTests.Fixture(seedBinding: false)
            {
                AllowedTables = new[] { "account" },
            };
            f.SeedTemplate();
            var destination = f.Service.Rows.Values.Single(r => r.LogicalName == "asx_destination");
            var folder = f.Service.Rows.Values.Single(r => r.LogicalName == "asx_folder");
            for (int i = 1; i < destinations; i++)
            {
                f.Service.Seed(
                    new Entity("asx_destination", Guid.NewGuid())
                    {
                        ["asx_revisionid"] = destination["asx_revisionid"],
                        ["asx_key"] = "d" + i,
                        ["asx_libraryid"] = destination["asx_libraryid"],
                    }
                );
                f.Service.Seed(
                    new Entity("asx_folder", Guid.NewGuid())
                    {
                        ["asx_revisionid"] = folder["asx_revisionid"],
                        ["asx_key"] = "root",
                        ["asx_sectionkey"] = "d" + i,
                        ["asx_expression"] = "{root.name}",
                        ["asx_order"] = 0,
                    }
                );
            }
            var queued = f.Execute(
                new WorkerRequest
                {
                    Command = "Queue",
                    TemplateId = f.TemplateId,
                    RecordId = f.RecordId,
                    RequestId = Guid.NewGuid(),
                }
            );
            int calls = 0;
            f.Service.QueryHook = _ =>
            {
                calls++;
                return null;
            };
            f.Execute(new WorkerRequest { Command = "Plan", Key = queued.Key });
            f.Service.QueryHook = null;
            return calls;
        }
        int one = Calls(1),
            ten = Calls(10);
        int perDestination = (ten - one) / 9;
        output.WriteLine(
            "Plan queries: "
                + (one - perDestination)
                + " fixed + "
                + perDestination
                + " per destination."
        );
        Assert.True(perDestination > 0);
    }

    [Fact]
    public void TheDeepestConditionTreeAPreviewCarriesFitsTheWireNestingQuota()
    {
        GroupDto Nested(int depth) =>
            depth == 0
                ? new GroupDto
                {
                    Conditions = new[]
                    {
                        new ConditionDto { Column = "name", Operator = "IsNotNull" },
                    },
                }
                : new GroupDto { Groups = new[] { Nested(depth - 1) } };
        PreviewRequest Request(int depth) =>
            new PreviewRequest
            {
                RecordId = Guid.NewGuid().ToString("D"),
                Draft = new DraftDto
                {
                    Table = "account",
                    Sources = new[]
                    {
                        new SourceDto { Alias = "root", Table = "account" },
                    },
                    Destinations = new[]
                    {
                        new DestinationDto
                        {
                            Key = "main",
                            LibraryId = Guid.NewGuid().ToString("D"),
                            Folders = new[]
                            {
                                new FolderDto
                                {
                                    Key = "root",
                                    Name = "{root.name}",
                                    Condition = Nested(depth),
                                },
                            },
                        },
                    },
                },
            };
        Predicate Model(GroupDto group) =>
            new Ascentix.Documents.Conditions.ConditionGroup(
                group.All,
                group
                    .Conditions.Select(c => (Predicate)c.ToModel())
                    .Concat(group.Groups.Select(Model))
                    .ToArray()
            );
        void Validate(PreviewRequest request)
        {
            var section = new DestinationSection
            {
                Key = "main",
                Library = new ApprovedLibrary
                {
                    ApprovalId = Guid.NewGuid(),
                    WebId = Guid.NewGuid(),
                    ListId = Guid.NewGuid(),
                    EntryId = Guid.NewGuid(),
                    EntryUrl = "https://example.sharepoint.com/sites/proto/General",
                    Approved = true,
                },
            };
            section.Nodes.Add(
                new FolderNode
                {
                    Key = "root",
                    Name = "{root.name}",
                    Condition = Model(request.Draft!.Destinations[0].Folders[0].Condition!),
                }
            );
            var template = new DocumentTemplate
            {
                Id = Guid.NewGuid(),
                Table = "account",
                Revision = 1,
            };
            template.Sources.Add(
                new SourceDto
                {
                    Alias = "root",
                    Table = "account",
                    Columns = new[]
                    {
                        new ColumnDto { Name = "name", Kind = "Text" },
                    },
                }.ToModel()
            );
            template.Destinations.Add(section);
            TemplateValidator.Validate(template);
        }
        // Nested(n) is a folder's condition group with n more groups nested inside it, so
        // Nested(Bounds.ConditionDepth - 1) has Bounds.ConditionDepth group levels: the deepest
        // the domain allows. It reads back within JsonWire's quota and one level more is refused.
        Validate(JsonWire.Read<PreviewRequest>(JsonWire.Write(Request(Bounds.ConditionDepth - 1))));
        Assert.Throws<EvaluationBlockedException>(() => Validate(Request(Bounds.ConditionDepth)));
        // The measured quota: MaxDepth 32 reads 11 nested groups and refuses 12, for a save
        // (DraftDto) and a preview (PreviewRequest) alike. A change to MaxDepth fails here.
        JsonWire.Read<PreviewRequest>(JsonWire.Write(Request(11)));
        JsonWire.Read<DraftDto>(JsonWire.Write(Request(11).Draft));
        Assert.ThrowsAny<Exception>(() =>
            JsonWire.Read<PreviewRequest>(JsonWire.Write(Request(12)))
        );
        Assert.ThrowsAny<Exception>(() =>
            JsonWire.Read<DraftDto>(JsonWire.Write(Request(12).Draft))
        );
    }

    [Fact]
    public void ARevisionIsReadAcrossEveryPage()
    {
        var service = new DurableWorkerTests.MemoryService();
        var revision = Guid.NewGuid();
        var seeded = Enumerable
            .Range(0, 5)
            .Select(i =>
            {
                var row = new Entity("asx_folder", Guid.NewGuid())
                {
                    ["asx_revisionid"] = new EntityReference("asx_revision", revision),
                    ["asx_key"] = "f" + i,
                };
                service.Seed(row);
                return row.Id;
            })
            .ToArray();
        service.Seed(
            new Entity("asx_folder", Guid.NewGuid())
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", Guid.NewGuid()),
            }
        );
        int pages = 0;
        // Pages of two rows stand in for Dataverse's 5,000, so five rows take three pages.
        service.QueryHook = query =>
        {
            pages++;
            query.PageInfo.Count = 2;
            return null;
        };
        var rows = new TemplateStore(service).Children("asx_folder", revision, "asx_key");
        Assert.Equal(3, pages);
        Assert.Equal(seeded.OrderBy(id => id), rows.Select(r => r.Id).OrderBy(id => id));
    }

    [Fact]
    public void TheTeamEntryRefusalNamesTheBound()
    {
        var error = Assert.Throws<EvaluationBlockedException>(() =>
            SecurityAdministration.Validate(
                Enumerable
                    .Range(0, Bounds.TeamEntries + 1)
                    .Select(_ => new PolicyEntry { TeamId = Guid.NewGuid(), Access = "Read" })
                    .ToArray(),
                new PolicyRole(),
                new PolicyRole()
            )
        );
        Assert.Equal(
            "At most "
                + Bounds.TeamEntries
                + " unique team policy entries with named access levels required.",
            error.Message
        );
    }

    private static string SourceRoot()
    {
        for (
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            dir != null;
            dir = dir.Parent
        )
            if (File.Exists(Path.Combine(dir.FullName, "Ascentix.Documents.sln")))
                return dir.FullName;
        throw new InvalidOperationException(
            "The repository root was not found above the test output."
        );
    }
}
