using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Ascentix.Documents.Domain;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class AdapterTests
{
    [Fact]
    public void SdkValuesNormalizeWithoutDisplayLabels()
    {
        Assert.Equal("1.25", SnapshotReader.Normalize(ValueKind.Number, new Money(1.25m)).Format());
        Assert.Equal(
            "3",
            SnapshotReader.Normalize(ValueKind.Choice, new OptionSetValue(3)).Format()
        );
        var id = Guid.NewGuid();
        Assert.Equal(
            id.ToString(),
            SnapshotReader
                .Normalize(
                    ValueKind.Lookup,
                    new EntityReference("account", id) { Name = "Localized Name" }
                )
                .Format()
        );
        Assert.Equal(
            "1-3",
            SnapshotReader
                .Normalize(
                    ValueKind.MultiChoice,
                    new OptionSetValueCollection { new OptionSetValue(3), new OptionSetValue(1) }
                )
                .Format()
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            SnapshotReader.Normalize(ValueKind.Boolean, "true")
        );
    }

    [Fact]
    public void DateOnlyAndInstantRemainDifferentTypes()
    {
        Assert.Equal(
            "2026-09-08",
            SnapshotReader
                .Normalize(ValueKind.DateOnly, new DateTime(2026, 9, 8, 12, 30, 0))
                .Format()
        );
        Assert.Equal(
            "2026-09-08T12-30-00Z",
            SnapshotReader
                .Normalize(
                    ValueKind.DateTime,
                    new DateTime(2026, 9, 8, 12, 30, 0, DateTimeKind.Unspecified)
                )
                .Format()
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            SnapshotReader.Normalize(ValueKind.DateTime, DateTime.Now)
        );
    }

    [Fact]
    public void WireRejectsAmbiguousTypesAndOffsets()
    {
        Assert.Throws<EvaluationBlockedException>(() =>
            new ConditionDto
            {
                Column = "name",
                Operator = "Bad",
                LiteralKind = "Text",
                Literal = "a",
            }.ToModel()
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            new ConditionDto
            {
                Column = "name",
                Operator = "0",
                LiteralKind = "Text",
                Literal = "a",
            }.ToModel()
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            new ConditionDto
            {
                Column = "when",
                Operator = "Equal",
                LiteralKind = "DateTime",
                Literal = "2026-09-08T12:00:00",
            }.ToModel()
        );
        Assert.Throws<FormatException>(() =>
            new ConditionDto
            {
                Column = "number",
                Operator = "Equal",
                LiteralKind = "Number",
                Literal = "1,2",
            }.ToModel()
        );
    }

    private static DocumentTemplate Template()
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
                Columns = new Dictionary<string, ValueKind> { ["name"] = ValueKind.Text },
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
        template.Destinations.Add(
            new DestinationSection
            {
                Key = "general",
                Library = new ApprovedLibrary
                {
                    ApprovalId = Guid.NewGuid(),
                    WebId = Guid.NewGuid(),
                    ListId = Guid.NewGuid(),
                    EntryId = Guid.NewGuid(),
                    PolicyRevision = Guid.NewGuid(),
                    EntryUrl = "https://example.sharepoint.com/sites/proto/General",
                    Approved = true,
                    PolicyApplied = true,
                },
                Nodes = new List<FolderNode>
                {
                    new FolderNode { Key = "root", Name = "{root.name}" },
                },
            }
        );
        return template;
    }

    [Fact]
    public void ReaderProjectsOnlyApprovedColumnsAndOneLookup()
    {
        var service = new FakeService();
        var read = new SnapshotReader(service).Read(Template(), service.RootId);
        Assert.Equal(
            "Parent",
            read.Values.Resolve(new FieldReference("customer", "name")).Format()
        );
        Assert.Equal(2, service.Retrieves.Count);
        Assert.All(service.Retrieves, c => Assert.False(c.AllColumns));
        Assert.Equal(
            new[] { "name", "parentaccountid" },
            service.Retrieves[0].Columns.OrderBy(v => v)
        );
    }

    [Fact]
    public void DeniedLookupDoesNotBecomeFalseOrNull()
    {
        var service = new FakeService { DenyParent = true };
        Assert.Throws<InvalidOperationException>(() =>
            new SnapshotReader(service).Read(Template(), service.RootId)
        );
    }

    [Fact]
    public void SecuredColumnsFailClosedBeforeReadingData()
    {
        var service = new FakeService { Secured = true };
        Assert.Throws<EvaluationBlockedException>(() =>
            new SnapshotReader(service).Read(Template(), service.RootId)
        );
        Assert.Empty(service.Retrieves);
    }

    [Fact]
    public void NullLookupHasExplicitTypedNullAndNoRemoteRead()
    {
        var service = new FakeService { NullParent = true };
        var read = new SnapshotReader(service).Read(Template(), service.RootId);
        Assert.True(read.Values.Resolve(new FieldReference("customer", "name")).IsNull);
        Assert.Single(service.Retrieves);
    }

    private sealed class FakeService : IOrganizationService
    {
        public Guid RootId { get; } = Guid.NewGuid();
        private Guid ParentId { get; } = Guid.NewGuid();
        public bool DenyParent,
            NullParent,
            Secured;
        public List<ColumnSet> Retrieves { get; } = new List<ColumnSet>();

        public Entity Retrieve(string entityName, Guid id, ColumnSet columns)
        {
            Retrieves.Add(columns);
            if (id == ParentId && DenyParent)
                throw new InvalidOperationException("Caller denied lookup access.");
            var row = new Entity(entityName, id)
            {
                RowVersion = "42",
                ["name"] = id == RootId ? "Root" : "Parent",
            };
            if (id == RootId && !NullParent)
                row["parentaccountid"] = new EntityReference("account", ParentId);
            return row;
        }

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            if (!(request is RetrieveAttributeRequest attribute))
                throw new NotSupportedException();
            AttributeMetadata metadata =
                attribute.LogicalName == "parentaccountid"
                    ? (AttributeMetadata)
                        new LookupAttributeMetadata { Targets = new[] { "account" } }
                    : new StringAttributeMetadata();
            metadata.LogicalName = attribute.LogicalName;
            metadata.IsSecured = Secured;
            typeof(AttributeMetadata).GetProperty("IsValidForRead")!.SetValue(metadata, true, null);
            var response = new RetrieveAttributeResponse();
            response.Results["AttributeMetadata"] = metadata;
            return response;
        }

        public Guid Create(Entity entity) => throw new NotSupportedException();

        public void Update(Entity entity) => throw new NotSupportedException();

        public void Delete(string name, Guid id) => throw new NotSupportedException();

        public EntityCollection RetrieveMultiple(QueryBase query) =>
            throw new NotSupportedException();

        public void Associate(
            string name,
            Guid id,
            Relationship relationship,
            EntityReferenceCollection rows
        ) => throw new NotSupportedException();

        public void Disassociate(
            string name,
            Guid id,
            Relationship relationship,
            EntityReferenceCollection rows
        ) => throw new NotSupportedException();
    }
}
