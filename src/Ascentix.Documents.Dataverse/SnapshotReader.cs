using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Domain;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public sealed class SnapshotRead
{
    public Snapshot Values { get; }
    public IReadOnlyList<EntityReference> Records { get; }
    public IReadOnlyList<string> Versions { get; }

    public SnapshotRead(
        Snapshot values,
        IReadOnlyList<EntityReference> records,
        IReadOnlyList<string> versions
    )
    {
        Values = values;
        Records = records;
        Versions = versions;
    }
}

public sealed class SnapshotReader
{
    private readonly IOrganizationService service;

    public SnapshotReader(IOrganizationService service)
    {
        this.service = service ?? throw new ArgumentNullException(nameof(service));
    }

    public SnapshotRead Read(DocumentTemplate template, Guid recordId)
    {
        TemplateValidator.Validate(template);
        var rootSource = template.Sources.Single(s => s.Alias == "root");
        var rootColumns = rootSource
            .Columns.Keys.Concat(
                template
                    .Sources.Where(s => s.RootLookupColumn != null)
                    .Select(s => s.RootLookupColumn!)
            )
            .Distinct()
            .ToArray();
        ValidateMetadata(template);
        // Runs under the explicitly supplied caller/planner service, never SYSTEM.
        var root = service.Retrieve(template.Table, recordId, new ColumnSet(rootColumns));
        var records = new List<EntityReference>();
        var versions = new List<string>();
        var values = new Dictionary<string, Value>();
        foreach (var source in template.Sources)
        {
            Entity? row = root;
            if (source.Alias != "root")
            {
                var lookup = root.GetAttributeValue<EntityReference>(source.RootLookupColumn!);
                if (lookup == null)
                    row = null;
                else
                {
                    if (lookup.LogicalName != source.Table)
                        throw new EvaluationBlockedException(
                            "Lookup target differs from the published source type."
                        );
                    // A denied retrieve is allowed to throw; it is never converted to a null row.
                    row = service.Retrieve(
                        source.Table,
                        lookup.Id,
                        new ColumnSet(source.Columns.Keys.ToArray())
                    );
                }
            }
            if (row != null)
            {
                records.Add(row.ToEntityReference());
                versions.Add(row.RowVersion ?? "");
            }
            foreach (var column in source.Columns)
            {
                object? raw = null;
                if (row != null)
                    row.Attributes.TryGetValue(column.Key, out raw);
                values.Add(source.Alias + "." + column.Key, Normalize(column.Value, raw));
            }
        }
        return new SnapshotRead(new Snapshot(values), records.AsReadOnly(), versions.AsReadOnly());
    }

    public void ValidateMetadata(DocumentTemplate template)
    {
        TemplateValidator.Validate(template);
        foreach (var source in template.Sources)
        {
            foreach (var column in source.Columns)
                ValidateAttribute(source.Table, column.Key, column.Value);
            if (source.RootLookupColumn != null)
            {
                var attribute =
                    ReadAttribute(template.Table, source.RootLookupColumn)
                    as LookupAttributeMetadata;
                if (
                    attribute == null
                    || attribute.IsSecured == true
                    || attribute.IsValidForRead != true
                    || attribute.Targets == null
                    || !attribute.Targets.Contains(source.Table)
                )
                    throw new EvaluationBlockedException(
                        "Direct lookup metadata is not approved or readable."
                    );
            }
        }
    }

    private AttributeMetadata ReadAttribute(string table, string column) =>
        (
            (RetrieveAttributeResponse)
                service.Execute(
                    new RetrieveAttributeRequest
                    {
                        EntityLogicalName = table,
                        LogicalName = column,
                        RetrieveAsIfPublished = false,
                    }
                )
        ).AttributeMetadata;

    private void ValidateAttribute(string table, string column, ValueKind kind)
    {
        var attribute = ReadAttribute(table, column);
        // Source fields must be readable, unsecured, and match the published value type.
        if (attribute.IsValidForRead != true || attribute.IsSecured == true)
            throw new EvaluationBlockedException(
                "Unreadable or field-secured source columns are not supported in this milestone."
            );
        if (Kind(attribute) != kind)
            throw new EvaluationBlockedException(
                "Published field type differs from current metadata."
            );
    }

    public static ValueKind Kind(AttributeMetadata metadata)
    {
        if (metadata is MultiSelectPicklistAttributeMetadata)
            return ValueKind.MultiChoice;
        switch (metadata.AttributeType)
        {
            case AttributeTypeCode.String:
            case AttributeTypeCode.Memo:
                return ValueKind.Text;
            case AttributeTypeCode.Decimal:
            case AttributeTypeCode.Double:
            case AttributeTypeCode.Integer:
            case AttributeTypeCode.BigInt:
            case AttributeTypeCode.Money:
                return ValueKind.Number;
            case AttributeTypeCode.Boolean:
                return ValueKind.Boolean;
            case AttributeTypeCode.Picklist:
            case AttributeTypeCode.State:
            case AttributeTypeCode.Status:
                return ValueKind.Choice;
            case AttributeTypeCode.Lookup:
            case AttributeTypeCode.Customer:
            case AttributeTypeCode.Owner:
            case AttributeTypeCode.Uniqueidentifier:
                return ValueKind.Lookup;
            case AttributeTypeCode.DateTime:
                var date = (DateTimeAttributeMetadata)metadata;
                if (date.DateTimeBehavior?.Value == DateTimeBehavior.DateOnly.Value)
                    return ValueKind.DateOnly;
                if (date.DateTimeBehavior?.Value == DateTimeBehavior.TimeZoneIndependent.Value)
                    throw new EvaluationBlockedException(
                        "Time-zone independent timestamp support requires an explicit value kind."
                    );
                return ValueKind.DateTime;
            default:
                throw new EvaluationBlockedException("Unsupported Dataverse field type.");
        }
    }

    public static Value Normalize(ValueKind kind, object? raw)
    {
        if (raw == null)
            return Value.Null(kind);
        if (raw is AliasedValue alias)
            raw = alias.Value;
        switch (kind)
        {
            case ValueKind.Text:
                if (raw is string text)
                    return Value.Text(text);
                break;
            case ValueKind.Number:
                if (raw is Money money)
                    return Value.Number(money.Value);
                if (raw is decimal || raw is double || raw is int || raw is long)
                    return Value.Number(
                        Convert.ToDecimal(raw, System.Globalization.CultureInfo.InvariantCulture)
                    );
                break;
            case ValueKind.Boolean:
                if (raw is bool flag)
                    return Value.Boolean(flag);
                break;
            case ValueKind.Choice:
                if (raw is OptionSetValue choice)
                    return Value.Choice(choice.Value);
                break;
            case ValueKind.Lookup:
                if (raw is EntityReference lookup)
                    return Value.Lookup(lookup.Id);
                if (raw is Guid id)
                    return Value.Lookup(id);
                break;
            case ValueKind.DateOnly:
                if (raw is DateTime day)
                    return Value.DateOnly(day);
                break;
            case ValueKind.DateTime:
                if (raw is DateTime date && date.Kind != DateTimeKind.Local)
                    return Value.Instant(
                        new DateTimeOffset(DateTime.SpecifyKind(date, DateTimeKind.Utc))
                    );
                break;
            case ValueKind.MultiChoice:
                if (raw is OptionSetValueCollection choices)
                    return Value.MultiChoice(choices.Select(c => c.Value));
                break;
        }
        throw new EvaluationBlockedException("SDK value does not match the published field type.");
    }
}
