using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Domain;

namespace Ascentix.Documents.Dataverse;

public static class JsonWire
{
    // Stored documents come from asx_payload columns of at most 500,000 characters (MaxLength in
    // the solution's Entity.xml files). Such a document holds fewer than 500,000 values, so the
    // item and array quotas follow the length and never refuse a document the length admits.
    private const int MaxLength = 500000;

    public static T Read<T>(string json)
    {
        if (string.IsNullOrWhiteSpace(json) || json.Length > MaxLength)
            throw new EvaluationBlockedException("JSON input missing or too large.");
        using (
            var reader = JsonReaderWriterFactory.CreateJsonReader(
                Encoding.UTF8.GetBytes(json),
                new XmlDictionaryReaderQuotas
                {
                    MaxDepth = 32,
                    MaxStringContentLength = MaxLength,
                    MaxArrayLength = MaxLength,
                    MaxBytesPerRead = 4096,
                    MaxNameTableCharCount = 20000,
                }
            )
        )
            return (T)
                new DataContractJsonSerializer(
                    typeof(T),
                    new DataContractJsonSerializerSettings { MaxItemsInObjectGraph = MaxLength }
                ).ReadObject(reader)!;
    }

    public static string Write<T>(T value)
    {
        using (var stream = new MemoryStream())
        {
            new DataContractJsonSerializer(typeof(T)).WriteObject(stream, value);
            return Encoding.UTF8.GetString(stream.ToArray());
        }
    }
}

/// <summary>The result of publishing a template revision.</summary>
[DataContract]
public sealed class PublishResult
{
    [DataMember]
    public string Status { get; set; } = "";

    /// <summary>Information for the author; a notice never means the publish failed.</summary>
    [DataMember]
    public string[] Notices { get; set; } = Array.Empty<string>();
}

[DataContract]
public sealed class SourceDto
{
    [DataMember]
    public string Alias { get; set; } = "root";

    [DataMember]
    public string Table { get; set; } = "";

    [DataMember]
    public string? Lookup { get; set; }

    [DataMember]
    public ColumnDto[] Columns { get; set; } = Array.Empty<ColumnDto>();

    public ValueSource ToModel() =>
        new ValueSource
        {
            Alias = Alias,
            Table = Table,
            RootLookupColumn = Lookup,
            Columns = Columns.ToDictionary(
                c => c.Name,
                c => ParseEnum<ValueKind>(c.Kind),
                StringComparer.Ordinal
            ),
        };

    internal static T ParseEnum<T>(string value)
        where T : struct
    {
        if (
            !Enum.TryParse<T>(value, false, out var result)
            || !Enum.IsDefined(typeof(T), result)
            || Enum.GetName(typeof(T), result) != value
        )
            throw new EvaluationBlockedException("Invalid named enum value.");
        return result;
    }
}

[DataContract]
public sealed class ColumnDto
{
    [DataMember]
    public string Name { get; set; } = "";

    [DataMember]
    public string Kind { get; set; } = "";
}

[DataContract]
public sealed class ConditionDto
{
    [DataMember]
    public string Source { get; set; } = "root";

    [DataMember]
    public string Column { get; set; } = "";

    [DataMember]
    public string Operator { get; set; } = "";

    [DataMember]
    public string? LiteralKind { get; set; }

    [DataMember]
    public string? Literal { get; set; }

    [DataMember]
    public string? RightSource { get; set; }

    [DataMember]
    public string? RightColumn { get; set; }

    /// <summary>Output only (LoadDraft): the name of the record a lookup condition compares with.</summary>
    [DataMember]
    public string? LiteralLabel { get; set; }

    /// <summary>Output only (LoadDraft): the table of that record.</summary>
    [DataMember]
    public string? LiteralTable { get; set; }

    public Condition ToModel()
    {
        Value? value = null;
        if (LiteralKind != null)
        {
            var kind = SourceDto.ParseEnum<ValueKind>(LiteralKind);
            if (Literal == null)
                throw new EvaluationBlockedException("Literal is required.");
            switch (kind)
            {
                case ValueKind.Text:
                    value = Value.Text(Literal);
                    break;
                case ValueKind.Number:
                    value = Value.Number(
                        decimal.Parse(
                            Literal,
                            NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                            CultureInfo.InvariantCulture
                        )
                    );
                    break;
                case ValueKind.Boolean:
                    value = Value.Boolean(bool.Parse(Literal));
                    break;
                case ValueKind.Choice:
                    value = Value.Choice(int.Parse(Literal, CultureInfo.InvariantCulture));
                    break;
                case ValueKind.Lookup:
                    value = Value.Lookup(Guid.Parse(Literal));
                    break;
                case ValueKind.DateOnly:
                    value = Value.DateOnly(
                        DateTime.ParseExact(Literal, "yyyy-MM-dd", CultureInfo.InvariantCulture)
                    );
                    break;
                case ValueKind.DateTime:
                    if (
                        !(
                            Literal.EndsWith("Z", StringComparison.Ordinal)
                            || System.Text.RegularExpressions.Regex.IsMatch(
                                Literal,
                                @"[+-]\d{2}:\d{2}$"
                            )
                        )
                    )
                        throw new EvaluationBlockedException(
                            "Timestamp requires an explicit offset."
                        );
                    value = Value.Instant(
                        DateTimeOffset.Parse(Literal, CultureInfo.InvariantCulture)
                    );
                    break;
                case ValueKind.MultiChoice:
                    value = Value.MultiChoice(
                        Literal.Split(',').Select(v => int.Parse(v, CultureInfo.InvariantCulture))
                    );
                    break;
            }
        }
        if ((RightSource == null) != (RightColumn == null))
            throw new EvaluationBlockedException("Right field reference is incomplete.");
        return new Condition(
            new FieldReference(Source, Column),
            SourceDto.ParseEnum<Comparison>(Operator),
            value,
            RightSource == null ? null : new FieldReference(RightSource, RightColumn!)
        );
    }
}

[DataContract]
public sealed class PreviewRequest
{
    /// <summary>The saved revision to preview; send it or Draft, not both.</summary>
    [DataMember]
    public string RevisionId { get; set; } = "";

    /// <summary>The editor's unsaved draft to preview; send it or RevisionId, not both.</summary>
    [DataMember]
    public DraftDto? Draft { get; set; }

    [DataMember]
    public string RecordId { get; set; } = "";
}

[DataContract]
public sealed class PlanDto
{
    [DataMember]
    public string Status { get; set; } = "PreviewOnly";

    [DataMember]
    public IntentDto[] Folders { get; set; } = Array.Empty<IntentDto>();

    /// <summary>Adjusted folder names and folders waiting for a value.</summary>
    [DataMember]
    public string[] Notices { get; set; } = Array.Empty<string>();

    public static PlanDto From(FolderPlan intents) =>
        new PlanDto
        {
            Notices = intents.Notices.ToArray(),
            Folders = intents
                .Select(i => new IntentDto
                {
                    Section = i.Section,
                    Node = i.Node,
                    Parent = i.Parent,
                    Name = i.Name,
                    RelativePath = i.RelativePath,
                    BindingKey = i.BindingKey,
                    LibraryApprovalId = i.LibraryApprovalId.ToString("D"),
                })
                .ToArray(),
        };
}

[DataContract]
public sealed class IntentDto
{
    [DataMember]
    public string Section { get; set; } = "";

    [DataMember]
    public string Node { get; set; } = "";

    [DataMember]
    public string? Parent { get; set; }

    [DataMember]
    public string Name { get; set; } = "";

    [DataMember]
    public string RelativePath { get; set; } = "";

    [DataMember]
    public string BindingKey { get; set; } = "";

    [DataMember]
    public string LibraryApprovalId { get; set; } = "";
}
