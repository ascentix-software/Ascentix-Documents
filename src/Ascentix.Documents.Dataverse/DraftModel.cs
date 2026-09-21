using System;
using System.Runtime.Serialization;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class DraftDto
{
    [DataMember]
    public string? TemplateId { get; set; }

    [DataMember]
    public string? Name { get; set; }

    [DataMember]
    public string? RevisionId { get; set; }

    [DataMember]
    public string? RowVersion { get; set; }

    [DataMember]
    public string Table { get; set; } = "";

    [DataMember]
    public SourceDto[] Sources { get; set; } = Array.Empty<SourceDto>();

    [DataMember]
    public DestinationDto[] Destinations { get; set; } = Array.Empty<DestinationDto>();
}

[DataContract]
public sealed class DestinationDto
{
    [DataMember]
    public string Key { get; set; } = "";

    [DataMember]
    public string? Name { get; set; }

    [DataMember]
    public string LibraryId { get; set; } = "";

    [DataMember]
    public FolderDto[] Folders { get; set; } = Array.Empty<FolderDto>();
}

[DataContract]
public sealed class FolderDto
{
    [DataMember]
    public string Key { get; set; } = "";

    [DataMember]
    public string? Parent { get; set; }

    [DataMember]
    public string Name { get; set; } = "";

    [DataMember]
    public GroupDto? Condition { get; set; }
}

[DataContract]
public sealed class GroupDto
{
    [DataMember]
    public bool All { get; set; } = true;

    [DataMember]
    public ConditionDto[] Conditions { get; set; } = Array.Empty<ConditionDto>();

    [DataMember]
    public GroupDto[] Groups { get; set; } = Array.Empty<GroupDto>();
}

[DataContract]
public sealed class DraftResult
{
    [DataMember]
    public string TemplateId { get; set; } = "";

    [DataMember]
    public string RevisionId { get; set; } = "";

    [DataMember]
    public string RowVersion { get; set; } = "";

    [DataMember]
    public string Status { get; set; } = "Draft";
}
