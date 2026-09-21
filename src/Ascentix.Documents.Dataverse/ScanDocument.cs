using System.Runtime.Serialization;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class ScanDocument : StoredDocument
{
    [DataMember]
    public int Page { get; set; } = 1;

    [DataMember]
    public string? Cookie { get; set; }
}
