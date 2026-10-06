using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using Ascentix.Documents.Dataverse;
using Xunit;

namespace Ascentix.Documents.SdkTests;

/// <summary>
/// Payloads saved by an earlier release lack members added since. The serializer skips
/// constructors and initializers, so every stored list must still read as present.
/// </summary>
public sealed class UpgradePayloadTests
{
    public static TheoryData<Type> StoredTypes()
    {
        var data = new TheoryData<Type>();
        foreach (
            var type in typeof(StoredDocument)
                .Assembly.GetTypes()
                .Where(t =>
                    typeof(StoredDocument).IsAssignableFrom(t)
                    && !t.IsAbstract
                    && t.GetConstructor(Type.EmptyTypes) != null
                )
                .OrderBy(t => t.Name)
        )
            data.Add(type);
        return data;
    }

    private static PropertyInfo[] Collections(Type type) =>
        type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p =>
                p.GetCustomAttribute<DataMemberAttribute>() != null
                && p.PropertyType != typeof(string)
                && typeof(IEnumerable).IsAssignableFrom(p.PropertyType)
            )
            .ToArray();

    [Theory]
    [MemberData(nameof(StoredTypes))]
    public void EveryStoredListReadsAsEmptyWhenAnEarlierReleaseDidNotWriteIt(Type type)
    {
        string json;
        using (var stream = new MemoryStream())
        {
            new DataContractJsonSerializer(type).WriteObject(
                stream,
                Activator.CreateInstance(type)
            );
            json = Encoding.UTF8.GetString(stream.ToArray());
        }
        var lists = Collections(type);
        var legacy = LegacyPayload.Strip(json, lists.Select(p => p.Name).ToArray());
        var read = typeof(JsonWire)
            .GetMethod(nameof(JsonWire.Read))!
            .MakeGenericMethod(type)
            .Invoke(null, new object[] { legacy })!;
        foreach (var list in lists)
            Assert.True(list.GetValue(read) != null, type.Name + "." + list.Name + " is null.");
    }

    [Fact]
    public void AMemberAnEarlierReleaseWroteButThisOneDroppedIsIgnored()
    {
        // 0.1.0.3 folder jobs stored ApprovedAclHash; 0.1.0.4 no longer has it.
        var json = JsonWire.Write(new OperationDocument { Key = "operation:legacy" });
        var legacy = json.Insert(1, "\"ApprovedAclHash\":\"acl-v1-hash\",");
        Assert.Equal("operation:legacy", JsonWire.Read<OperationDocument>(legacy).Key);
    }

    [Fact]
    public void TheAuditCoversTheListsAddedSince0103()
    {
        // The lists 0.1.0.4 added to stored documents; 0.1.0.3 rows have none of them.
        var added = new[]
        {
            (typeof(SecurityOperation), "Notices"),
            (typeof(SecurityOperation), "SkippedMembers"),
            (typeof(MembershipDocument), "Skipped"),
            (typeof(PolicyDocument), "Notices"),
            (typeof(ManagedGroup), "Principals"),
            (typeof(CatalogProbe), "Changes"),
            (typeof(RecordPlanDocument), "Notices"),
        };
        foreach (var (type, name) in added)
            Assert.Contains(Collections(type), p => p.Name == name);
    }
}
