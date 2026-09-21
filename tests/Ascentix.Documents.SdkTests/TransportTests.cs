using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class TransportTests
{
    private static SharePointTarget Target() =>
        new SharePointTarget(
            "https://example.sharepoint.com/sites/proto",
            Guid.NewGuid(),
            Guid.NewGuid(),
            "/sites/proto/General"
        );

    [Fact]
    public void FolderCreateUsesResourcePathWithoutOwnershipColumn()
    {
        var nonce = Guid.NewGuid();
        var target = Target();
        var request = SharePointRequests.CreateFolder(target, target.EntryPath, "O'Brien # 100%");
        var body = JsonWire.Read<FolderCreateBody>(request.Body!);
        Assert.Equal("POST", request.Method);
        Assert.Equal("None", request.RetryPolicy);
        Assert.Contains(target.ListId.ToString(), request.RelativeUri);
        Assert.Equal(1, body.Info.Type);
        Assert.Equal("https://example.sharepoint.com/sites/proto/General", body.Info.Path.Url);
        Assert.Equal("O'Brien # 100%", body.Info.Name.Url);
        Assert.Single(body.Values);
        Assert.Equal("FileLeafRef", body.Values.Single().Field);
        Assert.False(body.NewDocument);
    }

    [Fact]
    public void ExactLookupUsesResourcePathWithoutScanningLibraryItems()
    {
        var target = Target();
        var request = SharePointRequests.FindFolder(target, target.EntryPath, "O'Brien # 100%");
        Assert.DoesNotContain("/items?", request.RelativeUri);
        Assert.Contains(
            "GetFolderByServerRelativePath(decodedUrl='/sites/proto/General/O''Brien # 100%')?$select=Exists,UniqueId,ServerRelativeUrl",
            Uri.UnescapeDataString(request.RelativeUri)
        );
    }

    [Fact]
    public void FolderAbsenceRequiresExplicitEvidence()
    {
        var request = new WorkerRequest
        {
            HttpStatus = 200,
            ResponseBody = "{\"d\":{\"ListItemAllFields\":null}}",
        };
        Assert.Throws<EvaluationBlockedException>(() =>
            SharePointObservations.Find(request, "/sites/proto/General/Example")
        );
        request.ResponseBody = "{\"d\":{\"Exists\":false,\"ListItemAllFields\":null}}";
        Assert.Null(SharePointObservations.Find(request, "/sites/proto/General/Example"));
        request.ResponseBody = "{\"d\":{\"Exists\":true,\"ListItemAllFields\":null}}";
        Assert.Throws<EvaluationBlockedException>(() =>
            SharePointObservations.Find(request, "/sites/proto/General/Example")
        );
    }

    [Theory]
    [InlineData("/sites/other/General")]
    [InlineData("/sites/proto/General/../Other")]
    [InlineData("/sites/proto/General//nested")]
    public void UnapprovedParentIsRejected(string parent) =>
        Assert.Throws<EvaluationBlockedException>(() =>
            SharePointRequests.CreateFolder(Target(), parent, "Root")
        );

    [Fact]
    public void TransportNeverBuildsAutomaticRetriesOrCredentialPayloads()
    {
        var request = SharePointRequests.AddMember(
            16,
            "i:0#.f|membership|synthetic@example.invalid"
        );
        var body = JsonWire.Read<AddMemberBody>(request.Body!);
        Assert.Equal("SP.User", body.Metadata.Type);
        Assert.Equal("None", request.RetryPolicy);
        Assert.Throws<EvaluationBlockedException>(() => SharePointRequests.RemoveMember(0, 15));
        Assert.Throws<EvaluationBlockedException>(() =>
            SharePointRequests.ChangeOwnedGrant(Target(), 16, 0, true)
        );
    }
}
