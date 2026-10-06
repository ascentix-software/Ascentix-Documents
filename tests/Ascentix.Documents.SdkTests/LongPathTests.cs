using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

/// <summary>
/// Folder paths up to SharePoint's 400-character limit must reach SharePoint through the HTTP
/// with Microsoft Entra ID connector. The connector refuses a request whose URL path is longer
/// than its ASP.NET maxUrlLength (401 "The length of the URL for this request exceeds the
/// configured maxUrlLength value."). ASP.NET applies maxUrlLength to the URL path only, not the
/// query string, and its default is 260 characters
/// (https://learn.microsoft.com/dotnet/api/system.web.configuration.httpruntimesection.maxurllength).
/// Power Automate allows a request URL of 16,384 characters in all
/// (https://learn.microsoft.com/power-automate/limits-and-config).
/// </summary>
public sealed class LongPathTests
{
    private const int MaxUrlPath = 260;
    private const int MaxRequestUrl = 16384;
    private const int MaxQueryString = 2048;

    // "/sites/proto/General" (20) + "/" + 200 + "/" + 177 = 399 characters.
    private static readonly string RecordName = Words(200);
    private static readonly string ChildName = Words(177);

    /// <summary>Letters and spaces, so most characters are escaped in a URL.</summary>
    private static string Words(int length)
    {
        var text = string.Concat(Enumerable.Repeat("A B ", length)).Substring(0, length);
        return text.EndsWith(" ", StringComparison.Ordinal)
            ? text.Substring(0, length - 1) + "C"
            : text;
    }

    private static string PathPart(string url)
    {
        int query = url.IndexOf('?');
        return query < 0 ? url : url.Substring(0, query);
    }

    [Fact]
    public void ReadsOfA399CharacterPathKeepTheirUrlPathShortAndFixed()
    {
        var target = new SharePointTarget(
            "https://example.sharepoint.com/sites/proto",
            Guid.NewGuid(),
            Guid.NewGuid(),
            "/sites/proto/General"
        );
        string parent = target.EntryPath + "/" + RecordName;
        Assert.Equal(399, (parent + "/" + ChildName).Length);
        var urls = new[]
        {
            SharePointRequests.FindFolder(target, parent, ChildName).RelativeUri,
            SharePointRequests.FindFolder(target, target.EntryPath, "Short").RelativeUri,
            SharePointRequests.ByPath(
                "GetFolderByServerRelativePath",
                parent + "/" + ChildName,
                "$select=UniqueId,ServerRelativeUrl"
            ),
            SharePointRequests.ByPath(
                "GetFileByServerRelativePath",
                parent + "/" + ChildName,
                "$select=UniqueId,ServerRelativeUrl"
            ),
        }.Select(relative => target.Web.AbsoluteUri.TrimEnd('/') + "/" + relative);
        foreach (var url in urls)
        {
            Assert.True(PathPart(url).Length <= MaxUrlPath, PathPart(url));
            Assert.DoesNotContain("A%20B", PathPart(url));
            Assert.True(url.Length - PathPart(url).Length <= MaxQueryString, url);
            Assert.True(url.Length <= MaxRequestUrl, url);
        }
    }

    [Fact]
    public void AliasCarriesTheExactPathIncludingQuotesHashAndPercent()
    {
        string relative = SharePointRequests.ByPath(
            "GetFolderByServerRelativePath",
            "/sites/proto/General/O'Brien # 100% & Co",
            "$select=UniqueId"
        );
        Assert.StartsWith("_api/web/GetFolderByServerRelativePath(decodedUrl=@p)?", relative);
        Assert.Equal("/sites/proto/General/O'Brien # 100% & Co", SharePointDouble.Alias(relative));
    }

    [Fact]
    public void A399CharacterPathIsPlannedAndAppliedThroughTheConnector()
    {
        var f = new DurableWorkerTests.Fixture(seedBinding: false);
        f.SeedTemplate();
        f.Service.Rows[f.RecordId]["name"] = RecordName;
        f.Service.Seed(
            new Entity("asx_folder", Guid.NewGuid())
            {
                ["asx_revisionid"] = new EntityReference("asx_revision", f.RevisionId),
                ["asx_key"] = "deep",
                ["asx_sectionkey"] = "general",
                ["asx_parentkey"] = "root",
                ["asx_expression"] = ChildName,
                ["asx_order"] = 1,
            }
        );
        var planned = f.PlanRecord();
        var key = Assert.Single(planned.Keys);
        var sharePoint = new SharePointDouble(f);
        Assert.Equal("Applied", sharePoint.Drive(key));
        Assert.Contains("/sites/proto/General/" + RecordName + "/" + ChildName, sharePoint.Folders);
        Assert.Contains(
            sharePoint.Requests,
            url => url.Contains(Uri.EscapeDataString(RecordName + "/" + ChildName))
        );
        Assert.All(sharePoint.Requests, url => Assert.True(PathPart(url).Length <= MaxUrlPath));
    }

    [Fact]
    public void ConnectorUrlLengthRefusalIsNamedAndNotAnObservationMismatch()
    {
        var f = new DurableWorkerTests.Fixture();
        var read = f.Preflight(f.Claim());
        var refused = f.Call(
            "Observe",
            read,
            "{\"statusCode\":401,\"source\":\"canada.azure-apihub.net\",\"innerError\":\"The length of the URL for this request exceeds the configured maxUrlLength value.\"}",
            401
        );
        Assert.Equal("Blocked", refused.Status);
        Assert.Equal(
            "RequestUrlTooLong",
            f.Store.Require<OperationDocument>("asx_operation", f.Operation.Key).Value.ErrorCode
        );
    }

    /// <summary>
    /// A SharePoint library behind the connector: it refuses a URL path longer than the
    /// connector's maxUrlLength the way the connector does, and otherwise answers folder reads
    /// and creates from an in-memory folder tree.
    /// </summary>
    private sealed class SharePointDouble
    {
        private readonly DurableWorkerTests.Fixture f;
        private readonly Dictionary<string, Guid> folders = new Dictionary<string, Guid>(
            StringComparer.OrdinalIgnoreCase
        );
        public List<string> Requests { get; } = new List<string>();
        public IEnumerable<string> Folders => folders.Keys;

        public SharePointDouble(DurableWorkerTests.Fixture fixture)
        {
            f = fixture;
            folders["/sites/proto/General"] = f.EntryId;
        }

        /// <summary>The decoded value of the @p parameter alias.</summary>
        public static string Alias(string relative)
        {
            string query = relative.Substring(relative.IndexOf('?') + 1);
            string value = query
                .Split('&')
                .Single(p => p.StartsWith("@p=", StringComparison.Ordinal))
                .Substring(3);
            value = Uri.UnescapeDataString(value);
            Assert.True(value.Length >= 2 && value[0] == '\'' && value[value.Length - 1] == '\'');
            return value.Substring(1, value.Length - 2).Replace("''", "'");
        }

        private static string InlinePath(string relative)
        {
            const string start = "(decodedUrl='";
            int from = relative.IndexOf(start, StringComparison.Ordinal) + start.Length;
            int to = relative.IndexOf("')", from, StringComparison.Ordinal);
            return Uri.UnescapeDataString(relative.Substring(from, to - from)).Replace("''", "'");
        }

        private (int status, string body) Answer(WorkerResult work)
        {
            string url = work.SiteUrl + "/" + work.Http!.RelativeUri;
            Requests.Add(url);
            int query = url.IndexOf('?');
            if ((query < 0 ? url.Length : query) > MaxUrlPath)
                return (
                    401,
                    "{\"statusCode\":401,\"source\":\"canada.azure-apihub.net\",\"innerError\":\"The length of the URL for this request exceeds the configured maxUrlLength value.\"}"
                );
            string relative = work.Http.RelativeUri;
            if (work.Http.Method == "POST")
            {
                var create = JsonWire.Read<FolderCreateBody>(work.Http.Body!);
                string parent = new Uri(create.Info.Path.Url).AbsolutePath;
                parent = Uri.UnescapeDataString(parent);
                folders[parent + "/" + create.Info.Name.Url] = Guid.NewGuid();
                return (200, DurableWorkerTests.CreateBody());
            }
            if (relative.StartsWith("_api/web/lists(", StringComparison.Ordinal))
                return (200, f.LibraryBody());
            string path = relative.Contains("decodedUrl=@p")
                ? Alias(relative)
                : InlinePath(relative);
            if (!folders.TryGetValue(path, out var id))
                return (404, "{}");
            if (work.ProbeKind == "Folder")
                return (
                    200,
                    JsonWire.Write(
                        new ODataEnvelope<FolderLookupObservation>
                        {
                            Data = new FolderLookupObservation
                            {
                                Exists = true,
                                Id = id,
                                Path = path,
                                ListItemAllFields = new ItemObservation
                                {
                                    ItemId = 7,
                                    Id = id,
                                    Name = path.Substring(path.LastIndexOf('/') + 1),
                                    Path = path,
                                    Type = 1,
                                },
                            },
                        }
                    )
                );
            return (
                200,
                JsonWire.Write(
                    new ODataEnvelope<FolderObservation>
                    {
                        Data = new FolderObservation { Id = id, Path = path },
                    }
                )
            );
        }

        private WorkerResult Step(string command, string key, WorkerResult work, string run) =>
            f.Execute(
                new WorkerRequest
                {
                    Command = command,
                    Key = key,
                    RunId = run,
                    Token = work.Token,
                    ProbeId = work.ProbeId,
                    ProbeKind = work.ProbeKind,
                }
            );

        /// <summary>Drives the job the way the dispatch flow does, until it stops.</summary>
        public string Drive(string key)
        {
            for (int claim = 0; claim < 10; claim++)
            {
                string run = "run-" + claim;
                var work = f.Execute(
                    new WorkerRequest
                    {
                        Command = "Claim",
                        Key = key,
                        RunId = run,
                    }
                );
                for (int step = 0; step < 40; step++)
                {
                    if (work.Status == "Read" || work.Status == "Create")
                    {
                        var (status, body) = Answer(work);
                        work = f.Execute(
                            new WorkerRequest
                            {
                                Command = work.Status == "Read" ? "Observe" : "CreateResponse",
                                Key = key,
                                RunId = run,
                                Token = work.Token,
                                ProbeId = work.ProbeId,
                                ProbeKind = work.ProbeKind,
                                HttpStatus = status,
                                ResponseBody = body,
                            }
                        );
                    }
                    else if (work.Status == "ReadyToCreate")
                        work = Step("PrepareCreate", key, work, run);
                    else if (work.Status == "Verified")
                        work = Step("Complete", key, work, run);
                    else
                        break;
                }
                if (work.Status != "Pending")
                    return work.Status;
            }
            throw new InvalidOperationException("The job did not finish.");
        }
    }
}
