using System;
using System.Linq;
using Ascentix.Documents.Conditions;
using Ascentix.Documents.Dataverse;
using Microsoft.Xrm.Sdk;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class RuntimeScopeTests
{
    private static DurableWorkerTests.MemoryService Service()
    {
        var s = new DurableWorkerTests.MemoryService();
        RuntimeSeed.Seed(s, Guid.NewGuid(), "account");
        return s;
    }

    private static Guid Site(DurableWorkerTests.MemoryService s, string url, bool active = true)
    {
        var id = Guid.NewGuid();
        s.Seed(
            new Entity("sharepointsite", id)
            {
                ["absoluteurl"] = url,
                ["statecode"] = new OptionSetValue(active ? 0 : 1),
            }
        );
        return id;
    }

    private static WorkerResult Request(
        string url,
        string method = "GET",
        string endpoint = "_api/web"
    ) =>
        new WorkerResult
        {
            SiteUrl = url,
            Http = new HttpIntent { Method = method, RelativeUri = endpoint },
        };

    [Fact]
    public void MultipleRegisteredSitesUseOneHostConfigurationAndDeactivationIsRechecked()
    {
        var s = Service();
        string a = "https://example.sharepoint.com/sites/a",
            b = "https://example.sharepoint.com/teams/b";
        Site(s, a);
        var second = Site(s, b + "/");
        var profile = RuntimeProfile.Read(s);
        profile.ValidateTransport(Request(a));
        profile.ValidateTransport(Request(b, "POST", "_api/web/lists"));
        Assert.Single(profile.SharePointHosts);
        s.Rows[second]["statecode"] = new OptionSetValue(1);
        Assert.Throws<EvaluationBlockedException>(() => profile.ValidateTransport(Request(b)));
    }

    [Theory]
    [InlineData("https://example.sharepoint.com/sites/unregistered")]
    [InlineData("https://other.sharepoint.com/sites/registered")]
    [InlineData("https://example.sharepoint.com:444/sites/a")]
    [InlineData("https://example.sharepoint.com/sites/a?query=1")]
    [InlineData("https://example.sharepoint.com/sites/a#fragment")]
    [InlineData("https://user@example.sharepoint.com/sites/a")]
    [InlineData("http://example.sharepoint.com/sites/a")]
    public void HostScopeDoesNotAuthorizeUnregisteredOrInvalidTargets(string url)
    {
        var s = Service();
        Site(s, "https://example.sharepoint.com/sites/a");
        Site(s, "https://other.sharepoint.com/sites/registered");
        Assert.Throws<EvaluationBlockedException>(() =>
            RuntimeProfile.Read(s).ValidateTransport(Request(url))
        );
    }

    [Theory]
    [InlineData("example.sharepoint.com.evil.test")]
    [InlineData("*.sharepoint.com")]
    [InlineData("https://example.sharepoint.com")]
    [InlineData("example.sharepoint.com/sites/a")]
    [InlineData("EXAMPLE.sharepoint.com")]
    [InlineData("-example.sharepoint.com")]
    public void HostConfigurationRequiresCanonicalExactTenantNames(string host) =>
        Assert.Throws<EvaluationBlockedException>(() =>
            RuntimeProfile.ValidateHosts(new[] { host })
        );

    // Every Microsoft cloud's SharePoint Online domain: worldwide and GCC, GCC High, DoD (both
    // of its listed domains) and 21Vianet.
    [Theory]
    [InlineData("contoso.sharepoint.com")]
    [InlineData("contoso-my.sharepoint.com")]
    [InlineData("agency.sharepoint.us")]
    [InlineData("command.sharepoint-mil.us")]
    [InlineData("command.dps.mil")]
    [InlineData("contoso.sharepoint.cn")]
    public void HostConfigurationAcceptsEveryMicrosoftSharePointCloud(string host)
    {
        RuntimeProfile.ValidateHosts(new[] { host });
        var s = Service();
        s.Rows.Values.Single(r => r.LogicalName == "asx_runtime")["asx_sharepointhosts"] =
            "[\"" + host + "\"]";
        string url = "https://" + host + "/sites/a";
        Site(s, url);
        RuntimeProfile.Read(s).ValidateTransport(Request(url));
        // A listed host still needs a registered site.
        Assert.Throws<EvaluationBlockedException>(() =>
            RuntimeProfile.Read(s).ValidateTransport(Request("https://" + host + "/sites/b"))
        );
    }

    [Theory]
    [InlineData("contoso.sharepoint.de")]
    [InlineData("sharepoint.us")]
    [InlineData("contoso.sharepoint.us.evil.test")]
    [InlineData("contoso.mysharepoint.com")]
    [InlineData("contoso.dps.mil.evil.test")]
    [InlineData("evil-sharepoint.com")]
    [InlineData("contoso.sharepoint.com:444")]
    [InlineData("contoso.sharepoint.cоm")]
    [InlineData("xn--contso-6ve.sharepoint.com.evil.test")]
    [InlineData("contоso.sharepoint.com")]
    public void HostConfigurationRefusesOtherDomains(string host) =>
        Assert.Throws<EvaluationBlockedException>(() =>
            RuntimeProfile.ValidateHosts(new[] { host })
        );

    [Fact]
    public void AnyNumberOfHostsIsAcceptedWhileTheyFitTheirColumn()
    {
        RuntimeProfile.ValidateHosts(
            Enumerable.Range(0, 150).Select(i => "tenant" + i + ".sharepoint.com").ToArray()
        );
        // The hosts list is stored as one value of at most 5,000 characters; past that it cannot
        // be saved, and the refusal says so.
        var tooMany = Assert.Throws<EvaluationBlockedException>(() =>
            RuntimeProfile.ValidateHosts(
                Enumerable.Range(0, 400).Select(i => "tenant" + i + ".sharepoint.com").ToArray()
            )
        );
        Assert.Contains("5,000 characters", tooMany.Message);
    }

    [Fact]
    public void DuplicateHostsAndUnsupportedTransportRemainRejected()
    {
        Assert.Throws<EvaluationBlockedException>(() =>
            RuntimeProfile.ValidateHosts(
                new[] { "example.sharepoint.com", "example.sharepoint.com" }
            )
        );
        var s = Service();
        const string url = "https://example.sharepoint.com/sites/a";
        Site(s, url);
        var p = RuntimeProfile.Read(s);
        Assert.Throws<EvaluationBlockedException>(() =>
            p.ValidateTransport(Request(url, "DELETE"))
        );
        Assert.Throws<EvaluationBlockedException>(() =>
            p.ValidateTransport(Request(url, "GET", "https://other.test"))
        );
    }

    [Fact]
    public void IncompleteHostConfigurationCanBeDisplayedButCannotExecute()
    {
        var s = Service();
        foreach (var row in s.Rows.Values)
            if (row.LogicalName == "asx_runtime")
                row.Attributes.Remove("asx_sharepointhosts");
        Assert.Empty(
            RuntimeAdministration.Execute(s, new RuntimeRequest(), true, Guid.Empty).SharePointHosts
        );
        Assert.ThrowsAny<Exception>(() => RuntimeProfile.Read(s));
    }
}
