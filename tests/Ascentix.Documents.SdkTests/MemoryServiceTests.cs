using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class MemoryServiceTests
{
    [Fact]
    public void InConditionsAndPagingBehaveLikeDataverse()
    {
        var s = new DurableWorkerTests.MemoryService();
        var ids = Enumerable.Range(0, 7).Select(_ => Guid.NewGuid()).ToArray();
        for (int i = 0; i < ids.Length; i++)
            s.Seed(new Entity("thing", ids[i]) { ["n"] = i });
        var query = new QueryExpression("thing")
        {
            ColumnSet = new ColumnSet("n"),
            PageInfo = new PagingInfo { Count = 3, PageNumber = 1 },
        };
        query.Criteria.AddCondition(
            "thingid",
            ConditionOperator.In,
            ids.Take(5).Cast<object>().ToArray()
        );
        query.AddOrder("n", OrderType.Ascending);
        var first = s.RetrieveMultiple(query);
        Assert.Equal(new[] { 0, 1, 2 }, first.Entities.Select(e => (int)e["n"]));
        Assert.True(first.MoreRecords);
        Assert.False(string.IsNullOrEmpty(first.PagingCookie));
        query.PageInfo.PageNumber = 2;
        query.PageInfo.PagingCookie = first.PagingCookie;
        var second = s.RetrieveMultiple(query);
        Assert.Equal(new[] { 3, 4 }, second.Entities.Select(e => (int)e["n"]));
        Assert.False(second.MoreRecords);
    }

    [Fact]
    public void RuntimeSeedWritesRowsLegacyJsonAndWorker()
    {
        var s = new DurableWorkerTests.MemoryService();
        var worker = Guid.NewGuid();
        var runtime = RuntimeSeed.Seed(s, worker, "account", "contact");
        Assert.Equal(
            new[] { "account", "contact" },
            s.Rows.Values.Where(r => r.LogicalName == "asx_runtimetable")
                .Select(r => r.GetAttributeValue<string>("asx_logicalname"))
                .OrderBy(n => n)
        );
        Assert.Equal(
            "[\"account\",\"contact\"]",
            s.Rows[runtime].GetAttributeValue<string>("asx_allowedtables")
        );
        Assert.NotEqual(Guid.Empty, s.Rows[worker].GetAttributeValue<Guid>("applicationid"));
    }
}
