using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public sealed class TeamSnapshotReader
{
    private readonly IOrganizationService service;

    public TeamSnapshotReader(IOrganizationService service)
    {
        this.service = service;
    }

    public TeamPerson[] Read(Guid teamId)
    {
        var registration = new DocumentStore(service)
            .Require<TeamRegistration>("asx_teamregistration", "team:" + teamId.ToString("N"))
            .Value;
        if (registration.TeamId != teamId)
            throw new EvaluationBlockedException("Team registration identity mismatch.");
        if (!registration.Enabled)
            return Array.Empty<TeamPerson>(); // Explicit opt-out tombstone; a failed native read never means an empty team.
        var team = service.Retrieve("team", teamId, new ColumnSet("teamtype", "isdefault"));
        if (
            team.GetAttributeValue<OptionSetValue>("teamtype")?.Value != 0
            || team.GetAttributeValue<bool>("isdefault")
        )
            throw new EvaluationBlockedException(
                "Only opted-in, non-default manual owner teams are supported."
            );
        var query = new QueryExpression("systemuser")
        {
            ColumnSet = new ColumnSet(
                "domainname",
                "azureactivedirectoryobjectid",
                "isdisabled",
                "applicationid"
            ),
            PageInfo = new PagingInfo { Count = 500, PageNumber = 1 },
        };
        query.AddOrder("systemuserid", OrderType.Ascending);
        query
            .AddLink("teammembership", "systemuserid", "systemuserid")
            .LinkCriteria.AddCondition("teamid", ConditionOperator.Equal, teamId);
        var people = new List<TeamPerson>();
        var ids = new HashSet<Guid>();
        var logins = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var entra = new HashSet<Guid>();
        var cookies = new HashSet<string>();
        while (true)
        {
            var page = service.RetrieveMultiple(query);
            foreach (var row in page.Entities)
            {
                if (!ids.Add(row.Id))
                    throw new EvaluationBlockedException("Duplicate/changing membership page.");
                if (row.GetAttributeValue<bool>("isdisabled"))
                    continue;
                string upn = row.GetAttributeValue<string>("domainname") ?? "";
                Guid oid = row.GetAttributeValue<Guid>("azureactivedirectoryobjectid");
                if (
                    oid == Guid.Empty
                    || row.GetAttributeValue<Guid>("applicationid") != Guid.Empty
                    || upn.IndexOf("#EXT#", StringComparison.OrdinalIgnoreCase) >= 0
                    || !Regex.IsMatch(
                        upn,
                        @"\A[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\z"
                    )
                    || !entra.Add(oid)
                )
                    throw new EvaluationBlockedException(
                        "Unsupported or ambiguous team user identity."
                    );
                string login = "i:0#.f|membership|" + upn.ToLowerInvariant();
                if (!logins.Add(login))
                    throw new EvaluationBlockedException("Multiple users resolve to one login.");
                people.Add(
                    new TeamPerson
                    {
                        UserId = row.Id,
                        EntraId = oid,
                        Login = login,
                    }
                );
            }
            if (ids.Count > 2000)
                throw new EvaluationBlockedException(
                    "Team exceeds the supported complete snapshot bound."
                );
            if (!page.MoreRecords)
                break;
            if (
                page.Entities.Count == 0
                || string.IsNullOrEmpty(page.PagingCookie)
                || !cookies.Add(page.PagingCookie)
                || ++query.PageInfo.PageNumber > 20
            )
                throw new EvaluationBlockedException("Incomplete or cyclic team pagination.");
            query.PageInfo.PagingCookie = page.PagingCookie;
        }
        return people.OrderBy(p => p.UserId).ToArray();
    }

    public static string Hash(TeamPerson[] people)
    {
        string value =
            string.Join(
                "|",
                people
                    .OrderBy(p => p.UserId)
                    .Select(p =>
                        p.UserId.ToString("N") + ":" + p.EntraId.ToString("N") + ":" + p.Login
                    )
            ) + "|complete";
        using (var sha = System.Security.Cryptography.SHA256.Create())
            return BitConverter
                .ToString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(value)))
                .Replace("-", "")
                .ToLowerInvariant();
    }
}

public static class SecurityPaging
{
    public static string Next(Uri web, string first, string next, string[] visited)
    {
        if (visited.Length >= 20 || next.Length > 8000)
            throw new EvaluationBlockedException("Security pagination bound exceeded.");
        var endpoint = new Uri(web.AbsoluteUri.TrimEnd('/') + "/" + first);
        if (
            !Uri.TryCreate(next, UriKind.Absolute, out var candidate)
            || candidate.Scheme != "https"
            || candidate.Authority != endpoint.Authority
            || candidate.AbsolutePath != endpoint.AbsolutePath
            || candidate.UserInfo.Length != 0
            || candidate.Fragment.Length != 0
        )
            throw new EvaluationBlockedException("Security continuation left the issued endpoint.");
        string relative = candidate.AbsoluteUri.Substring(web.AbsoluteUri.TrimEnd('/').Length + 1);
        if (visited.Contains(relative, StringComparer.Ordinal))
            throw new EvaluationBlockedException("Cyclic security continuation.");
        // Preserve all original select/filter/order terms. Only opaque continuation parameters may differ.
        var original = endpoint.Query.TrimStart('?').Split('&').Where(v => v.Length > 0).ToArray();
        var actual = candidate.Query.TrimStart('?').Split('&').Where(v => v.Length > 0).ToArray();
        if (
            original.Any(v => !actual.Contains(v, StringComparer.Ordinal))
            || actual
                .Except(original)
                .Any(v =>
                    !v.StartsWith("$skiptoken=", StringComparison.Ordinal)
                    && !v.StartsWith("%24skiptoken=", StringComparison.Ordinal)
                )
        )
            throw new EvaluationBlockedException(
                "Security continuation changed the query contract."
            );
        return relative;
    }
}
