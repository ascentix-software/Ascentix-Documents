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

    public TeamPerson[] Read(Guid teamId) => Snapshot(teamId).People;

    /// <summary>
    /// Reads the team's members that SharePoint can take, and lists the ones it skipped and why.
    /// </summary>
    /// <param name="teamId">The opted-in Dataverse owner team.</param>
    /// <returns>The members to sync and one notice for each member that was skipped.</returns>
    public TeamSnapshot Snapshot(Guid teamId)
    {
        var registration = new DocumentStore(service)
            .Require<TeamRegistration>("asx_teamregistration", "team:" + teamId.ToString("N"))
            .Value;
        if (registration.TeamId != teamId)
            throw new EvaluationBlockedException("Team registration identity mismatch.");
        if (!registration.Enabled)
            return new TeamSnapshot(); // Explicit opt-out tombstone; a failed native read never means an empty team.
        var team = service.Retrieve("team", teamId, new ColumnSet("name", "teamtype", "isdefault"));
        if (
            team.GetAttributeValue<OptionSetValue>("teamtype")?.Value != 0
            || team.GetAttributeValue<bool>("isdefault")
        )
            throw new EvaluationBlockedException(
                "Only opted-in, non-default manual owner teams are supported."
            );
        string teamName = Clean(team.GetAttributeValue<string>("name"), teamId.ToString("D"));
        var query = new QueryExpression("systemuser")
        {
            ColumnSet = new ColumnSet(
                "fullname",
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
        var names = new Dictionary<Guid, string>();
        var skipped = new List<string>();
        var ids = new HashSet<Guid>();
        var cookies = new HashSet<string>();
        string Skip(string who, string reason) =>
            "Team '"
            + teamName
            + "': "
            + who
            + " was not added to the library group because "
            + reason
            + ".";
        while (true)
        {
            var page = service.RetrieveMultiple(query);
            foreach (var row in page.Entities)
            {
                if (!ids.Add(row.Id))
                    throw new EvaluationBlockedException("Duplicate/changing membership page.");
                if (row.GetAttributeValue<bool>("isdisabled"))
                    continue;
                // B2B guests are supported: their domainname is the #EXT# UPN that SharePoint
                // uses in their claims login.
                string upn = row.GetAttributeValue<string>("domainname") ?? "";
                Guid oid = row.GetAttributeValue<Guid>("azureactivedirectoryobjectid");
                string login = "i:0#.f|membership|" + upn.ToLowerInvariant();
                string? reason =
                    row.GetAttributeValue<Guid>("applicationid") != Guid.Empty
                        ? "it is an application user"
                    : oid == Guid.Empty ? "it has no Microsoft Entra object ID"
                    : !Regex.IsMatch(upn, Upn) ? "its sign-in name cannot be used in SharePoint"
                    : null;
                string who = Clean(
                    row.GetAttributeValue<string>("fullname"),
                    Clean(upn, "user " + row.Id.ToString("D"))
                );
                if (reason != null)
                {
                    skipped.Add(Skip(who, reason));
                    continue;
                }
                names[row.Id] = who;
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
        // Two members with one Entra object ID or one sign-in name cannot be told apart, so
        // every member in the conflict is skipped rather than guessing which one is right.
        var sameEntra = new HashSet<Guid>(
            people.GroupBy(p => p.EntraId).Where(g => g.Count() > 1).Select(g => g.Key)
        );
        var sameLogin = new HashSet<string>(
            people
                .GroupBy(p => p.Login, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key),
            StringComparer.OrdinalIgnoreCase
        );
        foreach (var person in people.OrderBy(p => p.UserId))
            if (sameEntra.Contains(person.EntraId))
                skipped.Add(
                    Skip(
                        names[person.UserId],
                        "another team member has the same Microsoft Entra object ID"
                    )
                );
            else if (sameLogin.Contains(person.Login))
                skipped.Add(
                    Skip(names[person.UserId], "another team member has the same sign-in name")
                );
        return new TeamSnapshot
        {
            People = people
                .Where(p => !sameEntra.Contains(p.EntraId) && !sameLogin.Contains(p.Login))
                .OrderBy(p => p.UserId)
                .ToArray(),
            Skipped = skipped.ToArray(),
        };
    }

    // An email-shaped UPN. It also matches B2B guests (name_domain#EXT#@tenant). '|' is left out
    // because it separates the parts of a SharePoint claims login.
    private const string Upn = @"\A[a-zA-Z0-9.!#$%&'*+/=?^_`{}~-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\z";

    private static string Clean(string? value, string fallback)
    {
        string text = new string((value ?? "").Where(c => !char.IsControl(c)).ToArray()).Trim();
        if (text.Length == 0)
            return fallback;
        return text.Length > 120 ? text.Substring(0, 120) + "..." : text;
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
