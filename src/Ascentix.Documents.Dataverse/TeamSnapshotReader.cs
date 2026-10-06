using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

/// <summary>
/// Decides how a Dataverse team gets library access. An owner team's members are synced into
/// the Documents group one by one. An Entra security group team or a Microsoft 365 group team is
/// represented by its group: the Documents group holds the group's SharePoint claim as its only
/// managed member, so new group members get access with no Documents run and Dataverse team
/// members are never read. Dataverse fills a group team's member list only when each person first
/// signs in, so it could not be copied reliably. Where the claim reaches more people than the
/// team holds, the admin must accept that when configuring access (<see cref="BroaderAccess"/>).
/// </summary>
public static class TeamPrincipal
{
    // team.teamtype and team.membershiptype choices:
    // https://learn.microsoft.com/power-apps/developer/data-platform/reference/entities/team
    // (TeamType 0 Owner, 1 Access, 2 Security Group, 3 Office Group; MembershipType 0 Members
    // and guests, 1 Members, 2 Owners, 3 Guests). How each membership type maps to the Entra
    // group: https://learn.microsoft.com/power-platform/admin/manage-group-teams
    private const int OwnerTeam = 0,
        SecurityGroupTeam = 2,
        Microsoft365GroupTeam = 3;
    private const int MembersOnly = 1,
        OwnersOnly = 2,
        GuestsOnly = 3;

    public static ColumnSet Columns() =>
        new ColumnSet(
            "name",
            "teamtype",
            "isdefault",
            "azureactivedirectoryobjectid",
            "membershiptype"
        );

    public static bool IsGroup(Entity team)
    {
        int? type = team.GetAttributeValue<OptionSetValue>("teamtype")?.Value;
        return type == SecurityGroupTeam || type == Microsoft365GroupTeam;
    }

    public const string GuestsAlsoWarning =
        "The group's guests will also have access to this library.";
    public const string AllMembersWarning =
        "All members of the group will have access to this library, not only its owners.";

    /// <summary>
    /// Refuses a team that cannot get library access, and a team whose group reaches more people
    /// than the team unless the admin accepted that.
    /// </summary>
    /// <param name="team">The team row, read with <see cref="Columns"/>.</param>
    /// <param name="acknowledged">The request's AcknowledgeBroaderAccess.</param>
    public static void Validate(Entity team, bool acknowledged)
    {
        RequireEligible(team);
        string? warning = BroaderAccess(team);
        if (warning != null && !acknowledged)
            throw new EvaluationBlockedException(warning);
    }

    /// <summary>Refuses a team that cannot get library access, with the reason.</summary>
    public static void RequireEligible(Entity team)
    {
        if (!Eligible(team, out var reason))
            throw new EvaluationBlockedException(reason!);
    }

    /// <summary>
    /// The warning an admin must accept when the team's group claim gives access to more people
    /// than the Dataverse team holds; null when it does not.
    /// - Members: SharePoint has no claim for a group's members without its guests, so the plain
    ///   group claim also admits the guests.
    /// - Security group Owners: SharePoint has no owners claim for a security group, so the whole
    ///   group is granted.
    /// </summary>
    public static string? BroaderAccess(Entity team)
    {
        if (!IsGroup(team))
            return null;
        if (Membership(team) == MembersOnly)
            return GuestsAlsoWarning;
        if (
            team.GetAttributeValue<OptionSetValue>("teamtype").Value == SecurityGroupTeam
            && Membership(team) == OwnersOnly
        )
            return AllMembersWarning;
        return null;
    }

    /// <summary>
    /// Whether the team can get library access; when it cannot, the reason an admin can act on.
    /// </summary>
    public static bool Eligible(Entity team, out string? reason)
    {
        int? type = team.GetAttributeValue<OptionSetValue>("teamtype")?.Value;
        string name = Name(team);
        reason = null;
        if (team.GetAttributeValue<bool>("isdefault"))
            reason =
                "Team '"
                + name
                + "' is a business unit's default team, which cannot get library access. Choose another team.";
        else if (type != OwnerTeam && type != SecurityGroupTeam && type != Microsoft365GroupTeam)
            reason =
                "Access teams cannot get library access. Choose an owner team, a Microsoft Entra security group team or a Microsoft 365 group team.";
        else if (type == OwnerTeam)
            return true;
        else if (team.GetAttributeValue<Guid>("azureactivedirectoryobjectid") == Guid.Empty)
            reason =
                "Team '"
                + name
                + "' has no Microsoft Entra group object ID, so SharePoint cannot identify its group. Create the group team in Dataverse with its group selected.";
        else if (Membership(team) == GuestsOnly)
            reason =
                "Team '"
                + name
                + "' includes only the guests of its group. SharePoint has no sign-in claim for only the guests of a group, so this team cannot get library access. Use a team whose membership type is Members, or Members and guests.";
        else
            return true;
        return false;
    }

    /// <summary>
    /// The SharePoint claim that stands for an eligible group team's group. SharePoint resolves it
    /// with _api/web/ensureuser before it is added to the Documents group.
    /// </summary>
    /// <remarks>
    /// Claim strings follow SharePoint's identity claim encoding
    /// (&lt;IdentityClaim&gt;:0&lt;ClaimType&gt;&lt;ClaimValueType&gt;&lt;AuthMode&gt;|&lt;OriginalIssuer&gt;|&lt;ClaimValue&gt;,
    /// https://learn.microsoft.com/previous-versions/office/developer/sharepoint-2010/gg481769(v=office.14)#appendix-a-identity-claim-encoding-characters-map).
    /// SharePoint Online's group claims are not in a reference page; they are given by Microsoft
    /// staff on Microsoft Q&amp;A:
    /// - Entra security group: c:0t.c|tenant|&lt;object ID&gt;
    ///   (https://learn.microsoft.com/answers/questions/412753, "This is the claim type for the group").
    /// - Microsoft 365 group members: c:0o.c|federateddirectoryclaimprovider|&lt;group ID&gt;
    ///   (https://learn.microsoft.com/answers/questions/349797 and /873831).
    /// - Microsoft 365 group owners: the same claim with the suffix _o. No Microsoft page states
    ///   it; community documentation of group-connected sites does, so it is verified live.
    /// Members and Members and guests share the plain group claim, and a security group's Owners
    /// team uses the security group claim, since SharePoint has no narrower claim for either; the
    /// admin accepts that when configuring access (<see cref="BroaderAccess"/>).
    /// </remarks>
    public static string Claim(Entity team)
    {
        RequireEligible(team);
        string id = team.GetAttributeValue<Guid>("azureactivedirectoryobjectid").ToString("D");
        if (team.GetAttributeValue<OptionSetValue>("teamtype").Value == SecurityGroupTeam)
            return "c:0t.c|tenant|" + id;
        return "c:0o.c|federateddirectoryclaimprovider|"
            + id
            + (Membership(team) == OwnersOnly ? "_o" : "");
    }

    private static int Membership(Entity team) =>
        team.GetAttributeValue<OptionSetValue>("membershiptype")?.Value ?? 0;

    internal static string Name(Entity team) =>
        TeamSnapshotReader.Clean(team.GetAttributeValue<string>("name"), team.Id.ToString("D"));
}

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
    /// <param name="teamId">The opted-in Dataverse team.</param>
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
        var team = service.Retrieve("team", teamId, TeamPrincipal.Columns());
        if (TeamPrincipal.IsGroup(team))
            return GroupSnapshot(team);
        TeamPrincipal.RequireEligible(team);
        string teamName = TeamPrincipal.Name(team);
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
        int used = 0;
        bool incomplete = false;
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
                var person = new TeamPerson
                {
                    UserId = row.Id,
                    EntraId = oid,
                    Login = login,
                };
                string? notice = reason == null ? null : Skip(who, reason);
                used += notice == null ? Footprint(person) : JsonWire.Write(notice).Length + 1;
                if (used > PeopleBudget)
                {
                    incomplete = true;
                    break;
                }
                if (notice != null)
                {
                    skipped.Add(notice);
                    continue;
                }
                names[row.Id] = who;
                people.Add(person);
            }
            if (incomplete || !page.MoreRecords)
                break;
            if (
                page.Entities.Count == 0
                || string.IsNullOrEmpty(page.PagingCookie)
                || !cookies.Add(page.PagingCookie)
            )
                throw new EvaluationBlockedException("Incomplete or cyclic team pagination.");
            query.PageInfo.PageNumber++;
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
        if (incomplete)
            skipped.Add(
                "Team '"
                    + teamName
                    + "' has more people than one access run can store (each person is kept twice in one 500,000-character Dataverse row), so Documents left the members of its SharePoint group as they are. Its access to the library is still applied, and every scheduled refresh tries again. A Microsoft Entra or Microsoft 365 group team is granted as one group and has no such limit."
            );
        return new TeamSnapshot
        {
            People = people
                .Where(p => !sameEntra.Contains(p.EntraId) && !sameLogin.Contains(p.Login))
                .OrderBy(p => p.UserId)
                .ToArray(),
            Skipped = skipped.ToArray(),
            Incomplete = incomplete,
        };
    }

    // asx_membership.asx_payload holds at most 500,000 characters (MaxLength in
    // Entities/asx_membership/Entity.xml, also enforced by JsonWire.Read). A team's row keeps
    // each person twice: as the team has them (Desired) and as SharePoint lists them in the
    // group (Observed). People are read while both copies, the skip notices and the row's own
    // keys and status (a few hundred characters; 2,000 are kept free) fit.
    private const int MembershipPayloadMaxLength = 500000;
    private const int PeopleBudget = MembershipPayloadMaxLength - 2000;

    /// <summary>The characters one person takes in the membership row: wanted and observed.</summary>
    private static int Footprint(TeamPerson person) =>
        JsonWire.Write(person).Length
        + JsonWire
            .Write(
                new SitePerson
                {
                    Id = int.MaxValue,
                    Login = person.Login,
                    Type = 1,
                }
            )
            .Length
        + 2;

    /// <summary>
    /// A group team's only managed member is its group. Dataverse team members are never read.
    /// A group team that changed after registration so SharePoint can no longer identify its
    /// group gets no member, with the reason as a notice, rather than stopping the run.
    /// </summary>
    private static TeamSnapshot GroupSnapshot(Entity team)
    {
        if (!TeamPrincipal.Eligible(team, out var reason))
            return new TeamSnapshot { Skipped = new[] { reason! } };
        return new TeamSnapshot
        {
            People = new[]
            {
                new TeamPerson
                {
                    EntraId = team.GetAttributeValue<Guid>("azureactivedirectoryobjectid"),
                    Login = TeamPrincipal.Claim(team),
                    Group = true,
                },
            },
        };
    }

    // An email-shaped UPN. It also matches B2B guests (name_domain#EXT#@tenant). '|' is left out
    // because it separates the parts of a SharePoint claims login.
    private const string Upn = @"\A[a-zA-Z0-9.!#$%&'*+/=?^_`{}~-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\z";

    internal static string Clean(string? value, string fallback)
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
        if (next.Length > 8000)
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
