using System;
using System.Runtime.Serialization;

namespace Ascentix.Documents.Dataverse;

[DataContract]
public sealed class TeamRegistration : StoredDocument
{
    [DataMember]
    public Guid TeamId { get; set; }

    [DataMember]
    public bool Enabled { get; set; }

    /// <summary>
    /// True for an Entra or Microsoft 365 group team, which is granted through its group, so a
    /// person signing in creates no membership work. teamtype is fixed when a team is created.
    /// Null for registrations written before this was stored, until their next registration or
    /// Apply; those keep queuing membership events, which then change nothing.
    /// </summary>
    [DataMember]
    public bool? Group { get; set; }

    /// <summary>
    /// The team's name when Documents last read it, so a team deleted in Dataverse is still
    /// shown by name. Null for registrations written before this was stored, until the team's
    /// next access run.
    /// </summary>
    [DataMember]
    public string? Name { get; set; }
}

[DataContract]
public sealed class TeamPerson
{
    [DataMember]
    public Guid UserId { get; set; }

    [DataMember]
    public Guid EntraId { get; set; }

    [DataMember]
    public string Login { get; set; } = "";

    /// <summary>
    /// True for the group claim of an Entra or Microsoft 365 group team, which SharePoint resolves
    /// with ensureuser before it is added. Missing in documents written before group teams.
    /// </summary>
    [DataMember]
    public bool Group { get; set; }
}

public sealed class TeamSnapshot
{
    public TeamPerson[] People { get; set; } = Array.Empty<TeamPerson>();

    public string[] Skipped { get; set; } = Array.Empty<string>();

    /// <summary>
    /// True when the team has more people than one access run can store; People then holds
    /// the first ones only, always the same ones while the team is unchanged.
    /// </summary>
    public bool Incomplete { get; set; }
}

[DataContract]
public sealed class PolicyEntry
{
    [DataMember]
    public Guid TeamId { get; set; }

    [DataMember]
    public string Access { get; set; } = "None";
}

[DataContract]
public sealed class PolicyRole
{
    [DataMember]
    public int Id { get; set; }

    [DataMember]
    public string High { get; set; } = "";

    [DataMember]
    public string Low { get; set; } = "";
}

[DataContract]
public sealed class PolicyTeamReference : StoredDocument
{
    [DataMember]
    public Guid TeamId { get; set; }

    [DataMember]
    public string PolicyKey { get; set; } = "";
}

[DataContract]
public sealed class PolicyDocument : StoredDocument
{
    [DataMember]
    public DateTime? NextReviewUtc { get; set; }

    [DataMember]
    public Guid LibraryId { get; set; }

    [DataMember]
    public Guid Generation { get; set; }

    [DataMember]
    public string? OperationKey { get; set; }

    [DataMember]
    public PolicyRole ApprovedReadRole { get; set; } = new PolicyRole();

    [DataMember]
    public PolicyRole ApprovedContributeRole { get; set; } = new PolicyRole();

    [DataMember]
    public Guid[] ManagedTeams { get; set; } = Array.Empty<Guid>();

    [DataMember]
    public PolicyEntry[] Desired { get; set; } = Array.Empty<PolicyEntry>();

    [DataMember]
    public PolicyEntry[] Approved { get; set; } = Array.Empty<PolicyEntry>();

    [DataMember]
    public PolicyEntry[] Queued { get; set; } = Array.Empty<PolicyEntry>();

    [DataMember]
    public PolicyEntry[] Applied { get; set; } = Array.Empty<PolicyEntry>();

    [DataMember]
    public PolicyRole ReadRole { get; set; } = new PolicyRole();

    [DataMember]
    public PolicyRole ContributeRole { get; set; } = new PolicyRole();

    [DataMember]
    public string[] ResidualAccess { get; set; } = Array.Empty<string>();

    /// <summary>What the last access run skipped or reconciled, for admins to review.</summary>
    [DataMember]
    public string[] Notices { get; set; } = Array.Empty<string>();

    /// <summary>
    /// The admin applied Desired while the queued run could not be replaced yet: a flow held it
    /// or SharePoint had not answered its write. The next access review replaces the run and
    /// queues Desired as soon as it can. Rows written before 0.1.0.4 read as false.
    /// </summary>
    [DataMember]
    public bool ApplyPending { get; set; }

    /// <summary>
    /// The last access run applied the library grants but left some team membership unsynced:
    /// members SharePoint refused, or a team too large for one run. Its notices say which.
    /// Rows written before 0.1.0.4 read as false.
    /// </summary>
    [DataMember]
    public bool MembershipIncomplete { get; set; }

    /// <summary>
    /// The admin accepted that Documents stops this library's permission inheritance. The next
    /// queued access run carries it and it is cleared here.
    /// </summary>
    [DataMember]
    public bool BreakInheritance { get; set; }

    /// <summary>
    /// The last access run stopped because the library inherits its site's permissions and no
    /// consent was given. Apply access with BreakInheritance resolves it.
    /// </summary>
    [DataMember]
    public bool Inherits { get; set; }
}

[DataContract]
public sealed class ManagedGroup : StoredDocument
{
    [DataMember]
    public Guid SiteId { get; set; }

    [DataMember]
    public Guid TeamId { get; set; }

    [DataMember]
    public Guid Nonce { get; set; }

    [DataMember]
    public int GroupId { get; set; }

    [DataMember]
    public string MembershipHash { get; set; } = "";

    [DataMember]
    public string Title { get; set; } = "";

    /// <summary>
    /// Group claims Documents put in this group for a group team. A claim no longer wanted is
    /// removed; any other non-person member was added by hand and is left in place. Documents
    /// written before group teams read it as empty (StoredDocument fills missing lists).
    /// </summary>
    [DataMember]
    public string[]? Principals { get; set; }
    public string Marker =>
        "Ascentix Documents v1; group=" + Nonce.ToString("D") + "; team=" + TeamId.ToString("D");
}

[DataContract]
public sealed class ManagedGrant : StoredDocument
{
    [DataMember]
    public Guid LibraryId { get; set; }

    [DataMember]
    public Guid TeamId { get; set; }

    [DataMember]
    public int GroupId { get; set; }

    [DataMember]
    public int RoleId { get; set; }

    [DataMember]
    public Guid Generation { get; set; }
}

[DataContract]
public sealed class MembershipDocument : StoredDocument
{
    [DataMember]
    public string GroupKey { get; set; } = "";

    [DataMember]
    public Guid Generation { get; set; }

    [DataMember]
    public TeamPerson[] Desired { get; set; } = Array.Empty<TeamPerson>();

    [DataMember]
    public SitePerson[] Observed { get; set; } = Array.Empty<SitePerson>();

    [DataMember]
    public string[] Skipped { get; set; } = Array.Empty<string>();

    [DataMember]
    public bool Complete { get; set; }

    /// <summary>
    /// The team had more people than one access run can store, so its group's members are
    /// left as they are this run. Rows written before 0.1.0.4 read as false.
    /// </summary>
    [DataMember]
    public bool Incomplete { get; set; }
}

[DataContract]
public sealed class SecurityRequest
{
    [DataMember]
    public string Command { get; set; } = "";

    [DataMember]
    public Guid LibraryId { get; set; }

    [DataMember]
    public Guid TeamId { get; set; }

    [DataMember]
    public bool Enabled { get; set; }

    [DataMember]
    public string? RowVersion { get; set; }

    [DataMember]
    public PolicyEntry[] Entries { get; set; } = Array.Empty<PolicyEntry>();

    [DataMember]
    public PolicyRole ReadRole { get; set; } = new PolicyRole();

    [DataMember]
    public PolicyRole ContributeRole { get; set; } = new PolicyRole();

    /// <summary>
    /// The admin saw and accepted that a group team gives access to more people than the
    /// Dataverse team holds (see TeamPrincipal.BroaderAccess).
    /// </summary>
    [DataMember]
    public bool AcknowledgeBroaderAccess { get; set; }

    /// <summary>
    /// ApplyPolicy: the admin saw and accepted that Documents stops the library's permission
    /// inheritance, keeping a copy of the site's permissions.
    /// </summary>
    [DataMember]
    public bool BreakInheritance { get; set; }

    /// <summary>
    /// RetryAccessRun and CancelAccessRun: the access run the admin saw on the library, so a
    /// newer run queued meanwhile is never retried or cancelled by mistake.
    /// </summary>
    [DataMember]
    public string? OperationKey { get; set; }
}

[DataContract]
public sealed class SecurityResult
{
    [DataMember]
    public string Status { get; set; } = "";

    [DataMember]
    public string RowVersion { get; set; } = "";

    [DataMember]
    public PolicyDocument? Policy { get; set; }

    [DataMember]
    public string[] Diff { get; set; } = Array.Empty<string>();

    /// <summary>The status of the library's queued access run, such as Blocked or RetryWait.</summary>
    [DataMember]
    public string? RunStatus { get; set; }

    /// <summary>What the queued access run reports first: why it stopped or why it waits.</summary>
    [DataMember]
    public string? RunNotice { get; set; }

    /// <summary>When a waiting access run checks again (UTC).</summary>
    [DataMember]
    public DateTime? RunNextAttemptUtc { get; set; }

    /// <summary>
    /// The teams of the policy's wanted and applied access, with their names, so a team deleted
    /// in Dataverse is shown as one instead of being read by ID.
    /// </summary>
    [DataMember]
    public PolicyTeam[] Teams { get; set; } = Array.Empty<PolicyTeam>();
}

/// <summary>A team of a library's access, as the library shows it.</summary>
[DataContract]
public sealed class PolicyTeam
{
    [DataMember]
    public Guid TeamId { get; set; }

    /// <summary>The team's name, or its last known name once deleted; null when none is known.</summary>
    [DataMember]
    public string? Name { get; set; }

    /// <summary>True when the team was deleted in Dataverse.</summary>
    [DataMember]
    public bool Deleted { get; set; }
}

[DataContract]
public sealed class SecurityOperation : OperationDocument
{
    [DataMember]
    public string? LibrarySetupKey { get; set; }

    [DataMember]
    public Guid LibraryId { get; set; }

    [DataMember]
    public Guid SiteId { get; set; }

    [DataMember]
    public string PolicyKey { get; set; } = "";

    [DataMember]
    public PolicyEntry[] Entries { get; set; } = Array.Empty<PolicyEntry>();

    [DataMember]
    public PolicyRole ReadRole { get; set; } = new PolicyRole();

    [DataMember]
    public PolicyRole ContributeRole { get; set; } = new PolicyRole();

    [DataMember]
    public int MutationMemberId { get; set; }

    [DataMember]
    public string MutationLogin { get; set; } = "";

    [DataMember]
    public Guid WebId { get; set; }

    [DataMember]
    public Guid ListId { get; set; }

    [DataMember]
    public string WebUrl { get; set; } = "";

    [DataMember]
    public int Steps { get; set; }

    [DataMember]
    public int TeamIndex { get; set; }

    [DataMember]
    public string GroupKey { get; set; } = "";

    [DataMember]
    public string MembershipKey { get; set; } = "";

    [DataMember]
    public string PageEndpoint { get; set; } = "";

    [DataMember]
    public string[] PageLinks { get; set; } = Array.Empty<string>();

    [DataMember]
    public AclAssignment[] Acl { get; set; } = Array.Empty<AclAssignment>();

    [DataMember]
    public SitePerson[] Members { get; set; } = Array.Empty<SitePerson>();

    [DataMember]
    public HttpIntent? Mutation { get; set; }

    [DataMember]
    public string MutationKind { get; set; } = "";

    [DataMember]
    public int MutationRole { get; set; }

    /// <summary>What this run skipped or reconciled; copied to the policy when it completes.</summary>
    [DataMember]
    public string[] Notices { get; set; } = Array.Empty<string>();

    /// <summary>Member changes SharePoint rejected in this run, keyed by group, so they are not retried.</summary>
    [DataMember]
    public string[] SkippedMembers { get; set; } = Array.Empty<string>();

    /// <summary>How many member changes were skipped in this run, including any not detailed.</summary>
    [DataMember]
    public int SkippedCount { get; set; }

    /// <summary>Consecutive read-backs that did not yet show a confirmed grant write.</summary>
    [DataMember]
    public int ReadbackMisses { get; set; }

    /// <summary>The group claim SharePoint resolved in this run, so it is added next.</summary>
    [DataMember]
    public string? EnsuredLogin { get; set; }

    /// <summary>
    /// The admin consented to stopping the library's permission inheritance for this run. Without
    /// it an inheriting library blocks the run with a notice.
    /// </summary>
    [DataMember]
    public bool BreakInheritance { get; set; }
}

[DataContract]
public sealed class SitePerson
{
    [DataMember(Name = "Id")]
    public int Id { get; set; }

    [DataMember(Name = "LoginName")]
    public string Login { get; set; } = "";

    [DataMember(Name = "PrincipalType")]
    public int Type { get; set; }
}

[DataContract]
public sealed class SiteGroup
{
    [DataMember(Name = "Id")]
    public int Id { get; set; }

    [DataMember(Name = "Title")]
    public string Title { get; set; } = "";

    [DataMember(Name = "Description")]
    public string Description { get; set; } = "";

    [DataMember(Name = "PrincipalType")]
    public int Type { get; set; }
}

[DataContract]
public sealed class GroupCreateBody
{
    [DataMember(Name = "__metadata")]
    public GroupMetadata Metadata { get; set; } = new GroupMetadata();

    [DataMember(Name = "Title")]
    public string Title { get; set; } = "";

    [DataMember(Name = "Description")]
    public string Description { get; set; } = "";
}

[DataContract]
public sealed class GroupMetadata
{
    [DataMember(Name = "type")]
    public string Type { get; set; } = "SP.Group";
}
