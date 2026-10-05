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
}

public sealed class TeamSnapshot
{
    public TeamPerson[] People { get; set; } = Array.Empty<TeamPerson>();

    public string[] Skipped { get; set; } = Array.Empty<string>();
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
