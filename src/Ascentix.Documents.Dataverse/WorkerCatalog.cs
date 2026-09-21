using System;
using System.Text.RegularExpressions;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public sealed class WorkerLibrary
{
    public Guid Id { get; set; }
    public SharePointTarget Target { get; set; } = null!;
    public Guid EntryId { get; set; }
    public Guid NativeParentId { get; set; }
    public Guid NativeSiteId { get; set; }
    public Guid PolicyRevision { get; set; }
    public string EntryUrl { get; set; } = "";
    public string AclHash { get; set; } = "";
}

public sealed class WorkerCatalog
{
    private readonly IOrganizationService service;

    public WorkerCatalog(IOrganizationService service)
    {
        this.service = service;
    }

    public WorkerLibrary Read(Guid id)
    {
        var row = service.Retrieve(
            "asx_library",
            id,
            new ColumnSet(
                "asx_siteid",
                "asx_listid",
                "asx_entryid",
                "asx_entryurl",
                "asx_nativeparentid",
                "asx_policyrevision",
                "asx_approved",
                "asx_policyapplied",
                "asx_aclhash"
            )
        );
        var siteRef =
            row.GetAttributeValue<EntityReference>("asx_siteid")
            ?? throw new EvaluationBlockedException("Library site approval missing.");
        var site = service.Retrieve(
            "asx_site",
            siteRef.Id,
            new ColumnSet("asx_nativeid", "asx_webid", "asx_url", "asx_approved")
        );
        if (
            !row.GetAttributeValue<bool>("asx_approved")
            || !row.GetAttributeValue<bool>("asx_policyapplied")
            || !site.GetAttributeValue<bool>("asx_approved")
        )
            throw new EvaluationBlockedException("Site/library/policy approval is not active.");
        var entry = new Uri(TemplateStore.Text(row, "asx_entryurl"));
        var aclHash = TemplateStore.Text(row, "asx_aclhash");
        if (!Regex.IsMatch(aclHash, "^[a-f0-9]{64}$"))
            throw new EvaluationBlockedException(
                "An approved complete library ACL fingerprint is required."
            );
        var target = new SharePointTarget(
            TemplateStore.Text(site, "asx_url"),
            Guid.Parse(TemplateStore.Text(site, "asx_webid")),
            Guid.Parse(TemplateStore.Text(row, "asx_listid")),
            Uri.UnescapeDataString(entry.AbsolutePath)
        );
        if (
            entry.GetLeftPart(UriPartial.Authority) != target.Web.GetLeftPart(UriPartial.Authority)
            || entry.Query.Length != 0
            || entry.Fragment.Length != 0
        )
            throw new EvaluationBlockedException("Entry URL is outside the approved origin.");
        var result = new WorkerLibrary
        {
            Id = id,
            Target = target,
            EntryId = Guid.Parse(TemplateStore.Text(row, "asx_entryid")),
            PolicyRevision = Guid.Parse(TemplateStore.Text(row, "asx_policyrevision")),
            EntryUrl = entry.AbsoluteUri,
            AclHash = aclHash,
            NativeParentId =
                row.GetAttributeValue<EntityReference>("asx_nativeparentid")?.Id ?? Guid.Empty,
            NativeSiteId =
                site.GetAttributeValue<EntityReference>("asx_nativeid")?.Id ?? Guid.Empty,
        };
        if (
            result.EntryId == Guid.Empty
            || result.PolicyRevision == Guid.Empty
            || result.NativeParentId == Guid.Empty
            || result.NativeSiteId == Guid.Empty
        )
            throw new EvaluationBlockedException(
                "Approved physical/native identities are incomplete."
            );
        return result;
    }
}
