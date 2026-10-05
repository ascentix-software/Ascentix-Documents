using System;
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
    public string EntryUrl { get; set; } = "";
}

public sealed class WorkerCatalog
{
    private readonly IOrganizationService service;

    public WorkerCatalog(IOrganizationService service)
    {
        this.service = service;
    }

    /// <summary>
    /// Reads where folder work for a library is written. The only access gate is that the site
    /// and library are approved (not suspended); folders inherit the library's permissions, so
    /// the library's policy-update state and ACL are not consulted.
    /// </summary>
    /// <param name="id">The approved library row.</param>
    /// <param name="requireApproved">
    /// False only for a write already sent to SharePoint, which may finish and record its result
    /// after the site or library is suspended. Nothing new starts without approval.
    /// </param>
    /// <returns>The library's SharePoint target and native identities.</returns>
    public WorkerLibrary Read(Guid id, bool requireApproved = true)
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
                "asx_approved"
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
            requireApproved
            && (
                !row.GetAttributeValue<bool>("asx_approved")
                || !site.GetAttributeValue<bool>("asx_approved")
            )
        )
            throw new EvaluationBlockedException("Site or library is suspended.");
        var entry = new Uri(TemplateStore.Text(row, "asx_entryurl"));
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
            EntryUrl = entry.AbsoluteUri,
            NativeParentId =
                row.GetAttributeValue<EntityReference>("asx_nativeparentid")?.Id ?? Guid.Empty,
            NativeSiteId =
                site.GetAttributeValue<EntityReference>("asx_nativeid")?.Id ?? Guid.Empty,
        };
        if (
            result.EntryId == Guid.Empty
            || result.NativeParentId == Guid.Empty
            || result.NativeSiteId == Guid.Empty
        )
            throw new EvaluationBlockedException(
                "Approved physical/native identities are incomplete."
            );
        return result;
    }
}
