using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Conditions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Ascentix.Documents.Dataverse;

public sealed class NativeLocations
{
    private readonly IOrganizationService service;

    public NativeLocations(IOrganizationService service)
    {
        this.service = service;
    }

    public Guid EnsureLibrary(
        Guid siteId,
        Guid webId,
        Guid listId,
        Guid nativeSite,
        string webUrl,
        string rootPath,
        string name
    )
    {
        new SharePointTarget(webUrl, webId, listId, rootPath);
        string relative = rootPath.Substring(
            Uri.UnescapeDataString(new Uri(webUrl).AbsolutePath).TrimEnd('/').Length + 1
        );
        var id = DocumentStore.StableId(
            "native-library:" + siteId.ToString("N") + ":" + listId.ToString("N")
        );
        var query = new QueryExpression("sharepointdocumentlocation")
        {
            ColumnSet = new ColumnSet(
                "relativeurl",
                "parentsiteorlocation",
                "statecode",
                "servicetype"
            ),
            TopCount = 2,
        };
        query.Criteria.AddCondition("parentsiteorlocation", ConditionOperator.Equal, nativeSite);
        query.Criteria.AddCondition("relativeurl", ConditionOperator.Equal, relative);
        var rows = service.RetrieveMultiple(query);
        if (rows.Entities.Count == 1)
            id = rows.Entities[0].Id;
        if (rows.Entities.Count == 0)
            service.Create(
                new Entity("sharepointdocumentlocation", id)
                {
                    ["name"] = name,
                    ["relativeurl"] = relative,
                    ["parentsiteorlocation"] = new EntityReference("sharepointsite", nativeSite),
                    ["servicetype"] = new OptionSetValue(0),
                }
            );
        else if (
            rows.MoreRecords
            || rows.Entities.Count != 1
            || rows.Entities[0].GetAttributeValue<string>("relativeurl") != relative
            || rows.Entities[0].GetAttributeValue<EntityReference>("parentsiteorlocation")?.Id
                != nativeSite
            || rows.Entities[0].GetAttributeValue<OptionSetValue>("statecode")?.Value != 0
            || rows.Entities[0].GetAttributeValue<OptionSetValue>("servicetype")?.Value != 0
        )
            throw new EvaluationBlockedException(
                "Native library navigation differs; no reassignment."
            );
        return id;
    }

    public Entity? Find(FolderStep folder, Guid nativeParentId, string entryUrl, Guid nativeSiteId)
    {
        ValidateParent(nativeParentId, entryUrl, nativeSiteId);
        var id = DocumentStore.StableId("location:" + folder.Key);
        var query = new QueryExpression("sharepointdocumentlocation")
        {
            ColumnSet = new ColumnSet(
                "relativeurl",
                "parentsiteorlocation",
                "regardingobjectid",
                "description",
                "statecode"
            ),
            TopCount = 2,
        };
        query.Criteria.AddCondition("sharepointdocumentlocationid", ConditionOperator.Equal, id);
        var rows = service.RetrieveMultiple(query);
        if (rows.MoreRecords || rows.Entities.Count > 1)
            throw new EvaluationBlockedException("Ambiguous record entry location.");
        if (rows.Entities.Count == 0)
            return null;
        var row = rows.Entities[0];
        var parent = row.GetAttributeValue<EntityReference>("parentsiteorlocation");
        var regarding = row.GetAttributeValue<EntityReference>("regardingobjectid");
        if (
            parent?.Id != nativeParentId
            || parent.LogicalName != "sharepointdocumentlocation"
            || regarding?.Id != folder.RecordId
            || regarding.LogicalName != folder.Table
            || row.GetAttributeValue<OptionSetValue>("statecode")?.Value != 0
            || row.GetAttributeValue<string>("description")
                != "AscentixDocuments:" + DocumentStore.Hash(folder.Key)
        )
            throw new EvaluationBlockedException(
                "Existing record entry differs from this destination. No navigation was redirected."
            );
        Domain.FolderNames.Validate(row.GetAttributeValue<string>("relativeurl"));
        return row;
    }

    /// <summary>
    /// Creates or verifies the native business-record location for a confirmed root folder.
    /// </summary>
    /// <param name="binding">The confirmed folder step and its business-record identity.</param>
    /// <param name="nativeParentId">The native document location for the library entry.</param>
    /// <param name="approvedEntryUrl">The approved absolute URL of the library entry.</param>
    /// <param name="nativeSiteId">The approved native SharePoint site registration.</param>
    /// <returns>The native location ID, or an empty GUID for a non-root folder step.</returns>
    public Guid Complete(
        FolderStep binding,
        Guid nativeParentId,
        string approvedEntryUrl,
        Guid nativeSiteId
    )
    {
        if (binding.ParentBinding != null)
            return Guid.Empty;
        if (binding.PhysicalId == Guid.Empty || string.IsNullOrEmpty(binding.PhysicalPath))
            throw new EvaluationBlockedException(
                "Physical folder observation required before navigation."
            );
        ValidateParent(nativeParentId, approvedEntryUrl, nativeSiteId);
        var id = DocumentStore.StableId("location:" + binding.Key);
        var query = new QueryExpression("sharepointdocumentlocation")
        {
            ColumnSet = new ColumnSet(
                "relativeurl",
                "parentsiteorlocation",
                "regardingobjectid",
                "description",
                "statecode"
            ),
            TopCount = 2,
        };
        query.Criteria.AddCondition("sharepointdocumentlocationid", ConditionOperator.Equal, id);
        var rows = service.RetrieveMultiple(query).Entities;
        var marker = "AscentixDocuments:" + DocumentStore.Hash(binding.Key);
        if (rows.Count > 1)
            throw new EvaluationBlockedException("Ambiguous native location.");
        if (rows.Count == 1)
        {
            var row = rows[0];
            var parent = row.GetAttributeValue<EntityReference>("parentsiteorlocation");
            var regarding = row.GetAttributeValue<EntityReference>("regardingobjectid");
            if (
                row.GetAttributeValue<string>("description") != marker
                || row.GetAttributeValue<string>("relativeurl") != binding.Candidate
                || parent?.LogicalName != "sharepointdocumentlocation"
                || parent.Id != nativeParentId
                || regarding?.LogicalName != binding.Table
                || regarding.Id != binding.RecordId
                || row.GetAttributeValue<OptionSetValue>("statecode")?.Value != 0
            )
                throw new EvaluationBlockedException(
                    "Managed native location differs; refusing reassignment."
                );
            return row.Id;
        }
        service.Create(
            new Entity("sharepointdocumentlocation", id)
            {
                ["name"] = "Documents " + binding.Section,
                ["description"] = marker,
                ["relativeurl"] = binding.Candidate,
                ["parentsiteorlocation"] = new EntityReference(
                    "sharepointdocumentlocation",
                    nativeParentId
                ),
                ["regardingobjectid"] = new EntityReference(binding.Table, binding.RecordId),
                ["servicetype"] = new OptionSetValue(0),
            }
        );
        return id;
    }

    /// <summary>
    /// Re-points a library's own document location chain after the library or its entry folder
    /// moved within the site. When the chain has as many path segments as the new path, each
    /// row takes its share of the new path, which keeps shared parent rows correct for everyone.
    /// Otherwise the library's own row is placed directly under the site with the full path.
    /// Record locations are relative to this row, so they follow it.
    /// </summary>
    /// <returns>The row updates to write, none when the chain already resolves to the entry.</returns>
    /// <remarks>
    /// Only reads and validates; Write applies the result, so a re-point validates every row
    /// before it writes any.
    /// </remarks>
    public Entity[] Follow(Guid locationId, Guid nativeSiteId, string webUrl, string entryPath)
    {
        string webPath = Uri.UnescapeDataString(new Uri(webUrl).AbsolutePath).TrimEnd('/');
        if (!entryPath.StartsWith(webPath + "/", StringComparison.Ordinal))
            throw new EvaluationBlockedException("Entry path is outside the site.");
        string expected = webUrl.TrimEnd('/') + "/" + entryPath.Substring(webPath.Length + 1);
        if (
            string.Equals(
                Uri.UnescapeDataString(ResolveParent(locationId, nativeSiteId)).TrimEnd('/'),
                Uri.UnescapeDataString(expected).TrimEnd('/'),
                StringComparison.OrdinalIgnoreCase
            )
        )
            return Array.Empty<Entity>();
        var parts = entryPath.Substring(webPath.Length + 1).Split('/');
        var chain = new List<Entity>();
        EntityReference current = new EntityReference("sharepointdocumentlocation", locationId);
        while (current.LogicalName == "sharepointdocumentlocation")
        {
            if (chain.Count >= 20)
                throw new EvaluationBlockedException("Native location ancestor cycle/depth.");
            var row = service.Retrieve(
                current.LogicalName,
                current.Id,
                new ColumnSet("relativeurl", "parentsiteorlocation")
            );
            chain.Insert(0, row);
            current =
                row.GetAttributeValue<EntityReference>("parentsiteorlocation")
                ?? throw new EvaluationBlockedException("Native parent chain is incomplete.");
        }
        if (current.Id != nativeSiteId)
            throw new EvaluationBlockedException("Native site differs from approval.");
        int[] counts = chain
            .Select(r => (r.GetAttributeValue<string>("relativeurl") ?? "").Split('/').Length)
            .ToArray();
        var updates = new List<Entity>();
        if (counts.Sum() == parts.Length)
        {
            int next = 0;
            for (int i = 0; i < chain.Count; i++)
            {
                string relative = string.Join("/", parts.Skip(next).Take(counts[i]));
                next += counts[i];
                if (relative != chain[i].GetAttributeValue<string>("relativeurl"))
                    updates.Add(Update(chain[i], relative, null));
            }
        }
        else
            updates.Add(
                Update(
                    chain[chain.Count - 1],
                    string.Join("/", parts),
                    new EntityReference("sharepointsite", nativeSiteId)
                )
            );
        return updates.ToArray();
    }

    /// <summary>Writes row updates prepared by Follow, each against the version it read.</summary>
    public void Write(IEnumerable<Entity> updates)
    {
        foreach (var target in updates)
            service.Execute(
                new UpdateRequest
                {
                    Target = target,
                    ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches,
                }
            );
    }

    private static Entity Update(Entity row, string relative, EntityReference? parent)
    {
        foreach (var segment in relative.Split('/'))
            Domain.FolderNames.Validate(segment);
        var target = new Entity("sharepointdocumentlocation", row.Id)
        {
            RowVersion = row.RowVersion,
            ["relativeurl"] = relative,
        };
        if (parent != null)
            target["parentsiteorlocation"] = parent;
        return target;
    }

    public void ValidateParent(Guid locationId, string expectedEntryUrl, Guid approvedNativeSiteId)
    {
        if (
            !string.Equals(
                Uri.UnescapeDataString(ResolveParent(locationId, approvedNativeSiteId))
                    .TrimEnd('/'),
                Uri.UnescapeDataString(expectedEntryUrl).TrimEnd('/'),
                StringComparison.OrdinalIgnoreCase
            )
        )
            throw new EvaluationBlockedException(
                "Native ancestor URL differs from approved entry."
            );
    }

    /// <summary>
    /// Resolves a bounded native location chain to its approved SharePoint site and relative path.
    /// </summary>
    /// <param name="locationId">The native document location at the end of the chain.</param>
    /// <param name="approvedNativeSiteId">The site registration that must terminate the chain.</param>
    /// <returns>The absolute URL resolved from the site and document-location path segments.</returns>
    public string ResolveParent(Guid locationId, Guid approvedNativeSiteId)
    {
        var visited = new HashSet<Guid>();
        var segments = new List<string>();
        EntityReference current = new EntityReference("sharepointdocumentlocation", locationId);
        while (current.LogicalName == "sharepointdocumentlocation")
        {
            if (!visited.Add(current.Id) || visited.Count > 20)
                throw new EvaluationBlockedException("Native location ancestor cycle/depth.");
            var row = service.Retrieve(
                current.LogicalName,
                current.Id,
                new ColumnSet("relativeurl", "parentsiteorlocation", "statecode", "servicetype")
            );
            if (
                row.GetAttributeValue<OptionSetValue>("statecode")?.Value != 0
                || row.GetAttributeValue<OptionSetValue>("servicetype")?.Value != 0
            )
                throw new EvaluationBlockedException(
                    "Native parent must be active SharePoint navigation."
                );
            var relative = row.GetAttributeValue<string>("relativeurl");
            if (
                string.IsNullOrWhiteSpace(relative)
                || relative.StartsWith("/", StringComparison.Ordinal)
                || relative.Contains("\\")
                || relative.Contains(":")
                || relative.Contains("?")
                || relative.Contains("#")
            )
                throw new EvaluationBlockedException("Invalid native relative URL.");
            foreach (var segment in relative.Split('/'))
                Domain.FolderNames.Validate(segment);
            segments.Insert(0, relative);
            current =
                row.GetAttributeValue<EntityReference>("parentsiteorlocation")
                ?? throw new EvaluationBlockedException("Native parent chain is incomplete.");
        }
        if (current.LogicalName != "sharepointsite" || current.Id != approvedNativeSiteId)
            throw new EvaluationBlockedException("Native site differs from approval.");
        var site = service.Retrieve(
            "sharepointsite",
            current.Id,
            new ColumnSet("absoluteurl", "statecode")
        );
        if (site.GetAttributeValue<OptionSetValue>("statecode")?.Value != 0)
            throw new EvaluationBlockedException("Native site is inactive.");
        var absolute =
            (site.GetAttributeValue<string>("absoluteurl") ?? "").TrimEnd('/')
            + "/"
            + string.Join("/", segments);
        if (
            !Uri.TryCreate(absolute, UriKind.Absolute, out var actual)
            || actual.Scheme != "https"
            || actual.UserInfo.Length != 0
            || actual.Query.Length != 0
            || actual.Fragment.Length != 0
        )
            throw new EvaluationBlockedException("Native ancestor URL is invalid.");
        return actual.AbsoluteUri;
    }
}
