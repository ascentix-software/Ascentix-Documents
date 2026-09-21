using System;
using Ascentix.Documents.Conditions;

namespace Ascentix.Documents.Dataverse;

public static class SiteIdentity
{
    public static string Key(string webUrl, Guid collectionId, Guid webId)
    {
        if (
            !Uri.TryCreate(webUrl, UriKind.Absolute, out var url)
            || url.Scheme != "https"
            || collectionId == Guid.Empty
            || webId == Guid.Empty
        )
            throw new EvaluationBlockedException(
                "Complete SharePoint site collection and web identity required."
            );
        return url.IdnHost.ToLowerInvariant()
            + ","
            + collectionId.ToString("D")
            + ","
            + webId.ToString("D");
    }

    public static Guid LibraryId(Guid siteId, Guid listId)
    {
        if (siteId == Guid.Empty || listId == Guid.Empty)
            throw new EvaluationBlockedException("Site-scoped library identity required.");
        return DocumentStore.StableId(
            "library:" + siteId.ToString("N") + ":" + listId.ToString("N")
        );
    }
}
