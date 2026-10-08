using System;

namespace Ascentix.Documents.Domain;

/// <summary>
/// How a folder path is written into a SharePoint read, and whether that address fits the HTTP
/// with Microsoft Entra ID connector. A path is never put in the URL path, which the connector
/// limits by its ASP.NET maxUrlLength; it goes in the query string as an OData parameter alias
/// ("Using parameter aliases in REST service calls",
/// https://learn.microsoft.com/sharepoint/dev/sp-add-ins/determine-sharepoint-rest-service-endpoint-uris).
/// The query string is limited by ASP.NET's maxQueryStringLength, 2,048 characters by default
/// (https://learn.microsoft.com/dotnet/api/system.web.configuration.httpruntimesection.maxquerystringlength);
/// the connector publishes no other value. Escaped, a character can take up to 9 characters, so
/// a path within SharePoint's 400 characters can still be too long as an address.
/// </summary>
public static class SharePointAddress
{
    /// <summary>The longest query string the connector accepts.</summary>
    public const int MaxQueryString = 2048;

    /// <summary>The fields a folder lookup by path reads: the longest read that carries a path.</summary>
    public const string FolderLookup =
        "$select=Exists,UniqueId,ServerRelativeUrl,ListItemAllFields/Id,ListItemAllFields/UniqueId,ListItemAllFields/FileLeafRef,ListItemAllFields/FileRef,ListItemAllFields/FSObjType&$expand=ListItemAllFields";

    /// <summary>The parameter alias that carries a decoded value, escaped for the query string.</summary>
    public static string Alias(string value) =>
        "@p='" + Uri.EscapeDataString(value.Replace("'", "''")) + "'";

    /// <summary>The length of the query string that looks a folder up by its path.</summary>
    public static int LookupLength(string path) => Alias(path).Length + 1 + FolderLookup.Length;

    /// <summary>
    /// Whether every read of the folder at this server-relative path fits the connector: the
    /// folder lookup is the longest, and its ancestors' paths are shorter.
    /// </summary>
    public static bool Fits(string path) => LookupLength(path) <= MaxQueryString;
}
