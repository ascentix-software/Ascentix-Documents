using System;
using System.Collections.Generic;
using System.Linq;

namespace Ascentix.Documents.SdkTests;

/// <summary>
/// Rewrites stored payloads to the shape an earlier release wrote: the named top-level members
/// are absent, as they are in rows saved before those members existed.
/// </summary>
internal static class LegacyPayload
{
    public static string Strip(string json, params string[] members)
    {
        var keep = new List<string>();
        foreach (var (name, text) in Members(json))
            if (!members.Contains(name))
                keep.Add(text);
        return "{" + string.Join(",", keep) + "}";
    }

    /// <summary>The top-level members of a JSON object, each with its "name":value text.</summary>
    private static IEnumerable<(string Name, string Text)> Members(string json)
    {
        json = json.Trim();
        if (json.Length < 2 || json[0] != '{' || json[json.Length - 1] != '}')
            throw new FormatException("A JSON object is required.");
        int depth = 0,
            start = 1;
        bool quoted = false;
        for (int i = 1; i < json.Length - 1; i++)
        {
            char c = json[i];
            if (quoted)
            {
                if (c == '\\')
                    i++;
                else if (c == '"')
                    quoted = false;
                continue;
            }
            if (c == '"')
                quoted = true;
            else if (c == '{' || c == '[')
                depth++;
            else if (c == '}' || c == ']')
                depth--;
            else if (c == ',' && depth == 0)
            {
                yield return Member(json.Substring(start, i - start));
                start = i + 1;
            }
        }
        if (json.Length - 1 > start)
            yield return Member(json.Substring(start, json.Length - 1 - start));
    }

    private static (string, string) Member(string text)
    {
        text = text.Trim();
        int end = text.IndexOf("\":", StringComparison.Ordinal);
        return (text.Substring(1, end - 1), text);
    }

    /// <summary>Strips the members from every row of the table, in place.</summary>
    /// <returns>How many rows were rewritten.</returns>
    public static int Strip(
        DurableWorkerTests.MemoryService service,
        string table,
        params string[] members
    )
    {
        int count = 0;
        foreach (var row in service.Rows.Values.Where(r => r.LogicalName == table).ToList())
        {
            var payload = row.GetAttributeValue<string>("asx_payload");
            if (string.IsNullOrEmpty(payload))
                continue;
            var legacy = Strip(payload, members);
            if (Members(legacy).Any(m => members.Contains(m.Name)))
                throw new InvalidOperationException("A stripped member is still present.");
            row["asx_payload"] = legacy;
            count++;
        }
        return count;
    }
}
