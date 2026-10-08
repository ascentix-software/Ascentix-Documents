using System;
using System.Reflection;
using Ascentix.Documents.Dataverse;

namespace Ascentix.Documents.SdkTests;

/// <summary>
/// Turns the "Update folders when records change" switch (RecordUpdates.Available) on or off for
/// the current thread, so tests keep covering the code kept for the setting while it is out of the
/// release. Dispose restores the release default.
/// </summary>
internal sealed class RecordUpdatesSwitch : IDisposable
{
    private static readonly FieldInfo Field =
        typeof(RecordUpdates).GetField("testOverride", BindingFlags.NonPublic | BindingFlags.Static)
        ?? throw new InvalidOperationException("RecordUpdates.testOverride is missing.");

    private RecordUpdatesSwitch(bool available) => Field.SetValue(null, available);

    /// <summary>Makes the record-update setting available, as before 0.1.0.4.</summary>
    internal static RecordUpdatesSwitch On() => new(true);

    /// <summary>Takes the record-update setting away, as in the 0.1.0.4 release.</summary>
    internal static RecordUpdatesSwitch Off() => new(false);

    public void Dispose() => Field.SetValue(null, null);
}
