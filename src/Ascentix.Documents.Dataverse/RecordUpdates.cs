using System;

namespace Ascentix.Documents.Dataverse;

/// <summary>
/// Whether Documents offers "Update folders when records change". Off in 0.1.0.4: Documents never
/// renames existing folders, so planning a record again after it changes creates only the folders
/// that are newly called for, which confuses users. The setting returns in a later release with a
/// choice to rename folders; turning this switch on brings back the code kept for it.
/// While off, a stored "on" reads as off, Save stores off, record Update events are not captured,
/// and the record Update steps are not registered: Repair removes any that remain.
/// </summary>
public static class RecordUpdates
{
    /// <summary>
    /// Set only by the SDK tests, through reflection, to cover the code kept for the setting.
    /// Per thread, so a test that turns it on cannot affect another test running in parallel.
    /// </summary>
#pragma warning disable CS0649 // Assigned only by the tests, through reflection.
    [ThreadStatic]
    private static bool? testOverride;
#pragma warning restore CS0649

    /// <summary>True when the record-update setting can be on.</summary>
    public static bool Available => testOverride ?? false;
}
