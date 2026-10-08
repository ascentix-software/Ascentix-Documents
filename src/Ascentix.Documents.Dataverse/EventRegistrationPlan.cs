using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace Ascentix.Documents.Dataverse;

public sealed class StepSpec
{
    public string Scope { get; set; } = "";
    public string Name { get; set; } = "";
    public Guid HandlerId { get; set; }
    public Guid MessageId { get; set; }
    public Guid? FilterId { get; set; }
    public bool Active { get; set; }

    /// <summary>
    /// A step that must not exist: a record Update step while "Update folders when records
    /// change" is out of the release (RecordUpdates.Available). Register deletes it, and a table
    /// that still has one needs repair.
    /// </summary>
    public bool Retired { get; set; }
}

public sealed class ExistingStep
{
    public Guid Id { get; set; }
    public string Name { get; set; } = "";
    public Guid HandlerId { get; set; }
    public Guid MessageId { get; set; }
    public Guid? FilterId { get; set; }
    public int Stage { get; set; }
    public int Mode { get; set; }
    public Guid? ImpersonatingUserId { get; set; }
    public bool Active { get; set; }
    public bool AsyncAutoDelete { get; set; }
    public string Description { get; set; } = "";
}

public sealed class RegistrationCatalog
{
    public Guid RecordHandler { get; set; }
    public Guid RetirementHandler { get; set; }
    public Guid MembershipHandler { get; set; }
    public Dictionary<string, Guid> Messages { get; } = new(StringComparer.Ordinal);

    /// <summary>Key "Message|table".</summary>
    public Dictionary<string, Guid> Filters { get; } = new(StringComparer.Ordinal);

    /// <summary>Read privilege id per table; null when metadata exposes none.</summary>
    public Dictionary<string, Guid?> ReadPrivileges { get; } = new(StringComparer.Ordinal);
}

public sealed class RegistrationChanges
{
    public List<StepSpec> Create { get; } = new();
    public List<KeyValuePair<Guid, StepSpec>> Update { get; } = new();
    public List<Guid> Delete { get; } = new();
    public bool IsEmpty => Create.Count == 0 && Update.Count == 0 && Delete.Count == 0;
}

[DataContract]
public sealed class TableReadiness
{
    [DataMember]
    public string Scope { get; set; } = "";

    [DataMember]
    public string Status { get; set; } = "";
}

/// <summary>Pure desired-state, diff and readiness logic for application-owned event steps.</summary>
public static class EventRegistrationPlan
{
    public const string Label = "Ascentix Documents event registration";
    public const string TeamScope = "team";

    /// <summary>Record messages registered per allowlisted table. Add CreateMultiple/UpdateMultiple only if single-record steps do not run for bulk items.</summary>
    public static readonly string[] RecordMessages = { "Create", "Update", "Delete" };

    public static List<StepSpec> Desired(
        RegistrationCatalog catalog,
        string[] tables,
        bool processUpdates
    )
    {
        var specs = new List<StepSpec>();
        foreach (var table in tables)
        foreach (var message in RecordMessages)
        {
            bool update = message.StartsWith("Update", StringComparison.Ordinal);
            specs.Add(
                new StepSpec
                {
                    Scope = table,
                    Name = "Ascentix Documents: event " + message + " " + table,
                    HandlerId = catalog.RecordHandler,
                    MessageId = catalog.Messages[message],
                    FilterId = catalog.Filters[message + "|" + table],
                    Active = !update || processUpdates,
                    Retired = update && !RecordUpdates.Available,
                }
            );
        }
        specs.Add(
            new StepSpec
            {
                Scope = TeamScope,
                Name = "Ascentix Documents: event Delete team",
                HandlerId = catalog.RetirementHandler,
                MessageId = catalog.Messages["Delete"],
                FilterId = catalog.Filters["Delete|team"],
                Active = true,
            }
        );
        foreach (var message in new[] { "Associate", "Disassociate" })
            specs.Add(
                new StepSpec
                {
                    Scope = TeamScope,
                    Name = "Ascentix Documents: membership " + message,
                    HandlerId = catalog.MembershipHandler,
                    MessageId = catalog.Messages[message],
                    FilterId = null,
                    Active = true,
                }
            );
        return specs;
    }

    public static RegistrationChanges Diff(
        IReadOnlyList<StepSpec> desired,
        IReadOnlyList<ExistingStep> existing,
        Guid worker
    )
    {
        var changes = new RegistrationChanges();
        var unmatched = existing.ToList();
        foreach (var spec in desired)
        {
            var match = unmatched.FirstOrDefault(s => Same(s, spec));
            if (match == null)
            {
                if (!spec.Retired)
                    changes.Create.Add(spec);
                continue;
            }
            unmatched.Remove(match);
            if (spec.Retired)
            {
                changes.Delete.Add(match.Id);
                continue;
            }
            if (!Correct(match, spec, worker))
                changes.Update.Add(new KeyValuePair<Guid, StepSpec>(match.Id, spec));
        }
        changes.Delete.AddRange(unmatched.Select(s => s.Id));
        return changes;
    }

    public static bool Correct(ExistingStep step, StepSpec spec, Guid worker) =>
        step.Name == spec.Name
        && step.Stage == 40
        && step.Mode == 1
        && step.ImpersonatingUserId == worker
        && step.AsyncAutoDelete
        && step.Description == Label
        && step.Active == spec.Active;

    public static List<TableReadiness> Readiness(
        IReadOnlyList<StepSpec> desired,
        IReadOnlyList<ExistingStep> existing,
        Guid worker
    ) =>
        desired
            .GroupBy(s => s.Scope, StringComparer.Ordinal)
            .Select(g => new TableReadiness
            {
                Scope = g.Key,
                Status = Status(g.ToList(), existing, worker),
            })
            .ToList();

    public static int Extra(
        IReadOnlyList<StepSpec> desired,
        IReadOnlyList<ExistingStep> existing
    ) => existing.Count(s => !desired.Any(spec => !spec.Retired && Same(s, spec)));

    private static string Status(
        List<StepSpec> specs,
        IReadOnlyList<ExistingStep> existing,
        Guid worker
    )
    {
        bool retired = specs.Any(spec => spec.Retired && existing.Any(s => Same(s, spec)));
        var pairs = specs
            .Where(spec => !spec.Retired)
            .Select(spec => new { Spec = spec, Step = existing.FirstOrDefault(s => Same(s, spec)) })
            .ToList();
        if (pairs.All(p => p.Step == null))
            return retired ? "Missing" : "Pending";
        if (pairs.Any(p => p.Step == null))
            return "Missing";
        if (pairs.Any(p => p.Step!.ImpersonatingUserId != worker))
            return "WrongWorker";
        if (pairs.Any(p => p.Step!.Mode != 1 || p.Step.Stage != 40))
            return "WrongMode";
        if (pairs.Any(p => p.Step!.Active != p.Spec.Active))
            return "WrongState";
        // A step that must no longer exist, such as a record Update step registered before
        // 0.1.0.4, makes the table need repair; Register deletes it.
        if (retired || pairs.Any(p => !Correct(p.Step!, p.Spec, worker)))
            return "Outdated";
        return "Ready";
    }

    private static bool Same(ExistingStep step, StepSpec spec) =>
        step.HandlerId == spec.HandlerId
        && step.MessageId == spec.MessageId
        && step.FilterId == spec.FilterId;
}
