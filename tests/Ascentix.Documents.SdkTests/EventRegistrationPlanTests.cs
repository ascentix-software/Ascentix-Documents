using System;
using System.Collections.Generic;
using System.Linq;
using Ascentix.Documents.Dataverse;
using Xunit;

namespace Ascentix.Documents.SdkTests;

public sealed class EventRegistrationPlanTests
{
    private static readonly Guid Worker = Guid.NewGuid();

    internal static RegistrationCatalog Catalog(params string[] tables)
    {
        var c = new RegistrationCatalog
        {
            RecordHandler = Guid.NewGuid(),
            RetirementHandler = Guid.NewGuid(),
            MembershipHandler = Guid.NewGuid(),
        };
        foreach (
            var m in EventRegistrationPlan
                .RecordMessages.Concat(new[] { "Delete", "Associate", "Disassociate" })
                .Distinct()
        )
            c.Messages[m] = Guid.NewGuid();
        foreach (var t in tables.Concat(new[] { "team" }))
        foreach (var m in EventRegistrationPlan.RecordMessages)
            c.Filters[m + "|" + t] = Guid.NewGuid();
        return c;
    }

    private static ExistingStep Matching(StepSpec spec, Guid? worker = null) =>
        new ExistingStep
        {
            Id = Guid.NewGuid(),
            Name = spec.Name,
            HandlerId = spec.HandlerId,
            MessageId = spec.MessageId,
            FilterId = spec.FilterId,
            Stage = 40,
            Mode = 1,
            ImpersonatingUserId = worker ?? Worker,
            Active = spec.Active,
            AsyncAutoDelete = true,
            Description = EventRegistrationPlan.Label,
        };

    [Fact]
    public void DesiredCoversTablesTeamAndMembershipWithUpdateFlag()
    {
        var c = Catalog("account", "contact");
        var off = EventRegistrationPlan.Desired(c, new[] { "account", "contact" }, false);
        Assert.Equal(2 * EventRegistrationPlan.RecordMessages.Length + 3, off.Count);
        Assert.All(off.Where(s => s.Name.Contains(" Update")), s => Assert.False(s.Active));
        Assert.Contains(off, s => s.Name == "Ascentix Documents: event Create account" && s.Active);
        Assert.Contains(off, s => s.Name == "Ascentix Documents: event Delete team");
        Assert.Contains(
            off,
            s => s.Name == "Ascentix Documents: membership Associate" && s.FilterId == null
        );
        var on = EventRegistrationPlan.Desired(c, new[] { "account" }, true);
        Assert.All(on.Where(s => s.Name.Contains(" Update")), s => Assert.True(s.Active));
        Assert.Equal(3, EventRegistrationPlan.Desired(c, Array.Empty<string>(), false).Count);
    }

    [Fact]
    public void DiffCreatesUpdatesAndDeletes()
    {
        var c = Catalog("account", "contact");
        var desired = EventRegistrationPlan.Desired(c, new[] { "account" }, false);
        var existing = new List<ExistingStep>
        {
            Matching(desired[0]),
            Matching(desired[1], Guid.NewGuid()), // wrong worker
        };
        var legacy = Matching(desired[2]);
        legacy.Mode = 0;
        legacy.Description = "Owned by standalone Ascentix Documents development setup v1";
        existing.Add(legacy);
        var removed = Matching(EventRegistrationPlan.Desired(c, new[] { "contact" }, false)[0]);
        existing.Add(removed);
        var duplicate = Matching(desired[0]);
        existing.Add(duplicate);
        var changes = EventRegistrationPlan.Diff(desired, existing, Worker);
        Assert.Equal(desired.Count - 3, changes.Create.Count);
        Assert.Equal(
            new[] { existing[1].Id, legacy.Id },
            changes.Update.Select(u => u.Key).ToArray()
        );
        Assert.Equal(
            new[] { removed.Id, duplicate.Id }.OrderBy(i => i),
            changes.Delete.OrderBy(i => i)
        );
    }

    [Fact]
    public void UnchangedRegistrationsProduceNoChanges()
    {
        var c = Catalog("account");
        var desired = EventRegistrationPlan.Desired(c, new[] { "account" }, true);
        var existing = desired.Select(s => Matching(s)).ToList();
        Assert.True(EventRegistrationPlan.Diff(desired, existing, Worker).IsEmpty);
    }

    [Fact]
    public void ReadinessReportsEachScope()
    {
        var c = Catalog("account", "contact", "lead");
        var desired = EventRegistrationPlan.Desired(
            c,
            new[] { "account", "contact", "lead" },
            false
        );
        var existing = desired
            .Where(s => s.Scope == "account" || s.Scope == "team")
            .Select(s => Matching(s))
            .ToList();
        existing.Add(Matching(desired.First(s => s.Scope == "contact")));
        var status = EventRegistrationPlan
            .Readiness(desired, existing, Worker)
            .ToDictionary(r => r.Scope, r => r.Status);
        Assert.Equal("Ready", status["account"]);
        Assert.Equal("Missing", status["contact"]);
        Assert.Equal("Pending", status["lead"]);
        Assert.Equal("Ready", status["team"]);
        existing.First(s => s.Name.EndsWith("Create account")).ImpersonatingUserId = Guid.NewGuid();
        Assert.Equal(
            "WrongWorker",
            EventRegistrationPlan
                .Readiness(desired, existing, Worker)
                .Single(r => r.Scope == "account")
                .Status
        );
        Assert.Equal(0, EventRegistrationPlan.Extra(desired, existing));
    }
}
