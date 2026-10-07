using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Ascentix.Documents.Conditions;

public sealed class FieldReference
{
    public string Source { get; }
    public string Column { get; }

    public FieldReference(string source, string column)
    {
        if (
            !Regex.IsMatch(source ?? "", "^[a-z][a-z0-9_]*$")
            || !Regex.IsMatch(column ?? "", "^[a-z][a-z0-9_]*$")
        )
            throw new EvaluationBlockedException("Invalid logical field reference.");
        Source = source!;
        Column = column!;
    }

    public override string ToString() => Source + "." + Column;
}

public sealed class Snapshot
{
    private readonly IReadOnlyDictionary<string, Value> values;

    public Snapshot(IDictionary<string, Value> values)
    {
        this.values = new Dictionary<string, Value>(values, StringComparer.Ordinal);
    }

    public Value Resolve(FieldReference field) =>
        values.TryGetValue(field.ToString(), out var value)
            ? value
            : throw new EvaluationBlockedException(
                "Field was not present in the authorized snapshot: " + field
            );
}

public abstract class Predicate
{
    public abstract bool Evaluate(Snapshot snapshot);
    public abstract IEnumerable<FieldReference> Fields { get; }
    internal abstract int Validate(int depth);

    /// <summary>
    /// Checks the tree's shape. Its size is bounded where it is saved (Bounds.ConfigurationRows,
    /// one row per group and condition). Its depth (groups nested at most 9 below a folder's
    /// group, conditions at depth 10) keeps the draft that carries it inside JsonWire's 32-level
    /// nesting quota, which refuses a save or a preview with groups nested 12 deep
    /// (CapTraceTests.TheDeepestConditionTreeAPreviewCarriesFitsTheWireNestingQuota).
    /// </summary>
    public void Validate() => Validate(0);
}

public sealed class Condition : Predicate
{
    public FieldReference Left { get; }
    public Comparison Operator { get; }
    public Value? Literal { get; }
    public FieldReference? Right { get; }

    public Condition(
        FieldReference left,
        Comparison op,
        Value? literal = null,
        FieldReference? right = null
    )
    {
        Left = left;
        Operator = op;
        Literal = literal;
        Right = right;
        Validate();
    }

    internal override int Validate(int depth)
    {
        if (depth > 10 || !Enum.IsDefined(typeof(Comparison), Operator))
            throw new EvaluationBlockedException("Invalid condition shape.");
        bool unary = Operator == Comparison.IsNull || Operator == Comparison.IsNotNull;
        if (unary ? Literal != null || Right != null : (Literal == null) == (Right == null))
            throw new EvaluationBlockedException(
                "Condition must have exactly one comparison source, or none for a null predicate."
            );
        return 1;
    }

    public override bool Evaluate(Snapshot snapshot) =>
        ValueComparer.Compare(
            snapshot.Resolve(Left),
            Operator,
            Right != null ? snapshot.Resolve(Right) : Literal
        );

    public override IEnumerable<FieldReference> Fields =>
        Right == null ? new[] { Left } : new[] { Left, Right! };
}

public sealed class ConditionGroup : Predicate
{
    public bool All { get; }
    public IReadOnlyList<Predicate> Children { get; }

    public ConditionGroup(bool all, params Predicate[] children)
    {
        All = all;
        Children = Array.AsReadOnly((Predicate[])children.Clone());
        Validate();
    }

    internal override int Validate(int depth)
    {
        if (depth > 10 || Children.Count == 0 || Children.Any(c => c == null))
            throw new EvaluationBlockedException("Invalid or empty condition group.");
        return Children.Sum(c => c.Validate(depth + 1));
    }

    public override bool Evaluate(Snapshot snapshot)
    {
        // Evaluates every child predicate before combining the verdicts.
        var verdicts = Children.Select(c => c.Evaluate(snapshot)).ToArray();
        return All ? verdicts.All(v => v) : verdicts.Any(v => v);
    }

    public override IEnumerable<FieldReference> Fields => Children.SelectMany(c => c.Fields);
}

public static class NameExpression
{
    private static readonly Regex Token = new Regex(
        @"\{([a-z][a-z0-9_]*)\.([a-z][a-z0-9_]*)\}",
        RegexOptions.CultureInvariant
    );

    public static IReadOnlyList<FieldReference> Fields(string expression)
    {
        if (string.IsNullOrWhiteSpace(expression) || expression.Length > 2000)
            throw new EvaluationBlockedException("A bounded naming expression is required.");
        var stripped = Token.Replace(expression, "");
        if (stripped.IndexOfAny(new[] { '{', '}' }) >= 0)
            throw new EvaluationBlockedException(
                "Names use {root.column} or {lookupAlias.column} tokens."
            );
        return Token
            .Matches(expression)
            .Cast<Match>()
            .Select(m => new FieldReference(m.Groups[1].Value, m.Groups[2].Value))
            .ToArray();
    }

    /// <summary>Renders a folder name from the snapshot.</summary>
    /// <param name="expression">The naming expression with its field tokens.</param>
    /// <param name="snapshot">The record's authorized values.</param>
    /// <param name="blank">The field the name waits for, when it returns null.</param>
    /// <returns>
    /// The rendered name, or null when a field is null, or when the whole name is blank because
    /// a field is blank. A blank field inside an otherwise filled name is kept as it is.
    /// </returns>
    public static string? Render(string expression, Snapshot snapshot, out FieldReference? blank)
    {
        Fields(expression);
        FieldReference? missing = null;
        FieldReference? empty = null;
        var name = Token.Replace(
            expression,
            m =>
            {
                var field = new FieldReference(m.Groups[1].Value, m.Groups[2].Value);
                var text = snapshot.Resolve(field).Format();
                if (text == null)
                {
                    missing ??= field;
                    return "";
                }
                if (string.IsNullOrWhiteSpace(text))
                    empty ??= field;
                return text;
            }
        );
        blank = missing ?? (string.IsNullOrWhiteSpace(name) ? empty : null);
        return blank == null ? name.Normalize(NormalizationForm.FormC) : null;
    }
}
