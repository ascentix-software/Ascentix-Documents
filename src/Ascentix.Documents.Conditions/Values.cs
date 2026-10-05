using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Ascentix.Documents.Conditions;

public enum ValueKind
{
    Text,
    Number,
    Boolean,
    Choice,
    Lookup,
    DateOnly,
    DateTime,
    MultiChoice,
}

public enum Comparison
{
    Equal,
    NotEqual,
    Greater,
    GreaterOrEqual,
    Less,
    LessOrEqual,
    Contains,
    DoesNotContain,
    IsNull,
    IsNotNull,
}

public sealed class EvaluationBlockedException : Exception
{
    public EvaluationBlockedException(string message)
        : base(message) { }
}

public sealed class Value
{
    public ValueKind Kind { get; }
    public object? Data { get; }
    public bool IsNull => Data == null;

    private Value(ValueKind kind, object? data)
    {
        Kind = kind;
        Data = data;
    }

    public static Value Null(ValueKind kind) => new Value(kind, null);

    public static Value Text(string? value) => new Value(ValueKind.Text, value);

    public static Value Number(decimal value) => new Value(ValueKind.Number, value);

    public static Value Boolean(bool value) => new Value(ValueKind.Boolean, value);

    public static Value Choice(int value) => new Value(ValueKind.Choice, value);

    public static Value Lookup(Guid value) => new Value(ValueKind.Lookup, value);

    public static Value DateOnly(DateTime value) =>
        new Value(ValueKind.DateOnly, DateTime.SpecifyKind(value.Date, DateTimeKind.Unspecified));

    public static Value Instant(DateTimeOffset value) =>
        new Value(ValueKind.DateTime, value.ToUniversalTime());

    public static Value MultiChoice(IEnumerable<int> values) =>
        new Value(
            ValueKind.MultiChoice,
            Array.AsReadOnly(values.Distinct().OrderBy(v => v).ToArray())
        );

    /// <summary>Formats the value for a folder name, culture-independently.</summary>
    /// <returns>The text, or null when the value is null so the planner can wait for it.</returns>
    public string? Format()
    {
        if (IsNull)
            return null;
        if (Kind == ValueKind.DateOnly)
            return ((DateTime)Data!).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        if (Kind == ValueKind.DateTime)
            return ((DateTimeOffset)Data!).ToString(
                "yyyy-MM-dd'T'HH-mm-ss'Z'",
                CultureInfo.InvariantCulture
            );
        if (Kind == ValueKind.MultiChoice)
            return string.Join(
                "-",
                ((IReadOnlyList<int>)Data!).Select(v => v.ToString(CultureInfo.InvariantCulture))
            );
        return Convert.ToString(Data, CultureInfo.InvariantCulture)!;
    }
}

public static class ValueComparer
{
    public static bool Supports(ValueKind kind, Comparison op)
    {
        if (!Enum.IsDefined(typeof(ValueKind), kind) || !Enum.IsDefined(typeof(Comparison), op))
            return false;
        if (
            op == Comparison.IsNull
            || op == Comparison.IsNotNull
            || op == Comparison.Equal
            || op == Comparison.NotEqual
        )
            return true;
        if (op == Comparison.Contains || op == Comparison.DoesNotContain)
            return kind == ValueKind.Text || kind == ValueKind.MultiChoice;
        return kind == ValueKind.Number || kind == ValueKind.DateOnly || kind == ValueKind.DateTime;
    }

    public static bool Compare(Value actual, Comparison op, Value? expected = null)
    {
        if (!Supports(actual.Kind, op))
            throw new EvaluationBlockedException("Operator is not supported for the field type.");
        if (op == Comparison.IsNull)
            return actual.IsNull;
        if (op == Comparison.IsNotNull)
            return !actual.IsNull;
        if (expected == null || expected.Kind != actual.Kind || expected.IsNull)
            throw new EvaluationBlockedException(
                "A non-null matching typed comparison value is required."
            );
        // Missing values match only explicit null predicates.
        if (actual.IsNull)
            return false;
        if (actual.Kind == ValueKind.MultiChoice)
        {
            var left = (IReadOnlyList<int>)actual.Data!;
            var right = (IReadOnlyList<int>)expected.Data!;
            if (op == Comparison.Contains || op == Comparison.DoesNotContain)
            {
                if (right.Count == 0)
                    throw new EvaluationBlockedException("Contains requires at least one choice.");
                bool contains = right.All(left.Contains);
                return op == Comparison.Contains ? contains : !contains;
            }
            bool equal = left.SequenceEqual(right);
            return op == Comparison.Equal ? equal : !equal;
        }
        if (
            actual.Kind == ValueKind.Text
            && (op == Comparison.Contains || op == Comparison.DoesNotContain)
        )
        {
            bool contains =
                ((string)actual.Data!).IndexOf(
                    (string)expected.Data!,
                    StringComparison.OrdinalIgnoreCase
                ) >= 0;
            return op == Comparison.Contains ? contains : !contains;
        }
        int comparison =
            actual.Kind == ValueKind.Text
                ? StringComparer.OrdinalIgnoreCase.Compare(actual.Data, expected.Data)
                : ((IComparable)actual.Data!).CompareTo(expected.Data);
        switch (op)
        {
            case Comparison.Equal:
                return comparison == 0;
            case Comparison.NotEqual:
                return comparison != 0;
            case Comparison.Greater:
                return comparison > 0;
            case Comparison.GreaterOrEqual:
                return comparison >= 0;
            case Comparison.Less:
                return comparison < 0;
            case Comparison.LessOrEqual:
                return comparison <= 0;
            default:
                throw new EvaluationBlockedException("Unsupported comparison.");
        }
    }
}
