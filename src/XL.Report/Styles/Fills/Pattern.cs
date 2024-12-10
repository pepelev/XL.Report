using XL.Report.Auxiliary;

namespace XL.Report.Styles.Fills;

public sealed class Pattern : IEquatable<Pattern>, ISpanFormattable
{
    public Pattern(string value)
    {
        Value = value;
    }

    public string Value { get; }

    public static Pattern MediumGray { get; } = new("mediumGray");
    public static Pattern DarkGray { get; } = new("darkGray");
    public static Pattern LightGray { get; } = new("lightGray");
    public static Pattern DarkHorizontal { get; } = new("darkHorizontal");
    public static Pattern DarkVertical { get; } = new("darkVertical");
    public static Pattern DarkDown { get; } = new("darkDown");
    public static Pattern DarkUp { get; } = new("darkUp");
    public static Pattern DarkGrid { get; } = new("darkGrid");
    public static Pattern DarkTrellis { get; } = new("darkTrellis");
    public static Pattern LightHorizontal { get; } = new("lightHorizontal");
    public static Pattern LightVertical { get; } = new("lightVertical");
    public static Pattern LightDown { get; } = new("lightDown");
    public static Pattern LightUp { get; } = new("lightUp");
    public static Pattern LightGrid { get; } = new("lightGrid");
    public static Pattern LightTrellis { get; } = new("lightTrellis");
    public static Pattern Gray125 { get; } = new("gray125");
    public static Pattern Gray0625 { get; } = new("gray0625");

    public bool Equals(Pattern? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Value == other.Value;
    }

    public string ToString(string? format, IFormatProvider? formatProvider) => Value;

    public bool TryFormat(
        Span<char> destination,
        out int charsWritten,
        ReadOnlySpan<char> format,
        IFormatProvider? provider)
    {
        var context = new FormatContext(destination);
        context.Write(Value);
        return context.Finish(out charsWritten);
    }

    public override string ToString() => ToString(null, null);

    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is Pattern other && Equals(other);
    public override int GetHashCode() => Value.GetHashCode();
}