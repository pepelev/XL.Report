using System.Globalization;

namespace XL.Report.Styles;

public readonly struct StyleId : IComparable<StyleId>, ISpanFormattable
{
    public int Index { get; }

    public StyleId(int index)
    {
        Index = index;
    }

    public int CompareTo(StyleId other)
    {
        return Index.CompareTo(other.Index);
    }

    public override string ToString() => ToString(null, null);

    public string ToString(string? format, IFormatProvider? formatProvider) =>
        Index.ToString(CultureInfo.InvariantCulture);

    public bool TryFormat(
        Span<char> destination,
        out int charsWritten,
        ReadOnlySpan<char> format,
        IFormatProvider? provider)
    {
        return Index.TryFormat(
            destination,
            out charsWritten,
            ReadOnlySpan<char>.Empty,
            CultureInfo.InvariantCulture
        );
    }
}