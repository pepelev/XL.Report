using System.Globalization;

namespace XL.Report;

internal readonly record struct FormatContext(bool Success, int CharsWritten)
{
    public static FormatContext Start => new(true, 0);
    public static FormatContext Fail => new(false, 0);

    public FormatContext Write<T>(
        ref Span<char> destination,
        T value,
        ReadOnlySpan<char> format,
        IFormatProvider? formatProvider) where T : ISpanFormattable
    {
        if (!Success)
        {
            return Fail;
        }

        if (value.TryFormat(destination, out var charsWritten, format, formatProvider))
        {
            destination = destination[charsWritten..];
            return new FormatContext(true, CharsWritten + charsWritten);
        }

        return Fail;
    }

    public FormatContext Write<T>(ref Span<char> destination, T value) where T : ISpanFormattable =>
        Write(ref destination, value, ReadOnlySpan<char>.Empty, CultureInfo.InvariantCulture);

    public FormatContext Write<T>(ref Span<char> destination, T value, Format<T> format)
    {
        if (!Success)
        {
            return Fail;
        }

        if (format(value, destination, out var charsWritten))
        {
            destination = destination[charsWritten..];
            return new FormatContext(true, CharsWritten + charsWritten);
        }

        return Fail;
    }

    public FormatContext Write(ref Span<char> destination, string value)
    {
        if (!Success)
        {
            return Fail;
        }

        if (value.Length <= destination.Length)
        {
            value.CopyTo(destination);
            destination = destination[value.Length..];
            return new FormatContext(true, CharsWritten + value.Length);
        }

        return Fail;
    }

    public bool Deconstruct(out int charsWritten)
    {
        charsWritten = CharsWritten;
        return Success;
    }

    public delegate bool Format<in T>(T value, Span<char> span, out int charsWritten);
}