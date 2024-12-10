using XL.Report.Auxiliary;

namespace XL.Report;

public readonly record struct SheetRelated<T>(string SheetName, T Value);

public static class SheetRelated
{
    public readonly struct Formattable<T>(SheetRelated<T> content) : ISpanFormattable
        where T : ISpanFormattable
    {
        public override string ToString() => ToString(null, null);

        public string ToString(string? format, IFormatProvider? formatProvider) =>
            FormatContext.ToString(this, format, formatProvider);

        public bool TryFormat(
            Span<char> destination,
            out int charsWritten,
            ReadOnlySpan<char> format,
            IFormatProvider? provider)
        {
            var context = new FormatContext(destination);
            context.Write(content.SheetName);
            context.Write("!");
            context.Write(content.Value, format, provider);
            return context.Finish(out charsWritten);
        }

        public static implicit operator Formattable<T>(SheetRelated<T> value) => value.ToFormattable();
    }

    public static Formattable<T> ToFormattable<T>(this SheetRelated<T> value)
        where T : ISpanFormattable => new(value);
}