namespace XL.Report;

public readonly record struct SheetRelated<T>(string SheetName, T Value);

public static class SheetRelated
{
    public readonly struct Formattable<T> : ISpanFormattable
        where T : ISpanFormattable
    {
        private readonly SheetRelated<T> content;

        public Formattable(SheetRelated<T> content)
        {
            this.content = content;
        }

        public override string ToString() => ToString(null, null);

        public string ToString(string? format, IFormatProvider? formatProvider) =>
            $"{content.SheetName}!{content.Value.ToString(format, formatProvider)}";

        public bool TryFormat(
            Span<char> destination,
            out int charsWritten,
            ReadOnlySpan<char> format,
            IFormatProvider? provider)
        {
            return FormatContext.Start
                .Write(ref destination, content.SheetName)
                .Write(ref destination, "!")
                .Write(ref destination, content.Value, format, provider)
                .Deconstruct(out charsWritten);
        }

        public static implicit operator Formattable<T>(SheetRelated<T> value) => value.ToFormattable();
    }

    public static Formattable<T> ToFormattable<T>(this SheetRelated<T> value)
        where T : ISpanFormattable => new(value);
}