namespace XL.Report;

public abstract class Expression : ISpanFormattable
{
    public abstract string ToString(string? format, IFormatProvider? formatProvider);

    public abstract bool TryFormat(
        Span<char> destination,
        out int charsWritten,
        ReadOnlySpan<char> format,
        IFormatProvider? provider);

    public sealed class Verbatim : Expression
    {
        private readonly string content;

        public Verbatim(string content)
        {
            this.content = content;
        }

        public override string ToString(string? format, IFormatProvider? formatProvider) => content;

        public override bool TryFormat(
            Span<char> destination,
            out int charsWritten,
            ReadOnlySpan<char> format,
            IFormatProvider? provider)
        {
            return FormatContext.Start.Write(ref destination, content).Deconstruct(out charsWritten);
        }
    }
}