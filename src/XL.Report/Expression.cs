using System.Globalization;
using XL.Report.Auxiliary;

namespace XL.Report;

public abstract class Expression : ISpanFormattable
{
    public abstract string ToString(string? format, IFormatProvider? formatProvider);

    public abstract bool TryFormat(
        Span<char> destination,
        out int charsWritten,
        ReadOnlySpan<char> format,
        IFormatProvider? provider);

    public sealed class Verbatim(string content) : Expression
    {
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

    public sealed class String(string content) : Expression
    {
        private const int EnclosingQuotesLength = 2;

        public override string ToString(string? format, IFormatProvider? formatProvider)
        {
            var quotes = content.AsSpan().Count('"');
            return string.Create(
                content.Length + quotes + EnclosingQuotesLength,
                content,
                (destination, source) => TryFormat(
                    destination,
                    out _,
                    "",
                    CultureInfo.InvariantCulture
                )
            );
        }

        public override bool TryFormat(
            Span<char> destination,
            out int charsWritten,
            ReadOnlySpan<char> format,
            IFormatProvider? provider)
        {
            if (destination.Length < content.Length + EnclosingQuotesLength)
            {
                charsWritten = 0;
                return false;
            }

            var source = content.AsSpan();
            var context = FormatContext.Start.Write(ref destination, "\"");
            while (source.Length > 0)
            {
                var quoteIndex = source.IndexOf('"');
                if (quoteIndex == -1)
                {
                    context = context.Write(ref destination, source);
                    break;
                }

                context = context
                    .Write(ref destination, source[..quoteIndex])
                    .Write(ref destination, "\"\"");
                source = source[(quoteIndex + 1)..];
            }

            context = context.Write(ref destination, "\"");
            return context.Deconstruct(out charsWritten);
        }
    }
}