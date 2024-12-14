#region Legal
// Copyright 2024 Pepelev Alexey
// 
// This file is part of XL.Report.
// 
// XL.Report is free software: you can redistribute it and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation, either
// version 3 of the License, or (at your option) any later version.
// 
// XL.Report is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
// without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU Lesser General Public License for more details.
// 
// You should have received a copy of the GNU Lesser General Public License along with XL.Report.
// If not, see <https://www.gnu.org/licenses/>.
#endregion

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
            var context = new FormatContext(destination);
            context.Write(content);
            return context.Finish(out charsWritten);
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
            var context = new FormatContext(destination);
            context.Write("\"");
            while (source.Length > 0)
            {
                var quoteIndex = source.IndexOf('"');
                if (quoteIndex == -1)
                {
                    context.Write(source);
                    break;
                }

                context.Write(source[..quoteIndex]);
                context.Write("\"\"");
                source = source[(quoteIndex + 1)..];
            }

            context.Write("\"");
            return context.Finish(out charsWritten);
        }
    }
}