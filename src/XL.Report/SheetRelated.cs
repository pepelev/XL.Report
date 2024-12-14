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

using XL.Report.Auxiliary;

namespace XL.Report;

public readonly struct SheetRelated<T>(string sheetName, T value)
{
    public string SheetName => sheetName;
    public T Value => value;
}

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