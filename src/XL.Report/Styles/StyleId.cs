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