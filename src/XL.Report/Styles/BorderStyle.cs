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

namespace XL.Report.Styles;

public sealed class BorderStyle : IEquatable<BorderStyle>, ISpanFormattable
{
    public BorderStyle(string value)
    {
        Value = value;
    }

    public string Value { get; }

    public static BorderStyle Thick { get; } = new("thick");
    public static BorderStyle MediumDashDotDot { get; } = new("mediumDashDotDot");
    public static BorderStyle Dashed { get; } = new("dashed"); //
    public static BorderStyle Hair { get; } = new("hair"); // 
    public static BorderStyle Dotted { get; } = new("dotted"); // 
    public static BorderStyle DashDotDot { get; } = new("dashDotDot"); //
    public static BorderStyle DashDot { get; } = new("dashDot"); //
    public static BorderStyle Thin { get; } = new("thin"); //
    public static BorderStyle SlantDashDot { get; } = new("slantDashDot");
    public static BorderStyle MediumDashDot { get; } = new("mediumDashDot");

    public bool Equals(BorderStyle? other)
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

    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is BorderStyle other && Equals(other);
    public override int GetHashCode() => Value.GetHashCode();
    public override string ToString() => ToString(null, null);
}