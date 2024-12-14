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

namespace XL.Report.Styles;

// todo make struct record
public sealed class Border : IEquatable<Border>, IBorder
{
    public Border(BorderStyle style, Color? color = null)
    {
        Color = color;
        Style = style;
    }

    public Color? Color { get; }
    public BorderStyle Style { get; }

    public bool Equals(Border? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Nullable.Equals(Color, other.Color) && Style.Equals(other.Style);
    }

    public override string ToString()
    {
        var parts = new[]
        {
            Style.ToString(),
            Color?.ToString()
        }.Where(part => !string.IsNullOrWhiteSpace(part));
        return string.Join(' ', parts);
    }

    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is Border other && Equals(other);

    public override int GetHashCode() => HashCode.Combine(Color, Style);
}