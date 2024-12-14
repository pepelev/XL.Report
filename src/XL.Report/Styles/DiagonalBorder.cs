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

public sealed class DiagonalBorders : IEquatable<DiagonalBorders>, IBorder
{
    public DiagonalBorders(BorderStyle style, Color? color, bool up, bool down)
    {
        if (!up && !down)
        {
            throw new ArgumentException(
                $"at least one of {nameof(up)}, {nameof(down)} must be true"
            );
        }

        Up = up;
        Down = down;
        Color = color;
        Style = style;
    }

    public bool Up { get; }
    public bool Down { get; }

    public Color? Color { get; }
    public BorderStyle Style { get; }

    public bool Equals(DiagonalBorders? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Up == other.Up &&
               Down == other.Down &&
               Nullable.Equals(Color, other.Color) &&
               Style.Equals(other.Style);
    }

    public override string ToString()
    {
        var diagonals = (UpDiagonal: Up, DownDiagonal: Down) switch
        {
            (true, true) => "Cross",
            (true, false) => "Up",
            (false, true) => "Down",
            _ => "bug"
        };

        var parts = new[]
        {
            diagonals,
            Style.ToString(),
            Color?.ToString()
        }.Where(part => !string.IsNullOrWhiteSpace(part));
        return string.Join(' ', parts);
    }

    public override bool Equals(object? obj) => 
        ReferenceEquals(this, obj) || obj is DiagonalBorders other && Equals(other);

    public override int GetHashCode() => HashCode.Combine(Up, Down, Color, Style);
}