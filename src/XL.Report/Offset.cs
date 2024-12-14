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

namespace XL.Report;

public readonly struct Offset(int x, int y) : IEquatable<Offset>
{
    public static Offset Zero => new(0, 0);

    public int X => x;
    public int Y => y;

    public bool Equals(Offset other) => X == other.X && Y == other.Y;
    public override bool Equals(object? obj) => obj is Offset offset && Equals(offset);
    public override int GetHashCode() => unchecked((X * 397) ^ Y);
    public static Offset operator +(Offset a, Offset b) => new(a.X + b.X, a.Y + b.Y);
    public static Location operator +(Location location, Offset b) => new(location.X + b.X, location.Y + b.Y);
    public static Location operator -(Location location, Offset offset) => new(location.X - offset.X, location.Y - offset.Y);
    public static Size operator -(Size size, Offset offset) => new(size.Width - offset.X, size.Height - offset.Y);
    public override string ToString() => $"{X}, {Y}";
}