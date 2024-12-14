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

namespace XL.Report.Styles;

public readonly struct Color : IEquatable<Color>
{
    public byte Red { get; }
    public byte Green { get; }
    public byte Blue { get; }

    public Color(byte red, byte green, byte blue)
    {
        Red = red;
        Green = green;
        Blue = blue;
    }

    public bool Equals(Color other) => Red == other.Red && Green == other.Green && Blue == other.Blue;
    public override bool Equals(object? obj) => obj is Color other && Equals(other);
    public override int GetHashCode() => HashCode.Combine(Red, Green, Blue);
    public static bool operator ==(Color left, Color right) => left.Equals(right);
    public static bool operator !=(Color left, Color right) => !left.Equals(right);

    public RgbHex ToRgbHex() => new(this);
    public override string ToString() => $"#{ToRgbHex()}";

    public readonly struct RgbHex(Color source) : ISpanFormattable
    {
        private readonly Color source = source;

        public string ToString(string? format, IFormatProvider? formatProvider) => ToString();

        public bool TryFormat(
            Span<char> destination,
            out int charsWritten,
            ReadOnlySpan<char> format,
            IFormatProvider? provider)
        {
            var context = new FormatContext(destination);
            context.Write(source.Red, "X2", CultureInfo.InvariantCulture);
            context.Write(source.Green, "X2", CultureInfo.InvariantCulture);
            context.Write(source.Blue, "X2", CultureInfo.InvariantCulture);
            return context.Finish(out charsWritten);
        }

        public override string ToString()
        {
            return string.Create(
                6,
                this,
                (span, @this) => @this.TryFormat(span, out _, "", CultureInfo.InvariantCulture)
            );
        }
    }
}