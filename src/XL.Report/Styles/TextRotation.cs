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

public readonly struct TextRotation : IEquatable<TextRotation>, ISpanFormattable
{
    public static TextRotation None => new(0);

    public int Value { get; }

    /// <param name="value">
    ///     0..90 match 0..90 degrees for polar angle
    ///     91..180 match -1..-90 degrees for polar angle
    /// </param>
    /// <exception cref="ArgumentOutOfRangeException">value outside of [0, 180] bounds</exception>
    public TextRotation(int value)
    {
        var correct = 0 <= value && value <= 180;
        if (!correct)
        {
            throw new ArgumentOutOfRangeException(
                nameof(value),
                value,
                "must be between [0, 180]"
            );
        }

        Value = value;
    }

    public static TextRotation FromPolarAngle(int degree)
    {
        degree %= 360;
        var correct = -90 <= degree && degree <= 90;
        if (!correct)
        {
            throw new ArgumentOutOfRangeException(
                nameof(degree),
                degree,
                "must be between [-90, 90]"
            );
        }

        return degree < 0
            ? new TextRotation(90 - degree)
            : new TextRotation(degree);
    }

    public override string ToString() => ToString(null, null);

    public string ToString(string? format, IFormatProvider? formatProvider) =>
        Value.ToString(CultureInfo.InvariantCulture);

    public bool TryFormat(
        Span<char> destination,
        out int charsWritten,
        ReadOnlySpan<char> format,
        IFormatProvider? provider)
    {
        return Value.TryFormat(
            destination,
            out charsWritten,
            ReadOnlySpan<char>.Empty,
            CultureInfo.InvariantCulture
        );
    }

    public bool Equals(TextRotation other) => Value == other.Value;
    public override bool Equals(object? obj) => obj is TextRotation other && Equals(other);
    public override int GetHashCode() => Value;
    public static bool operator ==(TextRotation left, TextRotation right) => left.Equals(right);
    public static bool operator !=(TextRotation left, TextRotation right) => !left.Equals(right);
}