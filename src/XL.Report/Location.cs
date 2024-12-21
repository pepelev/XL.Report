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

using System;
using System.Globalization;
using System.Text.RegularExpressions;

public readonly partial struct Location(int x, int y) : IEquatable<Location>, ISpanFormattable, IParsable<Location>
{
    private const char ZeroLetter = (char) (FirstLetter - 1);
    private const char FirstLetter = 'A';
    private const char LastLetter = 'Z';
    private const int AlphabetPower = LastLetter - ZeroLetter;

    public const int MinX = 1;
    public const int MinY = 1;

    // Most right bottom cell is XFD1048576
    public const int MaxX =
        ('X' - ZeroLetter) * AlphabetPower * AlphabetPower + // X
        ('F' - ZeroLetter) * AlphabetPower + // F
        'D' - ZeroLetter; // D
    public const int MaxY = 1048576;

    public const int XLength = MaxX - MinX + 1;
    public const int YLength = MaxY - MinY + 1;

    public int X { get; } = x;
    public int Y { get; } = y;

    public static bool TryParse(string? value, IFormatProvider? formatProvider, out Location result)
        => TryParse(value, out result);

    public static bool TryParse(string? value, out Location result)
    {
        if (value == null)
            return Fail(out result);

        if (Pattern().Match(value) is { Success: true } match)
        {
            var x = 0;
            foreach (var @char in match.Groups["X"].ValueSpan)
            {
                x = x * AlphabetPower + char.ToUpperInvariant(@char) - ZeroLetter;
            }

            var yParsed = int.TryParse(
                match.Groups["Y"].ValueSpan,
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out var y
            );
            result = new Location(x, y);
            return yParsed && result.IsCorrect();
        }

        result = default;
        return false;

        bool Fail(out Location result)
        {
            result = default;
            return false;
        }
    }

    public static Location Parse(string value, IFormatProvider? formatProvider) => Parse(value);

    public static Location Parse(string value)
    {
        return TryParse(value, out var result)
            ? result
            : throw new FormatException($"Value '{value}' has incorrect {nameof(Location)} format");
    }

    public bool Equals(Location other) => X == other.X && Y == other.Y;

    public override bool Equals(object? obj)
    {
        if (ReferenceEquals(null, obj))
            return false;

        return obj is Location location && Equals(location);
    }

    public override int GetHashCode() => unchecked((X * 397) ^ Y);

    private bool IsCorrect() => MinX <= X && X <= MaxX &&
                                MinY <= Y && Y <= MaxY;

    private static bool TryFormatCorrectX(uint x, Span<char> destination, out int charsWritten)
    {
        var xSize = x switch
        {
            <= AlphabetPower => 1,
            <= AlphabetPower * AlphabetPower + AlphabetPower => 2,
            _ => 3
        };
        if (destination.Length < xSize)
        {
            charsWritten = 0;
            return false;
        }

        var offset = xSize;
        while (--offset >= 0)
        {
            x--;
            var current = x % AlphabetPower;
            x /= AlphabetPower;
            destination[offset] = GetChar(current);
        }

        charsWritten = xSize;
        return true;
    }

    public string AsString() => FormatContext.ToString(this);

    public override string ToString() => AsString();
    public string ToString(string? format, IFormatProvider? formatProvider) => ToString();

    public bool TryFormat(
        Span<char> destination,
        out int charsWritten,
        ReadOnlySpan<char> format,
        IFormatProvider? provider)
    {
        if (IsCorrect())
        {
            var context = new FormatContext(destination);
            context.Write((uint)X, TryFormatCorrectX);
            context.Write(Y, ReadOnlySpan<char>.Empty, CultureInfo.InvariantCulture);
            return context.Finish(out charsWritten);
        }
        else
        {
            var context = new FormatContext(destination);
            context.Write(X, ReadOnlySpan<char>.Empty, CultureInfo.InvariantCulture);
            context.Write(", ");
            context.Write(Y, ReadOnlySpan<char>.Empty, CultureInfo.InvariantCulture);
            return context.Finish(out charsWritten);
        }
    }

    private static char GetChar(uint code) => (char) (code + FirstLetter);
    public static bool operator ==(Location a, Location b) => a.Equals(b);
    public static bool operator !=(Location a, Location b) => !(a == b);

    public readonly struct Reference : ISpanFormattable
    {
        private const string LockSign = "$";

        public Reference(Location location, bool columnLocked, bool rowLocked)
        {
            if (!Location.IsCorrect())
            {
                throw new ArgumentException();
            }

            Location = location;
            ColumnLocked = columnLocked;
            RowLocked = rowLocked;
        }

        public Location Location { get; }
        public bool ColumnLocked { get; }
        public bool RowLocked { get; }

        public string ToString(string? format, IFormatProvider? formatProvider) => FormatContext.ToString(this);
        public override string ToString() => ToString(null, null);

        public bool TryFormat(
            Span<char> destination,
            out int charsWritten,
            ReadOnlySpan<char> format,
            IFormatProvider? provider)
        {
            var context = new FormatContext(destination);
            if (ColumnLocked)
            {
                context.Write(LockSign);
            }

            context.Write((uint)Location.X, TryFormatCorrectX);

            if (RowLocked)
            {
                context.Write(LockSign);
            }

            context.Write(
                Location.Y,
                ReadOnlySpan<char>.Empty,
                CultureInfo.InvariantCulture
            );

            return context.Finish(out charsWritten);
        }
    }

    [GeneratedRegex(
        @"^\s*(?<X>[a-z]{1,3})(?<Y>[0-9]{1,7})\s*$",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant
    )]
    private static partial Regex Pattern();
}