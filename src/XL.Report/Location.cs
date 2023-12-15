﻿namespace XL.Report;

using System;
using System.Globalization;
using System.Text.RegularExpressions;

public readonly struct Location : IEquatable<Location>, IParsable<Location>
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

    public int X { get; }
    public int Y { get; }

    public Location(int x, int y)
    {
        X = x;
        Y = y;
    }

    public static bool TryParse(string? value, IFormatProvider? formatProvider, out Location result)
    {
        if (value == null)
            return Fail(out result);

        // todo to  static field
        var flags = RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant;
        var regex = new Regex(@"^\s*(?<X>[a-z]{1,3})(?<Y>[0-9]{1,7})\s*$", flags);
        if (regex.Match(value) is { Success: true } match)
        {
            var x = 0;
            foreach (var @char in match.Groups["X"].ValueSpan)
            {
                x = x * AlphabetPower + char.ToUpperInvariant(@char) - ZeroLetter;
            }

            var y = int.Parse(match.Groups["Y"].ValueSpan, CultureInfo.InvariantCulture);
            result = new Location(x, y);
            return result.IsCorrect();
        }

        result = default;
        return false;
    }

    private static bool Fail(out Location result)
    {
        result = default;
        return false;
    }

    public static Location Parse(string value, IFormatProvider? formatProvider)
    {
        return TryParse(value, formatProvider, out var result)
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

    public override string ToString() => IsCorrect()
        ? PrintCorrectValue()
        : $"{X}, {Y}";

    private string PrintCorrectValue() => PrintX() + PrintY();
    private string PrintX() => PrintX((uint) X);

    private static string PrintX(uint x)
    {
        if (x == 0)
            return "";

        x--;
        var reminder = x / AlphabetPower;
        var current = x % AlphabetPower;
        return PrintX(reminder) + GetChar(current);
    }

    private static char GetChar(uint code) => (char) (code + FirstLetter);
    private string PrintY() => Y.ToString(CultureInfo.InvariantCulture);
    public static bool operator ==(Location a, Location b) => a.Equals(b);
    public static bool operator !=(Location a, Location b) => !(a == b);
}