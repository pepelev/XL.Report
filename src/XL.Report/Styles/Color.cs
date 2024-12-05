using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;

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
            return FormatContext.Start
                .Write(ref destination, source.Red, "X2", CultureInfo.InvariantCulture)
                .Write(ref destination, source.Green, "X2", CultureInfo.InvariantCulture)
                .Write(ref destination, source.Blue, "X2", CultureInfo.InvariantCulture)
                .Deconstruct(out charsWritten);
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