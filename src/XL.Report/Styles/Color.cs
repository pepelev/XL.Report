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

    // todo span formattable
    public string ToRGBHex() => $"{Red:X2}{Green:X2}{Blue:X2}";
    public override string ToString() => $"#{ToRGBHex()}";
}