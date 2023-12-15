namespace XL.Report.Styles;

// todo union with Color
public readonly struct ColorWithAlpha : IEquatable<ColorWithAlpha>
{
    public byte Red { get; }
    public byte Green { get; }
    public byte Blue { get; }
    public byte Alpha { get; }

    public ColorWithAlpha(byte red, byte green, byte blue, byte alpha = byte.MaxValue)
    {
        Red = red;
        Green = green;
        Blue = blue;
        Alpha = alpha;
    }

    public bool Equals(ColorWithAlpha other)
        => Red == other.Red && Green == other.Green && Blue == other.Blue && Alpha == other.Alpha;

    public override bool Equals(object? obj) => obj is ColorWithAlpha other && Equals(other);
    public override int GetHashCode() => HashCode.Combine(Red, Green, Blue, Alpha);
    public static bool operator ==(ColorWithAlpha left, ColorWithAlpha right) => left.Equals(right);
    public static bool operator !=(ColorWithAlpha left, ColorWithAlpha right) => !left.Equals(right);
    public string ToARGBHex() => $"{Alpha:X2}{Red:X2}{Green:X2}{Blue:X2}";
    public override string ToString() => $"#{ToARGBHex()}";
}