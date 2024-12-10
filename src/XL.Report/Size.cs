using System.Diagnostics.Contracts;

namespace XL.Report;

public readonly struct Size(int width, int height) : IEquatable<Size>
{
    public override string ToString() => $"{Width}:{Height}";

    public static Size Cell => new(1, 1);
    public static Size Empty => new(0, 0);

    public static bool operator ==(Size a, Size b) => a.Equals(b);
    public static bool operator !=(Size a, Size b) => !(a == b);

    public bool IsCell => this == Cell;

    public int Width => width;
    public int Height => height;

    public bool IsDegenerate => Width < 0 || Height < 0;
    public bool HasArea => Width > 0 && Height > 0;

    public bool Equals(Size other)
    {
        return Width == other.Width && Height == other.Height;
    }

    public override bool Equals(object? obj)
    {
        if (ReferenceEquals(null, obj))
            return false;

        return obj is Size range && Equals(range);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            return (Width * 397) ^ Height;
        }
    }

    [Pure]
    public long GetArea()
    {
        return (long)Width * Height;
    }

    [Pure]
    public Offset AsOffset() => new(Width, Height);

    [Pure]
    public Size MultiplyWidth(int factor) => new(Width * factor, Height);

    [Pure]
    public Size MultiplyHeight(int factor) => new(Width, Height * factor);

    [Pure]
    public bool Contains(Size size) => size.Width <= Width && size.Height <= Height;

    public static Size operator *(Size size, int factor) => new(size.Width * factor, size.Height * factor);
}