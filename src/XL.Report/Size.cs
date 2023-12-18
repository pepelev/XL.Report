namespace XL.Report;

// todo make int
public readonly struct Size : IEquatable<Size>
{
    public override string ToString()
    {
        return $"{Width}:{Height}";
    }

    public static Size Cell => new(1, 1);
    public static Size Empty => new(0, 0);

    public static bool operator ==(Size a, Size b) => a.Equals(b);
    public static bool operator !=(Size a, Size b) => !(a == b);

    public bool IsCell => this == Cell;

    public Size(int width, int height)
    {
        Width = width;
        Height = height;
    }

    public int Width { get; }
    public int Height { get; }

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

    public long GetArea()
    {
        return (long)Width * Height;
    }

    public bool Contains(Size size) => size.Width <= Width && size.Height <= Height;
}