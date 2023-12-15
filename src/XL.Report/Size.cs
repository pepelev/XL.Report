namespace XL.Report;

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

    public Size(uint width, uint height)
    {
        Width = width;
        Height = height;
    }

    public uint Width { get; }
    public uint Height { get; }

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
            return ((int) Width * 397) ^ (int) Height;
        }
    }

    public ulong GetArea()
    {
        return (ulong)Width * Height;
    }
}