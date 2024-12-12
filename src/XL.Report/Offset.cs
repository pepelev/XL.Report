namespace XL.Report;

public readonly struct Offset(int x, int y) : IEquatable<Offset>
{
    public static Offset Zero => new(0, 0);

    public int X => x;
    public int Y => y;

    public bool Equals(Offset other) => X == other.X && Y == other.Y;
    public override bool Equals(object? obj) => obj is Offset offset && Equals(offset);
    public override int GetHashCode() => unchecked((X * 397) ^ Y);
    public static Offset operator +(Offset a, Offset b) => new(a.X + b.X, a.Y + b.Y);
    public static Location operator +(Location location, Offset b) => new(location.X + b.X, location.Y + b.Y);
    public static Location operator -(Location location, Offset offset) => new(location.X - offset.X, location.Y - offset.Y);
    public static Size operator -(Size size, Offset offset) => new(size.Width - offset.X, size.Height - offset.Y);
    public override string ToString() => $"{X}, {Y}";
}