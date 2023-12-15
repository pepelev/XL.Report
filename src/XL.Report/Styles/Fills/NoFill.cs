namespace XL.Report.Styles.Fills;

public sealed class NoFill : Fill, IEquatable<NoFill>
{
    public static NoFill Singleton { get; } = new();

    public bool Equals(NoFill? other)
    {
        return other != null;
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as NoFill);
    }

    public override int GetHashCode()
    {
        return 1299827;
    }

    public override string ToString()
    {
        return "No";
    }

    public override T Accept<T>(Visitor<T> visitor)
    {
        return visitor.Visit(this);
    }
}