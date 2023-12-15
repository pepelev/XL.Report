namespace XL.Report.Styles.Fills;

public sealed class SolidFill : Fill, IEquatable<SolidFill>
{
    public SolidFill(Color color)
    {
        Color = color;
    }

    public Color Color { get; }

    public bool Equals(SolidFill? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Color.Equals(other.Color);
    }

    public override T Accept<T>(Visitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is SolidFill other && Equals(other);

    public override int GetHashCode() => Color.GetHashCode();
}