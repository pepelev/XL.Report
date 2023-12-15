namespace XL.Report.Styles;

public sealed class Alignment : IEquatable<Alignment>
{
    public Alignment(
        HorizontalAlignment horizontal,
        VerticalAlignment vertical,
        bool wrapText)
    {
        Horizontal = horizontal;
        Vertical = vertical;
        WrapText = wrapText;
    }

    public static Alignment Default { get; } = new(
        HorizontalAlignment.Left,
        VerticalAlignment.Bottom,
        false
    );

    public HorizontalAlignment Horizontal { get; }
    public VerticalAlignment Vertical { get; }
    public bool WrapText { get; }

    public bool Equals(Alignment? other)
    {
        if (ReferenceEquals(null, other))
            return false;
        if (ReferenceEquals(this, other))
            return true;

        return Horizontal == other.Horizontal &&
               Vertical == other.Vertical &&
               WrapText == other.WrapText;
    }

    public override bool Equals(object? obj)
    {
        if (ReferenceEquals(null, obj))
            return false;
        if (ReferenceEquals(this, obj))
            return true;

        return obj is Alignment alignment && Equals(alignment);
    }

    public override string ToString()
    {
        return $"{Horizontal} {Vertical}{(WrapText ? " Wrap" : "")}";
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = (int)Horizontal;
            hashCode = (hashCode * 397) ^ (int)Vertical;
            hashCode = (hashCode * 397) ^ WrapText.GetHashCode();
            return hashCode;
        }
    }
}