using XL.Report.Styles.Fills;

namespace XL.Report.Styles;

public sealed class Appearance : IEquatable<Appearance>
{
    public Appearance(
        Alignment alignment,
        Font font,
        Fill fill,
        Borders borders)
    {
        Alignment = alignment ?? throw new ArgumentNullException(nameof(alignment));
        Font = font ?? throw new ArgumentNullException(nameof(font));
        Fill = fill ?? throw new ArgumentNullException(nameof(fill));
        Borders = borders ?? throw new ArgumentNullException(nameof(borders));
    }

    public Alignment Alignment { get; }
    public Font Font { get; }
    public Fill Fill { get; }
    public Borders Borders { get; }

    public static Appearance Default { get; } = new(
        Alignment.Default,
        Font.Default,
        NoFill.Singleton,
        Borders.None
    );

    public bool Equals(Appearance? other)
    {
        if (ReferenceEquals(null, other))
            return false;
        if (ReferenceEquals(this, other))
            return true;

        return Alignment.Equals(other.Alignment) &&
               Font.Equals(other.Font) &&
               Fill.Equals(other.Fill) &&
               Borders.Equals(other.Borders);
    }

    public override bool Equals(object? obj)
    {
        if (ReferenceEquals(null, obj))
            return false;
        if (ReferenceEquals(this, obj))
            return true;

        return obj is Appearance appearance && Equals(appearance);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = Alignment.GetHashCode();
            hashCode = (hashCode * 397) ^ Font.GetHashCode();
            hashCode = (hashCode * 397) ^ Fill.GetHashCode();
            hashCode = (hashCode * 397) ^ Borders.GetHashCode();
            return hashCode;
        }
    }
}