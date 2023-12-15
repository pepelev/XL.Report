namespace XL.Report.Styles;

public sealed partial class Style : IEquatable<Style>
{
    public Style(Appearance appearance, Format format)
    {
        Appearance = appearance ?? throw new ArgumentNullException(nameof(appearance));
        Format = format ?? throw new ArgumentNullException(nameof(format));
    }

    public static Style Default { get; } = new(
        Appearance.Default,
        Format.General
    );

    public Appearance Appearance { get; }
    public Format Format { get; }

    public bool Equals(Style? other)
    {
        if (ReferenceEquals(null, other))
            return false;
        if (ReferenceEquals(this, other))
            return true;

        return Appearance.Equals(other.Appearance) && Format.Equals(other.Format);
    }

    public override bool Equals(object? obj)
    {
        if (ReferenceEquals(null, obj))
            return false;
        if (ReferenceEquals(this, obj))
            return true;

        return obj is Style style && Equals(style);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            return (Appearance.GetHashCode() * 397) ^ Format.GetHashCode();
        }
    }
}