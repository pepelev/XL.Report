namespace XL.Report.Styles;

public sealed class DiagonalBorders : IEquatable<DiagonalBorders>, IBorder
{
    public DiagonalBorders(BorderStyle style, ColorWithAlpha? color, bool mainDiagonal, bool antidiagonal)
    {
        MainDiagonal = mainDiagonal;
        Antidiagonal = antidiagonal;
        Color = color;
        Style = style;
    }

    public static DiagonalBorders None { get; } = new(BorderStyle.None, color: null, false, false);

    public bool MainDiagonal { get; }
    public bool Antidiagonal { get; }

    public ColorWithAlpha? Color { get; }
    public BorderStyle Style { get; }

    public bool Equals(DiagonalBorders? other)
    {
        if (ReferenceEquals(null, other))
            return false;
        if (ReferenceEquals(this, other))
            return true;

        return MainDiagonal == other.MainDiagonal &&
               Antidiagonal == other.Antidiagonal &&
               Color.Equals(other.Color) &&
               Style == other.Style;
    }

    public override string ToString()
    {
        if (Equals(None) ||
            Style == BorderStyle.None ||
            !MainDiagonal && !Antidiagonal)
            return nameof(None);

        if (MainDiagonal && Antidiagonal)
            return PrintBorder("Both");

        if (MainDiagonal)
            return PrintBorder("Main");

        return PrintBorder("Anti");

        string PrintBorder(string type)
        {
            return $"{type} {Color} {Style}";
        }
    }

    public override bool Equals(object? obj)
    {
        if (ReferenceEquals(null, obj))
            return false;
        if (ReferenceEquals(this, obj))
            return true;

        return obj is DiagonalBorders borders && Equals(borders);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = MainDiagonal.GetHashCode();
            hashCode = (hashCode * 397) ^ Antidiagonal.GetHashCode();
            hashCode = (hashCode * 397) ^ Color.GetHashCode();
            hashCode = (hashCode * 397) ^ (int)Style;
            return hashCode;
        }
    }
}