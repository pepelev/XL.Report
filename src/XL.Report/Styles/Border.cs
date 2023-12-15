namespace XL.Report.Styles;

public sealed class Border : IEquatable<Border>, IBorder
{
    public Border(BorderStyle style, ColorWithAlpha? color = null)
    {
        Color = color;
        Style = style;
    }

    public static Border None { get; } = new(BorderStyle.None);

    public ColorWithAlpha? Color { get; }
    public BorderStyle Style { get; }

    public bool Equals(Border? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Color.Equals(other.Color) && Style == other.Style;
    }

    public override string ToString()
    {
        if (Equals(None))
        {
            return nameof(None);
        }

        if (Style == BorderStyle.None)
        {
            return Color.ToString();
        }

        return $"{Style} {Color}";
    }

    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is Border other && Equals(other);

    public override int GetHashCode() => HashCode.Combine(Color, (int)Style);
}