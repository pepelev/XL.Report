namespace XL.Report.Styles.Fills;

public sealed class PatternFill : Fill, IEquatable<PatternFill>
{
    public PatternFill(Pattern pattern, ColorWithAlpha color, ColorWithAlpha? background = null)
    {
        Pattern = pattern;
        Color = color;
        Background = background;
    }

    public ColorWithAlpha Color { get; }
    public ColorWithAlpha? Background { get; }
    public Pattern Pattern { get; }

    public bool Equals(PatternFill? other)
    {
        if (ReferenceEquals(null, other))
            return false;
        if (ReferenceEquals(this, other))
            return true;

        return Color.Equals(other.Color) &&
               Background.Equals(other.Background) &&
               Pattern == other.Pattern;
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as PatternFill);
    }

    public override T Accept<T>(Visitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    public override void Write(Xml xml)
    {
        throw new NotImplementedException();
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = Color.GetHashCode();
            hashCode = (hashCode * 397) ^ Background.GetHashCode();
            hashCode = (hashCode * 397) ^ Pattern.GetHashCode();
            return hashCode;
        }
    }

    public override string ToString()
    {
        return $"{Pattern}:({PrintColor(Color)}, {PrintColor(Background)})";
    }

    private static string PrintColor(ColorWithAlpha? color)
    {
        return color?.ToString() ?? "null";
    }
}