namespace XL.Report.Styles.Fills;

public sealed class PatternFill : Fill, IEquatable<PatternFill>
{
    public PatternFill(Pattern pattern)
        : this(pattern, Color.Auto)
    {
    }

    public PatternFill(Pattern pattern, Color color, Color? background = null)
    {
        Pattern = pattern;
        Color = color;
        Background = background;
    }

    public Color Color { get; }
    public Color? Background { get; }
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

    private static string PrintColor(Color? color)
    {
        return color?.ToString() ?? "null";
    }
}