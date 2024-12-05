namespace XL.Report.Styles.Fills;

public sealed class PatternFill : Fill, IEquatable<PatternFill>
{
    public PatternFill(Pattern pattern, Color? color, Color? background = null)
    {
        Pattern = pattern;
        Color = color;
        Background = background;
    }

    public Color? Color { get; }
    public Color? Background { get; }
    public Pattern Pattern { get; }

    public bool Equals(PatternFill? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Nullable.Equals(Color, other.Color) &&
               Nullable.Equals(Background, other.Background) &&
               Pattern.Equals(other.Pattern);
    }

    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is PatternFill other && Equals(other);
    public override T Accept<T>(Visitor<T> visitor) => visitor.Visit(this);

    public override void Write(Xml xml)
    {
        using (xml.WriteStartElement(XlsxStructure.Styles.Fills.Fill))
        using (xml.WriteStartElement(XlsxStructure.Styles.Fills.Pattern))
        {
            xml.WriteAttribute(XlsxStructure.Styles.Fills.PatternType, Pattern);

            if (Color is { } color)
            {
                using (xml.WriteStartElement("fgColor"))
                {
                    xml.WriteAttribute("rgb", color.ToRgbHex());
                }
            }

            if (Background is { } background)
            {
                using (xml.WriteStartElement("bgColor"))
                {
                    xml.WriteAttribute("rgb", background.ToRgbHex());
                }
            }
        }
    }

    public override int GetHashCode() => HashCode.Combine(Color, Background, Pattern);
    public override string ToString() => $"{Pattern}:({PrintColor(Color)}, {PrintColor(Background)})";
    private static string PrintColor(Color? color) => color?.ToString() ?? "null";
}