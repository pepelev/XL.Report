namespace XL.Report.Styles;

public sealed class Borders : IEquatable<Borders>
{
    public Borders(
        Border? left = null,
        Border? right = null,
        Border? top = null,
        Border? bottom = null,
        DiagonalBorders? diagonal = null)
    {
        Left = left;
        Right = right;
        Top = top;
        Bottom = bottom;
        Diagonal = diagonal;
    }

    public Border? Left { get; }
    public Border? Right { get; }
    public Border? Top { get; }
    public Border? Bottom { get; }
    public DiagonalBorders? Diagonal { get; }
    public static Borders Perimeter(Border border) => new(border, border, border, border, diagonal: null);

    public static Borders None { get; } = new();

    public void Write(Xml xml)
    {
        using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Border))
        {
            if (Diagonal?.Up == true)
            {
                xml.WriteAttribute("diagonalUp", "1");
            }

            if (Diagonal?.Down == true)
            {
                xml.WriteAttribute("diagonalDown", "1");
            }

            using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Left))
            {
                WriteElement(Left);
            }

            using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Right))
            {
                WriteElement(Right);
            }

            using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Top))
            {
                WriteElement(Top);
            }

            using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Bottom))
            {
                WriteElement(Bottom);
            }

            using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Diagonal))
            {
                WriteElement(Diagonal);
            }
        }

        void WriteElement(IBorder? border)
        {
            if (border == null)
            {
                return;
            }

            xml.WriteAttribute("style", border.Style);
            if (border.Color is { } color)
            {
                using (xml.WriteStartElement("color"))
                {
                    xml.WriteAttribute("rgb", color.ToRGBHex());
                }
            }
        }
    }

    public static void Write(Xml xml, IEnumerable<Borders> borders)
    {
        using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Collection))
        {
            foreach (var border in borders)
            {
                border.Write(xml);
            }
        }
    }

    public bool Equals(Borders? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Equals(Left, other.Left) &&
               Equals(Right, other.Right) &&
               Equals(Top, other.Top) &&
               Equals(Bottom, other.Bottom) &&
               Equals(Diagonal, other.Diagonal);
    }

    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is Borders other && Equals(other);
    public override int GetHashCode() => HashCode.Combine(Left, Right, Top, Bottom, Diagonal);

    public override string ToString()
    {
        if (Equals(None))
            return "None";

        var parts = new[]
            {
                (Name: nameof(Left), Value: Left?.ToString()),
                (Name: nameof(Right), Value: Right?.ToString()),
                (Name: nameof(Top), Value: Top?.ToString()),
                (Name: nameof(Bottom), Value: Bottom?.ToString()),
                (Name: nameof(Diagonal), Value: Diagonal?.ToString()),
            }
            .Where(part => !string.IsNullOrWhiteSpace(part.Value))
            .Select(part => $"{part.Name}: {part.Value}");
        return string.Join(' ', parts);
    }
}