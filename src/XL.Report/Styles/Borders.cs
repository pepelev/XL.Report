using System.Text;

namespace XL.Report.Styles;

public class Borders : IEquatable<Borders>
{
    public Borders(
        Border left,
        Border right,
        Border top,
        Border bottom,
        DiagonalBorders diagonal)
    {
        Left = left ?? throw new ArgumentNullException(nameof(left));
        Right = right ?? throw new ArgumentNullException(nameof(right));
        Top = top ?? throw new ArgumentNullException(nameof(top));
        Bottom = bottom ?? throw new ArgumentNullException(nameof(bottom));
        Diagonal = diagonal ?? throw new ArgumentNullException(nameof(diagonal));
    }

    public Border Left { get; }
    public Border Right { get; }
    public Border Top { get; }
    public Border Bottom { get; }
    public DiagonalBorders Diagonal { get; }

    public static Borders None { get; } = new(
        Border.None,
        Border.None,
        Border.None,
        Border.None,
        DiagonalBorders.None
    );

    public void Write(Xml xml)
    {
        using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Border))
        {
            // todo
            xml.WriteStartElement(XlsxStructure.Styles.Borders.Left);
            xml.WriteEndElement();

            xml.WriteStartElement(XlsxStructure.Styles.Borders.Right);
            xml.WriteEndElement();

            xml.WriteStartElement(XlsxStructure.Styles.Borders.Top);
            xml.WriteEndElement();

            xml.WriteStartElement(XlsxStructure.Styles.Borders.Bottom);
            xml.WriteEndElement();

            xml.WriteStartElement(XlsxStructure.Styles.Borders.Diagonal);
            xml.WriteEndElement();
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
            return false;
        if (ReferenceEquals(this, other))
            return true;

        return Equals(Left, other.Left) &&
               Equals(Right, other.Right) &&
               Equals(Top, other.Top) &&
               Equals(Bottom, other.Bottom) &&
               Equals(Diagonal, other.Diagonal);
    }

    public override bool Equals(object? obj)
    {
        if (ReferenceEquals(null, obj))
            return false;
        if (ReferenceEquals(this, obj))
            return true;
        if (obj.GetType() != GetType())
            return false;

        return Equals((Borders)obj);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = Left.GetHashCode();
            hashCode = (hashCode * 397) ^ Right.GetHashCode();
            hashCode = (hashCode * 397) ^ Top.GetHashCode();
            hashCode = (hashCode * 397) ^ Bottom.GetHashCode();
            hashCode = (hashCode * 397) ^ Diagonal.GetHashCode();
            return hashCode;
        }
    }

    public override string ToString()
    {
        if (Equals(None))
            return "None";

        var builder = new StringBuilder(128);
        var prependComma = false;
        AppendIfNotNone(Left, nameof(Left));
        AppendIfNotNone(Right, nameof(Right));
        AppendIfNotNone(Top, nameof(Top));
        AppendIfNotNone(Bottom, nameof(Bottom));

        if (!Diagonal.Equals(DiagonalBorders.None))
            WriteBorder(Diagonal, nameof(Diagonal));

        return builder.ToString();

        void AppendIfNotNone(Border border, string name)
        {
            if (border.Equals(Border.None))
                return;

            WriteBorder(border, name);
        }

        void WriteBorder(object border, string name)
        {
            if (prependComma)
                builder.Append(", ");

            builder.Append(name);
            builder.Append(": ");
            builder.Append(border);
            prependComma = true;
        }
    }
}