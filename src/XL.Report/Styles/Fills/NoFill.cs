using System.Xml;

namespace XL.Report.Styles.Fills;

public sealed class NoFill : Fill, IEquatable<NoFill>
{
    public bool Equals(NoFill? other)
    {
        return other != null;
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as NoFill);
    }

    public override int GetHashCode()
    {
        return 1299827;
    }

    public override string ToString()
    {
        return "No";
    }

    public override T Accept<T>(Visitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    public override void Write(XmlWriter xml)
    {
        xml.WriteStartElement(XlsxStructure.Styles.Fills.Fill);
        {
            xml.WriteStartElement(XlsxStructure.Styles.Fills.Pattern);
            xml.WriteAttributeString(XlsxStructure.Styles.Fills.PatternType, "none");
            xml.WriteEndElement();
        }
        xml.WriteEndElement();
    }
}