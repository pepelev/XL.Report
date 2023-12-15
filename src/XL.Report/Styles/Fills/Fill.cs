using System.Xml;

namespace XL.Report.Styles.Fills;

public abstract partial class Fill
{
    public static Fill No { get; } = new NoFill();

    public abstract T Accept<T>(Visitor<T> visitor);
    public abstract void Write(XmlWriter xml);

    public static void Write(XmlWriter xml, IEnumerable<Fill> fills)
    {
        xml.WriteStartElement(XlsxStructure.Styles.Fills.Collection);
        {
            foreach (var fill in fills)
            {
                fill.Write(xml);
            }
        }
        xml.WriteEndElement();
    }
}