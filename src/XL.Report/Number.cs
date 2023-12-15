using System.Xml;

namespace XL.Report;

public sealed class Number : Content
{
    private const string cell = "c";
    private const string reference = "r";
    private const string value = "v";

    private readonly long content;

    public Number(long content)
    {
        this.content = content;
    }

    public override void Write(XmlWriter xml, Location location)
    {
        xml.WriteStartElement(cell);
        xml.WriteStartAttribute(reference);
        // todo location can be incorrect
        xml.WriteValue(location.ToString());
        xml.WriteEndAttribute();
        {
            xml.WriteStartElement(value);
            {
                xml.WriteValue(content);
            }
            xml.WriteEndElement();
        }
        xml.WriteEndElement();
    }
}