using System.Xml;

namespace XL.Report;

public sealed class Number : Content
{
    // todo extract to consts
    private const string value = "v";

    private readonly long content;

    public Number(long content)
    {
        this.content = content;
    }

    public override void Write(XmlWriter xml)
    {
        xml.WriteStartElement(value);
        {
            xml.WriteValue(content);
        }
        xml.WriteEndElement();
    }
}