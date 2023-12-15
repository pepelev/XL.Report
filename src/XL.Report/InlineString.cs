using System.Xml;

namespace XL.Report;

public sealed class InlineString : Content
{
    private readonly string content;

    public InlineString(string content)
    {
        this.content = content;
    }

    public override void Write(XmlWriter xml)
    {
        // todo extract to consts
        xml.WriteAttributeString("t", "inlineStr");
        {
            xml.WriteStartElement("is");
            {
                xml.WriteStartElement("t");
                {
                    xml.WriteValue(content);
                }
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }
    }
}