using System.Xml;

namespace XL.Report;

public static class SharedString
{
    public sealed class ById : Content
    {
        private readonly SharedStringId id;

        public ById(SharedStringId id)
        {
            this.id = id;
        }

        public override void Write(XmlWriter xml)
        {
            xml.WriteAttributeString("t", "s");
            {
                xml.WriteStartElement("v");
                {
                    xml.WriteValue(id.Index);
                }
                xml.WriteEndElement();
            }
        }
    }

    public sealed class Force : Content
    {
        private readonly string content;
        private readonly SharedStrings sharedStrings;

        public Force(string content, SharedStrings sharedStrings)
        {
            this.content = content;
            this.sharedStrings = sharedStrings;
        }

        public override void Write(XmlWriter xml)
        {
            var id = sharedStrings.ForceRegister(content);
            var innerContent = new ById(id);
            innerContent.Write(xml);
        }
    }
}