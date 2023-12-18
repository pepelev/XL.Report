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

        public override void Write(Xml xml)
        {
            Write(xml, id);
        }

        internal static void Write(Xml xml, SharedStringId id)
        {
            xml.WriteAttribute("t", "s");
            {
                using (xml.WriteStartElement("v"))
                {
                    xml.WriteValueSpan(id.Index);
                }
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

        public override void Write(Xml xml)
        {
            var id = sharedStrings.ForceRegister(content);
            ById.Write(xml, id);
        }
    }
}