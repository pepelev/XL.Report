namespace XL.Report;

public sealed class InlineString : Content
{
    private readonly string content;

    public InlineString(string content)
    {
        this.content = content;
    }

    public override void Write(Xml xml)
    {
        Write(xml, content);
    }

    internal static void Write(Xml xml, string content)
    {
        // todo extract to consts
        xml.WriteAttribute("t", "inlineStr");
        {
            using (xml.WriteStartElement("is"))
            using (xml.WriteStartElement("t"))
            {
                xml.WriteValue(content);
            }
        }
    }
}