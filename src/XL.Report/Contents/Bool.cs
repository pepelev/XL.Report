namespace XL.Report.Contents;

public sealed class Bool : Content
{
    private readonly bool content;

    private Bool(bool content)
    {
        this.content = content;
    }

    public static Content True { get; } = new Bool(true);
    public static Content False { get; } = new Bool(false);

    public static Content From(bool value) => value
        ? True
        : False;

    public override void Write(Xml xml)
    {
        xml.WriteAttribute("t", "b");
        using (xml.WriteStartElement("v"))
        {
            xml.WriteValue(content ? 1 : 0);
        }
    }
}