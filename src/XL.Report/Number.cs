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

    public override void Write(Xml xml)
    {
        using (xml.WriteStartElement(value))
        {
            xml.WriteValueSpan(content);
        }
    }
}