namespace XL.Report;

public sealed class String : Content
{
    private readonly string content;
    private readonly SharedStrings sharedStrings;

    public String(string content, SharedStrings sharedStrings)
    {
        this.content = content;
        this.sharedStrings = sharedStrings;
    }

    public override void Write(Xml xml)
    {
        if (sharedStrings.TryRegister(content) is { } id)
        {
            SharedString.ById.Write(xml, id);
        }
        else
        {
            InlineString.Write(xml, content);
        }
    }
}