using System.Xml;

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

    public override void Write(XmlWriter xml)
    {
        Content innerContent = sharedStrings.TryRegister(content) is { } id
            ? new SharedString.ById(id)
            : new InlineString(content);

        innerContent.Write(xml);
    }
}