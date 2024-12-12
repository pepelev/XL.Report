namespace XL.Report.Contents;

public abstract class Content
{
    public static Content Blank { get; } = new BlankContent();

    public abstract void Write(Xml xml);

    private sealed class BlankContent : Content
    {
        public override void Write(Xml xml)
        {
        }
    }
}