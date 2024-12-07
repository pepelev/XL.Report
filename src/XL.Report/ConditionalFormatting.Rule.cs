namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public abstract class Rule
    {
        public abstract void Write(Xml xml, int priority);
    }
}