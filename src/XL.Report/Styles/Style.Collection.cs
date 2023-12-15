namespace XL.Report.Styles;

public sealed partial class Style
{
    public abstract class Collection
    {
        public abstract StyleId Register(Style style);
    }
}