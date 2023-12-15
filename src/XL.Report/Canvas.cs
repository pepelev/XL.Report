namespace XL.Report;

public abstract class Canvas
{
    public abstract Range Range { get; }
    public abstract void Place(Offset offset, Content content, Style style);
    public abstract void Merge(Offset offset, Size size, Content content, Style style);
}