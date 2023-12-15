using XL.Report.Styles;

namespace XL.Report;


public abstract class SheetWindow
{
    public abstract Range Range { get; }
    public abstract void Place(Content content, StyleId? styleId);
    public abstract void Merge(Size size, Content content, StyleId? styleId);
    public abstract void PushReduce(Offset offset, Size? newSize = null);
    public abstract void PopReduce();
}