using XL.Report.Styles;

namespace XL.Report;

public abstract class SheetWindow
{
    public abstract Range Range { get; }
    public abstract void TouchRow(int y);
    public abstract void Place(Content content, StyleId? styleId);
    public abstract void Merge(Size size, Content content, StyleId? styleId);
    public abstract IDisposable Reduce(Reduction reduction);

    public IDisposable Reduce(Offset offset, Size? newSize = null)
    {
        return Reduce(new Reduction(offset, newSize));
    }
}