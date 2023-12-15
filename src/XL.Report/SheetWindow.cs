namespace XL.Report;


public abstract class SheetWindow
{
    public abstract Range Range { get; }

    // todo maybe delete offset from here
    public abstract void Place(Offset offset, Content content, StyleId styleId);
    public abstract void Merge(Offset offset, Size size, Content content, StyleId styleId);
    public abstract void PushMove(Offset offset);
    public abstract void PopMove();
}