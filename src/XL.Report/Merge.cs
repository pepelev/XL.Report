using XL.Report.Contents;
using XL.Report.Styles;

namespace XL.Report;

public readonly struct Merge(Size size, Content content, StyleId? styleId = null) : IUnit<Range>
{
    public Range Write(SheetWindow window)
    {
        using (window.Reduce(Offset.Zero, size))
        {
            var merge = new WholeWindow(content, styleId);
            return merge.Write(window);
        }
    }

    public readonly struct WholeWindow(Content content, StyleId? styleId = null) : IUnit<Range>
    {
        public Range Write(SheetWindow window)
        {
            window.Merge(content, styleId);
            return window.Range;
        }
    }
}