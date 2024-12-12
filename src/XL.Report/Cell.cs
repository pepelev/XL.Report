using XL.Report.Contents;
using XL.Report.Styles;

namespace XL.Report;

public readonly struct Cell : IUnit<Location>, IUnit<Range>
{
    private readonly Content content;
    private readonly StyleId? styleId;

    public Cell(Content content, StyleId? styleId = null)
    {
        this.content = content;
        this.styleId = styleId;
    }

    public Location Write(SheetWindow window)
    {
        window.Place(content, styleId);
        return window.Range.LeftTop;
    }

    Range IUnit<Range>.Write(SheetWindow window)
    {
        var location = Write(window);
        return new Range(location, Size.Cell);
    }
}