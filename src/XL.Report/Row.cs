namespace XL.Report;

public readonly struct Row : IUnit<Range>
{
    private readonly IUnit<Range>[] units;

    public Row(params IUnit<Range>[] units)
    {
        this.units = units;
    }

    public Range Write(SheetWindow window)
    {
        var written = new Range(window.Range.LeftTop, Size.Empty);
        foreach (var unit in units ?? Array.Empty<IUnit<Range>>())
        {
            window.PushReduce(new Offset(written.Size.Width, 0));
            var range = unit.Write(window);
            window.PopReduce();
            written = Range.MinimalBounding(written, range);
        }

        return written;
    }
}