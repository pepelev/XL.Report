namespace XL.Report;

public readonly struct Bounded : IUnit<Range>
{
    private readonly Size bounds;
    private readonly IUnit<Range> unit;

    public Bounded(Size bounds, IUnit<Range> unit)
    {
        this.bounds = bounds;
        this.unit = unit;
    }

    public Range Write(SheetWindow window)
    {
        using (window.Reduce(Offset.Zero, bounds))
        {
            unit.Write(window);
        }

        return new Range(window.Range.LeftTop, bounds);
    }
}