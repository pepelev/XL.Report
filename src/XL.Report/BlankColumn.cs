namespace XL.Report;

public readonly struct BlankColumn : IUnit<Range>
{
    private readonly int cells;

    public BlankColumn()
        : this(1)
    {
    }

    public BlankColumn(int cells)
    {
        if (cells <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(cells),
                cells,
                "Must be positive"
            );
        }

        this.cells = cells;
    }

    public Range Write(SheetWindow window) => new(window.Range.LeftTop, new Size(cells, 0));
}