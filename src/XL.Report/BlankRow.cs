namespace XL.Report;

public readonly struct BlankRow : IUnit<Range>
{
    private readonly int cells;

    public BlankRow()
        : this(1)
    {
    }

    public BlankRow(int cells)
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

    public Range Write(SheetWindow window) => new(window.Range.LeftTop, new Size(0, cells));
}