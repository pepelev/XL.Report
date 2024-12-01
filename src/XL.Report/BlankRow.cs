namespace XL.Report;

public readonly struct BlankRow : IUnit<Range>
{
    private readonly int rows;

    public BlankRow(int rows = 1)
    {
        if (rows <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(rows),
                rows,
                "Must be positive"
            );
        }

        this.rows = rows;
    }

    public Range Write(SheetWindow window)
    {
        window.TouchRow(rows);
        return new Range(window.Range.LeftTop, new Size(0, rows));
    }
}