﻿namespace XL.Report;

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
            Range range;
            using (window.Reduce(new Offset(written.Size.Width, 0)))
            {
                range = unit.Write(window);
            }

            written = Range.MinimalBounding(written, range);
        }

        return written;
    }
}