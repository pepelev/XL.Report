namespace XL.Report;

public sealed class SimpleMatrix : IUnit<Range>
{
    private readonly Func<int, int, IUnit<Range>> cells;
    private readonly Size cellSize;
    private readonly int height;
    private readonly int width;

    public SimpleMatrix(int width, int height, Func<int, int, IUnit<Range>> cells)
        : this(width, height, cells, Size.Cell)
    {
    }

    public SimpleMatrix(int width, int height, Func<int, int, IUnit<Range>> cells, Size cellSize)
    {
        this.width = width;
        this.height = height;
        this.cells = cells;
        this.cellSize = cellSize;
    }

    public Range Write(SheetWindow window)
    {
        var totalSize = new Size(cellSize.Width * width, cellSize.Height * height);
        using (window.Reduce(Offset.Zero, totalSize))
        {
            for (var y = 0; y < height; y++)
            {
                for (var x = 0; x < width; x++)
                {
                    var cell = cells(x, y);
                    var offset = new Offset(cellSize.Width * x, cellSize.Height * y);
                    using (window.Reduce(offset, cellSize))
                    {
                        cell.Write(window);
                    }
                }
            }
        }

        return new Range(window.Range.LeftTop, totalSize);
    }
}