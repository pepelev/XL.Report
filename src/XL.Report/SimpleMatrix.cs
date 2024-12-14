#region Legal
// Copyright 2024 Pepelev Alexey
// 
// This file is part of XL.Report.
// 
// XL.Report is free software: you can redistribute it and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation, either
// version 3 of the License, or (at your option) any later version.
// 
// XL.Report is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
// without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU Lesser General Public License for more details.
// 
// You should have received a copy of the GNU Lesser General Public License along with XL.Report.
// If not, see <https://www.gnu.org/licenses/>.
#endregion

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