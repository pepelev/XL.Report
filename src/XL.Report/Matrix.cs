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

public sealed class Matrix<TX, TY> : IUnit<Range>
{
    private readonly IUnit<Range> xHeader;
    private readonly (IUnit<Range> Mark, TX Value)[] xAspects;
    private readonly IUnit<Range> yHeader;
    private readonly (IUnit<Range> Mark, TY Value)[] yAspects;
    private readonly Size cellSize;
    private readonly Func<TX, TY, IUnit<Range>> cells;

    public Matrix(
        IUnit<Range> xHeader,
        (IUnit<Range> Mark, TX Value)[] xAspects,
        IUnit<Range> yHeader,
        (IUnit<Range> Mark, TY Value)[] yAspects,
        Size cellSize,
        Func<TX, TY, IUnit<Range>> cells)
    {
        this.xHeader = xHeader;
        this.yHeader = yHeader;
        this.xAspects = xAspects;
        this.yAspects = yAspects;
        this.cellSize = cellSize;
        this.cells = cells;
    }

    public Range Write(SheetWindow window)
    {
        var totalSize = cellSize
            .MultiplyWidth(xAspects.Length + 1)
            .MultiplyHeight(yAspects.Length + 2);

        using (window.Reduce(Offset.Zero, totalSize))
        {
            using (window.Reduce(new Offset(cellSize.Width, 0), cellSize.MultiplyWidth(xAspects.Length)))
            {
                xHeader.Write(window);
            }

            using (window.Reduce(cellSize.AsOffset(), cellSize.MultiplyWidth(xAspects.Length)))
            {
                var xMarks = new Row(
                    xAspects.Select(aspect => new Bounded(cellSize, aspect.Mark) as IUnit<Range>).ToArray()
                );

                xMarks.Write(window);
            }

            using (window.Reduce(new Offset(0, cellSize.Height), cellSize))
            {
                yHeader.Write(window);
            }

            using (window.Reduce(new Offset(0, cellSize.Height * 2), cellSize.MultiplyHeight(yAspects.Length)))
            {
                var yMarks = new Column(
                    yAspects.Select(aspect => new Bounded(cellSize, aspect.Mark) as IUnit<Range>).ToArray()
                );

                yMarks.Write(window);
            }

            using (window.Reduce(cellSize.MultiplyHeight(2).AsOffset()))
            {
                var matrix = new SimpleMatrix(
                    xAspects.Length,
                    yAspects.Length,
                    (x, y) => cells(xAspects[x].Value, yAspects[y].Value),
                    cellSize
                );

                matrix.Write(window);
            }
        }

        return window.Range.ReduceLeftUp(totalSize);
    }
}