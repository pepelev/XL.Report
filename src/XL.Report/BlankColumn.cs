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