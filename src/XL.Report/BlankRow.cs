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

public readonly struct BlankRow : IUnit<Range>
{
    private readonly int rows;

    public BlankRow()
        : this(1)
    {
    }

    public BlankRow(int rows)
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