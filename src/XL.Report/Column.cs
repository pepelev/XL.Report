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

public readonly struct Column : IUnit<Range>
{
    private readonly IUnit<Range>[] units;

    public Column(params IUnit<Range>[] units)
    {
        this.units = units;
    }

    public Range Write(SheetWindow window)
    {
        var written = new Range(window.Range.LeftTop, Size.Empty);
        foreach (var unit in units ?? Array.Empty<IUnit<Range>>())
        {
            Range range;
            using (window.Reduce(new Offset(0, written.Size.Height)))
            {
                range = unit.Write(window);
            }

            written = Range.MinimalBounding(written, range);
        }

        return written;
    }
}