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

public sealed class Row(params IUnit<Range>[] units) : IUnit<Range>, IUnit<Range[]>
{
    public Range Write(SheetWindow window)
    {
        var written = new Range(window.Range.LeftTop, Size.Empty);
        foreach (var unit in units)
        {
            using (window.Reduce(new Offset(written.Size.Width, 0)))
            {
                var range = unit.Write(window);
                written = Range.MinimalBounding(written, range);
            }
        }

        return written;
    }

    Range[] IUnit<Range[]>.Write(SheetWindow window)
    {
        if (units is [])
        {
            return [];
        }

        var result = new Range[units.Length];
        var written = new Range(window.Range.LeftTop, Size.Empty);
        for (var i = 0; i < units.Length; i++)
        {
            var unit = units[i];
            using (window.Reduce(new Offset(written.Size.Width, 0)))
            {
                var range = unit.Write(window);
                written = Range.MinimalBounding(written, range);
                result[i] = range;
            }
        }

        return result;
    }
}