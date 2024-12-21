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

using System.Collections;

namespace XL.Report;

public sealed class RangeSet : IReadOnlyCollection<Range>
{
    private readonly List<Range> ranges = new();

    public void Add(Range range)
    {
        if (ranges.Count == 0)
        {
            ranges.Add(range);
            return;
        }

        var last = ranges[^1];
        if (last.TryUnion(range) is { } united)
        {
            ranges[^1] = united;
            return;
        }

        ranges.Add(range);
    }

    public IEnumerator<Range> GetEnumerator() => ranges.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    public int Count => ranges.Count;
}