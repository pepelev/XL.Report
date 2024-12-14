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

public sealed class BoundedSharedStrings : SharedStrings
{
    private readonly Dictionary<string, SharedStringId> index = new(StringComparer.Ordinal);
    private int totalLength;
    private readonly int maxCount;
    private readonly int maxSingleStringLength;
    private readonly int maxTotalLength;

    public BoundedSharedStrings(int maxCount, int maxSingleStringLength, int maxTotalLength)
    {
        this.maxCount = maxCount;
        this.maxTotalLength = maxTotalLength;
        this.maxSingleStringLength = maxSingleStringLength;
    }

    public override SharedStringId? TryRegister(string @string)
    {
        if (index.TryGetValue(@string, out var id))
        {
            return id;
        }

        if (maxSingleStringLength < @string.Length)
        {
            return null;
        }

        if (maxCount < index.Count + 1)
        {
            return null;
        }

        var newTotalLength = totalLength + @string.Length;
        if (maxTotalLength < newTotalLength)
        {
            return null;
        }

        return Add(@string);
    }

    public override SharedStringId ForceRegister(string @string)
    {
        if (index.TryGetValue(@string, out var id))
        {
            return id;
        }

        return Add(@string);
    }

    public override IEnumerator<(string String, SharedStringId Id)> GetEnumerator() =>
        index.Select(pair => (pair.Key, pair.Value)).GetEnumerator();

    private SharedStringId Add(string @string)
    {
        var sharedStringId = new SharedStringId(index.Count);
        index[@string] = sharedStringId;
        totalLength += @string.Length;
        return sharedStringId;
    }
}