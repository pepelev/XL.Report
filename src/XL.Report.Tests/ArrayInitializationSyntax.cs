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

namespace XL.Report.Tests;

[Explicit]
public sealed class ArrayInitializationSyntax
{
    [Test]
    public void TheOnlyByteTable()
    {
        var list = new List<string>();
        for (var start = 0; start < 8; start++)
        {
            for (var end = 0; end < 8; end++)
            {
                list.Add(Get(start, end));
            }
        }

        var join = string.Join(
            "," + Environment.NewLine,
            list
                .Select((@byte, index) => (Byte: @byte, Index: index))
                .GroupBy(pair => pair.Index / 8)
                .Select(group => string.Join(", ", group.Select(pair => pair.Byte)))
        );
        Console.WriteLine(join);

        string Get(int start, int end)
        {
            if (start > end)
                return "IllegalByte";

            var result = 0;
            for (var i = start; i <= end; i++)
            {
                result |= (1 << i);
            }

            var stringRepresentation = ((byte)result).ToString("b8");
            return $"0b{stringRepresentation[..4]}_{stringRepresentation[4..]}";
        }
    }

    [Test]
    public void Zeroes128()
    {
        var join = string.Join(
            $",{Environment.NewLine}",
            Enumerable
                .Repeat("0", 128)
                .Select((value, index) => (value, index))
                .GroupBy(pair => pair.index / 8)
                .Select(group => group.Select(pair => pair.value))
                .Select(group => string.Join(", ", group))
        );
        Console.WriteLine(join);
    }
}