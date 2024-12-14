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

using XL.Report.Auxiliary;

namespace XL.Report;

// 2 ^ 11

// 2 ^ 1
// 2 ^ 3
// 2 ^ 7 = 128
internal struct RowUsage
{
    private Level1 root;

    public bool TryMark(Interval<int> range)
    {
        var adjustedRange = new Interval<int>(
            range.LeftInclusive - Location.MinX,
            range.RightInclusive - Location.MinX
        );
        return root.TryMark(adjustedRange);
    }

    /// <summary>
    ///     [0, 2^14 - 1]
    /// </summary>
    private struct Level1
    {
        private const int LeftHandMinX = 0;
        private const int LeftHalfMaxX = Location.XLength / 2 - Location.MinX;

        private const int RightHalfMinX = LeftHalfMaxX + 1;
        private const int RightHalfMaxX = Location.XLength - Location.MinX;
        private static Interval<int> LeftHalf => new(LeftHandMinX, LeftHalfMaxX);
        private static Interval<int> RightHalf => new(RightHalfMinX, RightHalfMaxX);

        private Level2? a;
        private Level2? b;

        public bool TryMark(Interval<int> range)
        {
            var success = true;
            if (LeftHalf.Intersection(range) is { } leftPart)
            {
                a ??= new Level2();
                success &= a.TryMark(leftPart);
            }

            if (RightHalf.Intersection(range) is { } rightPart)
            {
                b ??= new Level2();
                var rightAdjustedRange = rightPart.Shift(-RightHalfMinX);
                success &= b.TryMark(rightAdjustedRange);
            }

            return success;
        }
    }

    /// <summary>
    ///     [0, 2^13 - 1]
    /// </summary>
    private sealed class Level2
    {
        private Array8<Level3?> next;
        private static Interval<int> FullLevel3Range => new(0, (1 << 10) - 1);

        public bool TryMark(Interval<int> range)
        {
            var firstEntryIndex = range.LeftInclusive >>> 10;
            var lastEntryIndex = range.RightInclusive >>> 10;
            var items = next.Content;
            var success = true;
            for (var i = firstEntryIndex; i <= lastEntryIndex && success; i++)
            {
                var leftBound = i << 10;
                var rightBound = ((i + 1) << 10) - 1;
                var entryRange = new Interval<int>(leftBound, rightBound);
                var rangeToMark = range.Intersection(entryRange)!.Value;
                var adjustedRangeToMark = rangeToMark.Shift(-leftBound);
                ref var entry = ref items[i];
                if (entry == null)
                {
                    if (adjustedRangeToMark == FullLevel3Range)
                    {
                        entry = Level3.Full;
                        continue;
                    }

                    entry = new Level3();
                    entry.TryMark(adjustedRangeToMark);
                    continue;
                }

                success &= entry.TryMark(adjustedRangeToMark);
            }

            return success;
        }
    }

    /// <summary>
    ///     [0, 2^10 - 1]
    /// </summary>
    public sealed class Level3
    {
        public const int BitsSize = 7;
        public const int Size = 128;
        public const int SizeMask = 127;

        public static Level3 Full { get; } = CreateFull();

        private static Level3 CreateFull()
        {
            var result = new Level3();
            result.map.Content.Fill(byte.MaxValue);
            return result;
        }

        private Array128<byte> map;

        public bool TryMark(Interval<int> range)
        {
            var firstByte = range.LeftInclusive >>> 3;
            var lastByte = range.RightInclusive >>> 3;
            var bytes = map.Content;
            if (firstByte == lastByte)
            {
                var index = (range.LeftInclusive & 7) * 8 + (range.RightInclusive & 7);
                var mask = TheOnlyByteTable[index];
                if ((bytes[firstByte] & mask) != 0)
                {
                    return false;
                }

                bytes[firstByte] |= mask;
                return true;

            }

            var firstByteMask = FirstByteTable[range.LeftInclusive & 7];
            if ((bytes[firstByte] & firstByteMask) != 0)
            {
                return false;
            }

            bytes[firstByte] |= firstByteMask;

            var middleFirstByte = firstByte + 1;
            var tg = bytes[middleFirstByte..lastByte];
            if (!tg.SequenceEqual(Zeroes128[..tg.Length]))
            {
                return false;
            }

            tg.Fill(byte.MaxValue);

            var lastByteMask = LastByteTable[range.RightInclusive & 7];
            if ((bytes[lastByte] & lastByteMask) != 0)
            {
                return false;
            }

            bytes[lastByte] |= lastByteMask;
            return true;
        }

        private static ReadOnlySpan<byte> FirstByteTable => new byte[]
        {
            0b0000_0001,
            0b0000_0011,
            0b0000_0111,
            0b0000_1111,

            0b0001_1111,
            0b0011_1111,
            0b0111_1111,
            0b1111_1111
        };

        private const byte IllegalByte = 0b10101010;

        /// <summary>
        ///     Generated by <see cref="XL.Report.Tests.ArrayInitializationSyntax.TheOnlyByteTable"/>
        /// </summary>
        private static ReadOnlySpan<byte> TheOnlyByteTable => new byte[]
        {
            0b0000_0001, 0b0000_0011, 0b0000_0111, 0b0000_1111, 0b0001_1111, 0b0011_1111, 0b0111_1111, 0b1111_1111,
            IllegalByte, 0b0000_0010, 0b0000_0110, 0b0000_1110, 0b0001_1110, 0b0011_1110, 0b0111_1110, 0b1111_1110,
            IllegalByte, IllegalByte, 0b0000_0100, 0b0000_1100, 0b0001_1100, 0b0011_1100, 0b0111_1100, 0b1111_1100,
            IllegalByte, IllegalByte, IllegalByte, 0b0000_1000, 0b0001_1000, 0b0011_1000, 0b0111_1000, 0b1111_1000,
            IllegalByte, IllegalByte, IllegalByte, IllegalByte, 0b0001_0000, 0b0011_0000, 0b0111_0000, 0b1111_0000,
            IllegalByte, IllegalByte, IllegalByte, IllegalByte, IllegalByte, 0b0010_0000, 0b0110_0000, 0b1110_0000,
            IllegalByte, IllegalByte, IllegalByte, IllegalByte, IllegalByte, IllegalByte, 0b0100_0000, 0b1100_0000,
            IllegalByte, IllegalByte, IllegalByte, IllegalByte, IllegalByte, IllegalByte, IllegalByte, 0b1000_0000
        };

        private static ReadOnlySpan<byte> LastByteTable => new byte[]
        {
            0b1000_0000,
            0b1100_0000,
            0b1110_0000,
            0b1111_0000,

            0b1111_1000,
            0b1111_1100,
            0b1111_1110,
            0b1111_1111
        };

        private static ReadOnlySpan<byte> Zeroes128 => new byte[Size];
    }
}