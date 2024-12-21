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

using FluentAssertions;

namespace XL.Report.Tests;

public sealed class LocationShould
{
    private sealed record Matching(string Text, Location Location, bool Clear);

    private static readonly Matching[] Matchings =
    [
        new Matching("A1", new Location(1, 1), Clear: true),
        new Matching("B20", new Location(2, 20), Clear: true),
        new Matching("XFD1048576", new Location(16384, 1048576), Clear: true),
        new Matching("XFD1", new Location(16384, 1), Clear: true),
        new Matching("AA400", new Location(27, 400), Clear: true),
        new Matching("A1048576", new Location(1, 1048576), Clear: true),

        new Matching(" A200\t", new Location(1, 200), Clear: false),
    ];

    private static TestCaseData[] parseCases =
    [
        new TestCaseData(null).Returns(null),
        new TestCaseData("A1").Returns(new Location(1, 1)),
        new TestCaseData("B20").Returns(new Location(2, 20)),
        new TestCaseData("C0").Returns(null),
        new TestCaseData("XFD1048576").Returns(new Location(16384, 1048576)),
        new TestCaseData("AXFD1048576").Returns(null),
        new TestCaseData("XFE1048576").Returns(null),
        new TestCaseData("XFD1048577").Returns(null),
        new TestCaseData("100").Returns(null),
        new TestCaseData("A").Returns(null),
        new TestCaseData("A 1").Returns(null),
        new TestCaseData("A1 0").Returns(null),
        new TestCaseData("A B20").Returns(null),
        new TestCaseData(" A200\t").Returns(new Location(1, 200)),
    ];

    private static IEnumerable<TestCaseData> PrintCases => Matchings
        .Where(matching => matching.Clear)
        .Select(matching => new TestCaseData(matching.Location).Returns(matching.Text));

    private static IEnumerable<TestCaseData> CompactTryFormatCases => Matchings
        .Where(matching => matching.Clear)
        .Select(matching => new TestCaseData(matching.Location, matching.Text.Length));

    [Test]
    [TestCaseSource(nameof(parseCases))]
    public Location? Parse(string? input)
    {
        if (Location.TryParse(input, out var result))
        {
            return result;
        }

        return null;
    }

    [Test]
    [TestCaseSource(nameof(PrintCases))]
    public string ToString(Location sut)
    {
        return sut.ToString(null, null);
    }

    [Test]
    [TestCaseSource(nameof(PrintCases))]
    public string TryFormat(Location sut)
    {
        const int size = 16;
        Span<char> buffer = stackalloc char[size];
        var results = new HashSet<string>();
        for (var i = 0; i < size; i++)
        {
            var activeBuffer = buffer[..i];
            activeBuffer.Clear();
            var success = sut.TryFormat(activeBuffer, out var charsWritten, format: "", provider: null);
            if (success)
            {
                results.Add(new string(activeBuffer[..charsWritten]));
            }
        }

        return results.Should().ContainSingle().Subject;
    }

    [Test]
    [TestCaseSource(nameof(CompactTryFormatCases))]
    public void UseSmallestPossibleBufferForTryFormat(Location sut, int chars)
    {
        Span<char> buffer = stackalloc char[chars];
        var success = sut.TryFormat(buffer, out var charsWritten, format: "", provider: null);
        (success, charsWritten).Should().Be((true, chars));
    }
}