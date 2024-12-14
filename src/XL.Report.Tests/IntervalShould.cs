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

internal sealed class IntervalShould
{
    public static TestCaseData[] Cases =
    [
        new TestCaseData(
            new Interval<int>(1, 1),
            new Interval<int>(1, 1)
        ).Returns(true),
        new TestCaseData(
            new Interval<int>(0, 1),
            new Interval<int>(1, 1)
        ).Returns(true),
        new TestCaseData(
            new Interval<int>(1, 2),
            new Interval<int>(1, 1)
        ).Returns(true),
        new TestCaseData(
            new Interval<int>(1, 2),
            new Interval<int>(0, 1)
        ).Returns(true),
        new TestCaseData(
            new Interval<int>(0, 1),
            new Interval<int>(1, 2)
        ).Returns(true),
        new TestCaseData(
            new Interval<int>(0, 1),
            new Interval<int>(1, 1)
        ).Returns(true),
        new TestCaseData(
            new Interval<int>(0, 1),
            new Interval<int>(0, 0)
        ).Returns(true)
    ];

    [Test]
    [TestCaseSource(nameof(Cases))]
    public bool Give_Intersection(Interval<int> a, Interval<int> b)
    {
        return a.Intersect(b);
    }

    [Test]
    [TestCaseSource(nameof(Cases))]
    public bool Be_Transitive(Interval<int> a, Interval<int> b)
    {
        var direct = a.Intersect(b);
        var reverse = b.Intersect(a);

        direct.Should().Be(reverse);
        return direct;
    }
}