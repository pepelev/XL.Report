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