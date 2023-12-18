using System.Numerics;

namespace XL.Report;

internal readonly struct Interval<T> where T : IComparisonOperators<T, T, bool>
{
    public Interval(T leftInclusive, T rightInclusive)
    {
        if (rightInclusive < leftInclusive)
        {
            throw new ArgumentException("rightInclusive < leftInclusive");
        }

        LeftInclusive = leftInclusive;
        RightInclusive = rightInclusive;
    }

    public bool Intersect(Interval<T> other)
    {
        return LeftInclusive <= other.RightInclusive && other.LeftInclusive <= RightInclusive;
    }

    public bool ToRightOf(Interval<T> other) => LeftInclusive > other.RightInclusive;

    public T LeftInclusive { get; }
    public T RightInclusive { get; }
}