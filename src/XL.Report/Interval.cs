using System.Numerics;

namespace XL.Report;

internal static class Interval
{
    public static Interval<T>? MergeIfAdjacent<T>(Interval<T> a, Interval<T> b) 
        where T : IComparisonOperators<T, T, bool>, INumber<T>
    {
        if (a.RightInclusive + T.One == b.LeftInclusive)
        {
            return new Interval<T>(a.LeftInclusive, b.RightInclusive);
        }

        if (b.RightInclusive + T.One == a.LeftInclusive)
        {
            return new Interval<T>(b.LeftInclusive, a.RightInclusive);
        }

        return null;
    }

    public static Interval<T> Shift<T>(this Interval<T> interval, T shift) where T : INumber<T> =>
        new(interval.LeftInclusive + shift, interval.RightInclusive + shift);
}

internal readonly struct Interval<T> : IEquatable<Interval<T>>
    where T : IComparisonOperators<T, T, bool>
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

    public Interval<T>? Intersection(Interval<T> other)
    {
        var left = LeftInclusive < other.LeftInclusive
            ? other.LeftInclusive
            : LeftInclusive;
        var right = RightInclusive < other.RightInclusive
            ? RightInclusive
            : other.RightInclusive;
        return left <= right
            ? new Interval<T>(left, right)
            : null;
    }

    public bool ToRightOf(Interval<T> other) => LeftInclusive > other.RightInclusive;
    public bool Contains(T item) => LeftInclusive <= item && item <= RightInclusive;

    public T LeftInclusive { get; }
    public T RightInclusive { get; }

    public static bool operator ==(Interval<T> left, Interval<T> right) => left.Equals(right);
    public static bool operator !=(Interval<T> left, Interval<T> right) => !left.Equals(right);

    public bool Equals(Interval<T> other)
    {
        return EqualityComparer<T>.Default.Equals(LeftInclusive, other.LeftInclusive) &&
               EqualityComparer<T>.Default.Equals(RightInclusive, other.RightInclusive);
    }

    public override bool Equals(object? obj) => obj is Interval<T> other && Equals(other);
    public override int GetHashCode() => HashCode.Combine(LeftInclusive, RightInclusive);
    public override string ToString() => $"[{LeftInclusive}, {RightInclusive}]";
}