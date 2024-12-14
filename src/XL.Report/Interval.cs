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

using System.Numerics;

namespace XL.Report;

internal static class Interval
{
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