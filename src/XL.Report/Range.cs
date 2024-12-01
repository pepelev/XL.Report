using System.Globalization;

namespace XL.Report;

public readonly struct Range
    : IEquatable<Range>, ISpanFormattable
#if NET7_0_OR_GREATER
    , IParsable<Range>
#endif
{
    public static Range EntireSheet => new(
        new Location(Location.MinX, Location.MinY),
        new Size(
            Location.XLength,
            Location.YLength
        )
    );

    public Range(Location leftTop, Size size)
    {
        LeftTop = leftTop;
        Size = size;
    }

    public static Range Create(Location leftTop, Location rightBottom)
    {
        var size = new Size(
            rightBottom.X - leftTop.X + 1,
            rightBottom.Y - leftTop.Y + 1
        );

        if (size.IsDegenerate)
        {
            throw new ArgumentException();
        }

        return new Range(leftTop, size);
    }

    public bool IsDegenerate => Size.IsDegenerate;

    public bool Equals(Range other)
    {
        return LeftTop == other.LeftTop && Size == other.Size;
    }

    public override bool Equals(object? obj)
    {
        if (ReferenceEquals(null, obj))
            return false;

        return obj is Range rectangle && Equals(rectangle);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            return (LeftTop.GetHashCode() * 397) ^ Size.GetHashCode();
        }
    }

    public static bool operator ==(Range a, Range b) => a.Equals(b);
    public static bool operator !=(Range a, Range b) => !(a == b);

    // todo this is not a cell, because Size can be empty
    public Location LeftTop { get; }
    public Size Size { get; }

    public bool IsEmpty => Size.Width == 0 || Size.Height == 0;

    public int Left => LeftTop.X;
    public int Right => LeftTop.X + Size.Width - 1;
    public int Top => LeftTop.Y;
    public int Bottom => LeftTop.Y + Size.Height - 1;

    public Location RightBottom => new(Right, Bottom);
    public Location RightTop => new(Right, Top);
    public Location LeftBottom => new(Left, Bottom);

    public bool Intersects(in Range range)
    {
        return Intersects(LeftTop.X, RightBottom.X, range.LeftTop.X, range.RightBottom.X) &&
               Intersects(LeftTop.Y, RightBottom.Y, range.LeftTop.Y, range.RightBottom.Y);
    }

    private static bool Intersects(int l1, int r1, int l2, int r2)
    {
        return r1 >= l2 && l1 <= r2;
    }

    public static bool TryParse(string? input, IFormatProvider? formatProvider, out Range result)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            result = default;
            return false;
        }

        var parts = input.Trim().Split(':');

        if (parts.Length != 2)
        {
            result = default;
            return false;
        }

        if (Location.TryParse(parts[0], formatProvider, out var topLeft) && Location.TryParse(parts[1], formatProvider, out var bottomRight))
        {
            var size = new Size( bottomRight.X - topLeft.X + 1, bottomRight.Y - topLeft.Y + 1);
            result = new Range(topLeft, size);
            return true;
        }

        result = default;
        return false;
    }

    public static Range Parse(string input, IFormatProvider? formatProvider)
    {
        if (TryParse(input, formatProvider, out var result))
        {
            return result;
        }

        throw new ArgumentException();
    }

    public static Range Parse(string input) => Parse(input, CultureInfo.InvariantCulture);

    private int TopHeight => (Size.Height + 1) >> 1;
    private int BottomHeight => Size.Height >> 1;
    private int LeftWidth => (Size.Width + 1) >> 1;
    private int RightWidth => Size.Width >> 1;

    public Range ReduceDown(int offset)
    {
        if (Size.Height < offset)
        {
            throw new InvalidOperationException();
        }

        return new Range(
            LeftTop + new Offset(0, offset),
            Size - new Offset(0, offset)
        );
    }

    public Range ReduceLeftUp(Size size)
    {
        var reducedRange = new Range(LeftTop, size);
        if (Contains(reducedRange))
        {
            return reducedRange;
        }

        throw new ArgumentException();
    }

    public static Range MinimalBounding(Range a, Range b)
    {
        var leftTop = new Location(
            x: Math.Min(a.Left, b.Left),
            y: Math.Min(a.Top, b.Top)
        );
        var rightBottom = new Location(
            x: Math.Max(a.Right, b.Right),
            y: Math.Max(a.Bottom, b.Bottom)
        );
        return Create(leftTop, rightBottom);
    }

    public Range SplitLeftTop()
    {
        return new Range(LeftTop, new Size(LeftWidth, TopHeight));
    }

    public Range SplitRightTop()
    {
        return new Range(
            new Location(LeftTop.X + LeftWidth, LeftTop.Y),
            new Size(RightWidth, TopHeight));
    }

    public Range SplitLeftBottom()
    {
        return new Range(
            new Location(LeftTop.X, LeftTop.Y + TopHeight),
            new Size(LeftWidth, BottomHeight));
    }

    public Range SplitRightBottom()
    {
        return new Range(
            new Location(LeftTop.X + LeftWidth, LeftTop.Y + TopHeight),
            new Size(RightWidth, BottomHeight));
    }

    public bool Contains(in Location location)
    {
        return LeftTop.X <= location.X && location.X <= RightBottom.X &&
               LeftTop.Y <= location.Y && location.Y <= RightBottom.Y;
    }

    public bool Contains(in Range range)
    {
        return Contains(range.LeftTop) &&
               Contains(range.RightBottom);
    }

    public string AsString() => $"{LeftTop}:{RightBottom}";
    public override string ToString() => AsString();
    public string ToString(string? format, IFormatProvider? formatProvider) => AsString();

    public bool TryFormat(
        Span<char> destination,
        out int charsWritten,
        ReadOnlySpan<char> format,
        IFormatProvider? provider)
    {
        return FormatContext.Start
            .Write(ref destination, LeftTop)
            .Write(ref destination, ":")
            .Write(ref destination, RightBottom)
            .Deconstruct(out charsWritten);
    }
}