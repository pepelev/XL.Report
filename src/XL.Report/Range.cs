namespace XL.Report;

public readonly struct Range : IEquatable<Range>, IParsable<Range>
{
    public override string ToString() => $"{LeftTopCell}:{RightBottomCell}";

    public static Range Whole => new Range(
        new Location(Location.MinX, Location.MinY),
        new Size(
            Location.MaxX - Location.MinX + 1,
            Location.MaxY - Location.MinY + 1
        )
    );

    public Range(Location leftTopCell, Size size)
    {
        LeftTopCell = leftTopCell;
        Size = size;
    }

    public bool Equals(Range other)
    {
        return LeftTopCell == other.LeftTopCell && Size == other.Size;
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
            return (LeftTopCell.GetHashCode() * 397) ^ Size.GetHashCode();
        }
    }

    public static bool operator ==(Range a, Range b) => a.Equals(b);
    public static bool operator !=(Range a, Range b) => !(a == b);

    public Location LeftTopCell { get; }
    public Size Size { get; }

    public Location RightBottomCell => new(
        LeftTopCell.X + (int)Size.Width - 1,
        LeftTopCell.Y + (int)Size.Height - 1
    );

    public Location RightTopCell => new(
        LeftTopCell.X + (int)Size.Width - 1,
        LeftTopCell.Y
    );

    public Location LeftBottomCell => new(
        LeftTopCell.X,
        LeftTopCell.Y + (int)Size.Height - 1
    );

    public bool Intersects(in Range range)
    {
        return Intersects(LeftTopCell.X, RightBottomCell.X, range.LeftTopCell.X, range.RightBottomCell.X) &&
               Intersects(LeftTopCell.Y, RightBottomCell.Y, range.LeftTopCell.Y, range.RightBottomCell.Y);
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
            var size = new Size((uint) (bottomRight.X - topLeft.X + 1), (uint) (bottomRight.Y - topLeft.Y + 1));
            result = new Range(topLeft, size);
            return true;
        }

        result = default(Range);
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

    private uint TopHeight => (Size.Height + 1) >> 1;
    private uint BottomHeight => Size.Height >> 1;
    private uint LeftWidth => (Size.Width + 1) >> 1;
    private uint RightWidth => Size.Width >> 1;

    public Range SplitLeftTop()
    {
        return new Range(LeftTopCell, new Size(LeftWidth, TopHeight));
    }

    public Range SplitRightTop()
    {
        return new Range(
            new Location(LeftTopCell.X + (int) LeftWidth, LeftTopCell.Y),
            new Size(RightWidth, TopHeight));
    }

    public Range SplitLeftBottom()
    {
        return new Range(
            new Location(LeftTopCell.X, LeftTopCell.Y + (int) TopHeight),
            new Size(LeftWidth, BottomHeight));
    }

    public Range SplitRightBottom()
    {
        return new Range(
            new Location(LeftTopCell.X + (int) LeftWidth, LeftTopCell.Y + (int) TopHeight),
            new Size(RightWidth, BottomHeight));
    }

    public bool Contains(in Location location)
    {
        return LeftTopCell.X <= location.X && location.X <= RightBottomCell.X &&
               LeftTopCell.Y <= location.Y && location.Y <= RightBottomCell.Y;
    }

    public bool Contains(in Range range)
    {
        return Contains(range.LeftTopCell) &&
               Contains(range.RightBottomCell);
    }
}