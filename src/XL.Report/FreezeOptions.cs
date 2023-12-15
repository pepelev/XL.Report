namespace XL.Report;

public sealed class FreezeOptions : IEquatable<FreezeOptions>
{
    public static FreezeOptions None { get; } = new(0, 0);

    public FreezeOptions(int freezeByX, int freezeByY)
    {
        if (freezeByX < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(freezeByY), freezeByX, "must be non-negative");
        }

        if (freezeByY < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(freezeByY), freezeByY, "must be non-negative");
        }

        FreezeByX = freezeByX;
        FreezeByY = freezeByY;
    }

    public int FreezeByX { get; }
    public int FreezeByY { get; }

    public bool Equals(FreezeOptions? other)
    {
        if (ReferenceEquals(null, other)) return false;
        if (ReferenceEquals(this, other)) return true;
        return FreezeByX == other.FreezeByX && FreezeByY == other.FreezeByY;
    }

    public override bool Equals(object? obj)
    {
        return ReferenceEquals(this, obj) || obj is FreezeOptions other && Equals(other);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(FreezeByX, FreezeByY);
    }

    public static bool operator ==(FreezeOptions? left, FreezeOptions? right)
    {
        return Equals(left, right);
    }

    public static bool operator !=(FreezeOptions? left, FreezeOptions? right)
    {
        return !Equals(left, right);
    }
}