namespace XL.Report;

public readonly struct SharedStringId : IEquatable<SharedStringId>, IComparable<SharedStringId>
{
    public int Index { get; }

    public SharedStringId(int index)
    {
        Index = index;
    }

    public bool Equals(SharedStringId other) => Index == other.Index;
    public override bool Equals(object? obj) => obj is SharedStringId other && Equals(other);
    public override int GetHashCode() => Index;
    public static bool operator ==(SharedStringId left, SharedStringId right) => left.Equals(right);
    public static bool operator !=(SharedStringId left, SharedStringId right) => !left.Equals(right);
    public int CompareTo(SharedStringId other) => Index.CompareTo(other.Index);
}