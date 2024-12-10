namespace XL.Report;

public readonly struct Reduction
{
    public Offset Offset { get; }
    public Size? NewSize { get; }

    public Reduction(Offset offset, Size? newSize)
    {
        if (offset.X < 0 || offset.Y < 0)
        {
            throw new ArgumentException("is negative (X < 0 || Y < 0)", nameof(offset));
        }

        if (newSize is { IsDegenerate: true })
        {
            throw new ArgumentException("is degenerate", nameof(newSize));
        }

        Offset = offset;
        NewSize = newSize;
    }
}