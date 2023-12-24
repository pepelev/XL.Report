namespace XL.Report;

public readonly struct Reduction
{
    public Offset Offset { get; }
    public Size? NewSize { get; }

    public Reduction(Offset offset, Size? newSize)
    {
        if (newSize?.IsDegenerate == true)
        {
            throw new ArgumentException("is degenerate", nameof(newSize));
        }
        
        Offset = offset;
        NewSize = newSize;
    }
}