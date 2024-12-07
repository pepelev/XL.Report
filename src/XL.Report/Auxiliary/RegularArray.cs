namespace XL.Report.Auxiliary;

internal readonly struct RegularArray<T>(T[] content) : IArray<T>
{
    public Span<T> Content => content;
}