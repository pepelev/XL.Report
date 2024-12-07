namespace XL.Report.Auxiliary;

internal interface IArray<T>
{
    Span<T> Content { get; }
}