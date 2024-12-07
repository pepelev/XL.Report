namespace XL.Report.Auxiliary;

public interface IBuffer<T>
{
    int Capacity { get; }
    int Count { get; set; }
    Span<T> Content { get; }
}