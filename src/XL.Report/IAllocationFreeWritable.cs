namespace XL.Report;

public interface IAllocationFreeWritable
{
    bool TryFormat(Span<char> destination, out int charsWritten);
    string AsString();
}