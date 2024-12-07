using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace XL.Report.Auxiliary;

public struct Buffer8<T> : IBuffer<T>
{
    public int Capacity => 8;
    public int Count { get; set; }
    public Span<T> Content => MemoryMarshal.CreateSpan(ref Unsafe.As<Array8<T>, T>(ref content), Capacity);
    private Array8<T> content;
}