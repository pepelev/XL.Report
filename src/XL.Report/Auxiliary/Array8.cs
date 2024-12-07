using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace XL.Report.Auxiliary;

[InlineArray(8)]
internal struct Array8<T>
{
    private T value0;
    public Span<T> Content => MemoryMarshal.CreateSpan(ref value0, 8);
}