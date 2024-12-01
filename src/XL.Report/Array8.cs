#if NET8_0_OR_GREATER

using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace XL.Report;

internal static class Default<T>
{
    private static T Value = default!;

    public static ref T Ref() => ref Value;
}

internal interface IArray<T>
{
    Span<T> Content { get; }
}

internal readonly struct RegularArray<T>(T[] content) : IArray<T>
{
    public Span<T> Content => content;
}

[InlineArray(4)]
internal struct Array4<T> : IArray<T>
{
    private T value0;
    public Span<T> Content => MemoryMarshal.CreateSpan(ref value0, 4);
}

[InlineArray(8)]
internal struct Array8<T> : IArray<T>
{
    private T value0;
    public Span<T> Content => MemoryMarshal.CreateSpan(ref value0, 8);
}

[InlineArray(16)]
internal struct Array16<T> : IArray<T>
{
    private T value0;
    public Span<T> Content => MemoryMarshal.CreateSpan(ref value0, 16);
}

[InlineArray(64)]
internal struct Array64<T> : IArray<T>
{
    private T value0;
    public Span<T> Content => MemoryMarshal.CreateSpan(ref value0, 64);
}

[InlineArray(128)]
internal struct Array128<T> : IArray<T>
{
    private T value0;
    public Span<T> Content => MemoryMarshal.CreateSpan(ref value0, 128);
}

public interface IBuffer<T>
{
    int Capacity { get; }
    int Count { get; set; }
    Span<T> Content { get; }
}

public struct Buffer8<T> : IBuffer<T>
{
    public int Capacity => 8;
    public int Count { get; set; }
    public Span<T> Content => MemoryMarshal.CreateSpan(ref Unsafe.As<Array8<T>, T>(ref content), Capacity);
    private Array8<T> content;
}

public struct Stack<T, TBuffer> where TBuffer : IBuffer<T>
{
    private TBuffer buffer;

    public Stack(TBuffer buffer)
    {
        this.buffer = buffer;
    }

    public void Push(T item)
    {
        buffer.Content[buffer.Count++] = item;
    }

    public ref T Peek() => ref buffer.Content[buffer.Count - 1];

    public void Pop()
    {
        buffer.Content[--buffer.Count] = default!;
    }

    public bool IsEmpty => buffer.Count == 0;

    public void Clear()
    {
        buffer.Content[..buffer.Count].Clear();
        buffer.Count = 0;
    }
}

#endif