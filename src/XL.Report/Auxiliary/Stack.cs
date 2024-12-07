namespace XL.Report.Auxiliary;

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