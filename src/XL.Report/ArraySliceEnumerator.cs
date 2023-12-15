using System.Collections;

namespace XL.Report;

internal struct ArraySliceEnumerator<T> : IEnumerator<T>
{
    private readonly T[] array;
    private readonly int count;
    private int index = 0;

    public ArraySliceEnumerator(T[] array, int count)
    {
        this.array = array;
        this.count = count;
    }

    public bool MoveNext()
    {
        if (index < count)
        {
            Current = array[index++];
            return true;
        }

        return false;
    }

    public void Reset() => throw new NotSupportedException();
    public T Current { get; private set; } = default!;

    object IEnumerator.Current => Current!;

    public void Dispose()
    {
    }

    public bool IsInitialized => array != null;
}