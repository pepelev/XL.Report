using System.Collections;

namespace XL.Report;

public struct Sorted<T> : IReadOnlyCollection<T> where T : struct, IComparable<T>
{
    public int Count { get; private set; }

    // T[] | SortedSet<T>
    private object? value;

    public void Add(T item)
    {
        if (value == null)
        {
            var newValue = new T[4];
            newValue[0] = item;
            value = newValue;
            Count = 1;
        }
        else if (value is T[] array)
        {
            var comparison = array[Count - 1].CompareTo(item);
            if (comparison < 0)
            {
                if (array.Length <= Count)
                {
                    // todo pool
                    var newSize = Math.Max(array.Length * 2, 4);
                    Array.Resize(ref array, newSize);
                }

                array[Count++] = item;
                value = array;
            }
            else if (comparison == 0)
            {
                throw new InvalidOperationException();
            }
            else
            {
                var set = new SortedSet<T>();
                for (var i = 0; i < Count; i++)
                {
                    set.Add(array[i]);
                }

                if (!set.Add(item))
                {
                    throw new InvalidOperationException();
                }

                value = set;
                Count++;
            }
        }
        else
        {
            var set = (SortedSet<T>)value;
            if (!set.Add(item))
            {
                throw new InvalidOperationException();
            }

            Count++;
        }
    }

    public Enumerator GetEnumerator() => new(this);
    IEnumerator<T> IEnumerable<T>.GetEnumerator() => GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public struct Enumerator : IEnumerator<T>
    {
        private ArraySliceEnumerator<T> arrayEnumerator;
        private SortedSet<T>.Enumerator setEnumerator;

        public Enumerator(Sorted<T> source)
        {
            if (source.value == null)
            {
                arrayEnumerator = new ArraySliceEnumerator<T>(Array.Empty<T>(), 0);
                setEnumerator = default;
            }
            else if (source.value is T[] array)
            {
                arrayEnumerator = new ArraySliceEnumerator<T>(array, source.Count);
                setEnumerator = default;
            }
            else
            {
                arrayEnumerator = default;
                var set = (SortedSet<T>)source.value;
                setEnumerator = set.GetEnumerator();
            }
        }

        public bool MoveNext()
        {
            if (arrayEnumerator.IsInitialized)
            {
                var result = arrayEnumerator.MoveNext();
                Current = arrayEnumerator.Current;
                return result;
            }
            else
            {
                var result = setEnumerator.MoveNext();
                Current = setEnumerator.Current;
                return result;
            }
        }

        public void Reset() => throw new NotSupportedException();

        public T Current { get; private set; } = default;
        object IEnumerator.Current => Current;

        public void Dispose()
        {
            // ArraySliceEnumerator<T> and SortedSet<T>.Enumerator not need to be disposed
        }
    }
}