using System.Collections;
using System.Runtime.CompilerServices;

namespace XL.Report;

public interface IKeyed<out T>
{
    T Key { get; }
}

// todo return arrays into pool
// todo make internal
// public struct Sorted<TKey, T> : IReadOnlyCollection<T>
//     where TKey : IComparable<TKey>
//     where T : struct, IKeyed<TKey>
// {
//     public int Count { get; private set; }
//
//     // T[] | SortedSet<T>
//     private object? value;
//
//     public void Add(T item)
//     {
//         var added = TryAdd(item);
//         if (!added.IsSuccess)
//         {
//             throw new InvalidOperationException();
//         }
//     }
//
// #if NET7_0_OR_GREATER
//     public readonly ref struct AdditionResult
//     {
//         // todo move to specific type with single generic
//         private static T @default;
//
//         private readonly ref T conflictingValue;
//
//         public static AdditionResult Success => new(ref @default);
//
//         public AdditionResult(ref T conflictingValue)
//         {
//             this.conflictingValue = ref conflictingValue;
//         }
//
//         public ref T ConflictingValue => ref conflictingValue;
//         public bool IsSuccess => Unsafe.AreSame(ref @default, ref conflictingValue);
//
//         public void ThrowOnConflict()
//         {
//             if (!IsSuccess)
//             {
//                 throw new InvalidOperationException();
//             }
//         }
//     }
// #endif
//
//     public AdditionResult TryAdd(T item)
//     {
//         if (value == null)
//         {
//             var newValue = new T[4];
//             newValue[0] = item;
//             value = newValue;
//             Count = 1;
//             return true;
//         }
//
//         if (value is T[] array)
//         {
//             var comparison = array[Count - 1].Key.CompareTo(item.Key);
//             if (comparison < 0)
//             {
//                 if (array.Length <= Count)
//                 {
//                     // todo pool
//                     var newSize = Math.Max(array.Length * 2, 4);
//                     Array.Resize(ref array, newSize);
//                 }
//
//                 array[Count++] = item;
//                 value = array;
//                 return true;
//             }
//
//             if (comparison == 0)
//             {
//                 return false;
//             }
//
//             var set = new SortedSet<T>();
//             for (var i = 0; i < Count; i++)
//             {
//                 set.Add(array[i]);
//             }
//
//             if (!set.Add(item))
//             {
//                 return false;
//             }
//
//             value = set;
//             Count++;
//             return true;
//         }
//         else
//         {
//             var set = (SortedSet<T>)value;
//             if (!set.Add(item))
//             {
//                 return false;
//             }
//
//             Count++;
//             return true;
//         }
//     }
//
//     public Enumerator GetEnumerator() => new(this);
//     IEnumerator<T> IEnumerable<T>.GetEnumerator() => GetEnumerator();
//     IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
//
//     public struct Enumerator : IEnumerator<T>
//     {
//         private ArraySliceEnumerator<T> arrayEnumerator;
//         private SortedSet<T>.Enumerator setEnumerator;
//
//         public Enumerator(Sorted<TKey, T> source)
//         {
//             if (source.value == null)
//             {
//                 arrayEnumerator = new ArraySliceEnumerator<T>(Array.Empty<T>(), 0);
//                 setEnumerator = default;
//             }
//             else if (source.value is T[] array)
//             {
//                 arrayEnumerator = new ArraySliceEnumerator<T>(array, source.Count);
//                 setEnumerator = default;
//             }
//             else
//             {
//                 arrayEnumerator = default;
//                 var set = (SortedSet<T>)source.value;
//                 setEnumerator = set.GetEnumerator();
//             }
//         }
//
//         public bool MoveNext()
//         {
//             if (arrayEnumerator.IsInitialized)
//             {
//                 var result = arrayEnumerator.MoveNext();
//                 Current = arrayEnumerator.Current;
//                 return result;
//             }
//             else
//             {
//                 var result = setEnumerator.MoveNext();
//                 Current = setEnumerator.Current;
//                 return result;
//             }
//         }
//
//         public void Reset() => throw new NotSupportedException();
//
//         public T Current { get; private set; } = default;
//         object IEnumerator.Current => Current;
//
//         public void Dispose()
//         {
//             // ArraySliceEnumerator<T> and SortedSet<T>.Enumerator not need to be disposed
//         }
//     }
// }