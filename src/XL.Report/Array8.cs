#if NET8_0_OR_GREATER

using System.Collections;
using System.Diagnostics.Contracts;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace XL.Report;

internal static class Default<T>
{
    private static T Value = default!;

    public static ref T Ref() => ref Value;
}

[InlineArray(8)]
public struct Array8<T>
{
    public T Value0;
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



public enum IteratorState
{
    BeforeTree,
    InsideTree,
    AfterTree
}

internal enum PushStrategy
{
    First,
    Last
}

internal enum Move
{
    Next,
    Previous
}

public sealed class BTreeSlim<TKey, T> : IEnumerable<T>
    where TKey : IComparable<TKey>
    where T : IKeyed<TKey>
{
    private const int t = 4;

    // todo use pool of nodes
    // https://learn.microsoft.com/ru-ru/dotnet/api/microsoft.extensions.objectpool.defaultobjectpool-1
    private Node root = new Node.Leaf();

    public bool IsEmpty => root is Node.Leaf { Items.Count: 0 };

    public AdditionResult TryAdd(T item)
    {
        if (root.IsFull)
        {
            var newRoot = new Node.NonLeaf();
            newRoot.Children.Content[0] = root;
            newRoot.Children.Count = 1;
            newRoot.SplitChild(0);
            root = newRoot;
        }

        return root.TryAdd(item);
    }

    public FindResult Find(TKey key)
    {
        return root.Find(key);
    }

    public Iterator CreateIterator() => new(this);

    public void Clear()
    {
        root = new Node.Leaf();
    }

    public struct Iterator
    {
        private readonly BTreeSlim<TKey, T> tree;
        private Move lastMove;

        // todo Buffer8 is too small
        private Stack<(Node Node, int Index), Buffer8<(Node Node, int Index)>> stack = new(new Buffer8<(Node Node, int Index)>());

        public Iterator(BTreeSlim<TKey, T> tree)
        {
            this.tree = tree;
            lastMove = Move.Previous;
        }

        public void ToBeforeTree()
        {
            stack.Clear();
            lastMove = Move.Previous;
        }

        public void ToAfterTree()
        {
            stack.Clear();
            lastMove = Move.Next;
        }

        public ref T Current
        {
            get
            {
                if (State == IteratorState.InsideTree)
                {
                    ref var entry = ref stack.Peek();
                    if (entry.Node is Node.Leaf leaf)
                    {
                        return ref leaf.Items.Content[entry.Index];
                    }

                    var nonLeaf = (Node.NonLeaf)entry.Node;
                    return ref nonLeaf.Items.Content[entry.Index];
                }

                throw new InvalidOperationException($"{nameof(Iterator)} is in wrong state: {State}");
            }
        }

        public void MoveNext()
        {
            MoveNextInternal();
            lastMove = Move.Next;
        }

        private void MoveNextInternal()
        {
            if (State == IteratorState.BeforeTree)
            {
                Push(tree.root, PushStrategy.First);
                return;
            }

            if (State == IteratorState.InsideTree)
            {
                if (stack.IsEmpty)
                {
                    return;
                }

                ref var currentFrame = ref stack.Peek();
                if (currentFrame.Node is Node.Leaf leaf)
                {
                    currentFrame.Index++;
                    if (leaf.Items.Count <= currentFrame.Index)
                    {
                        stack.Pop();
                    }
                }
                else
                {
                    var nonLeaf = (Node.NonLeaf)currentFrame.Node;
                    var newFrameIndex = currentFrame.Index + 1;
                    if (nonLeaf.Items.Count <= newFrameIndex)
                    {
                        stack.Pop();
                    }

                    Push(nonLeaf.Children.Content[newFrameIndex], PushStrategy.First);
                }
            }
        }

        private void Push(Node node, PushStrategy strategy)
        {
            while (true)
            {
                if (node is Node.Leaf leaf)
                {
                    if (leaf.Items.Count > 0)
                    {
                        var index = strategy == PushStrategy.First
                            ? 0
                            : leaf.Items.Count - 1;
                        stack.Push((node, index));
                    }

                    return;
                }
                else
                {
                    var nonLeaf = (Node.NonLeaf)node;
                    var index = strategy == PushStrategy.First
                        ? 0
                        : nonLeaf.Items.Count - 1;
                    var childIndex = strategy == PushStrategy.First
                        ? 0
                        : index + 1;

                    stack.Push((nonLeaf, index));

                    node = nonLeaf.Children.Content[childIndex];
                }
            }
        }

        public void MovePrevious()
        {
            MovePreviousInternal();
            lastMove = Move.Previous;
        }

        private void MovePreviousInternal()
        {
            if (State == IteratorState.AfterTree)
            {
                Push(tree.root, PushStrategy.Last);
                return;
            }

            if (State == IteratorState.InsideTree)
            {
                if (stack.IsEmpty)
                {
                    return;
                }

                ref var currentFrame = ref stack.Peek();
                if (currentFrame.Node is Node.Leaf)
                {
                    currentFrame.Index--;
                    if (currentFrame.Index < 0)
                    {
                        stack.Pop();
                    }
                }
                else
                {
                    var nonLeaf = (Node.NonLeaf)currentFrame.Node;
                    var childIndex = currentFrame.Index;
                    currentFrame.Index--;
                    if (currentFrame.Index < 0)
                    {
                        stack.Pop();
                    }

                    Push(nonLeaf.Children.Content[childIndex], PushStrategy.Last);
                }
            }
        }

        public IteratorState State => (stack.IsEmpty, lastMove) switch
        {
            (true, Move.Next) => IteratorState.AfterTree,
            (true, Move.Previous) => IteratorState.BeforeTree,
            _ => IteratorState.InsideTree
        };
    }

    public struct Enumerator : IEnumerator<T>
    {
        private Iterator iterator;

        internal Enumerator(BTreeSlim<TKey, T> tree)
        {
            iterator = tree.CreateIterator();
        }

        public bool MoveNext()
        {
            iterator.MoveNext();
            return iterator.State == IteratorState.InsideTree;
        }

        public void Reset() => throw new NotSupportedException();
        public T Current => iterator.Current;
        object IEnumerator.Current => Current;

        public void Dispose()
        {
        }
    }

    public readonly ref struct FindResult
    {
        private readonly ref T found;

        public static FindResult Fail => new(ref Default<T>.Ref());

        public FindResult(ref T found)
        {
            this.found = ref found;
        }

        public ref T Result => ref found;
        public bool NotFound => Unsafe.AreSame(ref Default<T>.Ref(), ref found);
        public bool Found => !NotFound;

        public void ThrowOnNotFound()
        {
            if (NotFound)
            {
                throw new InvalidOperationException();
            }
        }
    }

    public readonly ref struct AdditionResult
    {
        private readonly ref T conflictingValue;

        public static AdditionResult Success => new(ref Default<T>.Ref());

        public AdditionResult(ref T conflictingValue)
        {
            this.conflictingValue = ref conflictingValue;
        }

        public ref T ConflictingValue => ref conflictingValue;
        public bool IsSuccess => Unsafe.AreSame(ref Default<T>.Ref(), ref conflictingValue);

        public void ThrowOnConflict()
        {
            if (!IsSuccess)
            {
                throw new InvalidOperationException();
            }
        }
    }

    internal abstract class Node
    {
        public abstract bool IsFull { get; }
        public abstract AdditionResult TryAdd(T item);
        public abstract (T MiddleItem, Node NewRightNode) Split();
        public abstract FindResult Find(TKey key);

        public sealed class Leaf : Node
        {
            public Buffer8<T> Items;
            public override bool IsFull => Items.Count == 2 * t - 1;

            public override AdditionResult TryAdd(T item)
            {
                // todo check completely erased in Release
                Contract.Assert(!IsFull);

                var key = item.Key;
                var i = Items.Count - 1;
                for (; i >= 0; i--)
                {
                    var comparison = key.CompareTo(Items.Content[i].Key);
                    if (comparison == 0)
                    {
                        return new AdditionResult(ref Items.Content[i]);
                    }

                    if (comparison > 0)
                    {
                        break;
                    }
                }

                var itemIndex = i + 1;

                Items.Content[itemIndex..Items.Count].CopyTo(Items.Content[(itemIndex + 1)..]);
                Items.Content[itemIndex] = item;

                Items.Count++;
                return AdditionResult.Success;
            }

            public override (T MiddleItem, Node NewRightNode) Split()
            {
                // todo check completely erased in Release
                Contract.Assert(IsFull);

                var central = Items.Content[t - 1];
                var newRightNode = new Leaf();

                Items.Content[t..].CopyTo(newRightNode.Items.Content);
                Items.Content[(t - 1)..].Clear();
                Items.Count = t - 1;
                newRightNode.Items.Count = t - 1;
                return (central, newRightNode);
            }

            public override FindResult Find(TKey key)
            {
                for (var i = 0; i < Items.Count; i++)
                {
                    var comparison = Items.Content[i].Key.CompareTo(key);
                    if (comparison == 0)
                    {
                        return new FindResult(ref Items.Content[i]);
                    }
                }

                return FindResult.Fail;
            }
        }

        public sealed class NonLeaf : Node
        {
            public Buffer8<Node> Children;

            // todo Count stored twice
            public Buffer8<T> Items;

            public override bool IsFull => Items.Count == 2 * t - 1;

            // this                         this
            //   | (index)    ->           /    \
            //  targetChild       targetChild  secondPart
            public void SplitChild(int index)
            {
                var targetChild = Children.Content[index];
                var split = targetChild.Split();

                Items.Content[index..Items.Count].CopyTo(Items.Content[(index + 1)..]);
                Items.Content[index] = split.MiddleItem;
                Children.Content[(index + 1)..Children.Count].CopyTo(Children.Content[(index + 2)..]);
                Children.Content[index + 1] = split.NewRightNode;
                Items.Count++;
                Children.Count++;
            }

            public override AdditionResult TryAdd(T item)
            {
                // todo check completely erased in Release
                Contract.Assert(!IsFull);

                var key = item.Key;
                var i = Items.Count - 1;
                for (; i >= 0; i--)
                {
                    var comparison = key.CompareTo(Items.Content[i].Key);
                    if (comparison == 0)
                    {
                        return new AdditionResult(ref Items.Content[i]);
                    }

                    if (0 < comparison)
                    {
                        break;
                    }
                }

                var childIndex = i + 1;
                var node = Children.Content[childIndex];
                if (node.IsFull)
                {
                    SplitChild(childIndex);
                    if (Items.Content[childIndex].Key.CompareTo(key) < 0)
                    {
                        childIndex++;
                    }
                }

                return Children.Content[childIndex].TryAdd(item);
            }

            public override (T MiddleItem, Node NewRightNode) Split()
            {
                // todo check completely erased in Release
                Contract.Assert(Items.Count == 2 * t - 1);

                var centralItem = Items.Content[t - 1];
                var newRightNode = new NonLeaf();

                Items.Content[t..].CopyTo(newRightNode.Items.Content);
                Items.Content[(t - 1)..].Clear();
                Items.Count = t - 1;
                newRightNode.Items.Count = t - 1;

                Children.Content[t..].CopyTo(newRightNode.Children.Content);
                Children.Content[t..].Clear();
                Children.Count = t;
                newRightNode.Children.Count = t;

                return (centralItem, newRightNode);
            }

            public override FindResult Find(TKey key)
            {
                for (var i = 0; i < Items.Count; i++)
                {
                    var comparison = key.CompareTo(Items.Content[i].Key);
                    if (comparison == 0)
                    {
                        return new FindResult(ref Items.Content[i]);
                    }

                    if (comparison < 0)
                    {
                        return Children.Content[i].Find(key);
                    }
                }

                return Children.Content[Items.Count].Find(key);
            }
        }
    }

    public Enumerator GetEnumerator() => new(this);
    IEnumerator<T> IEnumerable<T>.GetEnumerator() => GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}

#endif