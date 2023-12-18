#if NET8_0_OR_GREATER

namespace XL.Report.Tests;

public sealed class BTreeSlimShould
{
    [Test]
    public void METHOD()
    {
        var sut = new BTreeSlim<int, Item>();
        
        sut.TryAdd(new Item(1, "a"));
        sut.TryAdd(new Item(2, "b"));
        sut.TryAdd(new Item(3, "c"));
        sut.TryAdd(new Item(4, "d"));
        sut.TryAdd(new Item(5, "e"));
        sut.TryAdd(new Item(6, "f"));
        sut.TryAdd(new Item(7, "g"));
        sut.TryAdd(new Item(8, "h"));

        var items = sut.ToList();

        var iterator = sut.CreateIterator();
        iterator.MoveNext();

        while (iterator.State != IteratorState.AfterTree)
        {
            Console.WriteLine(iterator.Current);
            iterator.MoveNext();
        }

        Console.WriteLine();
        Console.WriteLine("---");
        Console.WriteLine();

        iterator.MovePrevious();
        while (iterator.State != IteratorState.BeforeTree)
        {
            Console.WriteLine(iterator.Current);
            iterator.MovePrevious();
        }
    }

    private readonly record struct Item(int Key, string Value) : IKeyed<int>
    {
        int IKeyed<int>.Key => Key;
    }
}

#endif