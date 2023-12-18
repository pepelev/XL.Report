using FluentAssertions;

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

        var iterator = sut.BeforeTree();
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

    [Test]
    public void Bug37()
    {
        var sut = new BTreeSlim<int, Item>();

        for (var i = 1; i <= 37; i++)
        {
            sut.TryAdd(new Item(i, "a"));
        }

        sut.ToList().Should().HaveCount(37);
    }

    [Test]
    public void Left_Lower_Bound([Range(0, 16)] int key)
    {
        var sut = new BTreeSlim<int, Item>();

        sut.TryAdd(new Item(1, "a"));
        sut.TryAdd(new Item(3, "b"));
        sut.TryAdd(new Item(5, "c"));
        sut.TryAdd(new Item(7, "d"));
        sut.TryAdd(new Item(9, "e"));
        sut.TryAdd(new Item(11, "f"));
        sut.TryAdd(new Item(13, "g"));
        sut.TryAdd(new Item(15, "h"));

        var naiveResult = LeftLowerBound(sut, key);
        var result = sut.LeftLowerBound(key);

        Item? a = result.State == IteratorState.InsideTree
            ? result.Current
            : null;

        a.Should().Be(naiveResult);

        if (result.State == IteratorState.InsideTree)
        {
            Console.WriteLine($"{key}: {result.Current}");
        }
        else
        {
            Console.WriteLine($"{key}: {result.State}");
        }
    }

    [Test]
    public void Right_Lower_Bound([Range(0, 16)] int key)
    {
        var sut = new BTreeSlim<int, Item>();

        sut.TryAdd(new Item(1, "a"));
        sut.TryAdd(new Item(3, "b"));
        sut.TryAdd(new Item(5, "c"));
        sut.TryAdd(new Item(7, "d"));
        sut.TryAdd(new Item(9, "e"));
        sut.TryAdd(new Item(11, "f"));
        sut.TryAdd(new Item(13, "g"));
        sut.TryAdd(new Item(15, "h"));

        var naiveResult = RightLowerBound(sut, key);
        var result = sut.RightLowerBound(key);

        Item? a = result.State == IteratorState.InsideTree
            ? result.Current
            : null;

        a.Should().Be(naiveResult);

        if (result.State == IteratorState.InsideTree)
        {
            Console.WriteLine($"{key}: {result.Current}");
        }
        else
        {
            Console.WriteLine($"{key}: {result.State}");
        }
    }

    private Item? LeftLowerBound(IEnumerable<Item> items, int key)
    {
        foreach (var item in items)
        {
            if (item.Key >= key)
            {
                return item;
            }
        }

        return null;
    }

    private Item? RightLowerBound(IEnumerable<Item> items, int key)
    {
        foreach (var item in items.Reverse())
        {
            if (item.Key <= key)
            {
                return item;
            }
        }

        return null;
    }

    private readonly record struct Item(int Key, string Value) : IKeyed<int>
    {
        int IKeyed<int>.Key => Key;
    }
}

#endif