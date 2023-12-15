using System.Collections;
using System.Collections.Immutable;
using System.Diagnostics.Contracts;

namespace XL.Report;

public sealed class ColumnWidths : IReadOnlyCollection<(int X, float Width)>
{
    private readonly ImmutableDictionary<int, float> content;

    public static readonly ColumnWidths Default = new(ImmutableDictionary<int, float>.Empty);

    private ColumnWidths(ImmutableDictionary<int, float> content)
    {
        this.content = content;
    }

    public IEnumerator<(int X, float Width)> GetEnumerator() =>
        content.Select(pair => (pair.Key, pair.Value)).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    [Pure]
    public ColumnWidths With(int x, float width) => new(content.SetItem(x, width));

    public int Count => content.Count;
}