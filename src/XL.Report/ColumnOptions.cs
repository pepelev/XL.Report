using System.Collections;
using System.Collections.Immutable;
using System.Diagnostics.Contracts;
using XL.Report.Styles;

namespace XL.Report;

public sealed record ColumnOptions(float? Width = null, StyleId? StyleId = null, bool Hidden = false)
{
    public static ColumnOptions Default { get; } = new();

    public sealed class Collection : IReadOnlyCollection<(int X, ColumnOptions Options)>
    {
        private readonly ImmutableSortedDictionary<int, ColumnOptions> content;

        private Collection(ImmutableSortedDictionary<int, ColumnOptions> content)
        {
            this.content = content;
        }

        [Pure]
        public Collection With(int x, ColumnOptions options)
        {
            if (options == ColumnOptions.Default)
            {
                return content.ContainsKey(x)
                    ? new Collection(content.Remove(x))
                    : this;
            }

            return new Collection(content.SetItem(x, options));
        }

        // ReSharper disable MemberHidesStaticFromOuterClass
        public static Collection Default { get; } = new(ImmutableSortedDictionary<int, ColumnOptions>.Empty);
        // ReSharper restore MemberHidesStaticFromOuterClass

        public IEnumerator<(int X, ColumnOptions Options)> GetEnumerator() =>
            content
                .Select(pair => (pair.Key, pair.Value))
                .GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        public int Count => content.Count;
    }
}