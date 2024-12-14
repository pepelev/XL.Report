#region Legal
// Copyright 2024 Pepelev Alexey
// 
// This file is part of XL.Report.
// 
// XL.Report is free software: you can redistribute it and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation, either
// version 3 of the License, or (at your option) any later version.
// 
// XL.Report is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
// without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU Lesser General Public License for more details.
// 
// You should have received a copy of the GNU Lesser General Public License along with XL.Report.
// If not, see <https://www.gnu.org/licenses/>.
#endregion

using System.Collections;
using System.Collections.Immutable;
using System.Diagnostics.Contracts;
using XL.Report.Styles;

namespace XL.Report;

public sealed record ColumnOptions(float? Width = null, StyleId? StyleId = null, bool Hidden = false)
{
    public static ColumnOptions Default { get; } = new();

    public void Write(Xml xml, int x)
    {
        using (xml.WriteStartElement("col"))
        {
            xml.WriteAttribute("min", x);
            xml.WriteAttribute("max", x);
            if (Width is { } width)
            {
                xml.WriteAttribute("width", width, "N6");
            }

            if (StyleId is { } styleId)
            {
                xml.WriteAttribute("style", styleId);
            }

            if (Hidden)
            {
                xml.WriteAttribute("hidden", "true");
            }
        }
    }

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

        public void Write(Xml xml)
        {
            if (Count > 0)
            {
                using (xml.WriteStartElement("cols"))
                {
                    foreach (var (x, column) in content)
                    {
                        column.Write(xml, x);
                    }
                }
            }
        }

        public IEnumerator<(int X, ColumnOptions Options)> GetEnumerator() =>
            content
                .Select(pair => (pair.Key, pair.Value))
                .GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        public int Count => content.Count;
    }
}