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

namespace XL.Report;

public sealed class Table<T>(params Table<T>.Column[] columns)
{
    private Column[] Columns => columns;

    private sealed class HeaderUnit(Table<T> table) : IUnit<Body>
    {
        public Body Write(SheetWindow window)
        {
            var row = new XL.Report.Row(
                table.Columns
                    .Select(column => column.Header)
                    .ToArray()
            );

            var ranges = (row as IUnit<Range[]>).Write(window);
            var columns = table.Columns.Zip(
                ranges,
                (column, range) => column.Col(range.Size.Width)
            ).ToArray();
            return new Body(columns);
        }
    }

    public IUnit<Body> Header => new HeaderUnit(this);

    public sealed class Column(IUnit<Range> header, Func<T, IUnit<Range>> value)
    {
        public IUnit<Range> Header => header;
        public ValueColumn Col(int width) => new(width, value);
    }

    public sealed class ValueColumn(int width, Func<T, IUnit<Range>> value)
    {
        public int Width => width;
        public IUnit<Range> Value(T item) => value(item);
    }

    public sealed class Body(ValueColumn[] columns)
    {
        // ReSharper disable MemberHidesStaticFromOuterClass
        public Row Row(T item) => new(columns, item);
        // ReSharper restore MemberHidesStaticFromOuterClass
    }

    public sealed class Row(ValueColumn[] columns, T item)
        : IUnit<Range>, IUnit<Range[]>
    {
        public Range Write(SheetWindow window)
        {
            var usedWidth = 0;
            var usedHeight = 0;
            foreach (var column in columns)
            {
                var windowHeight = window.Range.Size.Height;
                var columnArea = new Reduction(
                    new Offset(usedWidth, 0),
                    new Size(column.Width, windowHeight)
                );
                using (window.Reduce(columnArea))
                {
                    var value = column.Value(item);
                    var range = value.Write(window);
                    usedHeight = Math.Max(usedHeight, range.Bottom);
                }

                usedWidth += column.Width;
            }

            var usedSize = new Size(usedWidth, usedHeight);
            return new Range(window.Range.LeftTop, usedSize);
        }

        Range[] IUnit<Range[]>.Write(SheetWindow window)
        {
            var result = new Range[columns.Length];
            var usedWidth = 0;
            var usedHeight = 0;
            for (var i = 0; i < columns.Length; i++)
            {
                var column = columns[i];
                var windowHeight = window.Range.Size.Height;
                var columnArea = new Reduction(
                    new Offset(usedWidth, 0),
                    new Size(column.Width, windowHeight)
                );
                using (window.Reduce(columnArea))
                {
                    var value = column.Value(item);
                    var range = value.Write(window);
                    usedHeight = Math.Max(usedHeight, range.Bottom);
                    result[i] = Range.Create(
                        window.Range.LeftTop,
                        new Location(
                            window.Range.LeftTop.X + column.Width - 1,
                            range.Bottom
                        )
                    );
                }

                usedWidth += column.Width;
            }

            AdjustHeight();
            return result;

            void AdjustHeight()
            {
                for (var i = 0; i < result.Length; i++)
                {
                    var range = result[i];
                    result[i] = Range.Create(
                        range.LeftTop,
                        new Location(
                            range.LeftTop.X + range.Width - 1,
                            usedHeight
                        )
                    );
                }
            }
        }
    }
}