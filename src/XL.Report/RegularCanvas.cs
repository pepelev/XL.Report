namespace XL.Report;

public sealed class RegularCanvas : Canvas
{
    private readonly Dictionary<Offset, Content> placed = new();

    public RegularCanvas(Range range)
    {
        Range = range;
    }

    public override Range Range { get; }

    public IEnumerable<Row> Rows => placed
        .GroupBy(pair => pair.Key.Y)
        .OrderBy(row => row.Key)
        .Select(
            row =>
            {
                var y = row.Key + Range.LeftTopCell.Y;
                var contents = row
                    .Select(pair => (pair.Key.X + Range.LeftTopCell.X, pair.Value))
                    .OrderBy(pair => pair.Item1);
                return new Row(y, contents);
            }
        );

    public override void Place(Offset offset, Content content, Style style)
    {
        placed[offset] = content;
    }

    public override void Merge(Offset offset, Size size, Content content, Style style)
    {
        throw new NotImplementedException();
    }

    public sealed class Row
    {
        public Row(int y, IEnumerable<(int X, Content Content)> contents)
        {
            Y = y;
            Contents = contents;
        }

        public int Y { get; }
        public IEnumerable<(int X, Content Content)> Contents { get; }
    }
}