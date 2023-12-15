using System.Xml;
using XL.Report.Styles;

namespace XL.Report;

// todo delete
public sealed class RegularSheetWindow : SheetWindow
{
    private readonly Dictionary<Offset, Cell> placed = new();
    private readonly Range startRange;
    private readonly Stack<Offset> moves = new();

    public RegularSheetWindow(Range range)
    {
        startRange = range;
    }

    public override Range Range
    {
        get
        {
            var currentOffset = CurrentOffset;
            return new Range(startRange.LeftTop + currentOffset, startRange.Size - currentOffset);
        }
    }

    public IEnumerable<Row> Rows => placed
        .GroupBy(pair => pair.Key.Y)
        .OrderBy(row => row.Key)
        .Select(
            row =>
            {
                var y = row.Key + Range.LeftTop.Y;
                var contents = row
                    .Select(pair => (pair.Key.X + Range.LeftTop.X, pair.Value))
                    .OrderBy(pair => pair.Item1);
                return new Row(y, contents);
            }
        );

    private Offset CurrentOffset => moves.Count == 0
        ? Offset.Zero
        : moves.Peek();

    public override void Place(Content content, StyleId? styleId)
    {
        if (Range.IsEmpty)
        {
            throw new InvalidOperationException();
        }

        placed[CurrentOffset] = new Cell(content, styleId);
    }

    public override void Merge(Size size, Content content, StyleId? styleId)
    {
        // todo check range can contain
        
        throw new NotImplementedException();
    }

    public override void PushReduce(Offset offset, Size? newSize = null)
    {
        // todo use newSize
        // todo check offset moves to right or bottom

        moves.Push(CurrentOffset + offset);
    }

    public override void PopReduce()
    {
        moves.Pop();
    }

    public readonly record struct Cell(Content Content, StyleId? StyleId)
    {
        public void Write(XmlWriter xml, Location location)
        {
            xml.WriteStartElement(XlsxStructure.Worksheet.Cell);
            xml.WriteAttributeString(XlsxStructure.Worksheet.Reference, location.ToString());
            if (StyleId is { } styleId)
            {
                xml.WriteAttributeInt(XlsxStructure.Worksheet.Style, styleId.Index);
            }

            {
                Content.Write(xml);
            }
            xml.WriteEndElement();
        }
    }

    // todo write content here instead of return
    public sealed class Row
    {
        public Row(int y, IEnumerable<(int X, Cell Cell)> contents)
        {
            Y = y;
            Contents = contents;
        }

        public int Y { get; }
        public IEnumerable<(int X, Cell Cell)> Contents { get; }
    }
}