using System.Text;
using System.Xml;
using XL.Report.Styles;

namespace XL.Report;

public sealed class StreamSheetWindow : SheetWindow, IDisposable
{
    private readonly SheetOptions options;
    private readonly BTreeSlim<int, Row> rows = new();
    private readonly Stack<Range> reductions = new();
    private readonly XmlWriter xml;
    private Range activeRange;
    private bool started;

    public StreamSheetWindow(Stream stream, SheetOptions options)
    {
        this.options = options;
        activeRange = Range.EntireSheet;
        var settings = new XmlWriterSettings
        {
            Encoding = Encoding.UTF8,
            CloseOutput = true
        };
        xml = XmlWriter.Create(stream, settings);
    }

    public override Range Range => reductions.Count > 0
        ? reductions.Peek()
        : activeRange;

    public void Dispose()
    {
        xml.Dispose();
    }

    public override void Place(Content content, StyleId? styleId)
    {
        var range = Range;
        if (range.IsEmpty)
        {
            throw new InvalidOperationException();
        }

        var cell = new Cell(content, styleId);
        var result = rows.Find(range.Top);
        if (result.Found)
        {
            ref var row = ref result.Result;
            row.Add(new CellX(range.Left, cell));
        }
        else
        {
            var newRow = new Row(range.Top);
            newRow.Add(new CellX(range.Left, cell));
            rows.TryAdd(newRow);
        }
    }

    public override void Merge(Size size, Content content, StyleId? styleId)
    {
        // todo check range can contain

        throw new NotImplementedException();
    }

    public override void PushReduce(Offset offset, Size? newSize = null)
    {
        var current = Range;
        // todo check size nonnegative
        var @new = new Range(current.LeftTop + offset, newSize ?? current.Size - offset);
        if (!current.Contains(@new))
        {
            throw new InvalidOperationException();
        }

        reductions.Push(@new);
    }

    public override void PopReduce()
    {
        reductions.Pop();
    }

    public void Flush()
    {
        int Write()
        {
            var mostDownY = activeRange.LeftTop.Y;
            
            foreach (var row in rows)
            {
                xml.WriteStartElement(XlsxStructure.Worksheet.Row);
                xml.WriteStartAttribute("r");
                xml.WriteValue(row.Y);
                {
                    foreach (var (x, content) in row)
                    {
                        var location = new Location(x, row.Y);
                        content.Write(xml, location);
                    }
                }
                xml.WriteEndElement();

                mostDownY = Math.Max(mostDownY, row.Y);
            }

            return mostDownY;
        }

        if (reductions.Count > 0)
        {
            throw new InvalidOperationException();
        }

        WriteStartOnlyFirstTime();
        var mostDownY = Write();
        var newActiveRange = activeRange.ReduceDown(mostDownY + 1 - activeRange.LeftTop.Y);

        activeRange = newActiveRange;
        rows.Clear();
    }

    private void WriteStartOnlyFirstTime()
    {
        if (started)
        {
            return;
        }

        xml.WriteStartDocument(true);
        xml.WriteStartElement("worksheet", XlsxStructure.Namespaces.Spreadsheet.Main);
        xml.WriteAttributeString("xmlns", "r", "", XlsxStructure.Namespaces.OfficeDocuments.Relationships);

        var freeze = options.Freeze;
        if (freeze != FreezeOptions.None)
        {
            xml.WriteStartElement("sheetViews");
            {
                xml.WriteStartElement("sheetView");
                xml.WriteAttributeString("workbookViewId", "0");
                {
                    xml.WriteStartElement("pane");
                    if (freeze.FreezeByX > 0)
                    {
                        xml.WriteAttributeInt("xSplit", freeze.FreezeByX);
                    }

                    if (freeze.FreezeByY > 0)
                    {
                        xml.WriteAttributeInt("ySplit", freeze.FreezeByY);
                    }

                    var topLeftCell = new Location(freeze.FreezeByX + 1, freeze.FreezeByY + 1);
                    xml.WriteAttributeString("topLeftCell", topLeftCell.AsString());
                    xml.WriteAttributeString("state", "frozen");
                    xml.WriteEndElement();
                }
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }

        if (options.Columns.Count > 0)
        {
            xml.WriteStartElement("cols");
            {
                foreach (var (x, width) in options.Columns)
                {
                    xml.WriteStartElement("col");
                    xml.WriteAttributeInt("min", x);
                    xml.WriteAttributeInt("max", x);
                    xml.WriteAttributeString("width", width.ToString("N6"));
                    xml.WriteEndElement();
                }
            }
            xml.WriteEndElement();
        }

        xml.WriteStartElement("sheetData");

        started = true;
    }

    public void Complete()
    {
        WriteStartOnlyFirstTime();
        xml.WriteEndDocument();
    }

    private readonly struct Row : IKeyed<int>
    {
        public int Y { get; }
        int IKeyed<int>.Key => Y;
        private readonly BTreeSlim<int, CellX> cells = new();

        public Row(int y)
        {
            Y = y;
        }

        public void Add(CellX cell)
        {
            cells.TryAdd(cell).ThrowOnConflict();
        }

        public BTreeSlim<int, CellX>.Enumerator GetEnumerator() => cells.GetEnumerator();
    }

    // todo pack compact (style? and int (actually short) can be packed into long)
    private readonly record struct CellX(int X, Cell Cell) : IKeyed<int>
    {
        int IKeyed<int>.Key => X;
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
}