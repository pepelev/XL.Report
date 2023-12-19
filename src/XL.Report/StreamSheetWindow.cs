using System.Text;
using System.Xml;
using XL.Report.Styles;

namespace XL.Report;

public sealed class StreamSheetWindow : SheetWindow, IDisposable
{
    private readonly SheetOptions options;
    private readonly BTreeSlim<int, Row> rows = new();
    private readonly List<Range> mergedRanges = new();
    private readonly Stack<Range> reductions = new();
    private readonly Xml xml;
    private Range activeRange;
    private (Xml.Block Document, Xml.Block SheetData)? started;

    public StreamSheetWindow(Stream stream, SheetOptions options)
    {
        this.options = options;
        activeRange = Range.EntireSheet;
        var settings = new XmlWriterSettings
        {
            Encoding = Encoding.UTF8,
            CloseOutput = true
        };
        xml = new Xml(XmlWriter.Create(stream, settings));
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
        if (!size.HasArea)
        {
            throw new ArgumentException($"must {nameof(size.HasArea)}", nameof(size));
        }

        var range = Range;
        if (!range.Size.Contains(size))
        {
            throw new ArgumentException($"is bigger than {nameof(range)}", nameof(size));
        }

        var top = range.Top;
        var mergeRow = new MergeRow(
            top,
            new Interval<int>(range.Left, range.Left + size.Width)
        );
        for (var i = 0; i < size.Height; i++)
        {
            var result = rows.Find(top + i);
            if (result.Found)
            {
                ref var row = ref result.Result;
                if (!row.CanAdd(mergeRow.Span))
                {
                    throw new ArgumentException("overlaps with already placed content", nameof(size));
                }
            }
        }

        for (var i = 0; i < size.Height; i++)
        {
            var result = rows.Find(top + i);
            if (result.Found)
            {
                ref var row = ref result.Result;
                row.Add(mergeRow, content, styleId);
            }
            else
            {
                var newRow = new Row(range.Top);
                newRow.Add(mergeRow, content, styleId);
                rows.TryAdd(newRow);
            }
        }

        mergedRanges.Add(new Range(range.LeftTop, size));
    }

    public override void PushReduce(Offset offset, Size? newSize = null)
    {
        if (newSize?.IsDegenerate == true)
        {
            throw new ArgumentException("is degenerate", nameof(newSize));
        }

        var current = Range;
        var @new = new Range(current.LeftTop + offset, newSize ?? current.Size - offset);
        if (!current.Contains(@new))
        {
            throw current.Contains(new Range(current.LeftTop + offset, Size.Empty))
                ? new ArgumentException("is too big", nameof(newSize))
                : new ArgumentException($"is out of current {nameof(Range)}", nameof(offset));
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
                using (xml.WriteStartElement(XlsxStructure.Worksheet.Row))
                {
                    xml.WriteAttributeSpan("r", row.Y);
                    foreach (var (x, content) in row)
                    {
                        var location = new Location(x, row.Y);
                        content.Write(xml, location);
                    }
                }

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

    private (Xml.Block Document, Xml.Block SheetData) WriteStartOnlyFirstTime()
    {
        if (started.HasValue)
        {
            return started.Value;
        }

        var document = xml.WriteStartDocument("worksheet", XlsxStructure.Namespaces.Spreadsheet.Main);
        xml.WriteAttribute("xmlns", "r", XlsxStructure.Namespaces.OfficeDocuments.Relationships);

        var freeze = options.Freeze;
        if (freeze != FreezeOptions.None)
        {
            using (xml.WriteStartElement("sheetViews"))
            {
                using (xml.WriteStartElement("sheetView"))
                {
                    xml.WriteAttribute("workbookViewId", "0");
                    using (xml.WriteStartElement("pane"))
                    {
                        if (freeze.FreezeByX > 0)
                        {
                            xml.WriteAttributeSpan("xSplit", freeze.FreezeByX);
                        }

                        if (freeze.FreezeByY > 0)
                        {
                            xml.WriteAttributeSpan("ySplit", freeze.FreezeByY);
                        }

                        var topLeftCell = new Location(freeze.FreezeByX + 1, freeze.FreezeByY + 1);
                        xml.WriteAttribute("topLeftCell", topLeftCell);
                        xml.WriteAttribute("state", "frozen");
                    }
                }
            }
        }

        if (options.Columns.Count > 0)
        {
            using (xml.WriteStartElement("cols"))
            {
                foreach (var (x, width) in options.Columns)
                {
                    using (xml.WriteStartElement("col"))
                    {
                        xml.WriteAttributeSpan("min", x);
                        xml.WriteAttributeSpan("max", x);
                        xml.WriteAttributeSpan("width", width, "N6");
                    }
                }
            }
        }

        var sheetData = xml.WriteStartElement("sheetData");
        started = (document, sheetData);
        return started.Value;
    }

    public void Complete()
    {
        var (document, sheetData) = WriteStartOnlyFirstTime();
        sheetData.Dispose();

        if (mergedRanges.Count > 0)
        {
            using (xml.WriteStartElement("mergeCells"))
            {
                foreach (var range in mergedRanges)
                {
                    using (xml.WriteStartElement("mergeCell"))
                    {
                        xml.WriteAttribute("ref", range);
                    }
                }
            }
        }

        document.Dispose();
    }

    private readonly struct Row : IKeyed<int>
    {
        public int Y { get; }
        int IKeyed<int>.Key => Y;
        private readonly BTreeSlim<int, CellX> cells = new();
        private readonly BTreeSlim<int, MergeRow> merges = new();

        public Row(int y)
        {
            Y = y;
        }

        public void Add(CellX cell)
        {
            if (!CanAdd(new Interval<int>(cell.X, cell.X)))
            {
                throw new InvalidOperationException();
            }

            cells.TryAdd(cell).ThrowOnConflict();
        }

        public bool CanAdd(Interval<int> span)
        {
            var mergeBoundIterator = merges.RightLowerBound(span.LeftInclusive);
            while (mergeBoundIterator.State == IteratorState.InsideTree)
            {
                if (mergeBoundIterator.Current.Span.RightThan(span))
                {
                    break;
                }

                if (mergeBoundIterator.Current.Span.Intersect(span))
                {
                    return false;
                }

                mergeBoundIterator.MoveNext();
            }

            var cellBoundIterator = cells.LeftLowerBound(span.LeftInclusive);
            if (cellBoundIterator.State == IteratorState.InsideTree)
            {
                if (cellBoundIterator.Current.X <= span.RightInclusive)
                {
                    return false;
                }
            }

            return true;
        }

        public void Add(MergeRow merge, Content content, StyleId? styleId)
        {
            if (!CanAdd(merge.Span))
            {
                throw new InvalidOperationException();
            }

            if (Y == merge.Top)
            {
                var cellX = new CellX(merge.Span.LeftInclusive, new Cell(content, styleId));
                cells.TryAdd(cellX).ThrowOnConflict();
            }

            merges.TryAdd(merge).ThrowOnConflict();
        }

        public BTreeSlim<int, CellX>.Enumerator GetEnumerator() => cells.GetEnumerator();
    }

    // todo pack left and right to shorts
    private readonly record struct MergeRow(int Top, Interval<int> Span) : IKeyed<int>
    {
        int IKeyed<int>.Key => Span.LeftInclusive;
    }

    // todo pack compact (style? and int (actually short) can be packed into long)
    private readonly record struct CellX(int X, Cell Cell) : IKeyed<int>
    {
        int IKeyed<int>.Key => X;
    }

    public readonly record struct Cell(Content Content, StyleId? StyleId)
    {
        public void Write(Xml xml, Location location)
        {
            using (xml.WriteStartElement(XlsxStructure.Worksheet.Cell))
            {
                xml.WriteAttribute(XlsxStructure.Worksheet.Reference, location);

                if (StyleId is { } styleId)
                {
                    xml.WriteAttributeSpan(XlsxStructure.Worksheet.Style, styleId.Index);
                }

                Content.Write(xml);
            }
        }
    }
}