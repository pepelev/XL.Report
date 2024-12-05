using System.Buffers;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using XL.Report.Styles;

namespace XL.Report;

internal sealed class StreamSheetWindow : SheetWindow, IDisposable
{
    private readonly List<Range> mergedRanges = new();
    private readonly SheetOptions options;
    private readonly Stack<Range> reductions = new();
    private readonly ReductionStage?[] reductionStages = new ReductionStage?[32];
    private readonly Dictionary<int, Row> rows = new();
    private readonly Xml xml;
    private int maxTouchedY = -1;
    private Range activeRange;
    private (Xml.Block Document, Xml.Block SheetData)? started;
    private bool valid = true; // todo set and use

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

    public override void TouchRow(int y)
    {
        var absoluteY = Range.LeftTop.Y + y - 1;
        maxTouchedY = Math.Max(maxTouchedY, absoluteY);
    }

    public override void Place(Content content, StyleId? styleId)
    {
        var range = Range;
        if (range.IsEmpty)
        {
            throw new InvalidOperationException();
        }

        if (!valid)
        {
            throw new InvalidOperationException();
        }

        var cell = new Cell(content, styleId);
        ref var row = ref CollectionsMarshal.GetValueRefOrAddDefault(rows, range.Top, out var exists);
        if (!exists)
        {
            row = new Row();
        }

        row.Place(range.Left, cell);
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

        if (!valid)
        {
            throw new InvalidOperationException();
        }

        var top = range.Top;
        var interval = new Interval<int>(range.Left, range.Left + size.Width - 1);
        for (var i = 0; i < size.Height; i++)
        {
            ref var row = ref CollectionsMarshal.GetValueRefOrAddDefault(rows, top + i, out var exists);
            if (!exists)
            {
                row = new Row();
            }

            var isMain = i == 0;
            row.Merge(isMain, interval, content, styleId);
        }

        mergedRanges.Add(new Range(range.LeftTop, size));
    }

    public override IDisposable Reduce(Reduction reduction)
    {
        var current = Range;
        var @new = new Range(current.LeftTop + reduction.Offset, reduction.NewSize ?? current.Size - reduction.Offset);
        if (!current.Contains(@new))
        {
            throw current.Contains(new Range(current.LeftTop + reduction.Offset, Size.Empty))
                ? new ArgumentException("is too big", nameof(reduction.NewSize))
                : new ArgumentException($"is out of current {nameof(Range)}", nameof(reduction.Offset));
        }

        var index = reductions.Count;
        var stage = Stage();
        reductions.Push(@new);
        return stage;

        ReductionStage Stage()
        {
            if (index < reductionStages.Length)
            {
                return reductionStages[index] ??= new ReductionStage(this, index);
            }

            return new ReductionStage(this, index);
        }
    }

    public void Flush(RowOptions? rowOptions = null)
    {
        if (reductions.Count > 0)
        {
            throw new InvalidOperationException();
        }

        WriteStartOnlyFirstTime();
        var mostDownY = Write();
        var newActiveRange = activeRange.ReduceDown(mostDownY + 1 - activeRange.LeftTop.Y);

        activeRange = newActiveRange;
        rows.Clear();
        return;

        int Write()
        {
            if (rows.Count <= 4)
            {
                Array4<KeyValuePair<int, Row>> buffer = default;
                return WriteRows(buffer.Content);
            }

            if (rows.Count <= 16)
            {
                Array16<KeyValuePair<int, Row>> buffer = default;
                return WriteRows(buffer.Content);
            }

            if (rows.Count <= 64)
            {
                Array64<KeyValuePair<int, Row>> buffer = default;
                return WriteRows(buffer.Content);
            }
            else
            {
                var pool = ArrayPool<KeyValuePair<int, Row>>.Shared;
                var buffer = pool.Rent(rows.Count);
                var mostDownY = WriteRows(buffer);
                pool.Return(buffer);
                return mostDownY;
            }
        }

        int WriteRows(Span<KeyValuePair<int, Row>> buffer)
        {
            var index = 0;
            foreach (var pair in rows)
            {
                buffer[index++] = pair;
            }

            var sortedRows = buffer[..index];
            sortedRows.Sort(default(ByKey<int, Row>));

            foreach (var (y, row) in sortedRows)
            {
                using (xml.WriteStartElement(XlsxStructure.Worksheet.Row))
                {
                    xml.WriteAttribute("r", y);
                    rowOptions?.WriteAttributes(xml);

                    if (row.CellsCount <= 4)
                    {
                        Array4<KeyValuePair<int, Cell>> cellBuffer = default;
                        WriteCells(cellBuffer.Content, row, y);
                    }
                    else if (row.CellsCount <= 16)
                    {
                        Array16<KeyValuePair<int, Cell>> cellBuffer = default;
                        WriteCells(cellBuffer.Content, row, y);
                    }
                    else if (row.CellsCount <= 64)
                    {
                        Array64<KeyValuePair<int, Cell>> cellBuffer = default;
                        WriteCells(cellBuffer.Content, row, y);
                    }
                    else
                    {
                        var pool = ArrayPool<KeyValuePair<int, Cell>>.Shared;
                        var cellBuffer = pool.Rent(rows.Count);
                        WriteCells(cellBuffer, row, y);
                        pool.Return(cellBuffer);
                    }
                }
            }

            int? mostDownWrittenRow = sortedRows.Length == 0
                ? null
                : sortedRows[^1].Key;

            if (rowOptions != null && rowOptions != RowOptions.Default)
            {
                for (var y = mostDownWrittenRow ?? activeRange.LeftTop.Y; y <= maxTouchedY; y++)
                {
                    using (xml.WriteStartElement(XlsxStructure.Worksheet.Row))
                    {
                        xml.WriteAttribute("r", y);
                        rowOptions.WriteAttributes(xml);
                    }
                }
            }

            return Math.Max(mostDownWrittenRow ?? activeRange.LeftTop.Y, maxTouchedY);
        }

        void WriteCells(Span<KeyValuePair<int, Cell>> span, Row row, int y)
        {
            row.CopyCellsTo(span);
            var sortedCells = span[..row.CellsCount];
            sortedCells.Sort(default(ByKey<int, Cell>));

            foreach (var (x, cell) in sortedCells)
            {
                var location = new Location(x, y);
                cell.Write(xml, location);
            }
        }
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
                            xml.WriteAttribute("xSplit", freeze.FreezeByX);
                        }

                        if (freeze.FreezeByY > 0)
                        {
                            xml.WriteAttribute("ySplit", freeze.FreezeByY);
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
                foreach (var (x, column) in options.Columns)
                {
                    column.Write(xml, x);
                }
            }
        }

        var sheetData = xml.WriteStartElement("sheetData");
        started = (document, sheetData);
        return started.Value;
    }

    public void Complete(XmlHyperlinks hyperlinks, IReadOnlyCollection<ConditionalFormatting> formattings)
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

        var priority = 1;
        foreach (var formatting in formattings)
        {
            priority = formatting.Write(xml, priority);
        }

        hyperlinks.WriteSheetPart(xml);
        document.Dispose();
    }

    private sealed class ReductionStage(StreamSheetWindow window, int index) : IDisposable
    {
        public void Dispose()
        {
            if (window.reductions.Count != index + 1)
            {
                throw new InvalidOperationException();
            }

            window.reductions.Pop();
        }
    }

    private struct Row
    {
        private readonly Dictionary<int, Cell> cells = new();
        private RowUsage usage;

        public Row()
        {
        }

        public void Place(int x, Cell cell)
        {
            if (!usage.TryMark(new Interval<int>(x, x)))
            {
                throw new InvalidOperationException();
            }

            cells[x] = cell;
        }

        public void Merge(bool isMain, Interval<int> range, Content content, StyleId? styleId)
        {
            if (!usage.TryMark(range))
            {
                throw new InvalidOperationException();
            }

            if (isMain)
            {
                cells[range.LeftInclusive] = new Cell(content, styleId);
            }
        }

        public int CellsCount => cells.Count;

        public void CopyCellsTo(Span<KeyValuePair<int, Cell>> destination)
        {
            var index = 0;
            foreach (var pair in cells)
            {
                destination[index++] = pair;
            }
        }
    }

    private readonly struct Cell(Content content, StyleId? styleId)
    {
        public void Write(Xml xml, Location location)
        {
            using (xml.WriteStartElement(XlsxStructure.Worksheet.Cell))
            {
                xml.WriteAttribute(XlsxStructure.Worksheet.Reference, location);

                if (styleId is { } value)
                {
                    xml.WriteAttribute(XlsxStructure.Worksheet.Style, value);
                }

                content.Write(xml);
            }
        }
    }
}