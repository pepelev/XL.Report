﻿#region Legal
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

using System.Buffers;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using XL.Report.Auxiliary;
using XL.Report.Contents;
using XL.Report.Styles;

namespace XL.Report;

internal sealed partial class StreamSheetWindow : SheetWindow, IDisposable
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

    public override void Merge(Content content, StyleId? styleId)
    {
        if (!valid)
        {
            throw new InvalidOperationException();
        }

        var range = Range;
        var interval = new Interval<int>(range.Left, range.Right);
        for (var y = range.Top; y <= range.Bottom; y++)
        {
            ref var row = ref CollectionsMarshal.GetValueRefOrAddDefault(rows, y, out var exists);
            if (!exists)
            {
                row = new Row();
            }

            var isMain = y == range.Top;
            row.Merge(isMain, interval, content, styleId);
        }

        mergedRanges.Add(range);
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

    public void Flush(RowOptions rowOptions)
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

            var nextRowToWrite = activeRange.LeftTop.Y;
            var maxUsedRow = activeRange.LeftTop.Y;

            foreach (var (y, row) in sortedRows)
            {
                if (rowOptions != RowOptions.Default)
                {
                    for (; nextRowToWrite < y; nextRowToWrite++)
                    {
                        using (xml.WriteStartElement(XlsxStructure.Worksheet.Row))
                        {
                            xml.WriteAttribute("r", nextRowToWrite);
                            rowOptions.WriteAttributes(xml);
                        }
                    }
                }

                using (xml.WriteStartElement(XlsxStructure.Worksheet.Row))
                {
                    xml.WriteAttribute("r", y);
                    rowOptions.WriteAttributes(xml);

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

                nextRowToWrite = y + 1;
                maxUsedRow = y;
            }

            if (rowOptions != RowOptions.Default)
            {
                for (; nextRowToWrite <= maxTouchedY; nextRowToWrite++)
                {
                    using (xml.WriteStartElement(XlsxStructure.Worksheet.Row))
                    {
                        xml.WriteAttribute("r", nextRowToWrite);
                        rowOptions.WriteAttributes(xml);
                    }
                }
            }

            return Math.Max(maxUsedRow, maxTouchedY);
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

        options.Write(xml);

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
}