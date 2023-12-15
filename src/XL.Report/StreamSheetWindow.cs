﻿using System.Text;
using System.Xml;
using XL.Report.Styles;

namespace XL.Report;

public sealed class StreamSheetWindow : SheetWindow, IDisposable
{
    private readonly Dictionary<Location, Cell> placed = new();
    private Range activeRange;
    private readonly Stack<Range> reductions = new();
    private bool started = false;
    private readonly XmlWriter xml;

    public StreamSheetWindow(Stream stream)
    {
        activeRange = Range.EntireSheet;
        var settings = new XmlWriterSettings
        {
            Indent = true,
            Encoding = Encoding.UTF8,
            NewLineChars = "\n",
            Async = true,
            CloseOutput = true
        };
        xml = XmlWriter.Create(stream, settings);
    }

    public override Range Range => reductions.Count > 0
        ? reductions.Peek()
        : activeRange;

    public IEnumerable<Row> Rows => placed
        .GroupBy(pair => pair.Key.Y)
        .OrderBy(row => row.Key)
        .Select(
            row =>
            {
                var y = row.Key;
                var contents = row
                    .Select(pair => (pair.Key.X, pair.Value))
                    .OrderBy(pair => pair.Item1);
                return new Row(y, contents);
            }
        );

    public override void Place(Content content, StyleId? styleId)
    {
        var range = Range;
        if (range.IsEmpty)
        {
            throw new InvalidOperationException();
        }

        placed[range.LeftTopCell] = new Cell(content, styleId);
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
        var @new = new Range(current.LeftTopCell + offset, newSize ?? current.Size - offset);
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

    public async Task FlushAsync()
    {
        int Write()
        {
            var mostDownY = activeRange.LeftTopCell.Y;
            foreach (var row in Rows)
            {
                xml.WriteStartElement(XlsxStructure.Worksheet.Row);
                xml.WriteStartAttribute("r");
                xml.WriteValue(row.Y);
                {
                    foreach (var (x, content) in row.Contents)
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
        var newActiveRange = activeRange.ReduceDown(mostDownY + 1 - activeRange.LeftTopCell.Y);
        // todo continueOnCapturedContext
        await xml.FlushAsync().ConfigureAwait(false);

        activeRange = newActiveRange;
        placed.Clear();
    }

    private void WriteStartOnlyFirstTime()
    {
        if (started)
        {
            return;
        }

        xml.WriteStartDocument(standalone: true);
        xml.WriteStartElement("worksheet", ns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        xml.WriteAttributeString("xmlns", "r", "", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        xml.WriteStartElement("sheetData");

        started = true;
    }

    public async Task CompleteAsync()
    {
        WriteStartOnlyFirstTime();
        // todo continueOnCapturedContext
        await xml.WriteEndDocumentAsync().ConfigureAwait(false);
        await xml.FlushAsync().ConfigureAwait(false);
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

    public void Dispose()
    {
        xml.Dispose();
    }
}