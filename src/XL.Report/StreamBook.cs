using System.Diagnostics.Contracts;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using XL.Report.Styles;

namespace XL.Report;

public sealed class Hyperlinks
{
    public sealed record Hyperlink<T>(Range Range, T Target, string? Tooltip);

    private readonly Book.SheetBuilder sheet;
    private readonly List<Hyperlink<string>> urlHyperlinks = new();
    private readonly List<Hyperlink<string>> definedNameHyperlinks = new();
    private readonly List<Hyperlink<Range>> thisSheetRangeHyperlinks = new();
    private readonly List<Hyperlink<SheetRelated<Range>>> rangeHyperlinks = new();

    public void WriteSheetPart(Xml xml)
    {
        // if (not empty)
        using (xml.WriteStartElement("hyperlinks"))
        {
            foreach (var (range, target, tooltip) in thisSheetRangeHyperlinks)
            {
                using (xml.WriteStartElement("hyperlink"))
                {
                    xml.WriteAttribute("ref", range);
                    xml.WriteAttribute("location", new SheetRelated<Range>(sheet.Name, target).ToFormattable());
                    if (tooltip != null)
                    {
                        xml.WriteAttribute("tooltip", tooltip);
                    }
                }
            }

            
        }
    }

    public bool RequireRelsPart => throw new NotImplementedException();

    public void WriteRelsPart(Xml xml)
    {
        asdad
    }
}

public sealed class StreamBook : Book
{
    private sealed record DefinedName(string Id, string? Comment, StreamSheetBuilder Sheet, Range Range);

    private readonly ZipArchive archive;
    private readonly CompressionLevel compressionLevel;
    private readonly List<StreamSheetBuilder> sheets = new();
    private readonly Dictionary<string, DefinedName> definedNames = new();

    public StreamBook(Stream output, CompressionLevel compressionLevel, bool leaveOpen)
        : this(
            output,
            new BoundedSharedStrings(
                64 * 1024,
                32,
                64 * 1024 * 16
            ),
            compressionLevel,
            leaveOpen)
    {
    }

    public StreamBook(
        Stream output,
        SharedStrings sharedStrings,
        CompressionLevel compressionLevel,
        bool leaveOpen)
    {
        Strings = sharedStrings;
        this.compressionLevel = compressionLevel;
        archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen);
    }

    public override Style.Collection Styles { get; } = new();
    public override SharedStrings Strings { get; }

    public override void Dispose()
    {
        archive.Dispose();
    }

    public override SheetBuilder CreateSheet(string name, SheetOptions options)
    {
        EnsureThereIsNoCurrentSheet();

        var path = new SheetPath(sheets.Count);
        var entry = archive.CreateEntry(path.AsString(), compressionLevel);
        var builder = new StreamSheetBuilder(this, path, name, entry, options);
        sheets.Add(builder);
        return builder;
    }

    private void DefineName(StreamSheetBuilder sheet, string name, Range range, string? comment = null)
    {
        // todo check name(id) with rules
        var definedName = new DefinedName(name, comment, sheet, range);
        definedNames.Add(name, definedName);
    }

    private void EnsureThereIsNoCurrentSheet()
    {
        if (sheets.Count > 0 && sheets.Last().IsOpen)
        {
            throw new InvalidOperationException();
        }
    }

    public override void Complete()
    {
        EnsureThereIsNoCurrentSheet();

        Write("xl/workbook.xml", Workbook);
        Write("xl/styles.xml", Styles.Write);
        Write("xl/sharedStrings.xml", SharedStrings);
        Write("[Content_Types].xml", ContentTypes);
        Write("_rels/.rels", Rels);
        Write("xl/_rels/workbook.xml.rels", WorkbookRels);
    }

    private void Write(string name, Action<Xml> act)
    {
        var entry = archive.CreateEntry(name, compressionLevel);
        using var stream = entry.Open();
        var settings = new XmlWriterSettings
        {
            Encoding = Encoding.UTF8
        };
        using var xml = new Xml(XmlWriter.Create(stream, settings));
        act(xml);
    }

    private void Workbook(Xml xml)
    {
        using (xml.WriteStartDocument("workbook", XlsxStructure.Namespaces.Spreadsheet.Main))
        {
            xml.WriteAttribute("xmlns", "r", XlsxStructure.Namespaces.OfficeDocuments.Relationships);
            using (xml.WriteStartElement("sheets"))
            {
                // todo check zero sheets
                for (var i = 0; i < sheets.Count; i++)
                {
                    var sheet = sheets[i];
                    var id = (i + 1).ToString(CultureInfo.InvariantCulture);
                    using (xml.WriteStartElement("sheet"))
                    {
                        xml.WriteAttribute("name", sheet.Name);
                        xml.WriteAttribute("sheetId", id);
                        xml.WriteAttribute("r", "id", $"rId{id}");
                    }
                }
            }
        }
    }

    private void SharedStrings(Xml xml)
    {
        using (xml.WriteStartDocument("sst", XlsxStructure.Namespaces.Spreadsheet.Main))
        {
            var expectedId = new SharedStringId(0);
            foreach (var (@string, id) in Strings.OrderBy(pair => pair.Id))
            {
                if (id != expectedId)
                {
                    throw new InvalidOperationException();
                }

                using (xml.WriteStartElement("si"))
                {
                    using (xml.WriteStartElement("t"))
                    {
                        xml.WriteValue(@string);
                    }
                }

                expectedId = new SharedStringId(id.Index + 1);
            }
        }
    }

    public void ContentTypes(Xml xml)
    {
        using (xml.WriteStartDocument("Types", "http://schemas.openxmlformats.org/package/2006/content-types"))
        {
            using (xml.WriteStartElement("Default"))
            {
                xml.WriteAttribute("Extension", "rels");
                xml.WriteAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
            }

            using (xml.WriteStartElement("Default"))
            {
                xml.WriteAttribute("Extension", "xml");
                xml.WriteAttribute("ContentType", "application/xml");
            }

            using (xml.WriteStartElement("Override"))
            {
                xml.WriteAttribute("PartName", "/xl/workbook.xml");
                xml.WriteAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            }

            foreach (var builder in sheets)
            {
                using (xml.WriteStartElement("Override"))
                {
                    xml.WriteAttribute("PartName", builder.Path.SlashLeadingString());
                    xml.WriteAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                }
            }

            using (xml.WriteStartElement("Override"))
            {
                xml.WriteAttribute("PartName", "/xl/styles.xml");
                xml.WriteAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            }

            using (xml.WriteStartElement("Override"))
            {
                xml.WriteAttribute("PartName", "/xl/sharedStrings.xml");
                xml.WriteAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
            }
        }
    }

    public static void Rels(Xml xml)
    {
        using (xml.WriteStartDocument("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships"))
        {
            using (xml.WriteStartElement("Relationship"))
            {
                xml.WriteAttribute("Id", "rId1");
                xml.WriteAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
                xml.WriteAttribute("Target", "xl/workbook.xml");
            }
        }
    }

    public void WorkbookRels(Xml xml)
    {
        using (xml.WriteStartDocument("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships"))
        {
            var id = 1;

            foreach (var sheet in sheets)
            {
                using (xml.WriteStartElement("Relationship"))
                {
                    xml.WriteAttribute("Id", $"rId{id.ToString(CultureInfo.InvariantCulture)}");
                    xml.WriteAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                    xml.WriteAttribute("Target", sheet.Path.AsString()[3..]); // todo crutch
                }

                id++;
            }

            using (xml.WriteStartElement("Relationship"))
            {
                xml.WriteAttribute("Id", $"rId{id.ToString(CultureInfo.InvariantCulture)}");
                xml.WriteAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
                xml.WriteAttribute("Target", "styles.xml");
            }

            id++;

            using (xml.WriteStartElement("Relationship"))
            {
                xml.WriteAttribute("Id", $"rId{id.ToString(CultureInfo.InvariantCulture)}");
                xml.WriteAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
                xml.WriteAttribute("Target", "sharedStrings.xml");
            }
        }
    }

    private readonly struct SheetPath
    {
        private readonly int index;

        public SheetPath(int index)
        {
            this.index = index;
        }

        [Pure]
        public string AsString()
        {
            var number = (index + 1).ToString(CultureInfo.InvariantCulture);
            return $"xl/worksheets/sheet{number}.xml";
        }

        [Pure]
        public string SlashLeadingString() => $"/{AsString()}";
    }

    private sealed class StreamSheetBuilder : SheetBuilder
    {
        
        private readonly StreamBook book;
        private readonly Stream entryStream;
        private readonly StreamSheetWindow window;
        private readonly List<Hyperlink<string>> urlHyperlinks = new();
        private readonly List<Hyperlink<string>> definedNameHyperlinks = new();
        private readonly List<Hyperlink<Range>> thisSheetRangeHyperlinks = new();
        private readonly List<Hyperlink<SheetRelated<Range>>> rangeHyperlinks = new();
        private volatile bool open = true;

        public StreamSheetBuilder(
            StreamBook book,
            SheetPath path,
            string name,
            ZipArchiveEntry entry,
            SheetOptions options)
        {
            this.book = book;
            Path = path;
            Name = name;
            entryStream = entry.Open();
            window = new StreamSheetWindow(entryStream, options);
        }

        public SheetPath Path { get; }

        public bool IsOpen => open;

        // todo add checks
        public override string Name { get; set; }

        public override void Dispose()
        {
            open = false;
            window.Dispose();
        }

        public override T WriteRow<T>(IUnit<T> unit)
        {
            var result = unit.Write(window);
            window.Flush();
            return result;
        }

        public override void DefineName(string name, Range range, string? comment = null)
        {
            book.DefineName(this, name, range, comment);
        }

        // todo check range not overlaps
        public override void AddHyperlink(Range range, string url, string? tooltip = null)
        {
            var hyperlink = new Hyperlink<string>(range, url, tooltip);
            urlHyperlinks.Add(hyperlink);
        }

        public override void AddHyperlinkToDefinedName(Range range, string name, string? tooltip = null)
        {
            var hyperlink = new Hyperlink<string>(range, name, tooltip);
            definedNameHyperlinks.Add(hyperlink);
        }

        public override void AddHyperlinkToRange(Range range, Range target, string? tooltip = null)
        {
            var hyperlink = new Hyperlink<Range>(range, target, tooltip);
            thisSheetRangeHyperlinks.Add(hyperlink);
        }

        public override void AddHyperlinkToRange(Range range, SheetRelated<Range> target, string? tooltip = null)
        {
            var hyperlink = new Hyperlink<SheetRelated<Range>>(range, target, tooltip);
            rangeHyperlinks.Add(hyperlink);
        }

        public override void Complete()
        {
            window.Complete();
        }
    }
}