using System.Diagnostics.Contracts;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using XL.Report.Styles;

namespace XL.Report;

public sealed class StreamBook : Book
{
    private readonly ZipArchive archive;
    private readonly CompressionLevel compressionLevel;
    private readonly List<StreamSheetBuilder> sheets = new();

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

    public override SheetBuilder OpenSheet(string name, SheetOptions options)
    {
        EnsureThereIsNoCurrentSheet();

        var path = new SheetPath(sheets.Count);
        var entry = archive.CreateEntry(path.AsString(), compressionLevel);
        var builder = new StreamSheetBuilder(path, name, entry, options);
        sheets.Add(builder);
        return builder;
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

    private void Write(string name, Action<XmlWriter> act)
    {
        var entry = archive.CreateEntry(name, compressionLevel);
        using var stream = entry.Open();
        var settings = new XmlWriterSettings
        {
            Encoding = Encoding.UTF8,
        };
        using var xml = XmlWriter.Create(stream, settings);
        act(xml);
    }

    private void Workbook(XmlWriter xml)
    {
        xml.WriteStartDocument(true);
        {
            xml.WriteStartElement("workbook", XlsxStructure.Namespaces.Spreadsheet.Main);
            xml.WriteAttributeString("xmlns", "r", "", XlsxStructure.Namespaces.OfficeDocuments.Relationships);
            {
                xml.WriteStartElement("sheets");
                {
                    for (var i = 0; i < sheets.Count; i++)
                    {
                        var sheet = sheets[i];
                        var id = (i + 1).ToString(CultureInfo.InvariantCulture);
                        xml.WriteStartElement("sheet");
                        xml.WriteAttributeString("name", sheet.Name);
                        xml.WriteAttributeString("sheetId", id);
                        xml.WriteAttributeString("r", "id", null, $"rId{id}");
                        xml.WriteEndElement();
                    }
                }
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }
    }

    private void SharedStrings(XmlWriter xml)
    {
        xml.WriteStartDocument(true);
        {
            xml.WriteStartElement("sst", XlsxStructure.Namespaces.Spreadsheet.Main);
            {
                var expectedId = new SharedStringId(0);
                foreach (var (@string, id) in Strings.OrderBy(pair => pair.Id))
                {
                    if (id != expectedId)
                    {
                        throw new InvalidOperationException();
                    }

                    xml.WriteStartElement("si");
                    {
                        xml.WriteElementString("t", @string);
                    }
                    xml.WriteEndElement();

                    expectedId = new SharedStringId(id.Index + 1);
                }
            }
            xml.WriteEndElement();
        }
    }

    public void ContentTypes(XmlWriter xml)
    {
        xml.WriteStartDocument(true);
        {
            xml.WriteStartElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");
            {
                xml.WriteStartElement("Default");
                xml.WriteAttributeString("Extension", "rels");
                xml.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
                xml.WriteEndElement();

                xml.WriteStartElement("Default");
                xml.WriteAttributeString("Extension", "xml");
                xml.WriteAttributeString("ContentType", "application/xml");
                xml.WriteEndElement();

                xml.WriteStartElement("Override");
                xml.WriteAttributeString("PartName", "/xl/workbook.xml");
                xml.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
                xml.WriteEndElement();

                foreach (var builder in sheets)
                {
                    xml.WriteStartElement("Override");
                    xml.WriteAttributeString("PartName", builder.Path.SlashLeadingString());
                    xml.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                    xml.WriteEndElement();
                }

                xml.WriteStartElement("Override");
                xml.WriteAttributeString("PartName", "/xl/styles.xml");
                xml.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
                xml.WriteEndElement();

                xml.WriteStartElement("Override");
                xml.WriteAttributeString("PartName", "/xl/sharedStrings.xml");
                xml.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }
    }

    public static void Rels(XmlWriter xml)
    {
        xml.WriteStartDocument(true);
        {
            xml.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
            {
                xml.WriteStartElement("Relationship");
                xml.WriteAttributeString("Id", "rId1");
                xml.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
                xml.WriteAttributeString("Target", "xl/workbook.xml");
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }
    }

    public void WorkbookRels(XmlWriter xml)
    {
        xml.WriteStartDocument(true);
        {
            xml.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
            {
                var id = 1;

                foreach (var sheet in sheets)
                {
                    xml.WriteStartElement("Relationship");
                    xml.WriteAttributeString("Id", $"rId{id.ToString(CultureInfo.InvariantCulture)}");
                    xml.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                    xml.WriteAttributeString("Target", sheet.Path.AsString()[3..]); // todo crutch
                    xml.WriteEndElement();

                    id++;
                }

                xml.WriteStartElement("Relationship");
                xml.WriteAttributeString("Id", $"rId{id.ToString(CultureInfo.InvariantCulture)}");
                xml.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
                xml.WriteAttributeString("Target", "styles.xml");
                xml.WriteEndElement();

                id++;

                xml.WriteStartElement("Relationship");
                xml.WriteAttributeString("Id", $"rId{id.ToString(CultureInfo.InvariantCulture)}");
                xml.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
                xml.WriteAttributeString("Target", "sharedStrings.xml");
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
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
        private readonly Stream entryStream;
        private readonly StreamSheetWindow window;
        private volatile bool open = true;

        public StreamSheetBuilder(
            SheetPath path,
            string name,
            ZipArchiveEntry entry,
            SheetOptions options)
        {
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

        public override void Complete()
        {
            window.Complete();
        }
    }
}