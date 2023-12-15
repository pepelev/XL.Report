using System.Diagnostics.Contracts;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using XL.Report.Styles;

namespace XL.Report;

public abstract class Book : IAsyncDisposable
{
    public abstract Style.Collection Styles { get; }
    public abstract SharedStrings Strings { get; }

    public abstract ValueTask DisposeAsync();

    // todo check name constraints
    public abstract SheetBuilder OpenSheet(string name);

    public abstract Task CompleteAsync();

    public abstract class SheetBuilder : IAsyncDisposable
    {
        public abstract string Name { get; set; }
        public abstract ValueTask DisposeAsync();
        public abstract Task<T> WriteRowAsync<T>(IUnit<T> unit);
        public abstract Task CompleteAsync();
    }
}

public sealed class StreamBook : Book
{
    private readonly Stream output;
    private readonly CompressionLevel compressionLevel;
    private readonly bool continueOnCapturedContext;
    private readonly ZipArchive archive;

    public StreamBook(Stream output, CompressionLevel compressionLevel, bool leaveOpen, bool continueOnCapturedContext = false)
    {
        this.output = output;
        this.compressionLevel = compressionLevel;
        this.continueOnCapturedContext = continueOnCapturedContext;
        archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen);
    }

    private readonly List<StreamSheetBuilder> sheets = new();

    public override Style.Collection Styles { get; } = new();
    public override SharedStrings Strings => throw new NotImplementedException();

    // todo make sync
    public override async ValueTask DisposeAsync()
    {
        archive.Dispose();
    }

    public override SheetBuilder OpenSheet(string name)
    {
        EnsureThereIsNoCurrentSheet();

        var path = new SheetPath(sheets.Count);
        var entry = archive.CreateEntry(path.AsString(), compressionLevel);
        var builder = new StreamSheetBuilder(path, name, entry, continueOnCapturedContext);
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

    public override async Task CompleteAsync()
    {
        EnsureThereIsNoCurrentSheet();

        await WriteAsync("xl/workbook.xml", Workbook).ConfigureAwait(continueOnCapturedContext);
        await WriteAsync("xl/styles.xml", Styles.Write).ConfigureAwait(continueOnCapturedContext);
        await WriteAsync("[Content_Types].xml", ContentTypes).ConfigureAwait(continueOnCapturedContext);
        await WriteAsync("_rels/.rels", Rels).ConfigureAwait(continueOnCapturedContext);
        await WriteAsync("xl/_rels/workbook.xml.rels", WorkbookRels).ConfigureAwait(continueOnCapturedContext);
    }

    private async Task WriteAsync(string name, Action<XmlWriter> act)
    {
        var entry = archive.CreateEntry(name, compressionLevel);
        await using var stream = entry.Open();
        var settings = new XmlWriterSettings
        {
            Indent = true,
            Encoding = Encoding.UTF8,
            NewLineChars = "\n",
            Async = true
        };
        await using var xml = XmlWriter.Create(stream, settings);
        act(xml);
    }

    private void Workbook(XmlWriter xml)
    {
        xml.WriteStartDocument(standalone: true);
        {
            xml.WriteStartElement("workbook", ns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            xml.WriteAttributeString("xmlns", "r", "", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
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

    public void ContentTypes(XmlWriter xml)
    {
        xml.WriteStartDocument(standalone: true);
        {
            xml.WriteStartElement("Types", ns: "http://schemas.openxmlformats.org/package/2006/content-types");
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
            }
            xml.WriteEndElement();
        }
    }

    public static void Rels(XmlWriter xml)
    {
        xml.WriteStartDocument(standalone: true);
        {
            xml.WriteStartElement("Relationships", ns: "http://schemas.openxmlformats.org/package/2006/relationships");
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
        xml.WriteStartDocument(standalone: true);
        {
            xml.WriteStartElement("Relationships", ns: "http://schemas.openxmlformats.org/package/2006/relationships");
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
        public SheetPath Path { get; }
        private readonly Stream entryStream;
        private volatile bool open = true;
        private readonly bool continueOnCapturedContext;
        private readonly StreamSheetWindow window;

        public StreamSheetBuilder(SheetPath path, string name, ZipArchiveEntry entry, bool continueOnCapturedContext)
        {
            Path = path;
            Name = name;
            this.continueOnCapturedContext = continueOnCapturedContext;
            entryStream = entry.Open();
            window = new StreamSheetWindow(entryStream);
        }

        public bool IsOpen => open;

        // todo add checks
        public override string Name { get; set; }

        public override async ValueTask DisposeAsync()
        {
            // todo decide use sync dispose or async for xml
            open = false;
            window.Dispose();
        }

        public override async Task<T> WriteRowAsync<T>(IUnit<T> unit)
        {
            var result = unit.Write(window);
            await window.FlushAsync().ConfigureAwait(continueOnCapturedContext);
            return result;
        }

        public override async Task CompleteAsync()
        {
            await window.CompleteAsync().ConfigureAwait(false);
        }
    }
}