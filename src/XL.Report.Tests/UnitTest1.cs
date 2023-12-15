using System.IO.Compression;
using System.Text;
using System.Xml;

namespace XL.Report.Tests;

public sealed class Tests
{
    /*
        "_rels/.rels" -> "xl/workbook.xml"
        "[Content_Types].xml" -> "xl/workbook.xml"
        "xl/workbook.xml" -> "xl/worksheets/sheet1.xml"
        "[Content_Types].xml" -> "xl/worksheets/sheet1.xml"
        "xl/_rels/workbook.xml.rels" -> "xl/worksheets/sheet1.xml"
        todo add sharedStrings, styles
        todo unit, newSheet, flush-all
     */

    [Test]
    public void Test1()
    {
        var canvas = new RegularCanvas(Range.Whole);
        canvas.Place(Offset.Zero, new InlineString("Hello world!"), null);

        using var archive = new ZipArchive(
            new FileStream("D:/archives/test.xlsx", FileMode.Create),
            ZipArchiveMode.Create
        );

        Write("xl/worksheets/sheet1.xml", xml => Sheet(xml, canvas));
        Write("xl/workbook.xml", xml => Workbook(xml, "Щит1"));
        Write("[Content_Types].xml", ContentTypes);
        Write("_rels/.rels", Rels);
        Write("xl/_rels/workbook.xml.rels", WorkbookRels);

        void Write(string name, Action<XmlWriter> act)
        {
            var entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
            using var stream = entry.Open();
            var settings = new XmlWriterSettings
            {
                Indent = true,
                Encoding = Encoding.UTF8,
                NewLineChars = "\n",
                Async = true
            };
            using var xml = XmlWriter.Create(stream, settings);
            act(xml);
        }

        return;
        
        var stream = new MemoryStream();
        using (var xml = XmlWriter.Create(
            stream,
            new XmlWriterSettings
            {
                Indent = true,
                Encoding = Encoding.UTF8,
                NewLineChars = "\n"
            }
        ))
        {
            Sheet(xml, canvas);
        }

        var s = Encoding.UTF8.GetString(stream.ToArray());
        Console.WriteLine(s);
    }

    private static void Sheet(XmlWriter xml, RegularCanvas canvas)
    {
        xml.WriteStartDocumentAsync(standalone: true);
        {
            xml.WriteStartElement("worksheet", ns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            // xml.WriteAttributeString("xmlns", "");
            xml.WriteAttributeString("xmlns", "r", "", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            // xml.WriteAttributeString("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            // xml.WriteAttributeString("mc", "Ignorable", "x14ac");
            // xml.WriteAttributeString("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            {
                xml.WriteStartElement("sheetData");
                {
                    foreach (var row in canvas.Rows)
                    {
                        xml.WriteStartElement("row");
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
                    }
                }
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }
        xml.WriteEndDocument();
    }

    public static void Workbook(XmlWriter xml, string sheetName)
    {
        xml.WriteStartDocument(standalone: true);
        {
            xml.WriteStartElement("workbook", ns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            xml.WriteAttributeString("xmlns", "r", "", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            {
                xml.WriteStartElement("sheets");
                {
                    xml.WriteStartElement("sheet");
                    xml.WriteAttributeString("name", sheetName);
                    xml.WriteAttributeString("sheetId", "1");
                    xml.WriteAttributeString("r", "id", null, "rId1");
                    xml.WriteEndElement();
                }
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }
    }

    public static void WorkbookRels(XmlWriter xml)
    {
        xml.WriteStartDocument(standalone: true);
        {
            xml.WriteStartElement("Relationships", ns: "http://schemas.openxmlformats.org/package/2006/relationships");
            {
                xml.WriteStartElement("Relationship");
                xml.WriteAttributeString("Id", "rId1");
                xml.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                xml.WriteAttributeString("Target", "worksheets/sheet1.xml");
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

    public static void ContentTypes(XmlWriter xml)
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

                xml.WriteStartElement("Override");
                xml.WriteAttributeString("PartName", "/xl/worksheets/sheet1.xml");
                xml.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }
    }
}