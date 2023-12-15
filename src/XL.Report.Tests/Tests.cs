using System.IO.Compression;
using XL.Report.Styles;
using XL.Report.Styles.Fills;

namespace XL.Report.Tests;

public sealed class Tests
{
    /*
        "_rels/.rels" -> "xl/workbook.xml"
        "_rels/.rels" -> "xl/styles.xml"
        "[Content_Types].xml" -> "xl/workbook.xml"
        "xl/workbook.xml" -> "xl/worksheets/sheet1.xml"
        "[Content_Types].xml" -> "xl/worksheets/sheet1.xml"
        "[Content_Types].xml" -> "xl/styles.xml"
        "xl/styles.xml" -> "xl/worksheets/sheet1.xml"
        "xl/sharedStrings.xml" -> "xl/worksheets/sheet1.xml"
        "xl/_rels/workbook.xml.rels" -> "xl/worksheets/sheet1.xml"

        todo add
        - styles,
        - Book class,
        - flush rows,
        - imperative-programing(new sheet, print),
        - conditional-formatting,
        - hyper-links,
        - sharedStrings,
        - закрепление областей,
        - column width and row height
        - column and row styles,
        - many styles within one cell
        - merged cells
     */

    [Test]
    public void Book_Proto()
    {
        using var book = new StreamBook(
            new WriteOnlyStream(
                new FileStream("D:/archives/test-book.xlsx", FileMode.Create)
            ),
            CompressionLevel.Optimal,
            leaveOpen: false
        );

        using (var sheet = book.OpenSheet("Prototype", SheetOptions.Default))
        {
            var bigRed = new Style(
                new Appearance(
                    Alignment.Default,
                    new Font("Times New Roman", 40, new Color(255, 100, 15), FontStyle.Regular),
                    Fill.No,
                    Borders.None
                ),
                Format.General
            );
            var row = new Row(
                new Cell(book.Strings.String("Hello world!")),
                new Cell(new Number(509), book.Styles.Register(bigRed))
            );

            sheet.WriteRow(row);
            sheet.WriteRow(row);
            sheet.Complete();
        }

        using (var sheet = book.OpenSheet("Prototype 2", SheetOptions.Default))
        {
            var bigRed = new Style(
                new Appearance(
                    Alignment.Default,
                    new Font("Times New Roman", 40, new Color(255, 100, 15), FontStyle.Regular),
                    Fill.No,
                    Borders.None
                ),
                Format.General
            );
            var row = new Row(
                new Cell(book.Strings.String("Hello world!")),
                new Cell(new Number(511), book.Styles.Register(bigRed))
            );

            sheet.WriteRow(row);
            sheet.WriteRow(row);
            sheet.Complete();
        }

        book.Complete();
    }
}