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

        added:
        - Book class
        - flush rows
        - imperative-programing(new sheet print)
        - defined-names
        - hyper-links
        - sharedStrings
        - закрепление областей
        - column width
        - merged cells
        - styles

        partially added:
        - formulas
        - conditional-formatting

        todo add:
        - BlankContent : Content and others
        - formulas.row and column references
        - formulas.functions and combinations
        - formulas.range-reference
        - conditional-formatting.rest forms
        - row height
        - column styles
        - row styles
        - many styles within one cell
        - netstandard2.1 support
        - reorder sheets
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

        using (var sheet = book.CreateSheet("Prototype", SheetOptions.Default))
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

        using (var sheet = book.CreateSheet("Prototype 2", SheetOptions.Default))
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