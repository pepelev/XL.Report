#region Legal
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

using System.IO.Compression;
using XL.Report.Contents;
using XL.Report.Styles;
using XL.Report.Styles.Fills;

namespace XL.Report.Tests;

[Explicit]
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
        - row/column freezing
        - column widths
        - column styles
        - row height
        - row styles
        - merged cells
        - styles
        - contents
        - conditional-formatting
            - value forms
            - timePeriod forms
            - several rules
            - several ranges
        - handy way to create cells and merges
            - registration of styles inside
            - selection of content type by parameter type
        - tests
            - verify infrastructure
        - make ConditionalFormatting.Rule abstract class

        partially added:
        - formulas
        - conditional-formatting

        todo add:
        - todos
        - Rule.(Contains, NotContains, StartsWith, EndsWith, IsError, NotIsError) is not working in excel
        - sheet (range) name consistency
        - tests
        - docs
        - formulas.row and column references
        - formulas.functions and combinations
        - formulas.range-reference
        - conditional-formatting.rest forms
        - many styles within one cell
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
                new Cell(new Number.Integral(509), book.Styles.Register(bigRed))
            );

            sheet.WriteRow<Range>(row);
            sheet.WriteRow<Range>(row);
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
                new Cell(new Number.Integral(511), book.Styles.Register(bigRed))
            );

            sheet.WriteRow<Range>(row);
            sheet.WriteRow<Range>(row);
            sheet.Complete();
        }

        book.Complete();
    }

    [Test]
    public void METHOD()
    {
        var days = 14918;
        var dateTime = new DateTime(1940, 11, 3).AddDays(-days);
        Console.WriteLine(dateTime.ToString("O"));

        var s = 8.33333333333333333333.ToString("0.000000000000");
        Console.WriteLine(s);
    }
}