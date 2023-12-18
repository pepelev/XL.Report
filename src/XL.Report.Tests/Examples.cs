﻿using System.IO.Compression;
using System.Text;
using JetBrains.Annotations;
using XL.Report.Styles;

namespace XL.Report.Tests;

public sealed class Examples
{
    private const string resultsDirectory = "results";
    private static string TestName => TestContext.CurrentContext.Test.MethodName ?? throw new Exception("no test");

    [OneTimeSetUp]
    public void OneTimeSetUp()
    {
        Directory.CreateDirectory(resultsDirectory);
    }

    private Stream ResultStream() => new WriteOnlyStream(
        new FileStream($"{resultsDirectory}/{TestName}.xlsx", FileMode.Create)
    );

    [Test]
    public void Simplest_Book()
    {
        using var book = new StreamBook(ResultStream(), CompressionLevel.Optimal, false);
        using (var sheet = book.OpenSheet(TestName, SheetOptions.Default))
        {
            IUnit<Location> cell = new Cell(new Number(42));
            sheet.WriteRow(cell);
            sheet.Complete();
        }

        book.Complete();
    }

    [Test]
    public void Font_Styling()
    {
        using var book = new StreamBook(ResultStream(), CompressionLevel.Optimal, false);
        var styles = book.Styles;
        using (var sheet = book.OpenSheet(TestName, SheetOptions.Default))
        {
            var times = Style.Default.WithFontFamily("Times New Roman");
            var row = new Row(
                new Cell(new InlineString("Without style")),
                new Cell(new InlineString("Default"), styles.Register(Style.Default)),
                new Cell(new InlineString("Times New Roman"), styles.Register(times)),
                new Cell(new InlineString("Big times"), styles.Register(times.WithFontSize(20))),
                new Cell(
                    new InlineString("Orange"),
                    styles.Register(
                        Style.Default.WithFontColor(new Color(255, 100, 15))
                    )
                ),
                new Cell(
                    new InlineString("Green"),
                    styles.Register(
                        Style.Default.WithFontColor(new Color(15, 200, 15))
                    )
                )
            );

            sheet.WriteRow(row);
            sheet.Complete();
        }

        book.Complete();
    }

    [Test]
    public void Multiple_Sheets()
    {
        using var book = new StreamBook(ResultStream(), CompressionLevel.Optimal, false);

        for (var i = 0; i < 10; i++)
        {
            using var sheet = book.OpenSheet($"{TestName} {i + 1}", SheetOptions.Default);
            var row = new Row(
                new Cell(new InlineString("Sheet index:")),
                new Cell(new Number(i))
            );
            sheet.WriteRow(row);
            sheet.Complete();
        }

        book.Complete();
    }

    [Test]
    public void Merged_Cell()
    {
        using var book = new StreamBook(ResultStream(), CompressionLevel.Optimal, false);

        using (var sheet = book.OpenSheet(TestName, SheetOptions.Default))
        {
            var row = new Merge(
                new InlineString("Merged cell"),
                new Size(5, 2)
            );
            sheet.WriteRow(row);
            sheet.Complete();
        }

        book.Complete();
    }

    [Test]
    public void Large()
    {
        var random = new Random(592739);
        var styles = Array(64, RandomStyle);
        var contents = Array(1024, RandomWord);

        using var book = new StreamBook(ResultStream(), CompressionLevel.Optimal, false);

        using (var sheet = book.OpenSheet(TestName, SheetOptions.Default))
        {
            for (var i = 0; i < 100_000; i++)
            {
                var rowLength = random.Next(10, 40);
                var cells = Array(
                    rowLength,
                    () => new Cell(
                        book.Strings.String(random.Pick(contents)),
                        book.Styles.Register(random.Pick(styles))
                    ) as IUnit<Range>
                );
                var row = new Row(cells);
                sheet.WriteRow(row);
            }

            sheet.Complete();
        }

        book.Complete();

        Style RandomStyle()
        {
            var color = new Color((byte)random.Next(256), (byte)random.Next(256), (byte)random.Next(256));
            return Style.Default.WithFontColor(color);
        }

        string RandomWord()
        {
            var length = random.Next(3, 64);
            var builder = new StringBuilder(length);
            for (var i = 0; i < length; i++)
            {
                builder.Append((char)random.Next('а', 'я' + 1));
            }

            return builder.ToString();
        }
    }

    private static T[] Array<T>(int count, [InstantHandle] Func<T> factory)
    {
        var result = new T[count];
        for (var i = 0; i < count; i++)
        {
            result[i] = factory();
        }

        return result;
    }
}