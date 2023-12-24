using System.IO.Compression;
using System.Text;
using JetBrains.Annotations;
using XL.Report.Styles;
using XL.Report.Styles.Fills;

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
        new FileStream(Path(TestName), FileMode.Create)
    );

    private static string Path(string testName) => $"{resultsDirectory}/{testName}.xlsx";

    [Test]
    public void Simplest_Book()
    {
        using var book = new StreamBook(ResultStream(), CompressionLevel.Optimal, false);
        using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
        {
            IUnit<Location> cell = new Cell(new Number(42));
            sheet.WriteRow(cell);
            sheet.Complete();
        }

        book.Complete();
    }

    [Test]
    public void Formulas()
    {
        using var book = new StreamBook(ResultStream(), CompressionLevel.Optimal, false);
        using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
        {
            var row = new Row(
                new Cell(new Number(42)),
                new Cell(new Number(24)),
                new Cell(new Formula(new Expression.Verbatim("A1+B1")))
            );
            sheet.WriteRow(row);
            sheet.Complete();
        }

        book.Complete();
    }

    [Test]
    public void Conditional_Formattings()
    {
        using var book = new StreamBook(ResultStream(), CompressionLevel.Optimal, false);

        var random = new Random(Seed: 162317036);

        using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
        {
            for (var y = 0; y < 3; y++)
            {
                var row = new Row(
                    new Cell(new Number(random.Next(100))),
                    new Cell(new Number(random.Next(100))),
                    new Cell(new Number(random.Next(100))),
                    new Cell(new Number(random.Next(100))),
                    new Cell(new Number(random.Next(100))),
                    new Cell(new Number(random.Next(100))),
                    new Cell(new Number(random.Next(100))),
                    new Cell(new Number(random.Next(100))),
                    new Cell(new Number(random.Next(100))),
                    new Cell(new Number(random.Next(100)))
                );

                sheet.WriteRow(row);
            }

            var formatting = new ConditionalFormatting(
                Range.Parse("A1:J10"),
                ConditionalFormatting.Rule.Duplicates,
                book.Styles.Register(
                    new Style.Diff(
                        Fill: new SolidFill(new Color(10, 234, 196))
                    )
                )
            );
            sheet.AddConditionalFormatting(formatting);
            sheet.Complete();
        }

        book.Complete();
    }

    [Test]
    public void Font_Styling()
    {
        using var book = new StreamBook(ResultStream(), CompressionLevel.Optimal, false);
        var styles = book.Styles;
        using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
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
    public void Styling()
    {
        using var book = new StreamBook(ResultStream(), CompressionLevel.Optimal, false);
        var styles = book.Styles;
        var bold = styles.Register(Style.Default.Bold());
        var options = new SheetOptions(
            FreezeOptions.None,
            ColumnWidths.Default
                .With(1, width: 24)
                .With(2, width: 24)
                .With(3, width: 24)
                .With(4, width: 24)
                .With(5, width: 24)
                .With(6, width: 24)
                .With(7, width: 24)
                .With(8, width: 24)
                .With(9, width: 24)
                .With(10, width: 24)
                .With(11, width: 24)
                .With(12, width: 24)
                .With(13, width: 24)
                .With(14, width: 24)
                .With(15, width: 24)
                .With(16, width: 24)
                .With(17, width: 24)
                .With(18, width: 24)
        );
        using (var sheet = book.CreateSheet(TestName, options))
        {
            sheet.WriteRow<Range>(new Cell(new InlineString("Fonts:"), bold));
            sheet.WriteRow(
                new Row(
                    new Cell(
                        new InlineString("Default")
                    ),
                    new Cell(
                        new InlineString("Arial"),
                        styles.Register(Style.Default.WithFontFamily("Arial"))
                    ),
                    new Cell(
                        new InlineString("Arial 14"),
                        styles.Register(Style.Default.WithFontFamily("Arial").WithFontSize(14))
                    ),
                    new Cell(
                        new InlineString("Arial red"),
                        styles.Register(Style.Default.WithFontFamily("Arial").WithFontColor(new Color(240, 0, 0)))
                    ),
                    new Cell(
                        new InlineString("Bold"),
                        styles.Register(Style.Default.Bold())
                    ),
                    new Cell(
                        new InlineString("Italic"),
                        styles.Register(Style.Default.Italic())
                    ),
                    new Cell(
                        new InlineString("Strikethrough"),
                        styles.Register(Style.Default.Strikethrough())
                    ),
                    new Cell(
                        new InlineString("Superscript"),
                        styles.Register(Style.Default.Superscript())
                    ),
                    new Cell(
                        new InlineString("Subscript"),
                        styles.Register(Style.Default.Subscript())
                    ),
                    new Cell(
                        new InlineString("Underline single"),
                        styles.Register(Style.Default.Underline())
                    ),
                    new Cell(
                        new InlineString("Underline double"),
                        styles.Register(Style.Default.Underline(Underline.Double))
                    ),
                    new Cell(
                        new InlineString("Underline single by cell"),
                        styles.Register(Style.Default.Underline(Underline.SingleByCell))
                    ),
                    new Cell(
                        new InlineString("Underline double by cell"),
                        styles.Register(Style.Default.Underline(Underline.DoubleByCell))
                    ),
                    new Cell(
                        new InlineString("Composite"),
                        styles.Register(Style.Default.Bold().Italic().Strikethrough().Subscript().Underline().WithFontColor(new Color(240, 0, 0)))
                    )
                )
            );

            // todo new BlankRow(3) is not working for moving
            sheet.WriteRow(new BlankRow(3));

            var matrix = new Matrix<(HorizontalAlignment HorizontalAlignment, Content Content), VerticalAlignment>(
                xHeader: new Merge.All(new InlineString("Alignments.Horizontal:")),
                xAspects: new[]
                {
                    (
                        new Cell(new InlineString("By content, string")) as IUnit<Range>,
                        (HorizontalAlignment.ByContent, new InlineString("value") as Content)
                    ),
                    (
                        new Cell(new InlineString("By content, number")) as IUnit<Range>,
                        (HorizontalAlignment.ByContent, new Number(42) as Content)
                    ),
                    (
                        new Cell(new InlineString("Left, no indent")) as IUnit<Range>,
                        (HorizontalAlignment.Left(), new InlineString("value") as Content)
                    ),
                    (
                        new Cell(new InlineString("Left, indent")) as IUnit<Range>,
                        (HorizontalAlignment.Left(indent: 3), new InlineString("value") as Content)
                    ),
                    (
                        new Cell(new InlineString("Center")) as IUnit<Range>,
                        (HorizontalAlignment.Center, new InlineString("value") as Content)
                    ),
                    (
                        new Cell(new InlineString("Center continuous")) as IUnit<Range>,
                        (HorizontalAlignment.CenterContinuous, new InlineString("value") as Content)
                    ),
                    (
                        new Cell(new InlineString("Right, no indent")) as IUnit<Range>,
                        (HorizontalAlignment.Right(), new InlineString("value") as Content)
                    ),
                    (
                        new Cell(new InlineString("Right, indent")) as IUnit<Range>,
                        (HorizontalAlignment.Right(indent: 3), new InlineString("value") as Content)
                    ),
                    (
                        new Cell(new InlineString("Fill")) as IUnit<Range>,
                        (HorizontalAlignment.Fill, new InlineString("value") as Content)
                    ),
                    (
                        new Cell(new InlineString("Justify")) as IUnit<Range>,
                        (HorizontalAlignment.Justify, new InlineString("The quick brown fox jumps over the lazy dog") as Content)
                    ),
                    (
                        new Cell(new InlineString("Distributed, no indent")) as IUnit<Range>,
                        (HorizontalAlignment.Distributed(), new InlineString("The quick brown fox jumps over the lazy dog") as Content)
                    ),
                    (
                        new Cell(new InlineString("Distributed, indent")) as IUnit<Range>,
                        (HorizontalAlignment.Distributed(indent: 3), new InlineString("The quick brown fox jumps over the lazy dog") as Content)
                    ),
                    (
                        new Cell(new InlineString("Distributed, JustifyLastLine")) as IUnit<Range>,
                        (HorizontalAlignment.DistributedJustifyLastLine, new InlineString("The quick brown fox jumps over the lazy dog") as Content)
                    ),
                },
                yHeader: new Cell(new InlineString("Alignments.Vertical")),
                yAspects: new[]
                {
                    (new Cell(new InlineString("Bottom")) as IUnit<Range>, VerticalAlignment.Bottom),
                    (new Cell(new InlineString("Center")) as IUnit<Range>, VerticalAlignment.Center),
                    (new Cell(new InlineString("Top")) as IUnit<Range>, VerticalAlignment.Top),
                    (new Cell(new InlineString("Distributed")) as IUnit<Range>, VerticalAlignment.Distributed),
                    (new Cell(new InlineString("Justify")) as IUnit<Range>, VerticalAlignment.Justify),
                },
                cellSize: Size.Cell,
                (pair, verticalAlignment) => new Cell(
                    pair.Content,
                    styles.Register(
                        Style.Default.With(pair.HorizontalAlignment).With(verticalAlignment)
                    )
                )
            );

            sheet.WriteRow<Range>(new Cell(new InlineString("Alignments:"), bold));
            sheet.WriteRow(matrix);

            sheet.WriteRow(new BlankRow(3));

            sheet.WriteRow<Range>(new Cell(new InlineString("Fills:"), bold));

            sheet.WriteRow(
                new Row(
                    new Column(
                        new Cell(new InlineString("Solid")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new SolidFill(new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern MediumGray")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.MediumGray, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern DarkGray")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.DarkGray, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern LightGray")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.LightGray, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern DarkHorizontal")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.DarkHorizontal, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern DarkVertical")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.DarkVertical, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern DarkDown")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.DarkDown, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern DarkUp")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.DarkUp, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern DarkGrid")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.DarkGrid, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern DarkTrellis")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.DarkTrellis, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern LightHorizontal")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.LightHorizontal, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern LightVertical")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.LightVertical, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern LightDown")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.LightDown, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern LightUp")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.LightUp, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern LightGrid")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.LightGrid, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern LightTrellis")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.LightTrellis, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern Gray125")),
                        new Cell(new InlineString(""), styles.Register(Style.Default.With(new PatternFill(Pattern.Gray125, new Color(20, 20, 70), background: new Color(10, 234, 196)))))
                    ),
                    new Column(
                        new Cell(new InlineString("Pattern Gray0625")),
                        new Cell(new InlineString(""),
                            styles.Register(Style.Default.With(new PatternFill(Pattern.Gray0625, new Color(20, 20, 70), background: new Color(10, 234, 196))))
                        )
                    )
                )
            );

            sheet.WriteRow(new BlankRow(3));

            sheet.WriteRow<Range>(new Cell(new InlineString("Borders:"), bold));
            sheet.WriteRow(new BlankRow());
            sheet.WriteRow(
                new Row(
                    new BlankColumn(),
                    new Cell(
                        new InlineString("Left Top"),
                        styles.Register(
                            Style.Default.With(
                                new Borders(
                                    left: new Border(BorderStyle.Thick, new Color(20, 150, 20)),
                                    top: new Border(BorderStyle.Thick, new Color(240, 20, 20))
                                )
                            )
                        )
                    ),
                    new BlankColumn(),
                    new Cell(
                        new InlineString("Diagonal"),
                        styles.Register(
                            Style.Default.With(
                                new Borders(
                                    diagonal: new DiagonalBorders(BorderStyle.Thin, new Color(240, 150, 150), down: true, up: false)
                                )
                            )
                        )
                    ),
                    new BlankColumn(),
                    new Cell(
                        new InlineString("Thick"),
                        styles.Register(Style.Default.With(Borders.Perimeter(new Border(BorderStyle.Thick, new Color(240, 20, 20)))))
                    ),
                    new BlankColumn(),
                    new Cell(
                        new InlineString("MediumDashDotDot"),
                        styles.Register(Style.Default.With(Borders.Perimeter(new Border(BorderStyle.MediumDashDotDot, new Color(240, 20, 20)))))
                    ),
                    new BlankColumn(),
                    new Cell(
                        new InlineString("Dashed"),
                        styles.Register(Style.Default.With(Borders.Perimeter(new Border(BorderStyle.Dashed, new Color(240, 20, 20)))))
                    ),
                    new BlankColumn(),
                    new Cell(
                        new InlineString("Hair"),
                        styles.Register(Style.Default.With(Borders.Perimeter(new Border(BorderStyle.Hair, new Color(240, 20, 20)))))
                    ),
                    new BlankColumn(),
                    new Cell(
                        new InlineString("Dotted"),
                        styles.Register(Style.Default.With(Borders.Perimeter(new Border(BorderStyle.Dotted, new Color(240, 20, 20)))))
                    ),
                    new BlankColumn(),
                    new Cell(
                        new InlineString("DashDotDot"),
                        styles.Register(Style.Default.With(Borders.Perimeter(new Border(BorderStyle.DashDotDot, new Color(240, 20, 20)))))
                    ),
                    new BlankColumn(),
                    new Cell(
                        new InlineString("DashDot"),
                        styles.Register(Style.Default.With(Borders.Perimeter(new Border(BorderStyle.DashDot, new Color(240, 20, 20)))))
                    ),
                    new BlankColumn(),
                    new Cell(
                        new InlineString("Thin"),
                        styles.Register(Style.Default.With(Borders.Perimeter(new Border(BorderStyle.Thin, new Color(240, 20, 20)))))
                    ),
                    new BlankColumn(),
                    new Cell(
                        new InlineString("SlantDashDot"),
                        styles.Register(Style.Default.With(Borders.Perimeter(new Border(BorderStyle.SlantDashDot, new Color(240, 20, 20)))))
                    ),
                    new BlankColumn(),
                    new Cell(
                        new InlineString("MediumDashDot"),
                        styles.Register(Style.Default.With(Borders.Perimeter(new Border(BorderStyle.MediumDashDot, new Color(240, 20, 20)))))
                    )
                )
            );

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
            using var sheet = book.CreateSheet($"{TestName} {i + 1}", SheetOptions.Default);
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

        using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
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

        using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
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

    [Test]
    public void Hyperlinks()
    {
        using var book = new StreamBook(ResultStream(), CompressionLevel.Optimal, false);

        using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
        {
            var row = new Row(
                new Cell(new InlineString("example.com")),
                new Cell(new InlineString("mail")),
                new Cell(new InlineString("file")),
                new Cell(new InlineString("range")),
                new Cell(new InlineString("defined name"))
            );
            sheet.WriteRow(row);
            sheet.DefineName("area", Range.Parse("E2:E5"));
            sheet.Hyperlinks.Add(Range.Parse("A1:A1"), "https://example.com", tooltip: "https");
            sheet.Hyperlinks.Add(Range.Parse("B1:B1"), Hyperlink.Mailto("who@example.com", "Hello"), tooltip: "mail");
            sheet.Hyperlinks.Add(Range.Parse("C1:C1"), "new-file.xlsx", tooltip: "new-file");
            sheet.Hyperlinks.AddToRange(Range.Parse("D1:D1"), target: Range.Parse("D2:D5"), tooltip: "range");
            sheet.Hyperlinks.AddToDefinedName(Range.Parse("E1:E1"), "area", tooltip: "defined-name");
            sheet.Complete();
        }

        book.Complete();
    }

    [Test]
    public void Inspect_Archive_Read([Values(nameof(Large))] string testName)
    {
        using var file = new FileStream(Path(testName), FileMode.Open);
        Console.WriteLine("Before open");
        using var zipArchive = new ZipArchive(new ReadTracingStream(file));
        Console.WriteLine("After open");
        var entry = zipArchive.Entries.ElementAt(4);
        Console.WriteLine("After entry select");
        var stream = entry.Open();
        Console.WriteLine("After entry open");
        _ = stream.Read(new byte[1024]);
        Console.WriteLine("After entry read");
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