using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;
using FluentAssertions.Extensions;
using JetBrains.Annotations;
using XL.Report.Styles;
using XL.Report.Styles.Fills;
using static XL.Report.ConditionalFormatting;
using static XL.Report.Number;

namespace XL.Report.Tests;

public sealed class Examples
{
    private static string TestName => TestContext.CurrentContext.Test.FullName ?? throw new Exception("no test");

    [Test]
    public async Task Simplest_Book()
    {
        using var output = Output.Prepare(TestName);
        using (var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false))
        {
            var units = new Units(book);
            using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
            {
                IUnit<Location> cell = units.Cell(42);
                sheet.WriteRow(cell);
                sheet.Complete();
            }

            book.Complete();
        }

        await output.VerifyAsync();
    }

    [Test]
    public async Task Contents()
    {
        using var output = Output.Prepare(TestName);
        using (var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false))
        {
            var units = new Units(book);
            using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
            {
                var content = new Column(
                    new Row(
                        units.Cell("Int"),
                        units.Cell(42)
                    ),
                    new Row(
                        units.Cell("DateTime"),
                        units.Cell(1.January(2025), Style.Default with { Format = Format.IsoDate })
                    ),
                    new Row(
                        units.Cell("Bool"),
                        units.Cell(true)
                    ),
                    new Row(
                        units.Cell("Decimal"),
                        units.Cell(4.2m)
                    ),
                    new Row(
                        units.Cell("Double"),
                        units.Cell(4.2)
                    ),
                    new Row(
                        units.Cell("Long"),
                        units.Cell(422L)
                    ),
                    new Row(
                        units.Cell("String"),
                        units.Cell("String from units.Cell(string)")
                    ),
                    new Row(
                        units.Cell("String"),
                        units.Cell(Repeat("Long string from units.Cell(string). ", 25))
                    ),
                    new Row(
                        units.Cell("String"),
                        units.Cell(new InlineString("Inline string"))
                    ),
                    new Row(
                        units.Cell("String"),
                        units.Cell(book.Strings.Force(Repeat("Long SharedString.Force.", 25)))
                    ),
                    new Row(
                        units.Cell("String"),
                        units.Cell(
                            new SharedString.ById(
                                book.Strings.ForceRegister("SharedString.ById")
                            )
                        )
                    )
                );
                sheet.WriteRow(content);
                sheet.Complete();
            }

            book.Complete();
        }

        await output.VerifyAsync();
    }

    [Test]
    public async Task Blank_Line_Within_RowOptions()
    {
        using var output = Output.Prepare(TestName);
        using (var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false))
        {
            var units = new Units(book);
            using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
            {
                sheet.WriteRow(
                    new Column(
                        units.Cell(1),
                        new BlankRow(),
                        units.Cell(3)
                    ),
                    new RowOptions(Height: 24)
                );

                sheet.Complete();
            }

            book.Complete();
        }

        await output.VerifyAsync();
    }

    [Test]
    public async Task Data_Types()
    {
        using var output = Output.Prepare(TestName);
        using (var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false))
        {
            var units = new Units(book);
            using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
            {
                var dateStyle = Style.Default.With(Format.IsoDateTime);
                var blueBackground = Style.Default.With(new SolidFill(new Color(45, 45, 180)));
                var row = new Row(
                    units.Cell(true),
                    units.Cell(42),
                    units.Cell("Hello"),
                    units.Cell(new DateTime(2012, 07, 15, 17, 30, 42), dateStyle),
                    units.Cell(new Expression.Verbatim("TODAY()"), dateStyle),
                    units.BlankCell(blueBackground)
                );
                sheet.WriteRow(row);
                sheet.Complete();
            }

            book.Complete();
        }

        await output.VerifyAsync();
    }

    [Test]
    public async Task Formulas()
    {
        using var output = Output.Prepare(TestName);
        using (var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false))
        {
            var units = new Units(book);
            using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
            {
                var row = new Row(
                    units.Cell(42),
                    units.Cell(24),
                    units.Cell(new Expression.Verbatim("A1+B1"))
                );
                sheet.WriteRow(row);
                sheet.Complete();
            }

            book.Complete();
        }

        await output.VerifyAsync();
    }

    [Pure]
    private static string CutSheetName(string name) => name.Length > 31
            ? name[^31..]
            : name;

    private static (string Name, Condition Condition) Named(
        Condition condition,
        [CallerArgumentExpression(nameof(condition))]
        string expression = ""
    )
    {
        return (expression, condition);
    }
    
    [Test]
    public async Task Conditional_Formattings()
    {
        using var output = Output.Prepare(TestName);
        using (var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false))
        {
            var units = new Units(book);

            var a = new Expression.Verbatim("10");
            var b = new Expression.Verbatim("25");
            var rules = new[]
            {
                Named(Condition.Duplicates),
                Named(Condition.Unique),
                Named(Condition.Contains("3")),
                Named(Condition.NotContains("3")),
                Named(Condition.StartsWith("1")),
                Named(Condition.EndsWith("1")),
                Named(Condition.IsError),
                Named(Condition.IsNotError),
                Named(Condition.Between(a, b)),
                Named(Condition.NotBetween(a, b)),
                Named(Condition.Equal(b)),
                Named(Condition.NotEqual(b)),
                Named(Condition.GreaterThan(b)),
                Named(Condition.GreaterThanOrEqual(b)),
                Named(Condition.LessThan(b)),
                Named(Condition.LessThanOrEqual(b)),
                Named(new Condition.ExtremeValues(Condition.Target.Highest, 3)),
                Named(new Condition.ExtremePercentValues(Condition.Target.Lowest, 35)),
                Named(new Condition.AverageRelatedValues(Condition.AverageRelation.AboveAverage)),
                Named(new Condition.AverageRelatedValues(Condition.AverageRelation.BelowAverage)),
                Named(new Condition.AverageRelatedValues(Condition.AverageRelation.BelowOrEqualAverage)),
                Named(new Condition.AverageRelatedValues(Condition.AverageRelation.AboveOrEqualAverage))
            };

            var brightTurquoise = new Color(10, 234, 196);
            var mainStyleId = book.Styles.Register(
                new Style.Diff(
                    Fill: new SolidFill(brightTurquoise)
                )
            );
            var cerulean = new Color(10, 171, 234);
            var alternativeStyleId = book.Styles.Register(
                new Style.Diff(
                    Fill: new SolidFill(cerulean)
                )
            );
            foreach (var (name, condition) in rules)
            {
                using var sheet = book.CreateSheet(CutSheetName(name), SheetOptions.Default);
                FillData(name, sheet);
                var formatting = new ConditionalFormatting(
                    new[] { Range.Parse("A2:J5"), Range.Parse("A6:F6") },
                    new[] { new ValueRule(condition, mainStyleId) }
                );
                sheet.AddConditionalFormatting(formatting);
                sheet.Complete();
            }

            {
                var name = "Condition.Duplicates -> brightTurquoise; Condition.Between(10, 25) -> cerulean";
                using var sheet = book.CreateSheet(CutSheetName(name), SheetOptions.Default);
                FillData(name, sheet);
                var formatting = new ConditionalFormatting(
                    new[] { Range.Parse("A2:J5"), Range.Parse("A6:F6") },
                    new[]
                    {
                        new ValueRule(Condition.Duplicates, mainStyleId),
                        new ValueRule(Condition.Between(a, b), alternativeStyleId)
                    }
                );
                sheet.AddConditionalFormatting(formatting);
                sheet.Complete();
            }

            book.Complete();

            void FillData(string name, Book.SheetBuilder sheet)
            {
                var random = new Random(Seed: 162317036);
                var title = units.Cell(name);
                sheet.WriteRow<Range>(title);
                for (var y = 0; y < 3; y++)
                {
                    var row = new Row(
                        units.Cell(random.Next(100)),
                        units.Cell(random.Next(100)),
                        units.Cell(random.Next(100)),
                        units.Cell(random.Next(100)),
                        units.Cell(random.Next(100)),
                        units.Cell(random.Next(100)),
                        units.Cell(random.Next(100)),
                        units.Cell(random.Next(100)),
                        units.Cell(random.Next(100)),
                        units.Cell(random.Next(100))
                    );

                    sheet.WriteRow(row);
                }

                sheet.WriteRow<Range>(units.Cell(new Formula(new Expression.Verbatim("10/0"))));
            }
        }

        await output.VerifyAsync();
    }

    [Test]
    public async Task Font_Styling()
    {
        using var output = Output.Prepare(TestName);
        using (var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false))
        {
            var units = new Units(book);
            using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
            {
                var times = Style.Default.WithFontFamily("Times New Roman");
                var row = new Row(
                    units.Cell("Without style"),
                    units.Cell("Default", Style.Default),
                    units.Cell("Times New Roman", times),
                    units.Cell("Big times", times.WithFontSize(20)),
                    units.Cell("Orange", Style.Default.WithFontColor(new Color(255, 100, 15))),
                    units.Cell("Green", Style.Default.WithFontColor(new Color(15, 200, 15)))
                );

                sheet.WriteRow(row);
                sheet.Complete();
            }

            book.Complete();
        }

        await output.VerifyAsync();
    }

    [Test]
    public async Task Styling()
    {
        using var output = Output.Prepare(TestName);
        using (var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false))
        {
            var units = new Units(book);
            var styles = book.Styles;
            var bold = styles.Register(Style.Default.Bold());
            var options = new SheetOptions(
                FreezeOptions.None,
                ColumnOptions.Collection.Default
                    .With(1, new ColumnOptions(Width: 24))
                    .With(2, new ColumnOptions(Width: 24))
                    .With(3, new ColumnOptions(Width: 24))
                    .With(4, new ColumnOptions(Width: 24))
                    .With(5, new ColumnOptions(Width: 24))
                    .With(6, new ColumnOptions(Width: 24))
                    .With(7, new ColumnOptions(Width: 24))
                    .With(8, new ColumnOptions(Width: 24))
                    .With(9, new ColumnOptions(Width: 24))
                    .With(10, new ColumnOptions(Width: 24))
                    .With(11, new ColumnOptions(Width: 24))
                    .With(12, new ColumnOptions(Width: 24))
                    .With(13, new ColumnOptions(Width: 24))
                    .With(14, new ColumnOptions(Width: 24))
                    .With(15, new ColumnOptions(Width: 24))
                    .With(16, new ColumnOptions(Width: 24))
                    .With(17, new ColumnOptions(Width: 24))
                    .With(18, new ColumnOptions(Width: 24))
                    .With(19, new ColumnOptions(Hidden: true))
            );
            using (var sheet = book.CreateSheet(TestName, options))
            {
                sheet.WriteRow<Range>(units.Cell("Fonts:", bold));
                sheet.WriteRow(
                    new Row(
                        units.Cell("Default"),
                        units.Cell("Arial", Style.Default.WithFontFamily("Arial")),
                        units.Cell("Arial 14", Style.Default.WithFontFamily("Arial").WithFontSize(14)),
                        units.Cell("Arial red",
                            Style.Default.WithFontFamily("Arial").WithFontColor(new Color(240, 0, 0))),
                        units.Cell("Bold", Style.Default.Bold()),
                        units.Cell("Italic", Style.Default.Italic()),
                        units.Cell("Strikethrough", Style.Default.Strikethrough()),
                        units.Cell("Superscript", Style.Default.Superscript()),
                        units.Cell("Subscript", Style.Default.Subscript()),
                        units.Cell("Underline single", Style.Default.Underline()),
                        units.Cell("Underline double", Style.Default.Underline(Underline.Double)),
                        units.Cell("Underline single by cell", Style.Default.Underline(Underline.SingleByCell)),
                        units.Cell("Underline double by cell", Style.Default.Underline(Underline.DoubleByCell)),
                        units.Cell(
                            "Composite",
                            Style.Default
                                .Bold()
                                .Italic()
                                .Strikethrough()
                                .Subscript()
                                .Underline()
                                .WithFontColor(new Color(240, 0, 0))
                        )
                    )
                );

                sheet.WriteRow(new BlankRow(3));

                var matrix = new Matrix<(HorizontalAlignment HorizontalAlignment, Content Content), VerticalAlignment>(
                    xHeader: new Merge.All(new InlineString("Alignments.Horizontal:")),
                    xAspects:
                    [
                        (
                            units.Cell("By content, string"),
                            (HorizontalAlignment.ByContent, new InlineString("value"))
                        ),
                        (
                            units.Cell("By content, number"),
                            (HorizontalAlignment.ByContent, new Integral(42))
                        ),
                        (
                            units.Cell("Left, no indent"),
                            (HorizontalAlignment.Left(), new InlineString("value"))
                        ),
                        (
                            units.Cell("Left, indent"),
                            (HorizontalAlignment.Left(indent: 3), new InlineString("value"))
                        ),
                        (
                            units.Cell("Center"),
                            (HorizontalAlignment.Center, new InlineString("value"))
                        ),
                        (
                            units.Cell("Center continuous"),
                            (HorizontalAlignment.CenterContinuous, new InlineString("value"))
                        ),
                        (
                            units.Cell("Right, no indent"),
                            (HorizontalAlignment.Right(), new InlineString("value"))
                        ),
                        (
                            units.Cell("Right, indent"),
                            (HorizontalAlignment.Right(indent: 3), new InlineString("value"))
                        ),
                        (
                            units.Cell("Fill"),
                            (HorizontalAlignment.Fill, new InlineString("value"))
                        ),
                        (
                            units.Cell("Justify"),
                            (HorizontalAlignment.Justify,
                                new InlineString("The quick brown fox jumps over the lazy dog"))
                        ),
                        (
                            units.Cell("Distributed, no indent"),
                            (HorizontalAlignment.Distributed(),
                                new InlineString("The quick brown fox jumps over the lazy dog"))
                        ),
                        (
                            units.Cell("Distributed, indent"),
                            (HorizontalAlignment.Distributed(indent: 3),
                                new InlineString("The quick brown fox jumps over the lazy dog"))
                        ),
                        (
                            units.Cell("Distributed, JustifyLastLine"),
                            (HorizontalAlignment.DistributedJustifyLastLine,
                                new InlineString("The quick brown fox jumps over the lazy dog"))
                        )
                    ],
                    yHeader: units.Cell("Alignments.Vertical"),
                    yAspects:
                    [
                        (units.Cell("Bottom"), VerticalAlignment.Bottom),
                        (units.Cell("Center"), VerticalAlignment.Center),
                        (units.Cell("Top"), VerticalAlignment.Top),
                        (units.Cell("Distributed"), VerticalAlignment.Distributed),
                        (units.Cell("Justify"), VerticalAlignment.Justify)
                    ],
                    cellSize: Size.Cell,
                    (pair, verticalAlignment) => units.Cell(
                        pair.Content,
                        Style.Default.With(pair.HorizontalAlignment).With(verticalAlignment)
                    )
                );

                sheet.WriteRow<Range>(units.Cell("Alignments:", bold));
                sheet.WriteRow(matrix);

                sheet.WriteRow(new BlankRow(3));

                sheet.WriteRow<Range>(units.Cell("Fills:", bold));

                sheet.WriteRow(
                    new Row(
                        new Column(
                            units.Cell("Solid"),
                            units.Cell("", Style.Default.With(new SolidFill(new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern MediumGray"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.MediumGray, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern DarkGray"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.DarkGray, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern LightGray"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.LightGray, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern DarkHorizontal"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.DarkHorizontal, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern DarkVertical"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.DarkVertical, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern DarkDown"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.DarkDown, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern DarkUp"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.DarkUp, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern DarkGrid"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.DarkGrid, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern DarkTrellis"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.DarkTrellis, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern LightHorizontal"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.LightHorizontal, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern LightVertical"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.LightVertical, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern LightDown"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.LightDown, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern LightUp"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.LightUp, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern LightGrid"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.LightGrid, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern LightTrellis"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.LightTrellis, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern Gray125"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.Gray125, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        ),
                        new Column(
                            units.Cell("Pattern Gray0625"),
                            units.Cell("",
                                Style.Default.With(new PatternFill(Pattern.Gray0625, new Color(20, 20, 70),
                                    background: new Color(10, 234, 196))))
                        )
                    )
                );

                sheet.WriteRow(new BlankRow(3));

                sheet.WriteRow<Range>(units.Cell("Borders:", bold));
                sheet.WriteRow(new BlankRow());
                sheet.WriteRow(
                    new Row(
                        new BlankColumn(),
                        units.Cell(
                            "Left Top",
                            Style.Default.With(
                                new Borders(
                                    left: new Border(BorderStyle.Thick, new Color(20, 150, 20)),
                                    top: new Border(BorderStyle.Thick, new Color(240, 20, 20))
                                )
                            )
                        ),
                        new BlankColumn(),
                        units.Cell(
                            "Diagonal",
                            Style.Default.With(
                                new Borders(
                                    diagonal: new DiagonalBorders(BorderStyle.Thin, new Color(240, 150, 150),
                                        down: true, up: false)
                                )
                            )
                        ),
                        new BlankColumn(),
                        units.Cell(
                            "Thick",
                            Style.Default.With(Borders.Perimeter(new Border(BorderStyle.Thick, new Color(240, 20, 20))))
                        ),
                        new BlankColumn(),
                        units.Cell(
                            "MediumDashDotDot",
                            Style.Default.With(Borders.Perimeter(new Border(BorderStyle.MediumDashDotDot,
                                new Color(240, 20, 20))))
                        ),
                        new BlankColumn(),
                        units.Cell(
                            "Dashed",
                            Style.Default.With(
                                Borders.Perimeter(new Border(BorderStyle.Dashed, new Color(240, 20, 20))))
                        ),
                        new BlankColumn(),
                        units.Cell(
                            "Hair",
                            Style.Default.With(Borders.Perimeter(new Border(BorderStyle.Hair, new Color(240, 20, 20))))
                        ),
                        new BlankColumn(),
                        units.Cell(
                            "Dotted",
                            Style.Default.With(
                                Borders.Perimeter(new Border(BorderStyle.Dotted, new Color(240, 20, 20))))
                        ),
                        new BlankColumn(),
                        units.Cell(
                            "DashDotDot",
                            Style.Default.With(Borders.Perimeter(new Border(BorderStyle.DashDotDot,
                                new Color(240, 20, 20))))
                        ),
                        new BlankColumn(),
                        units.Cell(
                            "DashDot",
                            Style.Default.With(
                                Borders.Perimeter(new Border(BorderStyle.DashDot, new Color(240, 20, 20))))
                        ),
                        new BlankColumn(),
                        units.Cell(
                            "Thin",
                            Style.Default.With(Borders.Perimeter(new Border(BorderStyle.Thin, new Color(240, 20, 20))))
                        ),
                        new BlankColumn(),
                        units.Cell(
                            "SlantDashDot",
                            Style.Default.With(Borders.Perimeter(new Border(BorderStyle.SlantDashDot,
                                new Color(240, 20, 20))))
                        ),
                        new BlankColumn(),
                        units.Cell(
                            "MediumDashDot",
                            Style.Default.With(Borders.Perimeter(new Border(BorderStyle.MediumDashDot,
                                new Color(240, 20, 20))))
                        )
                    )
                );

                sheet.WriteRow(new BlankRow(), new RowOptions(Height: 24));
                sheet.WriteRow(new BlankRow(), new RowOptions(Hidden: true));

                sheet.Complete();
            }

            book.Complete();
        }

        await output.VerifyAsync();
    }

    [Test]
    public async Task Multiple_Sheets()
    {
        using var output = Output.Prepare(TestName);
        using (var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false))
        {
            var units = new Units(book);
            for (var i = 0; i < 10; i++)
            {
                using var sheet = book.CreateSheet($"{TestName} {i + 1}", SheetOptions.Default);
                var row = new Row(
                    units.Cell("Sheet index:"),
                    units.Cell(i)
                );
                sheet.WriteRow(row);
                sheet.Complete();
            }

            book.Complete();
        }

        await output.VerifyAsync();
    }

    [Test]
    public async Task Merged_Cell()
    {
        using var output = Output.Prepare(TestName);
        using (var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false))
        {
            var units = new Units(book);
            using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
            {
                var row = units.Merge("Merged cell", new Size(5, 2));
                sheet.WriteRow(row);
                sheet.Complete();
            }

            book.Complete();
        }

        await output.VerifyAsync();
    }

    [Test]
    [Explicit]
    public void Large()
    {
        var random = new Random(592739);
        var styles = Array(64, RandomStyle);
        var contents = Array(1024, RandomWord);

        using var output = Output.Prepare(TestName);
        using var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false);
        var units = new Units(book);
        using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
        {
            for (var i = 0; i < 100_000; i++)
            {
                var rowLength = random.Next(10, 40);
                var cells = Array(
                    rowLength,
                    () => units.Cell(
                        random.Pick(contents),
                        random.Pick(styles)
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
    public async Task Hyperlinks()
    {
        using var output = Output.Prepare(TestName);
        using (var book = new StreamBook(output.Stream, CompressionLevel.Optimal, false))
        {
            var units = new Units(book);
            using (var sheet = book.CreateSheet(TestName, SheetOptions.Default))
            {
                var row = new Row(
                    units.Cell("example.com"),
                    units.Cell("mail"),
                    units.Cell("file"),
                    units.Cell("range"),
                    units.Cell("defined name")
                );
                sheet.WriteRow(row);
                sheet.DefineName("area", Range.Parse("E2:E5").EnsureValid());
                sheet.Hyperlinks.Add(Range.Parse("A1:A1").EnsureValid(), "https://example.com", tooltip: "https");
                sheet.Hyperlinks.Add(Range.Parse("B1:B1").EnsureValid(), Hyperlink.Mailto("who@example.com", "Hello"), tooltip: "mail");
                sheet.Hyperlinks.Add(Range.Parse("C1:C1").EnsureValid(), "new-file.xlsx", tooltip: "new-file");
                sheet.Hyperlinks.AddToRange(Range.Parse("D1:D1").EnsureValid(), target: Range.Parse("D2:D5").EnsureValid(), tooltip: "range");
                sheet.Hyperlinks.AddToDefinedName(Range.Parse("E1:E1").EnsureValid(), "area", tooltip: "defined-name");
                sheet.Complete();
            }

            book.Complete();
        }

        await output.VerifyAsync();
    }

    [Test]
    [Explicit]
    public void Inspect_Archive_Read([Values("")] string testName)
    {
        using var file = new FileStream(testName, FileMode.Open);
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

    private static string Repeat(string @string, int times)
    {
        var builder = new StringBuilder();
        for (var i = 0; i < times; i++)
        {
            builder.Append(@string);
        }

        return builder.ToString();
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