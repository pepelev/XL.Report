using System.Diagnostics.Contracts;
using XL.Report.Styles.Fills;

namespace XL.Report.Styles;

public sealed partial record Style(Appearance Appearance, Format Format)
{
    public static Style Default { get; } = new(
        Appearance.Default,
        Format.General
    );

    [Pure]
    public Style With(Alignment alignment) => this with
    {
        Appearance = Appearance with
        {
            Alignment = alignment
        }
    };

    [Pure]
    public Style With(HorizontalAlignment alignment)
    {
        var currentAlignment = Appearance.Alignment;
        return this with
        {
            Appearance = Appearance with
            {
                Alignment = new Alignment(
                    horizontal: alignment,
                    currentAlignment.Vertical,
                    currentAlignment.OverflowBehavior,
                    currentAlignment.ReadingOrder,
                    currentAlignment.TextRotation
                )
            }
        };
    }

    [Pure]
    public Style With(VerticalAlignment alignment)
    {
        var currentAlignment = Appearance.Alignment;
        return this with
        {
            Appearance = Appearance with
            {
                Alignment = new Alignment(
                    currentAlignment.Horizontal,
                    vertical: alignment,
                    currentAlignment.OverflowBehavior,
                    currentAlignment.ReadingOrder,
                    currentAlignment.TextRotation
                )
            }
        };
    }

    [Pure]
    public Style With(OverflowBehavior overflowBehavior)
    {
        var currentAlignment = Appearance.Alignment;
        return this with
        {
            Appearance = Appearance with
            {
                Alignment = new Alignment(
                    currentAlignment.Horizontal,
                    currentAlignment.Vertical,
                    overflowBehavior,
                    currentAlignment.ReadingOrder,
                    currentAlignment.TextRotation
                )
            }
        };
    }

    [Pure]
    public Style With(ReadingOrder readingOrder)
    {
        var currentAlignment = Appearance.Alignment;
        return this with
        {
            Appearance = Appearance with
            {
                Alignment = new Alignment(
                    currentAlignment.Horizontal,
                    currentAlignment.Vertical,
                    currentAlignment.OverflowBehavior,
                    readingOrder,
                    currentAlignment.TextRotation
                )
            }
        };
    }

    [Pure]
    public Style With(TextRotation textRotation)
    {
        var currentAlignment = Appearance.Alignment;
        return this with
        {
            Appearance = Appearance with
            {
                Alignment = new Alignment(
                    currentAlignment.Horizontal,
                    currentAlignment.Vertical,
                    currentAlignment.OverflowBehavior,
                    currentAlignment.ReadingOrder,
                    textRotation
                )
            }
        };
    }

    [Pure]
    public Style With(Font font) => this with
    {
        Appearance = Appearance with
        {
            Font = font
        }
    };

    [Pure]
    public Style WithFontFamily(string fontFamily) => With(Appearance.Font with { Family = fontFamily });

    [Pure]
    public Style WithFontSize(float fontSize) => With(Appearance.Font with { Size = fontSize });

    [Pure]
    public Style WithFontColor(Color color) => With(Appearance.Font with { Color = color });

    [Pure]
    public Style WithFontStyle(FontStyle style) => With(Appearance.Font with { Style = style });

    [Pure]
    public Style Bold() => 
        WithFontStyle(Appearance.Font.Style with { IsBold = true });

    [Pure]
    public Style Italic() => 
        WithFontStyle(Appearance.Font.Style with { IsItalic = true });

    [Pure]
    public Style Strikethrough() => 
        WithFontStyle(Appearance.Font.Style with { IsStrikethrough = true });

    [Pure]
    public Style Underline(Underline underline = Styles.Underline.Single) => 
        WithFontStyle(Appearance.Font.Style with { Underline = underline });

    [Pure]
    public Style Superscript() => 
        WithFontStyle(Appearance.Font.Style with { Alignment = FontVerticalAlignment.Superscript});

    [Pure]
    public Style Subscript() => 
        WithFontStyle(Appearance.Font.Style with { Alignment = FontVerticalAlignment.Subscript });

    [Pure]
    public Style With(Fill fill) => this with
    {
        Appearance = Appearance with
        {
            Fill = fill
        }
    };

    [Pure]
    public Style With(Borders borders) => this with
    {
        Appearance = Appearance with
        {
            Borders = borders
        }
    };

    [Pure]
    public Style With(Format format) => this with { Format = format };
}