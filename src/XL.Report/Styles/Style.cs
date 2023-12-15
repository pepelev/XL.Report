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
    public Style WithFontColor(Color fontColor) => With(Appearance.Font with { Color = fontColor });

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
}