using System.Globalization;
using static XL.Report.XlsxStructure.Styles;

namespace XL.Report.Styles;

public sealed record Font(
    string Family,
    float Size,
    Color? Color,
    FontStyle Style
)
{
    public static Font Calibri11 { get; } = new(
        "Calibri",
        11,
        Color: null,
        FontStyle.Regular
    );

    public void Write(Xml xml)
    {
        using (xml.WriteStartElement(Fonts.Font))
        {
            using (xml.WriteStartElement(Fonts.Size))
            {
                xml.WriteAttributeSpan(Fonts.SizeValue, Size, "N3");
            }

            if (Color is { } color)
            {
                using (xml.WriteStartElement(Fonts.Color))
                {
                    // todo print color allocation free
                    xml.WriteAttribute(Fonts.ColorRgb, color.ToRGBHex());
                }
            }

            using (xml.WriteStartElement(Fonts.Name))
            {
                xml.WriteAttribute(Fonts.NameValue, Family);
            }

            // todo write Style
        }
    }

    public static void Write(Xml xml, IReadOnlyCollection<Font> fonts)
    {
        using (xml.WriteStartElement(Fonts.Collection))
        {
            foreach (var font in fonts)
            {
                font.Write(xml);
            }
        }
    }

    public override string ToString()
    {
        var segments = new[]
        {
            Family,
            Size.ToString(CultureInfo.InvariantCulture),
            Color?.ToString(),
            Style.ToString()
        }.Where(segment => !string.IsNullOrWhiteSpace(segment));
        return string.Join(' ', segments);
    }
}