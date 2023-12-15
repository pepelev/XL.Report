using System.Globalization;
using System.Xml;
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

    public void Write(XmlWriter xml)
    {
        xml.WriteStartElement(Fonts.Font);
        {
            xml.WriteStartElement(Fonts.Size);
            xml.WriteAttributeString(Fonts.SizeValue, Size.ToString("N3", CultureInfo.InvariantCulture));
            xml.WriteEndElement();

            if (Color is { } color)
            {
                xml.WriteStartElement(Fonts.Color);
                xml.WriteAttributeString(Fonts.ColorRgb, color.ToRGBHex());
                xml.WriteEndElement();
            }

            xml.WriteStartElement(Fonts.Name);
            xml.WriteAttributeString(Fonts.NameValue, Family);
            xml.WriteEndElement();

            // todo write Style
        }
        xml.WriteEndElement();
    }

    public static void Write(XmlWriter xml, IReadOnlyCollection<Font> fonts)
    {
        xml.WriteStartElement(Fonts.Collection);
        xml.WriteAttributeString(Fonts.CollectionCount, fonts.Count.ToString(CultureInfo.InvariantCulture));
        {
            foreach (var font in fonts)
            {
                font.Write(xml);
            }
        }
        xml.WriteEndElement();
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