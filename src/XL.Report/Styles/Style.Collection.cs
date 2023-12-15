using System.Globalization;
using System.Xml;
using XL.Report.Styles.Fills;
using static XL.Report.XlsxStructure.Styles;

namespace XL.Report.Styles;

public sealed partial class Style
{
    public sealed class Collection
    {
        private readonly Dictionary<Format, int> formats = new();
        private readonly Dictionary<Font, int> fonts = new();
        private readonly Dictionary<Fill, int> fills = new();
        private readonly Dictionary<Borders, int> borders = new();
        private readonly Dictionary<Style, StyleId> registeredStyles = new();

        public Collection()
        {
            _ = Register(Default);
        }

        public StyleId Register(Style style)
        {
            if (registeredStyles.TryGetValue(style, out var styleId))
            {
                return styleId;
            }

            formats.TryAdd(style.Format, formats.Count);
            fonts.TryAdd(style.Appearance.Font, fonts.Count);
            fills.TryAdd(style.Appearance.Fill, fills.Count);
            borders.TryAdd(style.Appearance.Borders, borders.Count);
            var newStyleId = new StyleId(registeredStyles.Count);
            registeredStyles.Add(style, newStyleId);
            return newStyleId;
        }

        public void Write(XmlWriter xml)
        {
            xml.WriteStartElement(RootElement, XlsxStructure.Namespaces.Main);
            {
                var orderedFonts = fonts
                    .OrderBy(font => font.Value)
                    .Select(pair => pair.Key)
                    .ToList();
                Font.Write(xml, orderedFonts);

                xml.WriteStartElement(CellFormats.StyleFormats);
                xml.WriteAttributeString(CellFormats.StyleFormatsCount, "1");
                {
                    xml.WriteStartElement(CellFormats.CellFormat);
                    xml.WriteAttributeString(CellFormats.NumberFormatIndex, "0");
                    xml.WriteAttributeString(CellFormats.FontIndex, "0");
                    xml.WriteAttributeString(CellFormats.FillIndex, "0");
                    xml.WriteAttributeString(CellFormats.BordersIndex, "0");
                    xml.WriteEndElement();
                }
                xml.WriteEndElement();

                xml.WriteStartElement(CellFormats.Collection);
                var count = registeredStyles.Count.ToString(CultureInfo.InvariantCulture);
                xml.WriteAttributeString(CellFormats.CollectionCount, count);
                {
                    var orderedStyles = registeredStyles
                        .OrderBy(pair => pair.Value)
                        .Select(pair => pair.Key)
                        .ToList();

                    foreach (var style in orderedStyles)
                    {
                        xml.WriteStartElement(CellFormats.CellFormat);
                        xml.WriteAttributeString(CellFormats.StyleFormatIndex, "0");
                        var formatIndex = formats[style.Format];
                        var fontIndex = fonts[style.Appearance.Font];
                        var fillIndex = fills[style.Appearance.Fill];
                        var bordersIndex = borders[style.Appearance.Borders];
                        xml.WriteAttributeInt(CellFormats.NumberFormatIndex, formatIndex);
                        xml.WriteAttributeInt(CellFormats.FontIndex, fontIndex);
                        xml.WriteAttributeInt(CellFormats.FillIndex, fillIndex);
                        xml.WriteAttributeInt(CellFormats.BordersIndex, bordersIndex);

                        if (formatIndex > 0)
                        {
                            xml.WriteAttributeString(CellFormats.ApplyFormat, "1");
                        }

                        if (fontIndex > 0)
                        {
                            xml.WriteAttributeString(CellFormats.FontIndex, "1");
                        }

                        if (fillIndex > 0)
                        {
                            xml.WriteAttributeString(CellFormats.FillIndex, "1");
                        }

                        if (bordersIndex > 0)
                        {
                            xml.WriteAttributeString(CellFormats.BordersIndex, "1");
                        }

                        // todo alignment

                        xml.WriteEndElement();
                    }
                }
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
        }
    }
}