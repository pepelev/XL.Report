using XL.Report.Styles.Fills;
using static XL.Report.XlsxStructure.Styles;

namespace XL.Report.Styles;

public sealed partial record Style
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

        public void Write(Xml xml)
        {
            using (xml.WriteStartDocument(RootElement, XlsxStructure.Namespaces.Spreadsheet.Main))
            {
                var orderedFonts = fonts
                    .OrderBy(font => font.Value)
                    .Select(pair => pair.Key)
                    .ToList();
                Font.Write(xml, orderedFonts);

                var orderedFills = fills
                    .OrderBy(font => font.Value)
                    .Select(pair => pair.Key);
                Fill.Write(xml, orderedFills);

                var orderedBorders = borders
                    .OrderBy(font => font.Value)
                    .Select(pair => pair.Key);
                Borders.Write(xml, orderedBorders);

                using (xml.WriteStartElement(CellFormats.StyleFormats))
                {
                    xml.WriteAttribute(CellFormats.StyleFormatsCount, "1");
                    using (xml.WriteStartElement(CellFormats.CellFormat))
                    {
                        xml.WriteAttribute(CellFormats.NumberFormatIndex, "0");
                        xml.WriteAttribute(CellFormats.FontIndex, "0");
                        xml.WriteAttribute(CellFormats.FillIndex, "0");
                        xml.WriteAttribute(CellFormats.BordersIndex, "0");
                    }
                }

                using (xml.WriteStartElement(CellFormats.Collection))
                {
                    var orderedStyles = registeredStyles
                        .OrderBy(pair => pair.Value)
                        .Select(pair => pair.Key)
                        .ToList();

                    foreach (var style in orderedStyles)
                    {
                        using (xml.WriteStartElement(CellFormats.CellFormat))
                        {
                            xml.WriteAttribute(CellFormats.StyleFormatIndex, "0");
                            var formatIndex = formats[style.Format];
                            var fontIndex = fonts[style.Appearance.Font];
                            var fillIndex = fills[style.Appearance.Fill];
                            var bordersIndex = borders[style.Appearance.Borders];
                            xml.WriteAttribute(CellFormats.NumberFormatIndex, formatIndex);
                            xml.WriteAttribute(CellFormats.FontIndex, fontIndex);
                            xml.WriteAttribute(CellFormats.FillIndex, fillIndex);
                            xml.WriteAttribute(CellFormats.BordersIndex, bordersIndex);

                            if (formatIndex > 0)
                            {
                                xml.WriteAttribute(CellFormats.ApplyFormat, "1");
                            }

                            if (fontIndex > 0)
                            {
                                xml.WriteAttribute(CellFormats.ApplyFont, "1");
                            }

                            if (fillIndex > 0)
                            {
                                xml.WriteAttribute(CellFormats.ApplyFill, "1");
                            }

                            if (bordersIndex > 0)
                            {
                                xml.WriteAttribute(CellFormats.ApplyBorders, "1");
                            }

                            // todo alignment
                        }
                    }
                }
            }
        }
    }
}