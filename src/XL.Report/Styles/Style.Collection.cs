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

using XL.Report.Styles.Fills;
using static XL.Report.XlsxStructure.Styles;

namespace XL.Report.Styles;

public sealed partial record Style
{
    public sealed class Collection
    {
        private readonly Dictionary<Format, int> formats = new();
        private readonly Dictionary<Alignment, int> alignments = new();
        private readonly Dictionary<Font, int> fonts = new();
        private readonly Dictionary<Fill, int> fills = new();
        private readonly Dictionary<Borders, int> borders = new();
        private readonly Dictionary<Style, StyleId> registeredStyles = new();
        private readonly Dictionary<Diff, StyleDiffId> registeredDiffs = new();

        public Collection()
        {
            fills.TryAdd(Fill.No, 0);
            fills.TryAdd(new PatternFill(Pattern.Gray125, color: null), 1);
            _ = Register(Default);
        }

        public StyleId Register(Style style)
        {
            if (registeredStyles.TryGetValue(style, out var styleId))
            {
                return styleId;
            }

            formats.TryAdd(style.Format, formats.Count + 256);
            alignments.TryAdd(style.Appearance.Alignment, alignments.Count);
            fonts.TryAdd(style.Appearance.Font, fonts.Count);
            fills.TryAdd(style.Appearance.Fill, fills.Count);
            borders.TryAdd(style.Appearance.Borders, borders.Count);
            var newStyleId = new StyleId(registeredStyles.Count);
            registeredStyles.Add(style, newStyleId);
            return newStyleId;
        }

        public StyleDiffId Register(Diff diff)
        {
            if (registeredDiffs.TryGetValue(diff, out var styleDiffId))
            {
                return styleDiffId;
            }

            var newStyleDiffId = new StyleDiffId(registeredDiffs.Count);
            registeredDiffs.Add(diff, newStyleDiffId);
            return newStyleDiffId;
        }

        public void Write(Xml xml)
        {
            using (xml.WriteStartDocument(RootElement, XlsxStructure.Namespaces.Spreadsheet.Main))
            {
                Format.Write(xml, formats.Select(pair => (pair.Key, pair.Value)));

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
                            var alignment = style.Appearance.Alignment;
                            var alignmentIndex = alignments[alignment];
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

                            if (alignmentIndex > 0 && !alignment.IsDefault)
                            {
                                xml.WriteAttribute(CellFormats.ApplyAlignment, "1");
                                alignment.Write(xml);
                            }
                        }
                    }
                }

                if (registeredDiffs.Count > 0)
                {
                    using (xml.WriteStartElement("dxfs"))
                    {
                        var orderedDiffs = registeredDiffs
                            .OrderBy(pair => pair.Value)
                            .Select(pair => pair.Key)
                            .ToList();

                        foreach (var diff in orderedDiffs)
                        {
                            diff.Write(xml);
                        }
                    }
                }
            }
        }
    }
}