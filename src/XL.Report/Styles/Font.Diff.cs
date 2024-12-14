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

namespace XL.Report.Styles;

public sealed partial record Font
{
    public sealed record Diff(
        string? Family = null,
        float? Size = null,
        Color? Color = null,
        FontStyle.Diff? Style = null)
    {
        public void Write(Xml xml)
        {
            using (xml.WriteStartElement(XlsxStructure.Styles.Fonts.Font))
            {
                if (Style is { } style)
                {
                    if (style.IsBold)
                    {
                        xml.WriteEmptyElement("b");
                    }

                    if (style.IsItalic)
                    {
                        xml.WriteEmptyElement("i");
                    }

                    if (style.IsStrikethrough)
                    {
                        xml.WriteEmptyElement("strike");
                    }

                    if (style.Underline is { } underline)
                    {
                        using (xml.WriteStartElement("u"))
                        {
                            var value = underline switch
                            {
                                UnderlineDiff.Single => "single",
                                UnderlineDiff.Double => "double",
                                _ => "single"
                            };
                            xml.WriteAttribute("val", value);
                        }
                    }

                    if (style.Alignment is { } alignment)
                    {
                        using (xml.WriteStartElement("vertAlign"))
                        {
                            var value = alignment switch
                            {
                                FontVerticalAlignmentDiff.Subscript => "subscript",
                                FontVerticalAlignmentDiff.Superscript => "superscript",
                                _ => "superscript"
                            };
                            xml.WriteAttribute("val", value);
                        }
                    }
                }

                if (Size is { } size)
                {
                    using (xml.WriteStartElement(XlsxStructure.Styles.Fonts.Size))
                    {
                        xml.WriteAttribute(XlsxStructure.Styles.Fonts.SizeValue, size, "N3");
                    }
                }

                if (Color is { } color)
                {
                    using (xml.WriteStartElement(XlsxStructure.Styles.Fonts.Color))
                    {
                        xml.WriteAttribute(XlsxStructure.Styles.Fonts.ColorRgb, color.ToRgbHex());
                    }
                }

                if (Family is { } family)
                {
                    using (xml.WriteStartElement(XlsxStructure.Styles.Fonts.Name))
                    {
                        xml.WriteAttribute(XlsxStructure.Styles.Fonts.NameValue, family);
                    }
                }
            }
        }
    }
}