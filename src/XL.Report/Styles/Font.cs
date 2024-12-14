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

using System.Globalization;
using static XL.Report.XlsxStructure.Styles;

namespace XL.Report.Styles;

public sealed partial record Font(
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
            if (Style.IsBold)
            {
                xml.WriteEmptyElement("b");
            }

            if (Style.IsItalic)
            {
                xml.WriteEmptyElement("i");
            }

            if (Style.IsStrikethrough)
            {
                xml.WriteEmptyElement("strike");
            }

            if (Style.Underline is { } underline)
            {
                using (xml.WriteStartElement("u"))
                {
                    var value = underline switch
                    {
                        Underline.Single => "single",
                        Underline.SingleByCell => "singleAccounting",
                        Underline.Double => "double",
                        Underline.DoubleByCell => "doubleAccounting",
                        _ => "single"
                    };
                    xml.WriteAttribute("val", value);
                }
            }

            if (Style.Alignment != FontVerticalAlignment.Regular)
            {
                using (xml.WriteStartElement("vertAlign"))
                {
                    var value = Style.Alignment switch
                    {
                        FontVerticalAlignment.Subscript => "subscript",
                        FontVerticalAlignment.Superscript => "superscript",
                        _ => "superscript"
                    };
                    xml.WriteAttribute("val", value);
                }
            }

            using (xml.WriteStartElement(Fonts.Size))
            {
                xml.WriteAttribute(Fonts.SizeValue, Size, "N3");
            }

            if (Color is { } color)
            {
                using (xml.WriteStartElement(Fonts.Color))
                {
                    xml.WriteAttribute(Fonts.ColorRgb, color.ToRgbHex());
                }
            }

            using (xml.WriteStartElement(Fonts.Name))
            {
                xml.WriteAttribute(Fonts.NameValue, Family);
            }
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