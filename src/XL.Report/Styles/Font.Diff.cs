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
                        xml.WriteAttribute(XlsxStructure.Styles.Fonts.ColorRgb, color.ToRGBHex());
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