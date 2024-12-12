using XL.Report.Contents;
using XL.Report.Styles;

namespace XL.Report;

internal sealed partial class StreamSheetWindow
{
    private readonly struct Cell(Content content, StyleId? styleId)
    {
        public void Write(Xml xml, Location location)
        {
            using (xml.WriteStartElement(XlsxStructure.Worksheet.Cell))
            {
                xml.WriteAttribute(XlsxStructure.Worksheet.Reference, location);

                if (styleId is { } value)
                {
                    xml.WriteAttribute(XlsxStructure.Worksheet.Style, value);
                }

                content.Write(xml);
            }
        }
    }
}