using XL.Report.Styles.Fills;

namespace XL.Report.Styles;

public sealed partial record Style
{
    public sealed record Diff(Font.Diff? Font = null, Fill? Fill = null, Borders.Diff? Borders = null, Format? Format = null)
    {
        public void Write(Xml xml)
        {
            using (xml.WriteStartElement("dxf"))
            {
                Font?.Write(xml);
                const int fakeIndex = 0;
                Format?.Write(xml, fakeIndex);
                Fill?.Write(xml);
                Borders?.Write(xml);
            }
        }
    }
}