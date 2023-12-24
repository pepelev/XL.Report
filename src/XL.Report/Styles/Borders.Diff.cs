namespace XL.Report.Styles;

public sealed partial class Borders
{
    public sealed record Diff(
        Border? Left = null,
        Border? Right = null,
        Border? Top = null,
        Border? Bottom = null
    )
    {
        public static Diff Perimeter(Border border) => new(border, border, border, border);

        public void Write(Xml xml)
        {
            using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Border))
            {
                using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Left))
                {
                    WriteElement(Left, xml);
                }

                using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Right))
                {
                    WriteElement(Right, xml);
                }

                using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Top))
                {
                    WriteElement(Top, xml);
                }

                using (xml.WriteStartElement(XlsxStructure.Styles.Borders.Bottom))
                {
                    WriteElement(Bottom, xml);
                }
            }
        }
    }
}