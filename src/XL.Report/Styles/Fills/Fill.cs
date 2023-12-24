namespace XL.Report.Styles.Fills;

// todo gradient fill
public abstract partial class Fill
{
    public static Fill No { get; } = new NoFill();

    public abstract T Accept<T>(Visitor<T> visitor);
    public abstract void Write(Xml xml);

    public static void Write(Xml xml, IEnumerable<Fill> fills)
    {
        xml.WriteStartElement(XlsxStructure.Styles.Fills.Collection);
        {
            foreach (var fill in fills)
            {
                fill.Write(xml);
            }
        }
        xml.WriteEndElement();
    }
}