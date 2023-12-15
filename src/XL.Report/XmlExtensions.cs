using System.Globalization;
using System.Xml;

namespace XL.Report;

internal static class XmlExtensions
{
    public static void WriteAttributeInt(this XmlWriter xml, string name, int value)
    {
        xml.WriteAttributeString(name, value.ToString(CultureInfo.InvariantCulture));
    }
}