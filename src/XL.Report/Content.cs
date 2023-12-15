using System.Xml;

namespace XL.Report;

public abstract class Content
{
    // todo styles
    public abstract void Write(XmlWriter xml, Location location);
}