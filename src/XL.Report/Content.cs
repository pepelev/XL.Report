using System.Xml;

namespace XL.Report;

public abstract class Content : IContent
{
    public abstract void Write(XmlWriter xml, StyleId styleId);
    public abstract void Write(XmlWriter xml);

    // todo styles
    public abstract void Write(XmlWriter xml, Location location);
}