using System.Xml;

namespace XL.Report;

public abstract class Content
{
    public abstract void Write(XmlWriter xml);
}