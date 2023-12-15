using System.Xml;

namespace XL.Report;

// todo make struct
public sealed class InlineString : IContent
{
    private readonly string content;

    public InlineString(string content)
    {
        this.content = content;
    }

    public void Write(XmlWriter xml, StyleId styleId)
    {
        throw new NotImplementedException();
    }

    public void Write(XmlWriter xml)
    {
        xml.WriteAttributeString("t", "inlineStr");
        xml.WriteStartElement("is");
        {
            xml.WriteStartElement("t");
            {
                xml.WriteValue(content);
            }
            xml.WriteEndElement();
        }
        xml.WriteEndElement();
    }
}

// todo make struct
public sealed class Number : Content
{
    private const string cell = "c";
    private const string reference = "r";
    private const string value = "v";

    private readonly long content;

    public Number(long content)
    {
        this.content = content;
    }

    public override void Write(XmlWriter xml, StyleId styleId)
    {
        
    }

    public override void Write(XmlWriter xml)
    {
        xml.WriteStartElement(value);
        {
            xml.WriteValue(content);
        }
        xml.WriteEndElement();
    }

    public override void Write(XmlWriter xml, Location location)
    {
        xml.WriteStartElement(cell);
        xml.WriteStartAttribute(reference);
        // todo location can be incorrect
        xml.WriteValue(location.ToString());
        xml.WriteEndAttribute();
        {
            xml.WriteStartElement(value);
            {
                xml.WriteValue(content);
            }
            xml.WriteEndElement();
        }
        xml.WriteEndElement();
    }
}

public interface IUnit<out T>
{
    public T Write(XmlWriter xml, Canvas canvas);
}

public interface IContent
{
    void Write(XmlWriter xml, StyleId styleId);
    void Write(XmlWriter xml);
}

public readonly struct StyleId
{
    // todo
    public bool IsDefault => false;

    public string AsString()
    {
        throw new NotImplementedException();
    }
}

public sealed class Cell : IUnit<Location>
{
    // todo extract to separate class
    private const string cell = "c";
    private const string reference = "r";
    private const string styleName = "s";

    private readonly IContent content;
    private readonly Style style;
    private readonly Style.Collection styles;

    public Location Write(XmlWriter xml, Canvas canvas)
    {
        var range = canvas.Range;
        if (range.IsEmpty)
        {
            throw new InvalidOperationException();
        }

        var location = range.LeftTopCell;

        xml.WriteStartElement(cell);
        xml.WriteAttributeString(reference, location.ToString());

        var styleId = styles.Register(style);
        if (!styleId.IsDefault)
        {
            xml.WriteAttributeString(styleName, styleId.AsString());
        }

        content.Write(xml);
        xml.WriteEndElement();
        return location;
    }
}