using System.Globalization;
using System.Xml;

namespace XL.Report;

public sealed class Xml : IDisposable
{
    private readonly char[] buffer = new char[64];
    private readonly Action endDocument;
    private readonly Action endElement;

    public Xml(XmlWriter xml)
    {
        this.Raw = xml;
        endDocument = xml.WriteEndDocument;
        endElement = xml.WriteEndElement;
    }

    public XmlWriter Raw { get; }

    public void Dispose()
    {
        Raw.Dispose();
    }

    public Block WriteStartDocument(string rootElementName, string @namespace)
    {
        Raw.WriteStartDocument(true);
        Raw.WriteStartElement(rootElementName, @namespace);
        return new Block(endDocument);
    }

    public void WriteEndDocument()
    {
        Raw.WriteEndDocument();
    }

    public Block WriteStartElement(string name)
    {
        Raw.WriteStartElement(name);
        return new Block(endElement);
    }

    public void WriteEndElement()
    {
        Raw.WriteEndElement();
    }

    public void WriteEmptyElement(string name)
    {
        Raw.WriteStartElement(name);
        Raw.WriteEndElement();
    }

    public void WriteAttribute(string name, string value)
    {
        Raw.WriteAttributeString(name, value);
    }

    public void WriteAttribute(string prefix, string name, string value)
    {
        Raw.WriteAttributeString(prefix, name, null, value);
    }

    public void WriteAttribute<T>(string name, T value, string? format = null) where T : ISpanFormattable
    {
        Raw.WriteStartAttribute(name);
        WriteValue(value, format);
    }

    public void WriteValue(string value)
    {
        Raw.WriteValue(value);
    }

    public void WriteValue<T>(T value, string? format = null) where T : ISpanFormattable
    {
        if (value.TryFormat(buffer, out var charsWritten, format, CultureInfo.InvariantCulture))
        {
            Raw.WriteChars(buffer, 0, charsWritten);
        }
        else
        {
            Raw.WriteValue(value.ToString(format, CultureInfo.InvariantCulture));
        }
    }

    public readonly struct Block : IDisposable
    {
        private readonly Action action;

        public Block(Action action)
        {
            this.action = action;
        }

        public void Dispose()
        {
            action();
        }
    }
}