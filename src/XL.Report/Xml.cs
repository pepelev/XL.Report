using System.Globalization;
using System.Xml;

namespace XL.Report;

public sealed class Xml : IDisposable
{
    private readonly XmlWriter xml;
    private readonly char[] buffer = new char[64];
    private readonly Action endDocument;
    private readonly Action endElement;

    public Xml(XmlWriter xml)
    {
        this.xml = xml;
        endDocument = xml.WriteEndDocument;
        endElement = xml.WriteEndElement;
    }

    public Block WriteStartDocument(string rootElementName, string @namespace)
    {
        xml.WriteStartDocument(standalone: true);
        xml.WriteStartElement(rootElementName, @namespace);
        return new Block(endDocument);
    }

    public void WriteEndDocument()
    {
        xml.WriteEndDocument();
    }

    public Block WriteStartElement(string name)
    {
        xml.WriteStartElement(name);
        return new Block(endElement);
    }

    public void WriteEndElement()
    {
        xml.WriteEndElement();
    }

    public void WriteAttribute(string name, string value)
    {
        xml.WriteAttributeString(name, value);
    }

    public void WriteAttribute(string prefix, string name, string value)
    {
        xml.WriteAttributeString(prefix, name, null, value);
    }

    public void WriteAttribute<T>(string name, T value) where T : IAllocationFreeWritable
    {
        xml.WriteStartAttribute(name);
        WriteValue(value);
    }

    public void WriteAttributeSpan<T>(string name, T value, string? format = null) where T : ISpanFormattable
    {
        xml.WriteStartAttribute(name);
        WriteValueSpan(value, format);
    }

    public void WriteValue(string value)
    {
        xml.WriteValue(value);
    }

    public void WriteValue<T>(T value) where T : IAllocationFreeWritable
    {
        if (value.TryFormat(buffer, out var charsWritten))
        {
            xml.WriteChars(buffer, 0, charsWritten);
        }
        else
        {
            xml.WriteValue(value.AsString());
        }
    }

    public void WriteValueSpan<T>(T value, string? format = null) where T : ISpanFormattable
    {
        if (value.TryFormat(buffer, out var charsWritten, format, CultureInfo.InvariantCulture))
        {
            xml.WriteChars(buffer, 0, charsWritten);
        }
        else
        {
            xml.WriteValue(value.ToString(format, CultureInfo.InvariantCulture));
        }
    }

    public XmlWriter Raw => xml;

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

    public void Dispose()
    {
        xml.Dispose();
    }
}