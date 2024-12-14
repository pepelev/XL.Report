#region Legal
// Copyright 2024 Pepelev Alexey
// 
// This file is part of XL.Report.
// 
// XL.Report is free software: you can redistribute it and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation, either
// version 3 of the License, or (at your option) any later version.
// 
// XL.Report is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
// without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU Lesser General Public License for more details.
// 
// You should have received a copy of the GNU Lesser General Public License along with XL.Report.
// If not, see <https://www.gnu.org/licenses/>.
#endregion

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

    public void WriteStartAttribute(string name)
    {
        Raw.WriteStartAttribute(name);
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