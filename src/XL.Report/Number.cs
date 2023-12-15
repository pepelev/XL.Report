using System.Xml;

namespace XL.Report;

public sealed class InlineString : Content
{
    private readonly string content;

    public InlineString(string content)
    {
        this.content = content;
    }

    public override void Write(XmlWriter xml)
    {
        // todo extract to consts
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

public sealed class Number : Content
{
    // todo extract to consts
    private const string value = "v";

    private readonly long content;

    public Number(long content)
    {
        this.content = content;
    }

    public override void Write(XmlWriter xml)
    {
        xml.WriteStartElement(value);
        {
            xml.WriteValue(content);
        }
        xml.WriteEndElement();
    }
}

public interface IUnit<out T>
{
    public T Write(SheetWindow window);
}

public readonly struct StyleId : IComparable<StyleId>
{
    public int Index { get; }

    public StyleId(int index)
    {
        Index = index;
    }

    // todo
    public bool IsDefault => Index == 0;

    public int CompareTo(StyleId other)
    {
        return Index.CompareTo(other.Index);
    }
}

public readonly struct Row : IUnit<Size>
{
    private readonly IUnit<Size>[] units;

    public Row(params IUnit<Size>[] units)
    {
        this.units = units;
    }

    public Size Write(SheetWindow window)
    {
        var offset = 0;
        var maxHeight = 0;
        foreach (var unit in units ?? Array.Empty<IUnit<Size>>())
        {
            window.PushMove(new Offset(offset, 0));
            var size = unit.Write(window);
            window.PopMove();
            offset += size.Width;
            maxHeight = Math.Max(maxHeight, size.Height);
        }

        return new Size(offset, maxHeight);
    }
}

public readonly struct Cell : IUnit<Location>, IUnit<Size>
{
    private readonly Content content;
    private readonly StyleId styleId;

    public Cell(Content content, StyleId styleId)
    {
        this.content = content;
        this.styleId = styleId;
    }

    public Location Write(SheetWindow window)
    {
        window.Place(Offset.Zero, content, styleId);
        return window.Range.LeftTopCell;
    }

    Size IUnit<Size>.Write(SheetWindow window)
    {
        Write(window);
        return Size.Cell;
    }
}

public readonly struct Merge : IUnit<Range>
{
    private readonly Content content;
    private readonly Size size;
    private readonly StyleId styleId;

    public Merge(Content content, Size size, StyleId styleId)
    {
        this.content = content;
        this.size = size;
        this.styleId = styleId;
    }

    public Range Write(SheetWindow window)
    {
        window.Merge(Offset.Zero, size, content, styleId);
        return new Range(window.Range.LeftTopCell, size);
    }
}