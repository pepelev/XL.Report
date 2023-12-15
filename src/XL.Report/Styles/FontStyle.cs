using System.Text;

namespace XL.Report.Styles;

public struct FontStyle : IEquatable<FontStyle>
{
    private const ushort UnderlineMask = 0x7F;
    private const ushort UnderlineNullBit = 0x80;

    private const ushort BoldBit = 1 << 8;
    private const ushort ItalicBit = 1 << 9;
    private const ushort CrossedBit = 1 << 10;

    private readonly ushort value;

    public FontStyle(
        Underline? underline,
        bool bold,
        bool italic,
        bool crossed
    )
    {
        value = (ushort)
        (
            GetUnderline()
            | GetBitValue(bold, BoldBit)
            | GetBitValue(italic, ItalicBit)
            | GetBitValue(crossed, CrossedBit)
        );

        int GetUnderline()
        {
            return underline is Underline underlineValue
                ? (int)underlineValue & UnderlineMask
                : UnderlineNullBit;
        }

        int GetBitValue(bool value, int bit)
        {
            return value
                ? bit
                : 0x0000;
        }
    }

    public static FontStyle Regular => new(null, false, false, false);
    public static FontStyle Bold => new(null, true, false, false);
    public static FontStyle Italic => new(null, false, true, false);
    public static FontStyle Crossed => new(null, false, false, true);

    public static FontStyle Underlined(Underline underline = Styles.Underline.Single)
    {
        return new FontStyle(underline, false, false, false);
    }

    public Underline? Underline
    {
        get
        {
            if ((value & UnderlineNullBit) == UnderlineNullBit)
                return null;

            return (Underline)(value & UnderlineMask);
        }
    }

    public bool IsBold => (value & BoldBit) == BoldBit;
    public bool IsItalic => (value & ItalicBit) == ItalicBit;
    public bool IsCrossed => (value & CrossedBit) == CrossedBit;

    public bool Equals(FontStyle other)
    {
        return value == other.value;
    }

    public override bool Equals(object obj)
    {
        return obj is FontStyle style && Equals(style);
    }

    public override string ToString()
    {
        if (Equals(Regular))
            return nameof(Regular);
        if (Equals(Bold))
            return nameof(Bold);
        if (Equals(Italic))
            return nameof(Italic);
        if (Equals(Crossed))
            return nameof(Crossed);

        var builder = new StringBuilder(64);
        var isFirst = true;

        if (IsBold)
        {
            builder.Append("Bold");
            isFirst = false;
        }

        if (IsItalic)
        {
            if (!isFirst)
                builder.Append(", ");
            builder.Append("Italic");
            isFirst = false;
        }

        if (IsCrossed)
        {
            if (!isFirst)
                builder.Append(", ");
            builder.Append("Crossed");
            isFirst = false;
        }

        if (Underline is Underline underline)
        {
            if (!isFirst)
                builder.Append(", ");
            builder.Append("Underline: ");
            builder.Append(underline);
        }

        return builder.ToString();
    }

    public override int GetHashCode()
    {
        return value;
    }
}