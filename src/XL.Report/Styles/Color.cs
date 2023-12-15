namespace XL.Report.Styles;

public readonly struct Color : IEquatable<Color>
{
    private const ulong AutoBit = 1UL << 32;

    internal ulong Value { get; }

    private Color(ulong value)
    {
        Value = value;
    }

    public Color(uint value)
        : this((ulong)value)
    {
    }

    public byte Red => (byte)(Value & 0xFF);
    public byte Green => (byte)((Value >> 8) & 0xFF);
    public byte Blue => (byte)((Value >> 16) & 0xFF);
    public byte Alpha => (byte)((Value >> 24) & 0xFF);

    public static Color Auto => new(AutoBit);

    public bool IsAuto => (AutoBit & Value) != 0;

    public Color(byte red, byte green, byte blue, byte alpha = byte.MaxValue)
        : this(BuildValue(red, green, blue, alpha))
    {
    }

    private static uint BuildValue(byte red, byte green, byte blue, byte alpha)
    {
        return (uint)
        (
            (red << 0) |
            (green << 8) |
            (blue << 16) |
            (alpha << 24)
        );
    }

    public static bool operator ==(Color a, Color b)
    {
        return a.Equals(b);
    }

    public static bool operator !=(Color a, Color b)
    {
        return !(a == b);
    }

    public bool Equals(Color other)
    {
        return Value == other.Value;
    }

    public override bool Equals(object obj)
    {
        if (ReferenceEquals(null, obj))
            return false;

        return obj is Color color && Equals(color);
    }

    public static bool TryParse(string value, out Color result)
    {
        if (value == null)
        {
            result = default(Color);
            return false;
        }

        if (value.Length < 3)
        {
            result = default(Color);
            return false;
        }

        var hex = 0U;
        var index = 0;
        var hexStarted = false;
        var sharpWas = false;
        const uint wrongHex = uint.MaxValue;
        foreach (var ch in value)
        {
            if (char.IsWhiteSpace(ch))
                continue;

            if (ch == '#')
            {
                if (sharpWas || hexStarted)
                {
                    result = default(Color);
                    return false;
                }

                sharpWas = true;
                continue;
            }

            var piece = GetHex(ch);
            if (piece == wrongHex)
            {
                result = default(Color);
                return false;
            }

            hexStarted = true;
            hex |= piece << (index++ * 4);
        }

        switch (index)
        {
            case 3:
            {
                var red = (byte)((hex & 15) * 17);
                var green = (byte)(((hex >> 4) & 15) * 17);
                var blue = (byte)((hex >> 8) * 17);
                result = new Color(red, green, blue);
                return true;
            }
            case 4:
            {
                var red = (byte)((hex & 15) * 17);
                var green = (byte)(((hex >> 4) & 15) * 17);
                var blue = (byte)(((hex >> 8) & 15) * 17);
                var alpha = (byte)(((hex >> 12) & 15) * 17);
                result = new Color(red, green, blue, alpha);
                return true;
            }
            case 6:
                ReverseByte(ref hex, 0);
                ReverseByte(ref hex, 1);
                ReverseByte(ref hex, 2);
                result = new Color(hex | (255U << 24));
                return true;
            case 8:
                ReverseByte(ref hex, 0);
                ReverseByte(ref hex, 1);
                ReverseByte(ref hex, 2);
                ReverseByte(ref hex, 3);
                result = new Color(hex);
                return true;
            default:
                result = default(Color);
                return false;
        }

        uint GetHex(char ch)
        {
            if (ch >= '0' && ch <= '9')
                return (uint)ch - '0';

            if (ch >= 'A' && ch <= 'F')
                return (uint)ch - 'A' + 10;

            if (ch >= 'a' && ch <= 'f')
                return (uint)ch - 'a' + 10;

            return wrongHex;
        }

        void ReverseByte(ref uint colorHex, int byteIndex)
        {
            var bitShift = 8 * byteIndex;
            var @byte = (colorHex >> bitShift) & 255;
            var high = @byte >> 4;
            var low = @byte & 15;
            var reverseByte = (low << 4) | high;
            colorHex = (colorHex & ~(255U << bitShift)) | (reverseByte << bitShift);
        }
    }

    public static Color Parse(string value)
    {
        if (TryParse(value, out var result))
            return result;

        throw new ArgumentException("Could not parse value. Value must have format '#000000'", nameof(value));
    }

    public override string ToString()
    {
        return IsAuto
            ? "Auto"
            : $"#{Red:X2}{Green:X2}{Blue:X2}{Alpha:X2}";
    }

    public override int GetHashCode()
    {
        return (int)Value;
    }

    public string ToHex()
    {
        return ToString();
    }
}