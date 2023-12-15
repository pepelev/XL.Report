namespace XL.Report.Styles;

public sealed class Format : IEquatable<Format>
{
    public Format(string code)
    {
        Code = code ?? throw new ArgumentNullException(nameof(code));
    }

    public string Code { get; }

    public bool Equals(Format? other)
    {
        if (ReferenceEquals(this, other))
            return true;

        if (ReferenceEquals(null, other))
            return false;

        return string.Equals(Code, other.Code);
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as Format);
    }

    public override int GetHashCode()
    {
        return Code.GetHashCode();
    }

    #region Formats

    public static Format General { get; } = new("General");
    public static Format IsoDate { get; } = new("yyyy-mm-dd");
    public static Format IsoTime { get; } = new("hh:mm:ss");
    public static Format IsoDateTime { get; } = new(@"yyyy-mm-dd""T""hh:mm:ss");
    public static Format Integer { get; } = new("0");

    public static Format Float(int decimals = 2)
    {
        if (decimals < 1)
        {
            throw new ArgumentOutOfRangeException(
                nameof(decimals),
                decimals,
                "A number that greater than zero is expected"
            );
        }

        return new Format($"0.${new string('0', decimals)}");
    }

    #endregion
}