﻿using System.Text;

namespace XL.Report.Styles;

public sealed class Font : IEquatable<Font>
{
    public Font(
        string family,
        float size,
        Color color,
        FontStyle style
    )
    {
        Family = family;
        Size = size;
        Color = color;
        Style = style;
    }

    public static Font Default { get; } = new(
        "Calibri",
        11,
        Color.Auto,
        FontStyle.Regular
    );

    public FontStyle Style { get; }
    public string Family { get; }
    public float Size { get; }
    public Color Color { get; }

    public bool Equals(Font other)
    {
        if (ReferenceEquals(null, other))
            return false;
        if (ReferenceEquals(this, other))
            return true;

        return Style.Equals(other.Style) &&
               string.Equals(Family, other.Family) &&
               Size.Equals(other.Size) &&
               Color.Equals(other.Color);
    }

    public override string ToString()
    {
        var builder = new StringBuilder(128);
        builder.Append(Family);
        builder.Append(" ");
        builder.Append(Size);
        if (Color.IsAuto)
            builder.Append(" Auto");

        else if (Color != new Color(0, 0, 0))
            builder.Append(Color.ToHex());

        builder.Append(" ");
        builder.Append(Style);

        return builder.ToString();
    }

    public override bool Equals(object obj)
    {
        if (ReferenceEquals(null, obj))
            return false;
        if (ReferenceEquals(this, obj))
            return true;

        return obj is Font font && Equals(font);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = Style.GetHashCode();
            hashCode = (hashCode * 397) ^ Family.GetHashCode();
            hashCode = (hashCode * 397) ^ Size.GetHashCode();
            hashCode = (hashCode * 397) ^ Color.GetHashCode();
            return hashCode;
        }
    }
}