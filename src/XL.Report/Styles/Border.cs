namespace XL.Report.Styles;

// todo make struct record
public sealed class Border : IEquatable<Border>, IBorder
{
    public Border(BorderStyle style, Color? color = null)
    {
        Color = color;
        Style = style;
    }

    public Color? Color { get; }
    public BorderStyle Style { get; }

    public bool Equals(Border? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Nullable.Equals(Color, other.Color) && Style.Equals(other.Style);
    }

    public override string ToString()
    {
        var parts = new[]
        {
            Style.ToString(),
            Color?.ToString()
        }.Where(part => !string.IsNullOrWhiteSpace(part));
        return string.Join(' ', parts);
    }

    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is Border other && Equals(other);

    public override int GetHashCode() => HashCode.Combine(Color, Style);
}