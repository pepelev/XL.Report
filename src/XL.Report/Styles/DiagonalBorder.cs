namespace XL.Report.Styles;

public sealed class DiagonalBorders : IEquatable<DiagonalBorders>, IBorder
{
    public DiagonalBorders(BorderStyle style, Color? color, bool up, bool down)
    {
        if (!up && !down)
        {
            throw new ArgumentException(
                $"at least one of {nameof(up)}, {nameof(down)} must be true"
            );
        }

        Up = up;
        Down = down;
        Color = color;
        Style = style;
    }

    public bool Up { get; }
    public bool Down { get; }

    public Color? Color { get; }
    public BorderStyle Style { get; }

    public bool Equals(DiagonalBorders? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Up == other.Up &&
               Down == other.Down &&
               Nullable.Equals(Color, other.Color) &&
               Style.Equals(other.Style);
    }

    public override string ToString()
    {
        var diagonals = (UpDiagonal: Up, DownDiagonal: Down) switch
        {
            (true, true) => "Cross",
            (true, false) => "Up",
            (false, true) => "Down",
            _ => "bug"
        };

        var parts = new[]
        {
            diagonals,
            Style.ToString(),
            Color?.ToString()
        }.Where(part => !string.IsNullOrWhiteSpace(part));
        return string.Join(' ', parts);
    }

    public override bool Equals(object? obj) => 
        ReferenceEquals(this, obj) || obj is DiagonalBorders other && Equals(other);

    public override int GetHashCode() => HashCode.Combine(Up, Down, Color, Style);
}