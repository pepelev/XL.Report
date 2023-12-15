using XL.Report.Styles.Fills;

namespace XL.Report.Styles;

public sealed record Appearance(
    Alignment Alignment,
    Font Font,
    Fill Fill,
    Borders Borders
)
{
    public static Appearance Default { get; } = new(
        Alignment.Default,
        Font.Calibri11,
        Fill.No,
        Borders.None
    );
}