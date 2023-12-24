namespace XL.Report.Styles;

public readonly partial record struct FontStyle
{
    public readonly record struct Diff(
        bool IsBold = false,
        bool IsItalic = false,
        bool IsStrikethrough = false,
        UnderlineDiff? Underline = null,
        FontVerticalAlignmentDiff? Alignment = null
    );
}