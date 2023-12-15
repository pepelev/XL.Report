namespace XL.Report;

public sealed record SheetOptions(FreezeOptions Freeze, ColumnWidths Columns)
{
    public static SheetOptions Default { get; } = new(
        FreezeOptions.None,
        ColumnWidths.Default
    );
}