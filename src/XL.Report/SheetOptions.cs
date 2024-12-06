namespace XL.Report;

public sealed record SheetOptions(FreezeOptions Freeze, ColumnOptions.Collection Columns)
{
    public static SheetOptions Default { get; } = new(
        FreezeOptions.None,
        ColumnOptions.Collection.Default
    );

    public void Write(Xml xml)
    {
        Freeze.WriteAsSingleSheetView(xml);
        Columns.Write(xml);
    }
}