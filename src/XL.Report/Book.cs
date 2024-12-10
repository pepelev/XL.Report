using XL.Report.Styles;

namespace XL.Report;

public abstract class Book : IDisposable
{
    public abstract Style.Collection Styles { get; }
    public abstract SharedStrings Strings { get; }

    // todo check name constraints, uniq
    public abstract SheetBuilder CreateSheet(string name, SheetOptions options);

    public abstract void Complete();

    public abstract class SheetBuilder : IDisposable
    {
        public abstract string Name { get; set; }
        public abstract void Dispose();
        public abstract T WriteRow<T>(IUnit<T> unit);
        public abstract T WriteRow<T>(IUnit<T> unit, RowOptions options);
        public abstract void DefineName(string name, ValidRange range, string? comment = null);
        public abstract Hyperlinks Hyperlinks { get; }
        public abstract void AddConditionalFormatting(ConditionalFormatting formatting);
        public abstract void Complete();
    }

    public abstract void Dispose();
}