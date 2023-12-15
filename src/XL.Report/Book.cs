using XL.Report.Styles;

namespace XL.Report;

public abstract class Book : IDisposable
{
    public abstract Style.Collection Styles { get; }
    public abstract SharedStrings Strings { get; }

    // todo check name constraints
    public abstract SheetBuilder OpenSheet(string name, SheetOptions options);

    public abstract void Complete();

    public abstract class SheetBuilder : IDisposable
    {
        public abstract string Name { get; set; }
        public abstract void Dispose();
        public abstract T WriteRow<T>(IUnit<T> unit);
        public abstract void Complete();
    }

    public abstract void Dispose();
}