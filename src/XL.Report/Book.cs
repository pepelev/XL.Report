using XL.Report.Styles;

namespace XL.Report;

public abstract class Book : IAsyncDisposable
{
    public abstract Style.Collection Styles { get; }
    public abstract SharedStrings Strings { get; }

    public abstract ValueTask DisposeAsync();

    // todo check name constraints
    public abstract SheetBuilder OpenSheet(string name);

    public abstract Task CompleteAsync();

    public abstract class SheetBuilder : IAsyncDisposable
    {
        public abstract string Name { get; set; }
        public abstract ValueTask DisposeAsync();
        public abstract Task<T> WriteRowAsync<T>(IUnit<T> unit);
        public abstract Task CompleteAsync();
    }
}