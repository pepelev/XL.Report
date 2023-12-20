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
        public abstract void DefineName(string name, Range range, string? comment = null);

        // link = url | range
        // url = http(s) | file | mail
        // range = defined-name | range
        public abstract void AddHyperlink(Range range, string url, string? tooltip = null);
        public abstract void AddHyperlinkToDefinedName(Range range, string name, string? tooltip = null);
        public abstract void AddHyperlinkToRange(Range range, Range target, string? tooltip = null);
        public abstract void AddHyperlinkToRange(Range range, SheetRelated<Range> target, string? tooltip = null);
        public abstract void Complete();
    }

    public abstract void Dispose();
}

public static class Hyperlink
{
    public static string Mailto(string email, string? subject = null) =>
        subject is { } value
            ? $"mailto:{email}?subject={value}"
            : $"mailto:{email}";
}