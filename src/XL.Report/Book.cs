using XL.Report.Styles;

namespace XL.Report;

public abstract class Book
{
    public abstract Style.Collection Styles { get; }
    public abstract SharedStrings Strings { get; }
}