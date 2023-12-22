namespace XL.Report;

public abstract class Hyperlinks
{
    public abstract void Add(Range range, string url, string? tooltip = null);
    public abstract void AddToDefinedName(Range range, string name, string? tooltip = null);
    public abstract void AddToRange(Range range, Range target, string? tooltip = null);
    public abstract void AddToRange(Range range, SheetRelated<Range> target, string? tooltip = null);
}