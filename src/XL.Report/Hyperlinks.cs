namespace XL.Report;

public abstract class Hyperlinks
{
    public abstract void Add(ValidRange range, string url, string? tooltip = null);
    public abstract void AddToDefinedName(ValidRange range, string name, string? tooltip = null);
    public abstract void AddToRange(ValidRange range, ValidRange target, string? tooltip = null);
    public abstract void AddToRange(ValidRange range, SheetRelated<ValidRange> target, string? tooltip = null);
}