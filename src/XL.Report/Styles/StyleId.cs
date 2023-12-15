namespace XL.Report.Styles;

public readonly struct StyleId : IComparable<StyleId>
{
    public int Index { get; }

    public StyleId(int index)
    {
        Index = index;
    }

    public int CompareTo(StyleId other)
    {
        return Index.CompareTo(other.Index);
    }
}