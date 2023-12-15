namespace XL.Report.Styles.Fills;

public abstract partial class Fill
{
    public abstract T Accept<T>(Visitor<T> visitor);
}