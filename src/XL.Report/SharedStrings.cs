namespace XL.Report;

public abstract class SharedStrings
{
    public abstract SharedStringId? TryRegister(string @string);
    public abstract SharedStringId ForceRegister(string @string);
}