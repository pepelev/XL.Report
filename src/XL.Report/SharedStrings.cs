using System.Collections;

namespace XL.Report;

public abstract class SharedStrings : IEnumerable<(string String, SharedStringId Id)>
{
    public abstract SharedStringId? TryRegister(string @string);
    public abstract SharedStringId ForceRegister(string @string);
    public abstract IEnumerator<(string String, SharedStringId Id)> GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}