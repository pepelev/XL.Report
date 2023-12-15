namespace XL.Report;

public sealed class BoundedSharedStrings : SharedStrings
{
    private readonly Dictionary<string, SharedStringId> index = new(StringComparer.Ordinal);
    private int totalLength;
    private readonly int maxCount;
    private readonly int maxSingleStringLength;
    private readonly int maxTotalLength;

    public BoundedSharedStrings(int maxCount, int maxSingleStringLength, int maxTotalLength)
    {
        this.maxCount = maxCount;
        this.maxTotalLength = maxTotalLength;
        this.maxSingleStringLength = maxSingleStringLength;
    }

    public override SharedStringId? TryRegister(string @string)
    {
        if (index.TryGetValue(@string, out var id))
        {
            return id;
        }

        if (maxSingleStringLength < @string.Length)
        {
            return null;
        }

        if (maxCount < index.Count + 1)
        {
            return null;
        }

        var newTotalLength = totalLength + @string.Length;
        if (maxTotalLength < newTotalLength)
        {
            return null;
        }

        return Add(@string);
    }

    public override SharedStringId ForceRegister(string @string)
    {
        if (index.TryGetValue(@string, out var id))
        {
            return id;
        }

        return Add(@string);
    }

    public override IEnumerator<(string String, SharedStringId Id)> GetEnumerator() =>
        index.Select(pair => (pair.Key, pair.Value)).GetEnumerator();

    private SharedStringId Add(string @string)
    {
        var sharedStringId = new SharedStringId(index.Count);
        index[@string] = sharedStringId;
        totalLength += @string.Length;
        return sharedStringId;
    }
}