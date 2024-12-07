namespace XL.Report.Auxiliary;

internal readonly struct ByKey<TKey, TValue> : IComparer<KeyValuePair<TKey, TValue>>
    where TKey : IComparable<TKey>
{
    public int Compare(KeyValuePair<TKey, TValue> x, KeyValuePair<TKey, TValue> y) =>
        Comparer<TKey>.Default.Compare(x.Key, y.Key);
}