namespace XL.Report.Tests;

public static class RandomExtensions
{
    public static T Pick<T>(this Random random, IReadOnlyList<T> items) =>
        items[random.Next(items.Count)];
}