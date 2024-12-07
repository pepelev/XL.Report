namespace XL.Report.Auxiliary;

internal static class Default<T>
{
    private static T Value = default!;

    public static ref T Ref() => ref Value;
}