namespace XL.Report;

internal readonly record struct FormatContext(bool Success, int CharsWritten)
{
    public static FormatContext Start => new(true, 0);

    public FormatContext Write<T>(ref Span<char> destination, T value) where T : IAllocationFreeWritable
    {
        if (!Success)
        {
            return new FormatContext(false, 0);
        }

        if (value.TryFormat(destination, out var charsWritten))
        {
            destination = destination[charsWritten..];
            return new FormatContext(true, CharsWritten + charsWritten);
        }

        return new FormatContext(false, 0);
    }

    public FormatContext Write(ref Span<char> destination, string value)
    {
        if (!Success)
        {
            return new FormatContext(false, 0);
        }

        if (value.Length <= destination.Length)
        {
            value.CopyTo(destination);
            destination = destination[value.Length..];
            return new FormatContext(true, CharsWritten + value.Length);
        }

        return new FormatContext(false, 0);
    }

    public bool Deconstruct(out int charsWritten)
    {
        charsWritten = CharsWritten;
        return Success;
    }
}