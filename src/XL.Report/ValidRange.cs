namespace XL.Report;

public readonly struct ValidRange
{
    public ValidRange(Range value)
    {
        if (!value.IsValid)
        {
            throw new ArgumentException($"has value {value} which is not valid", nameof(value));
        }

        Value = value;
    }

    public Range Value { get; }

    public static implicit operator Range(ValidRange valid) => valid.Value;
}