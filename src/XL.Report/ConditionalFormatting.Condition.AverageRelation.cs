namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public abstract partial class Condition
    {
        public enum AverageRelation
        {
            BelowAverage,
            BelowOrEqualAverage,
            AboveOrEqualAverage,
            AboveAverage
        }
    }
}