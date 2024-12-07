namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public abstract partial class Condition
    {
        public sealed class AverageRelatedValues(AverageRelation relation) : Condition
        {
            public override void WriteAttributes(Xml xml)
            {
                xml.WriteAttribute("type", "aboveAverage");
                var (above, equal) = relation switch
                {
                    AverageRelation.BelowAverage => (Above: 0, Equal: 0),
                    AverageRelation.BelowOrEqualAverage => (Above: 0, Equal: 1),
                    AverageRelation.AboveOrEqualAverage => (Above: 1, Equal: 1),
                    AverageRelation.AboveAverage => (Above: 1, Equal: 0),
                    _ => throw new ArgumentOutOfRangeException(nameof(relation), relation, null)
                };

                xml.WriteAttribute("aboveAverage", above);
                xml.WriteAttribute("equalAverage", equal);
            }

            public override void WriteBody(Xml xml)
            {
            }
        }
    }
}