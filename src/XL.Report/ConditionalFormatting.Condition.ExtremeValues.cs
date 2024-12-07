namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public abstract partial class Condition
    {
        public sealed class ExtremeValues(Target target, int count) : Condition
        {
            public override void WriteAttributes(Xml xml)
            {
                xml.WriteAttribute("type", "top10");
                if (target == Target.Lowest)
                {
                    xml.WriteAttribute("bottom", "1");
                }

                xml.WriteAttribute("rank", count);
            }

            public override void WriteBody(Xml xml)
            {
            }
        }
    }
}