namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public abstract partial class Condition
    {
        public sealed class ExtremePercentValues(Target target, int percent) : Condition
        {
            public override void WriteAttributes(Xml xml)
            {
                xml.WriteAttribute("type", "top10");
                xml.WriteAttribute("percent", "1");
                if (target == Target.Lowest)
                {
                    xml.WriteAttribute("bottom", "1");
                }

                xml.WriteAttribute("rank", percent);
            }

            public override void WriteBody(Xml xml)
            {
            }
        }
    }
}