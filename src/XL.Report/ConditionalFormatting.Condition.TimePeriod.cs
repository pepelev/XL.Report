namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public abstract partial class Condition
    {
        public sealed class TimePeriod(string period) : Condition
        {
            public override void WriteAttributes(Xml xml)
            {
                xml.WriteAttribute("type", "timePeriod");
                xml.WriteAttribute("timePeriod", period);
            }

            public override void WriteBody(Xml xml)
            {
            }
        }
    }
}