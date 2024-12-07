namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public abstract partial class Condition
    {
        public sealed class Parameterless(string type) : Condition
        {
            public override void WriteAttributes(Xml xml)
            {
                xml.WriteAttribute("type", type);
            }

            public override void WriteBody(Xml xml)
            {
            }
        }
    }
}