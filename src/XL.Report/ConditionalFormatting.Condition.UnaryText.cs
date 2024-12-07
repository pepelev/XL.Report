namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public abstract partial class Condition
    {
        public sealed class UnaryText(string type, string text) : Condition
        {
            public override void WriteAttributes(Xml xml)
            {
                xml.WriteAttribute("type", type);
                xml.WriteAttribute("text", text);
            }

            public override void WriteBody(Xml xml)
            {
            }
        }
    }
}