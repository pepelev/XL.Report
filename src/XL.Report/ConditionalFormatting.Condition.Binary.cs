namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public abstract partial class Condition
    {
        public sealed class Binary(string @operator, Expression a, Expression b) : Condition
        {
            public override void WriteAttributes(Xml xml)
            {
                xml.WriteAttribute("type", "cellIs");
                xml.WriteAttribute("operator", @operator);
            }

            public override void WriteBody(Xml xml)
            {
                using (xml.WriteStartElement("formula"))
                {
                    xml.WriteValue(a);
                }

                using (xml.WriteStartElement("formula"))
                {
                    xml.WriteValue(b);
                }
            }
        }
    }
}