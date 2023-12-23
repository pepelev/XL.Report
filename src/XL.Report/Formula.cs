namespace XL.Report;

public sealed class Formula : Content
{
    private readonly Expression expression;

    public Formula(Expression expression)
    {
        this.expression = expression;
    }

    public override void Write(Xml xml)
    {
        // todo extract to consts
        xml.WriteAttribute("t", "str");
        {
            using (xml.WriteStartElement("f"))
            {
                xml.WriteValue(expression);
            }
        }
    }
}