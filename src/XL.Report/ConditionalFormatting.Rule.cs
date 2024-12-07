using XL.Report.Styles;

namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public readonly record struct Rule(
        Condition Condition,
        StyleDiffId StyleId
    )
    {
        public void Write(Xml xml, int priority)
        {
            using (xml.WriteStartElement("cfRule"))
            {
                Condition.WriteAttributes(xml);
                xml.WriteAttribute("priority", priority);
                xml.WriteAttribute("dxfId", StyleId);
                Condition.WriteBody(xml);
            }
        }

        public static Rule Duplicates(StyleDiffId styleId) => new(Condition.Duplicates, styleId);
        public static Rule Unique(StyleDiffId styleId) => new(Condition.Unique, styleId);
        public static Rule IsError(StyleDiffId styleId) => new(Condition.IsError, styleId);
        public static Rule IsNotError(StyleDiffId styleId) => new(Condition.IsNotError, styleId);

        public static Rule Between(Expression a, Expression b, StyleDiffId styleId) =>
            new(Condition.Between(a, b), styleId);

        public static Rule NotBetween(Expression a, Expression b, StyleDiffId styleId) =>
            new(Condition.NotBetween(a, b), styleId);

        public static Rule StartsWith(string text, StyleDiffId styleId) => new(Condition.StartsWith(text), styleId);
        public static Rule EndsWith(string text, StyleDiffId styleId) => new(Condition.EndsWith(text), styleId);
        public static Rule Contains(string text, StyleDiffId styleId) => new(Condition.Contains(text), styleId);
        public static Rule NotContains(string text, StyleDiffId styleId) => new(Condition.NotContains(text), styleId);

        public static Rule GreaterThan(Expression expression, StyleDiffId styleId) =>
            new(Condition.GreaterThan(expression), styleId);

        public static Rule GreaterThanOrEqual(Expression expression, StyleDiffId styleId) =>
            new(Condition.GreaterThanOrEqual(expression), styleId);

        public static Rule LessThan(Expression expression, StyleDiffId styleId) =>
            new(Condition.LessThan(expression), styleId);

        public static Rule LessThanOrEqual(Expression expression, StyleDiffId styleId) =>
            new(Condition.LessThanOrEqual(expression), styleId);

        public static Rule Equal(Expression expression, StyleDiffId styleId) =>
            new(Condition.Equal(expression), styleId);

        public static Rule NotEqual(Expression expression, StyleDiffId styleId) =>
            new(Condition.NotEqual(expression), styleId);

        public static Rule Formula(Expression expression, StyleDiffId styleId) =>
            new(Condition.Formula(expression), styleId);

        // todo check all TimePeriods working
        public static Rule Yesterday(StyleDiffId styleId) => new(Condition.Yesterday, styleId);
        public static Rule Today(StyleDiffId styleId) => new(Condition.Today, styleId);
        public static Rule Tomorrow(StyleDiffId styleId) => new(Condition.Tomorrow, styleId);

        public static Rule Last7Days(StyleDiffId styleId) => new(Condition.Last7Days, styleId);

        public static Rule LastWeek(StyleDiffId styleId) => new(Condition.LastWeek, styleId);
        public static Rule ThisWeek(StyleDiffId styleId) => new(Condition.ThisWeek, styleId);
        public static Rule NextWeek(StyleDiffId styleId) => new(Condition.NextWeek, styleId);

        public static Rule LastMonth(StyleDiffId styleId) => new(Condition.LastMonth, styleId);
        public static Rule ThisMonth(StyleDiffId styleId) => new(Condition.ThisMonth, styleId);
        public static Rule NextMonth(StyleDiffId styleId) => new(Condition.NextMonth, styleId);
    }
}