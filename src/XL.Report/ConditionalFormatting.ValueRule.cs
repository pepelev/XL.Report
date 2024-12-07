using XL.Report.Styles;

namespace XL.Report;

public sealed partial class ConditionalFormatting
{
    public sealed class ValueRule(Condition condition, StyleDiffId styleId) : Rule
    {
        public override void Write(Xml xml, int priority)
        {
            using (xml.WriteStartElement("cfRule"))
            {
                condition.WriteAttributes(xml);
                xml.WriteAttribute("priority", priority);
                xml.WriteAttribute("dxfId", styleId);
                condition.WriteBody(xml);
            }
        }

        public static Rule Duplicates(StyleDiffId styleId) => new ValueRule(Condition.Duplicates, styleId);
        public static Rule Unique(StyleDiffId styleId) => new ValueRule(Condition.Unique, styleId);
        public static Rule IsError(StyleDiffId styleId) => new ValueRule(Condition.IsError, styleId);
        public static Rule IsNotError(StyleDiffId styleId) => new ValueRule(Condition.IsNotError, styleId);

        public static Rule Between(Expression a, Expression b, StyleDiffId styleId) =>
            new ValueRule(Condition.Between(a, b), styleId);

        public static Rule NotBetween(Expression a, Expression b, StyleDiffId styleId) =>
            new ValueRule(Condition.NotBetween(a, b), styleId);

        public static Rule StartsWith(string text, StyleDiffId styleId) =>
            new ValueRule(Condition.StartsWith(text), styleId);

        public static Rule EndsWith(string text, StyleDiffId styleId) =>
            new ValueRule(Condition.EndsWith(text), styleId);

        public static Rule Contains(string text, StyleDiffId styleId) =>
            new ValueRule(Condition.Contains(text), styleId);

        public static Rule NotContains(string text, StyleDiffId styleId) =>
            new ValueRule(Condition.NotContains(text), styleId);

        public static Rule GreaterThan(Expression expression, StyleDiffId styleId) =>
            new ValueRule(Condition.GreaterThan(expression), styleId);

        public static Rule GreaterThanOrEqual(Expression expression, StyleDiffId styleId) =>
            new ValueRule(Condition.GreaterThanOrEqual(expression), styleId);

        public static Rule LessThan(Expression expression, StyleDiffId styleId) =>
            new ValueRule(Condition.LessThan(expression), styleId);

        public static Rule LessThanOrEqual(Expression expression, StyleDiffId styleId) =>
            new ValueRule(Condition.LessThanOrEqual(expression), styleId);

        public static Rule Equal(Expression expression, StyleDiffId styleId) =>
            new ValueRule(Condition.Equal(expression), styleId);

        public static Rule NotEqual(Expression expression, StyleDiffId styleId) =>
            new ValueRule(Condition.NotEqual(expression), styleId);

        public static Rule Formula(Expression expression, StyleDiffId styleId) =>
            new ValueRule(Condition.Formula(expression), styleId);

        // todo check all TimePeriods working
        public static Rule Yesterday(StyleDiffId styleId) => new ValueRule(Condition.Yesterday, styleId);
        public static Rule Today(StyleDiffId styleId) => new ValueRule(Condition.Today, styleId);
        public static Rule Tomorrow(StyleDiffId styleId) => new ValueRule(Condition.Tomorrow, styleId);

        public static Rule Last7Days(StyleDiffId styleId) => new ValueRule(Condition.Last7Days, styleId);

        public static Rule LastWeek(StyleDiffId styleId) => new ValueRule(Condition.LastWeek, styleId);
        public static Rule ThisWeek(StyleDiffId styleId) => new ValueRule(Condition.ThisWeek, styleId);
        public static Rule NextWeek(StyleDiffId styleId) => new ValueRule(Condition.NextWeek, styleId);

        public static Rule LastMonth(StyleDiffId styleId) => new ValueRule(Condition.LastMonth, styleId);
        public static Rule ThisMonth(StyleDiffId styleId) => new ValueRule(Condition.ThisMonth, styleId);
        public static Rule NextMonth(StyleDiffId styleId) => new ValueRule(Condition.NextMonth, styleId);
    }
}