using XL.Report.Styles;

namespace XL.Report;

public sealed class ConditionalFormatting
{
    private readonly IReadOnlyCollection<Range> ranges;
    private readonly IReadOnlyCollection<Rule> rules;

    public ConditionalFormatting(Range range, Rule rule)
        : this(new[] { range }, new[] { rule })
    {
    }

    public ConditionalFormatting(IReadOnlyCollection<Range> ranges, IReadOnlyCollection<Rule> rules)
    {
        if (ranges.Count == 0)
        {
            throw new ArgumentException("empty", nameof(ranges));
        }

        if (rules.Count == 0)
        {
            throw new ArgumentException("empty", nameof(rules));
        }

        this.ranges = ranges;
        this.rules = rules;
    }

    public int Write(Xml xml, int firstPriority)
    {
        using (xml.WriteStartElement("conditionalFormatting"))
        {
            xml.WriteStartAttribute("sqref");

            var first = true;
            foreach (var range in ranges)
            {
                if (!first)
                {
                    xml.WriteValue(' ');
                }

                xml.WriteValue(range);
                first = false;
            }

            var priority = firstPriority;
            foreach (var rule in rules)
            {
                rule.Write(xml, priority);
                priority++;
            }

            return priority;
        }
    }

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
        public static Rule Between(Expression a, Expression b, StyleDiffId styleId) => new(Condition.Between(a, b), styleId);
        public static Rule NotBetween(Expression a, Expression b, StyleDiffId styleId) => new(Condition.NotBetween(a, b), styleId);
        public static Rule StartsWith(string text, StyleDiffId styleId) => new(Condition.StartsWith(text), styleId);
        public static Rule EndsWith(string text, StyleDiffId styleId) => new(Condition.EndsWith(text), styleId);
        public static Rule Contains(string text, StyleDiffId styleId) => new(Condition.Contains(text), styleId);
        public static Rule NotContains(string text, StyleDiffId styleId) => new(Condition.NotContains(text), styleId);
        public static Rule GreaterThan(Expression expression, StyleDiffId styleId) => new(Condition.GreaterThan(expression), styleId);
        public static Rule GreaterThanOrEqual(Expression expression, StyleDiffId styleId) => new(Condition.GreaterThanOrEqual(expression), styleId);
        public static Rule LessThan(Expression expression, StyleDiffId styleId) => new(Condition.LessThan(expression), styleId);
        public static Rule LessThanOrEqual(Expression expression, StyleDiffId styleId) => new(Condition.LessThanOrEqual(expression), styleId);
        public static Rule Equal(Expression expression, StyleDiffId styleId) => new(Condition.Equal(expression), styleId);
        public static Rule NotEqual(Expression expression, StyleDiffId styleId) => new(Condition.NotEqual(expression), styleId);
        public static Rule Formula(Expression expression, StyleDiffId styleId) => new(Condition.Formula(expression), styleId);

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

    public abstract class Condition
    {
        public abstract void WriteAttributes(Xml xml);
        public abstract void WriteBody(Xml xml);

        public sealed class Unary(string @operator, Expression operand) : Condition
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
                    xml.WriteValue(operand);
                }
            }
        }

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

        public static Condition Duplicates { get; } = new Parameterless("duplicateValues");
        public static Condition Unique { get; } = new Parameterless("uniqueValues");
        public static Condition Between(Expression a, Expression b) => new Binary("between", a, b);
        public static Condition NotBetween(Expression a, Expression b) => new Binary("notBetween", a, b);
        public static Condition IsError { get; } = new Parameterless("containsErrors");
        public static Condition IsNotError { get; } = new Parameterless("notContainsErrors");
        public static Condition StartsWith(string text) => new UnaryText("beginsWith", text);
        public static Condition EndsWith(string text) => new UnaryText("endsWith", text);
        public static Condition Contains(string text) => new UnaryText("containsText", text);
        public static Condition NotContains(string text) => new UnaryText("notContainsText", text);

        public static Condition GreaterThan(Expression expression) => new Unary(
            "greaterThan",
            expression
        );

        public static Condition GreaterThanOrEqual(Expression expression) => new Unary(
            "greaterThanOrEqual",
            expression
        );

        public static Condition LessThan(Expression expression) => new Unary(
            "lessThan",
            expression
        );

        public static Condition LessThanOrEqual(Expression expression) => new Unary(
            "lessThanOrEqual",
            expression
        );

        public static Condition Equal(Expression expression) => new Unary(
            "equal",
            expression
        );

        public static Condition NotEqual(Expression expression) => new Unary(
            "notEqual",
            expression
        );

        public static Condition Formula(Expression expression) => new Unary(
            "expression",
            expression
        );

        public static Condition Yesterday { get; } = new TimePeriod("yesterday");
        public static Condition Today { get; } = new TimePeriod("today");
        public static Condition Tomorrow { get; } = new TimePeriod("tomorrow");

        public static Condition Last7Days { get; } = new TimePeriod("last7Days");

        public static Condition LastWeek { get; } = new TimePeriod("lastWeek");
        public static Condition ThisWeek { get; } = new TimePeriod("thisWeek");
        public static Condition NextWeek { get; } = new TimePeriod("nextWeek");

        public static Condition LastMonth { get; } = new TimePeriod("lastMonth");
        public static Condition ThisMonth { get; } = new TimePeriod("thisMonth");
        public static Condition NextMonth { get; } = new TimePeriod("nextMonth");

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

        public enum Target
        {
            Highest,
            Lowest
        }

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

        public enum AverageRelation
        {
            BelowAverage,
            BelowOrEqualAverage,
            AboveOrEqualAverage,
            AboveAverage
        }

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