using XL.Report.Styles;

namespace XL.Report;

public sealed class ConditionalFormatting
{
    private readonly Range range;
    private readonly Rule rule;
    private readonly StyleDiffId styleId;

    public ConditionalFormatting(Range range, Rule rule, StyleDiffId styleId)
    {
        this.range = range;
        this.rule = rule;
        this.styleId = styleId;
    }

    public void Write(Xml xml, int priority)
    {
        using (xml.WriteStartElement("conditionalFormatting"))
        {
            xml.WriteAttribute("sqref", range);
            rule.Write(xml, styleId, priority);
        }
    }

    public abstract class Rule
    {
        public abstract void Write(Xml xml, StyleDiffId styleDiffId, int priority);

        public sealed class ValueIs : Rule
        {
            private readonly Expression operand;

            private readonly string @operator;

            public ValueIs(string @operator, Expression operand)
            {
                this.@operator = @operator;
                this.operand = operand;
            }

            // todo add or equal
            public static Rule GreaterThan(Expression expression) => new ValueIs(
                "greaterThan",
                expression
            );

            public static Rule LessThan(Expression expression) => new ValueIs(
                "lessThan",
                expression
            );

            public static Rule Equal(Expression expression) => new ValueIs(
                "equal",
                expression
            );

            public override void Write(Xml xml, StyleDiffId styleDiffId, int priority)
            {
                using (xml.WriteStartElement("cfRule"))
                {
                    xml.WriteAttribute("type", "cellIs");
                    xml.WriteAttribute("dxfId", styleDiffId);
                    xml.WriteAttribute("priority", priority);
                    xml.WriteAttribute("operator", @operator);
                    using (xml.WriteStartElement("formula"))
                    {
                        xml.WriteValue(operand);
                    }
                }
            }
        }

        public sealed class ValueIsBetween : Rule
        {
            private readonly Expression left;
            private readonly Expression right;

            public ValueIsBetween(Expression left, Expression right)
            {
                this.left = left;
                this.right = right;
            }

            public override void Write(Xml xml, StyleDiffId styleDiffId, int priority)
            {
                using (xml.WriteStartElement("cfRule"))
                {
                    xml.WriteAttribute("type", "containsText");
                    xml.WriteAttribute("dxfId", styleDiffId);
                    xml.WriteAttribute("priority", priority);
                    xml.WriteAttribute("operator", "between");
                    using (xml.WriteStartElement("formula"))
                    {
                        xml.WriteValue(left);
                    }

                    using (xml.WriteStartElement("formula"))
                    {
                        xml.WriteValue(right);
                    }
                }
            }
        }

        public static Rule Duplicates { get; } = new DuplicateValues();

        private sealed class DuplicateValues : Rule
        {
            public override void Write(Xml xml, StyleDiffId styleDiffId, int priority)
            {
                using (xml.WriteStartElement("cfRule"))
                {
                    xml.WriteAttribute("type", "duplicateValues");
                    xml.WriteAttribute("dxfId", styleDiffId);
                    xml.WriteAttribute("priority", priority);
                }
            }
        }

        public sealed class Contains : Rule
        {
            private readonly string text;

            public Contains(string text)
            {
                this.text = text;
            }

            private Expression Formula => new Expression.Verbatim(
                // todo A1 is correct?
                // todo escape
                $"NOT(ISERROR(SEARCH(\"{text}\",A1)))"
            );

            public override void Write(Xml xml, StyleDiffId styleDiffId, int priority)
            {
                using (xml.WriteStartElement("cfRule"))
                {
                    xml.WriteAttribute("type", "containsText");
                    xml.WriteAttribute("dxfId", styleDiffId);
                    xml.WriteAttribute("priority", priority);
                    xml.WriteAttribute("operator", "containsText");

                    // todo needed?
                    xml.WriteAttribute("text", text);

                    // todo needed?
                    using (xml.WriteStartElement("formula"))
                    {
                        xml.WriteValue(Formula);
                    }
                }
            }
        }
    }
}