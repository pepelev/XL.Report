namespace XL.Report;

public sealed partial class ConditionalFormatting
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
}