#region Legal
// Copyright 2024 Pepelev Alexey
// 
// This file is part of XL.Report.
// 
// XL.Report is free software: you can redistribute it and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation, either
// version 3 of the License, or (at your option) any later version.
// 
// XL.Report is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
// without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU Lesser General Public License for more details.
// 
// You should have received a copy of the GNU Lesser General Public License along with XL.Report.
// If not, see <https://www.gnu.org/licenses/>.
#endregion

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