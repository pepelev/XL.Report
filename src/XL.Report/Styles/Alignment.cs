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

namespace XL.Report.Styles;

public sealed class Alignment : IEquatable<Alignment>
{
    public Alignment(
        HorizontalAlignment horizontal,
        VerticalAlignment vertical,
        OverflowBehavior overflowBehavior = OverflowBehavior.Overflow,
        ReadingOrder readingOrder = ReadingOrder.ByContext,
        TextRotation textRotation = new())
    {
        var textRotationCorrect = textRotation == TextRotation.None || horizontal.SupportTextRotation;
        if (!textRotationCorrect)
        {
            throw new ArgumentException(
                $"{nameof(HorizontalAlignment)} {vertical} does not support {nameof(TextRotation)}",
                nameof(textRotation)
            );
        }

        var overflowBehaviorCorrect = overflowBehavior != OverflowBehavior.Shrink || vertical.SupportShrinkOnOverflow;
        if (!overflowBehaviorCorrect)
        {
            throw new ArgumentException(
                $"{nameof(VerticalAlignment)} {vertical} does not support " +
                $"{nameof(XL.Report.Styles.OverflowBehavior)}.{nameof(OverflowBehavior.Shrink)}",
                nameof(overflowBehavior)
            );
        }

        Horizontal = horizontal;
        Vertical = vertical;
        OverflowBehavior = overflowBehavior;
        ReadingOrder = readingOrder;
        TextRotation = textRotation;
    }

    public static Alignment Default { get; } = new(
        HorizontalAlignment.ByContent,
        VerticalAlignment.Bottom
    );

    public HorizontalAlignment Horizontal { get; }
    public VerticalAlignment Vertical { get; }
    public OverflowBehavior OverflowBehavior { get; }
    public ReadingOrder ReadingOrder { get; }
    public TextRotation TextRotation { get; }

    public bool IsDefault =>
        Horizontal.IsDefault &&
        Vertical.IsDefault &&
        OverflowBehavior == OverflowBehavior.Overflow &&
        ReadingOrder == ReadingOrder.ByContext &&
        TextRotation == TextRotation.None;

    public void Write(Xml xml)
    {
        using (xml.WriteStartElement("alignment"))
        {
            Horizontal.Write(xml);
            Vertical.Write(xml);
            if (TextRotation != TextRotation.None)
            {
                xml.WriteAttribute("textRotation", TextRotation);
            }

            if (OverflowBehavior == OverflowBehavior.Wrap)
            {
                xml.WriteAttribute("wrapText", "1");
            }

            if (OverflowBehavior == OverflowBehavior.Shrink)
            {
                xml.WriteAttribute("shrinkToFit", "1");
            }

            if (ReadingOrder == ReadingOrder.LeftToRight)
            {
                xml.WriteAttribute("readingOrder", "1");
            }

            if (ReadingOrder == ReadingOrder.RightToLeft)
            {
                xml.WriteAttribute("readingOrder", "2");
            }
        }
    }

    
    
    public bool Equals(Alignment? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Horizontal.Equals(other.Horizontal) &&
               Vertical.Equals(other.Vertical) &&
               OverflowBehavior == other.OverflowBehavior &&
               ReadingOrder == other.ReadingOrder &&
               TextRotation.Equals(other.TextRotation);
    }

    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is Alignment other && Equals(other);

    public override string ToString()
    {
        var parts = new[]
        {
            Horizontal.ToString(),
            Vertical.ToString(),
            OverflowBehavior switch
            {
                OverflowBehavior.Wrap => "wrap",
                OverflowBehavior.Shrink => "shrink-to-fit",
                _ => ""
            },
            ReadingOrder switch
            {
                ReadingOrder.LeftToRight => "left-to-right",
                ReadingOrder.RightToLeft => "right-to-left",
                _ => ""
            },
            TextRotation == TextRotation.None
                ? ""
                : $"{nameof(TextRotation)}: {TextRotation}"
        }.Where(part => !string.IsNullOrWhiteSpace(part));

        return string.Join(' ', parts);
    }

    public override int GetHashCode() =>
        HashCode.Combine(Horizontal, Vertical, (int)OverflowBehavior, (int)ReadingOrder, TextRotation);
}