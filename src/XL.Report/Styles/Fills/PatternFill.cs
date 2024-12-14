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

namespace XL.Report.Styles.Fills;

public sealed class PatternFill : Fill, IEquatable<PatternFill>
{
    public PatternFill(Pattern pattern, Color? color, Color? background = null)
    {
        Pattern = pattern;
        Color = color;
        Background = background;
    }

    public Color? Color { get; }
    public Color? Background { get; }
    public Pattern Pattern { get; }

    public bool Equals(PatternFill? other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Nullable.Equals(Color, other.Color) &&
               Nullable.Equals(Background, other.Background) &&
               Pattern.Equals(other.Pattern);
    }

    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is PatternFill other && Equals(other);
    public override T Accept<T>(Visitor<T> visitor) => visitor.Visit(this);

    public override void Write(Xml xml)
    {
        using (xml.WriteStartElement(XlsxStructure.Styles.Fills.Fill))
        using (xml.WriteStartElement(XlsxStructure.Styles.Fills.Pattern))
        {
            xml.WriteAttribute(XlsxStructure.Styles.Fills.PatternType, Pattern);

            if (Color is { } color)
            {
                using (xml.WriteStartElement("fgColor"))
                {
                    xml.WriteAttribute("rgb", color.ToRgbHex());
                }
            }

            if (Background is { } background)
            {
                using (xml.WriteStartElement("bgColor"))
                {
                    xml.WriteAttribute("rgb", background.ToRgbHex());
                }
            }
        }
    }

    public override int GetHashCode() => HashCode.Combine(Color, Background, Pattern);
    public override string ToString() => $"{Pattern}:({PrintColor(Color)}, {PrintColor(Background)})";
    private static string PrintColor(Color? color) => color?.ToString() ?? "null";
}