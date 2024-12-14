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

public sealed class NoFill : Fill, IEquatable<NoFill>
{
    public bool Equals(NoFill? other)
    {
        return other != null;
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as NoFill);
    }

    public override int GetHashCode()
    {
        return 1299827;
    }

    public override string ToString()
    {
        return "No";
    }

    public override T Accept<T>(Visitor<T> visitor)
    {
        return visitor.Visit(this);
    }

    public override void Write(Xml xml)
    {
        using (xml.WriteStartElement(XlsxStructure.Styles.Fills.Fill))
        {
            using (xml.WriteStartElement(XlsxStructure.Styles.Fills.Pattern))
            {
                xml.WriteAttribute(XlsxStructure.Styles.Fills.PatternType, "none");
            }
        }
    }
}