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
    public abstract partial class Condition
    {
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
    }
}