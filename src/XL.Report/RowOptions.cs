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

using XL.Report.Styles;

namespace XL.Report;

public sealed record RowOptions(float? Height = null, StyleId? StyleId = null, bool Hidden = false)
{
    public static RowOptions Default { get; } = new();

    public void WriteAttributes(Xml xml)
    {
        if (Height is { } height)
        {
            xml.WriteAttribute("ht", height, "N6");
            xml.WriteAttribute("customHeight", "true");
        }

        if (StyleId is { } styleId)
        {
            xml.WriteAttribute("s", styleId);
            xml.WriteAttribute("customFormat", "true");
        }

        if (Hidden)
        {
            xml.WriteAttribute("hidden", "true");
        }
    }
}