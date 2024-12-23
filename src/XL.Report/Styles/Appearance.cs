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

using XL.Report.Styles.Fills;

namespace XL.Report.Styles;

public sealed record Appearance(
    Alignment Alignment,
    Font Font,
    Fill Fill,
    Borders Borders
)
{
    public static Appearance Default { get; } = new(
        Alignment.Default,
        Font.Calibri11,
        Fill.No,
        Borders.None
    );
}