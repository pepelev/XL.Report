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

public sealed record SheetOptions(FreezeOptions Freeze, ColumnOptions.Collection Columns)
{
    public static SheetOptions Default { get; } = new(
        FreezeOptions.None,
        ColumnOptions.Collection.Default
    );

    public SheetOptions With(int x, ColumnOptions options) =>
        this with { Columns = Columns.With(x, options) };

    public void Write(Xml xml)
    {
        Freeze.WriteAsSingleSheetView(xml);
        Columns.Write(xml);
    }
}