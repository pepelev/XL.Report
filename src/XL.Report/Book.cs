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

public abstract class Book : IDisposable
{
    public abstract Style.Collection Styles { get; }
    public abstract SharedStrings Strings { get; }

    // todo check name constraints, uniq
    public abstract SheetBuilder CreateSheet(string name, SheetOptions options);

    public abstract void Complete();

    public abstract class SheetBuilder : IDisposable
    {
        public abstract string Name { get; set; }
        public abstract void Dispose();
        public abstract T WriteRow<T>(IUnit<T> unit);
        public abstract T WriteRow<T>(IUnit<T> unit, RowOptions options);
        public abstract void DefineName(string name, ValidRange range, string? comment = null);
        public abstract Hyperlinks Hyperlinks { get; }
        public abstract void AddConditionalFormatting(ConditionalFormatting formatting);
        public abstract void Complete();
    }

    public abstract void Dispose();
}