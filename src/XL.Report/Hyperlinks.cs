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

public abstract class Hyperlinks
{
    public abstract void Add(ValidRange range, string url, string? tooltip = null);
    public abstract void AddToDefinedName(ValidRange range, string name, string? tooltip = null);
    public abstract void AddToRange(ValidRange range, ValidRange target, string? tooltip = null);
    public abstract void AddToRange(ValidRange range, SheetRelated<ValidRange> target, string? tooltip = null);
}