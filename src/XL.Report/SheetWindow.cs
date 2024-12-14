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

using XL.Report.Contents;
using XL.Report.Styles;

namespace XL.Report;

public abstract class SheetWindow
{
    public abstract Range Range { get; }
    public abstract void TouchRow(int y);
    public abstract void Place(Content content, StyleId? styleId);
    public abstract void Merge(Content content, StyleId? styleId);
    public abstract IDisposable Reduce(Reduction reduction);

    public IDisposable Reduce(Offset offset, Size? newSize = null)
    {
        return Reduce(new Reduction(offset, newSize));
    }
}