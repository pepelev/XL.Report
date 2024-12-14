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

namespace XL.Report.Auxiliary;

internal readonly struct ByKey<TKey, TValue> : IComparer<KeyValuePair<TKey, TValue>>
    where TKey : IComparable<TKey>
{
    public int Compare(KeyValuePair<TKey, TValue> x, KeyValuePair<TKey, TValue> y) =>
        Comparer<TKey>.Default.Compare(x.Key, y.Key);
}