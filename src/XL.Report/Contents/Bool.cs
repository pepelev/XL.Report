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

namespace XL.Report.Contents;

public sealed class Bool : Content
{
    private readonly bool content;

    private Bool(bool content)
    {
        this.content = content;
    }

    public static Content True { get; } = new Bool(true);
    public static Content False { get; } = new Bool(false);

    public static Content From(bool value) => value
        ? True
        : False;

    public override void Write(Xml xml)
    {
        xml.WriteAttribute("t", "b");
        using (xml.WriteStartElement("v"))
        {
            xml.WriteValue(content ? 1 : 0);
        }
    }
}