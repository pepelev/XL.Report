﻿#region Legal
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

public sealed class String : Content
{
    private readonly string content;
    private readonly SharedStrings sharedStrings;

    public String(string content, SharedStrings sharedStrings)
    {
        this.content = content;
        this.sharedStrings = sharedStrings;
    }

    public override void Write(Xml xml)
    {
        if (sharedStrings.TryRegister(content) is { } id)
        {
            SharedString.ById.Write(xml, id);
        }
        else
        {
            InlineString.Write(xml, content);
        }
    }
}