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

using System.Text;

namespace XL.Report.Styles;

public readonly partial record struct FontStyle(
    bool IsBold = false,
    bool IsItalic = false,
    bool IsStrikethrough = false,
    Underline? Underline = null,
    FontVerticalAlignment Alignment = FontVerticalAlignment.Regular
)
{
    public static FontStyle Regular => new();
    public static FontStyle Bold => new(IsBold: true);
    public static FontStyle Italic => new(IsItalic: true);
    public static FontStyle Strikethrough => new(IsStrikethrough: true);
    public static FontStyle Underlined(Underline underline = Styles.Underline.Single) => new(Underline: underline);
    public static FontStyle Superscript => new(Alignment: FontVerticalAlignment.Superscript);
    public static FontStyle Subscript => new(Alignment: FontVerticalAlignment.Subscript);

    public override string ToString()
    {
        if (Equals(Regular))
            return nameof(Regular);
        if (Equals(Bold))
            return nameof(Bold);
        if (Equals(Italic))
            return nameof(Italic);
        if (Equals(Strikethrough))
            return nameof(Strikethrough);

        var builder = new StringBuilder(64);
        var isFirst = true;

        if (IsBold)
        {
            builder.Append("Bold");
            isFirst = false;
        }

        if (IsItalic)
        {
            if (!isFirst)
                builder.Append(", ");
            builder.Append("Italic");
            isFirst = false;
        }

        if (IsStrikethrough)
        {
            if (!isFirst)
                builder.Append(", ");
            builder.Append("Strikethrough");
            isFirst = false;
        }

        if (Underline is { } underline)
        {
            if (!isFirst)
                builder.Append(", ");
            builder.Append("Underline: ");
            builder.Append(underline);
        }

        if (Alignment != FontVerticalAlignment.Regular)
        {
            if (!isFirst)
                builder.Append(", ");
            builder.Append("Alignment: ");
            builder.Append(Alignment);
        }

        return builder.ToString();
    }
}