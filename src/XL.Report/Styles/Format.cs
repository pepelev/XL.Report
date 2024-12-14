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

namespace XL.Report.Styles;

public sealed class Format : IEquatable<Format>
{
    public Format(string code)
    {
        Code = code ?? throw new ArgumentNullException(nameof(code));
    }

    public string Code { get; }

    public bool Equals(Format? other)
    {
        if (ReferenceEquals(this, other))
            return true;

        if (ReferenceEquals(null, other))
            return false;

        return string.Equals(Code, other.Code);
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as Format);
    }

    public override int GetHashCode()
    {
        return Code.GetHashCode();
    }

    public void Write(Xml xml, int index)
    {
        using (xml.WriteStartElement("numFmt"))
        {
            xml.WriteAttribute("numFmtId", index);
            xml.WriteAttribute("formatCode", Code);
        }
    }

    public static void Write(Xml xml, IEnumerable<(Format Format, int Index)> formats)
    {
        using (xml.WriteStartElement("numFmts"))
        {
            foreach (var (format, index) in formats)
            {
                format.Write(xml, index);
            }
        }
    }

    #region Formats

    public static Format General { get; } = new("@");
    public static Format IsoDate { get; } = new("yyyy-mm-dd");
    public static Format IsoTime { get; } = new("hh:mm:ss");
    public static Format IsoDateTime { get; } = new(@"yyyy-mm-dd""T""hh:mm:ss");
    public static Format Integer { get; } = new("0");

    public static Format Float(int decimals = 2)
    {
        if (decimals < 1)
        {
            throw new ArgumentOutOfRangeException(
                nameof(decimals),
                decimals,
                "A number that greater than zero is expected"
            );
        }

        return new Format($"0.{new string('0', decimals)}");
    }

    #endregion
}