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

public static class Number
{
    // todo extract to consts
    private const string value = "v";

    private static void Write<T>(Xml xml, T content, string format)
        where T : ISpanFormattable
    {
        using (xml.WriteStartElement(value))
        {
            xml.WriteValue(content, format);
        }
    }

    public sealed class Integral(long content) : Content
    {
        public override void Write(Xml xml)
        {
            Number.Write(xml, content, "0");
        }
    }

    public sealed class Fractional(double content) : Content
    {
        public override void Write(Xml xml)
        {
            Write(xml, content);
        }

        internal static void Write(Xml xml, double content)
        {
            Number.Write(xml, content, "0.##############");
        }
    }

    public sealed class Financial(decimal content) : Content
    {
        public override void Write(Xml xml)
        {
            Number.Write(xml, content, "0.############################");
        }
    }

    public sealed class Instant(DateTime content) : Content
    {
        private static readonly DateTime Offset = new(1899, 12, 30);

        public override void Write(Xml xml)
        {
            var dateOffset = (content.Date - Offset).Ticks / TimeSpan.TicksPerDay;
            var timeOfDay = content.TimeOfDay.Ticks / (double)TimeSpan.TicksPerDay;
            var sum = dateOffset + timeOfDay;
            Fractional.Write(xml, sum);
        }
    }
}