namespace XL.Report;

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

    public sealed class Integral : Content
    {
        private readonly long content;

        public Integral(long content)
        {
            this.content = content;
        }

        public override void Write(Xml xml)
        {
            Number.Write(xml, content, "0");
        }
    }

    public sealed class Fractional : Content
    {
        private readonly double content;

        public Fractional(double content)
        {
            this.content = content;
        }

        public override void Write(Xml xml)
        {
            Write(xml, content);
        }

        internal static void Write(Xml xml, double content)
        {
            Number.Write(xml, content, "0.##############");
        }
    }

    public sealed class Financial : Content
    {
        private readonly decimal content;

        public Financial(decimal content)
        {
            this.content = content;
        }

        public override void Write(Xml xml)
        {
            Number.Write(xml, content, "0.############################");
        }
    }

    public sealed class Instant : Content
    {
        private static readonly DateTime offset = new(1899, 12, 30);

        private readonly DateTime content;

        public Instant(DateTime content)
        {
            this.content = content;
        }

        public override void Write(Xml xml)
        {
            var dateOffset = (content.Date - offset).Ticks / TimeSpan.TicksPerDay;
            var timeOfDay = content.TimeOfDay.Ticks / (double)TimeSpan.TicksPerDay;
            var sum = dateOffset + timeOfDay;
            Fractional.Write(xml, sum);
        }
    }
}