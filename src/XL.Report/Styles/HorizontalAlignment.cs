namespace XL.Report.Styles;

public abstract class HorizontalAlignment
{
    public static HorizontalAlignment ByContent { get; } = new ByContentAlignment();

    public static HorizontalAlignment Left(int indent = 0) => new Plain(
        "left",
        indent,
        supportTextRotation: true
    );

    public static HorizontalAlignment Center { get; } = new Plain(
        "center",
        0,
        supportTextRotation: true
    );

    public static HorizontalAlignment CenterContinuous { get; } = new Plain(
        "centerContinuous",
        0,
        supportTextRotation: false
    );

    public static HorizontalAlignment Right(int indent = 0) => new Plain(
        "right",
        indent,
        supportTextRotation: true
    );

    public static HorizontalAlignment Fill { get; } = new Plain(
        "fill",
        0,
        supportTextRotation: false
    );

    public static HorizontalAlignment Justify { get; } = new Plain(
        "justify",
        0,
        supportTextRotation: true
    );

    public static HorizontalAlignment Distributed(int indent = 0) => new Plain(
        "distributed",
        indent,
        supportTextRotation: true
    );

    public static HorizontalAlignment DistributedJustifyLastLine { get; } = new DistributedJustifyLastLineAlignment();

    public abstract bool IsDefault { get; }
    public abstract bool SupportTextRotation { get; }
    public abstract void Write(Xml xml);

    private sealed class Plain : HorizontalAlignment
    {
        private readonly string style;
        private readonly int indent;

        public Plain(string style, int indent, bool supportTextRotation)
        {
            if (this.indent < 0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(indent),
                    indent,
                    "must be non-negative"
                );
            }

            this.style = style;
            this.indent = indent;
            SupportTextRotation = supportTextRotation;
        }

        public override bool IsDefault => false;
        public override bool SupportTextRotation { get; }

        public override void Write(Xml xml)
        {
            xml.WriteAttribute("horizontal", style);
            if (indent > 0)
            {
                xml.WriteAttribute("indent", indent);
            }
        }

        public override string ToString() => indent > 0
            ? $"{style}, indent = {indent}"
            : style;

        private bool Equals(Plain other) => style == other.style &&
                                            indent == other.indent &&
                                            SupportTextRotation == other.SupportTextRotation;

        public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is Plain other && Equals(other);
        public override int GetHashCode() => HashCode.Combine(style, indent, SupportTextRotation);
    }

    private sealed class ByContentAlignment : HorizontalAlignment
    {
        public override bool IsDefault => true;
        public override bool SupportTextRotation => true;

        public override void Write(Xml xml)
        {
        }

        public override string ToString() => "by content";
        public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is ByContentAlignment;
        public override int GetHashCode() => 319;
    }

    private sealed class DistributedJustifyLastLineAlignment : HorizontalAlignment
    {
        public override bool IsDefault => false;
        public override bool SupportTextRotation => true;

        public override void Write(Xml xml)
        {
            xml.WriteAttribute("horizontal", "distributed");
            xml.WriteAttribute("justifyLastLine", "1");
        }

        public override bool Equals(object? obj) =>
            ReferenceEquals(this, obj) || obj is DistributedJustifyLastLineAlignment;

        public override int GetHashCode() => 317;
    }
}