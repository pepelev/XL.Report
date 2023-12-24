namespace XL.Report.Styles;

public abstract class VerticalAlignment
{
    public static VerticalAlignment Top { get; } = new Plain("top", supportShrinkOnOverflow: true);
    public static VerticalAlignment Center { get; } = new Plain("center", supportShrinkOnOverflow: true);
    public static VerticalAlignment Bottom { get; } = new BottomAlignment();
    public static VerticalAlignment Justify { get; } = new Plain("justify", supportShrinkOnOverflow: false);
    public static VerticalAlignment Distributed { get; } = new Plain("distributed", supportShrinkOnOverflow: false);

    public abstract bool IsDefault { get; }
    public abstract bool SupportShrinkOnOverflow { get; }
    public abstract void Write(Xml xml);

    private sealed class BottomAlignment : VerticalAlignment
    {
        public override bool IsDefault => true;
        public override bool SupportShrinkOnOverflow => true;

        public override void Write(Xml xml)
        {
        }

        public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is BottomAlignment;
        public override int GetHashCode() => 323;
    }

    private sealed class Plain : VerticalAlignment
    {
        private readonly string style;

        public Plain(string style, bool supportShrinkOnOverflow)
        {
            this.style = style;
            SupportShrinkOnOverflow = supportShrinkOnOverflow;
        }

        public override bool IsDefault => false;
        public override bool SupportShrinkOnOverflow { get; }

        public override void Write(Xml xml)
        {
            xml.WriteAttribute("vertical", style);
        }

        private bool Equals(Plain other) => 
            style == other.style && SupportShrinkOnOverflow == other.SupportShrinkOnOverflow;

        public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is Plain other && Equals(other);
        public override int GetHashCode() => HashCode.Combine(style, SupportShrinkOnOverflow);
    }
}