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