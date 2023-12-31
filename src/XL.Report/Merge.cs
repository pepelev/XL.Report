﻿using XL.Report.Styles;

namespace XL.Report;

public readonly struct Merge : IUnit<Range>
{
    private readonly Content content;
    private readonly Size size;
    private readonly StyleId? styleId;

    public Merge(Content content, Size size, StyleId? styleId = null)
    {
        this.content = content;
        this.size = size;
        this.styleId = styleId;
    }

    public Range Write(SheetWindow window)
    {
        window.Merge(size, content, styleId);
        return new Range(window.Range.LeftTop, size);
    }

    public readonly struct All : IUnit<Range>
    {
        private readonly Content content;
        private readonly StyleId? styleId;

        public All(Content content, StyleId? styleId = null)
        {
            this.content = content;
            this.styleId = styleId;
        }

        public Range Write(SheetWindow window)
        {
            window.Merge(window.Range.Size, content, styleId);
            return window.Range;
        }
    }
}