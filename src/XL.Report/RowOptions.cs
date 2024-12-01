using XL.Report.Styles;

namespace XL.Report;

public sealed record RowOptions(float? Height = null, StyleId? StyleId = null, bool Hidden = false)
{
    public static RowOptions Default { get; } = new();

    public void WriteAttributes(Xml xml)
    {
        if (Height is { } height)
        {
            xml.WriteAttribute("ht", height, "N6");
            xml.WriteAttribute("customHeight", "true");
        }

        if (StyleId is { } styleId)
        {
            xml.WriteAttribute("s", styleId);
            xml.WriteAttribute("customFormat", "true");
        }

        if (Hidden)
        {
            xml.WriteAttribute("hidden", "true");
        }
    }
}