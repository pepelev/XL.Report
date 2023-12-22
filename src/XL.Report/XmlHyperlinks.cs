using System.Globalization;

namespace XL.Report;

public sealed class XmlHyperlinks : Hyperlinks
{
    private sealed record Hyperlink<T>(Range Range, T Target, string? Tooltip);

    private readonly Book.SheetBuilder sheet;
    private readonly List<Hyperlink<string>> urlHyperlinks = new();
    private readonly List<Hyperlink<string>> definedNameHyperlinks = new();
    private readonly List<Hyperlink<Range>> thisSheetRangeHyperlinks = new();
    private readonly List<Hyperlink<SheetRelated<Range>>> rangeHyperlinks = new();

    public XmlHyperlinks(Book.SheetBuilder sheet)
    {
        this.sheet = sheet;
    }

    // todo check range not overlaps
    public override void Add(Range range, string url, string? tooltip = null)
    {
        var hyperlink = new Hyperlink<string>(range, url, tooltip);
        urlHyperlinks.Add(hyperlink);
    }

    public override void AddToDefinedName(Range range, string name, string? tooltip = null)
    {
        var hyperlink = new Hyperlink<string>(range, name, tooltip);
        definedNameHyperlinks.Add(hyperlink);
    }

    public override void AddToRange(Range range, Range target, string? tooltip = null)
    {
        var hyperlink = new Hyperlink<Range>(range, target, tooltip);
        thisSheetRangeHyperlinks.Add(hyperlink);
    }

    public override void AddToRange(Range range, SheetRelated<Range> target, string? tooltip = null)
    {
        var hyperlink = new Hyperlink<SheetRelated<Range>>(range, target, tooltip);
        rangeHyperlinks.Add(hyperlink);
    }

    public void WriteSheetPart(Xml xml)
    {
        var hasLinks =
            urlHyperlinks.Count > 0 ||
            definedNameHyperlinks.Count > 0 ||
            thisSheetRangeHyperlinks.Count > 0 ||
            rangeHyperlinks.Count > 0;

        if (!hasLinks)
        {
            return;
        }

        using (xml.WriteStartElement("hyperlinks"))
        {
            foreach (var (range, target, tooltip) in thisSheetRangeHyperlinks)
            {
                using (xml.WriteStartElement("hyperlink"))
                {
                    xml.WriteAttribute("ref", range);
                    xml.WriteAttribute("location", new SheetRelated<Range>(sheet.Name, target).ToFormattable());
                    if (tooltip != null)
                    {
                        xml.WriteAttribute("tooltip", tooltip);
                    }
                }
            }

            foreach (var (range, target, tooltip) in rangeHyperlinks)
            {
                using (xml.WriteStartElement("hyperlink"))
                {
                    xml.WriteAttribute("ref", range);
                    xml.WriteAttribute("location", target.ToFormattable());
                    if (tooltip != null)
                    {
                        xml.WriteAttribute("tooltip", tooltip);
                    }
                }
            }

            foreach (var (range, target, tooltip) in definedNameHyperlinks)
            {
                using (xml.WriteStartElement("hyperlink"))
                {
                    xml.WriteAttribute("ref", range);
                    xml.WriteAttribute("location", target);
                    if (tooltip != null)
                    {
                        xml.WriteAttribute("tooltip", tooltip);
                    }
                }
            }

            for (var i = 0; i < urlHyperlinks.Count; i++)
            {
                var (range, _, tooltip) = urlHyperlinks[i];
                using (xml.WriteStartElement("hyperlink"))
                {
                    xml.WriteAttribute("ref", range);

                    // todo synchronize rId here and in WriteRelsPart more explicit
                    // todo hyperlinks are not the only thing that resides in xl/worksheets/_rels/sheetX.xml.rels
                    xml.WriteAttribute("r", "id", $"rId{(i + 1).ToString(CultureInfo.InvariantCulture)}");
                    if (tooltip != null)
                    {
                        xml.WriteAttribute("tooltip", tooltip);
                    }
                }
            }
        }
    }

    public bool RequireRelsPart => urlHyperlinks.Count > 0;

    public void WriteRelsPart(Xml xml)
    {
        // using (xml.WriteStartDocument("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships"))
        for (var i = 0; i < urlHyperlinks.Count; i++)
        {
            var hyperlink = urlHyperlinks[i];
            using (xml.WriteStartElement("Relationship"))
            {
                xml.WriteAttribute("Id", $"rId{(i + 1).ToString(CultureInfo.InvariantCulture)}");
                xml.WriteAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
                xml.WriteAttribute("Target", hyperlink.Target);
                xml.WriteAttribute("TargetMode", "External");
            }
        }
    }
}