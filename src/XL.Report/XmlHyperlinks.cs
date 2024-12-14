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
    public override void Add(ValidRange range, string url, string? tooltip = null)
    {
        var hyperlink = new Hyperlink<string>(range, url, tooltip);
        urlHyperlinks.Add(hyperlink);
    }

    public override void AddToDefinedName(ValidRange range, string name, string? tooltip = null)
    {
        var hyperlink = new Hyperlink<string>(range, name, tooltip);
        definedNameHyperlinks.Add(hyperlink);
    }

    public override void AddToRange(ValidRange range, ValidRange target, string? tooltip = null)
    {
        var hyperlink = new Hyperlink<Range>(range, target, tooltip);
        thisSheetRangeHyperlinks.Add(hyperlink);
    }

    public override void AddToRange(ValidRange range, SheetRelated<ValidRange> target, string? tooltip = null)
    {
        var simplifiedTarget = new SheetRelated<Range>(target.SheetName, target.Value);
        var hyperlink = new Hyperlink<SheetRelated<Range>>(range, simplifiedTarget, tooltip);
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