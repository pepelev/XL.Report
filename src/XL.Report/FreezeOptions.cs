namespace XL.Report;

public sealed class FreezeOptions : IEquatable<FreezeOptions>
{
    public static FreezeOptions None { get; } = new(0, 0);

    public FreezeOptions(int columns, int rows)
    {
        // todo may exceed sheet size
        if (columns < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(rows), columns, "must be non-negative");
        }

        if (rows < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(rows), rows, "must be non-negative");
        }

        Columns = columns;
        Rows = rows;
    }

    public int Columns { get; }
    public int Rows { get; }

    public void WriteAsSingleSheetView(Xml xml)
    {
        if (this == None)
        {
            return;
        }

        using (xml.WriteStartElement("sheetViews"))
        {
            using (xml.WriteStartElement("sheetView"))
            {
                xml.WriteAttribute("workbookViewId", "0");
                using (xml.WriteStartElement("pane"))
                {
                    if (Columns > 0)
                    {
                        xml.WriteAttribute("xSplit", Columns);
                    }

                    if (Rows > 0)
                    {
                        xml.WriteAttribute("ySplit", Rows);
                    }

                    var topLeftCell = new Location(Columns + 1, Rows + 1);
                    xml.WriteAttribute("topLeftCell", topLeftCell);
                    xml.WriteAttribute("state", "frozen");
                }
            }
        }
    }

    public bool Equals(FreezeOptions? other)
    {
        if (ReferenceEquals(null, other)) return false;
        if (ReferenceEquals(this, other)) return true;
        return Columns == other.Columns && Rows == other.Rows;
    }

    public override bool Equals(object? obj)
    {
        return ReferenceEquals(this, obj) || obj is FreezeOptions other && Equals(other);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Columns, Rows);
    }

    public static bool operator ==(FreezeOptions? left, FreezeOptions? right)
    {
        return Equals(left, right);
    }

    public static bool operator !=(FreezeOptions? left, FreezeOptions? right)
    {
        return !Equals(left, right);
    }
}