using XL.Report.Styles;

namespace XL.Report;

internal sealed partial class StreamSheetWindow
{
    private struct Row
    {
        private readonly Dictionary<int, Cell> cells = new();
        private RowUsage usage;

        public Row()
        {
        }

        public void Place(int x, Cell cell)
        {
            if (!usage.TryMark(new Interval<int>(x, x)))
            {
                throw new InvalidOperationException();
            }

            cells[x] = cell;
        }

        public void Merge(bool isMain, Interval<int> range, Content content, StyleId? styleId)
        {
            if (!usage.TryMark(range))
            {
                throw new InvalidOperationException();
            }

            if (isMain)
            {
                cells[range.LeftInclusive] = new Cell(content, styleId);
            }
        }

        public int CellsCount => cells.Count;

        public void CopyCellsTo(Span<KeyValuePair<int, Cell>> destination)
        {
            var index = 0;
            foreach (var pair in cells)
            {
                destination[index++] = pair;
            }
        }
    }
}