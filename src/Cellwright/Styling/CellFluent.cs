namespace Cellwright;

/// <summary>
/// Fluent styling for <see cref="Cell"/> and <see cref="Range"/> facades (C# 14 extension
/// members). Transforms read the current style record, modify it immutably and re-intern —
/// per-range transforms memoize by source style so bulk restyling stays cheap.
/// </summary>
public static class CellFluent
{
    /// <summary>Extension members for single cells.</summary>
    extension(Cell cell)
    {
        /// <summary>Whether the cell carries a formula.</summary>
        public bool HasFormula => cell.Formula is not null;

        /// <summary>Applies an immutable transform to the cell's style.</summary>
        public Cell Style(Func<CellStyle, CellStyle> transform)
        {
            ArgumentNullException.ThrowIfNull(transform);
            return cell.SetStyle(transform(cell.StyleValue));
        }

        /// <summary>Makes the cell bold.</summary>
        public Cell Bold() => cell.Style(static s => s.Bold());

        /// <summary>Makes the cell italic.</summary>
        public Cell Italic() => cell.Style(static s => s.Italic());

        /// <summary>Applies a solid fill.</summary>
        public Cell Fill(ExcelColor color) => cell.Style(s => s.Fill(color));

        /// <summary>Applies a number format code.</summary>
        public Cell NumberFormat(string code) => cell.Style(s => s.NumberFormat(code));
    }

    /// <summary>Extension members for ranges.</summary>
    extension(CellRange range)
    {
        /// <summary>
        /// Applies an immutable transform to every cell's style. Distinct source styles are
        /// transformed once and the results reused across the range.
        /// </summary>
        public CellRange Style(Func<CellStyle, CellStyle> transform)
        {
            ArgumentNullException.ThrowIfNull(transform);
            Workbook workbook = range.Sheet.Workbook;
            var memo = new Dictionary<StyleId, StyleId>();
            foreach (Cell cell in range)
            {
                StyleId current = cell.Style;
                if (!memo.TryGetValue(current, out StyleId mapped))
                {
                    mapped = workbook.RegisterStyle(transform(workbook.GetStyle(current)));
                    memo[current] = mapped;
                }

                cell.SetStyle(mapped);
            }

            return range;
        }

        /// <summary>Makes the range bold.</summary>
        public CellRange Bold() => range.Style(static s => s.Bold());

        /// <summary>Applies a solid fill to the range.</summary>
        public CellRange Fill(ExcelColor color) => range.Style(s => s.Fill(color));

        /// <summary>Applies a number format code to the range.</summary>
        public CellRange NumberFormat(string code) => range.Style(s => s.NumberFormat(code));
    }
}
