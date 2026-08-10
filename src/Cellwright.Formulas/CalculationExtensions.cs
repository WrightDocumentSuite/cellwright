using System.Runtime.CompilerServices;
using Cellwright.Formulas.Engine;

namespace Cellwright.Formulas;

/// <summary>
/// The public surface of the formula engine: calculation and ad-hoc evaluation over the
/// Cellwright document model.
/// </summary>
public static class CalculationExtensions
{
    // The engine holds a dependency graph and dirty set per workbook; attaching it here keeps
    // the core free of any reference to the formula package.
    private static readonly ConditionalWeakTable<Workbook, CalculationEngine> Engines = [];

    extension(Workbook workbook)
    {
        /// <summary>
        /// Recalculates the workbook. The first call computes every formula; later calls only
        /// redo what changed, plus volatile formulas and anything downstream of them.
        /// </summary>
        public CalculationSummary Calculate(CalculateOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(workbook);
            return GetEngine(workbook).Calculate(options ?? CalculateOptions.Default);
        }

        /// <summary>
        /// Evaluates a formula that lives nowhere, such as <c>"SUM(Sales!B2:B100)*1.2"</c>.
        /// Relative references resolve against the first cell of the first sheet.
        /// </summary>
        /// <exception cref="FormulaParseException">The formula is malformed.</exception>
        public CellValue Evaluate(string formula)
        {
            ArgumentNullException.ThrowIfNull(workbook);
            ArgumentException.ThrowIfNullOrWhiteSpace(formula);
            if (workbook.Sheets.Count == 0)
                throw new InvalidOperationException("The workbook has no sheets to evaluate against.");

            return GetEngine(workbook).EvaluateAdHoc(formula, workbook.Sheets[0], new CellRef(1, 1));
        }
    }

    extension(Worksheet sheet)
    {
        /// <summary>Evaluates a formula as if it sat in <paramref name="originCell"/> of this sheet.</summary>
        /// <exception cref="FormulaParseException">The formula is malformed.</exception>
        public CellValue Evaluate(string formula, string originCell = "A1")
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentException.ThrowIfNullOrWhiteSpace(formula);
            return GetEngine(sheet.Workbook).EvaluateAdHoc(formula, sheet, CellRef.Parse(originCell));
        }

        /// <summary>
        /// Evaluates <paramref name="formula"/> as if it had been copied from
        /// <paramref name="sourceCell"/> to <paramref name="targetCell"/>. Relative reference
        /// axes move with the copy; absolute axes stay fixed.
        /// </summary>
        /// <remarks>
        /// The original text is compiled and cached once regardless of how many target cells are
        /// evaluated. This is the suitable callback for conditional-format formulas over large
        /// ranges.
        /// </remarks>
        public CellValue EvaluateCopied(
            string formula, CellRef sourceCell, CellRef targetCell)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentException.ThrowIfNullOrWhiteSpace(formula);
            if (!sourceCell.IsValid)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(sourceCell), "The source cell must lie inside the worksheet grid.");
            }
            if (!targetCell.IsValid)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(targetCell), "The target cell must lie inside the worksheet grid.");
            }

            return GetEngine(sheet.Workbook).EvaluateAdHocCopied(
                formula, sheet, sourceCell, targetCell);
        }
    }

    extension(Cell cell)
    {
        /// <summary>
        /// The formula's value, recalculating the workbook first when anything is stale.
        /// </summary>
        public CellValue CalculatedValue
        {
            get
            {
                Workbook workbook = cell.Sheet.Workbook;
                if (!workbook.IsFullyCalculated)
                    workbook.Calculate();
                return cell.Value;
            }
        }
    }

    private static CalculationEngine GetEngine(Workbook workbook) =>
        Engines.GetValue(workbook, static book => new CalculationEngine(book));

    /// <summary>The workbook's range cache, so tests and benchmarks can see hits and misses.</summary>
    internal static RangeCache RangeCacheOf(Workbook workbook) => GetEngine(workbook).Ranges;

    /// <summary>Compiled program count, exposed internally for cache-regression tests.</summary>
    internal static int ProgramCountOf(Workbook workbook) => GetEngine(workbook).ProgramCount;
}
