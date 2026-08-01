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
}
