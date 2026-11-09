# Cellwright runnable examples

Run from the repository root with the SDK pinned in `global.json`:

```sh
dotnet run -c Release --project samples/Cellwright.Samples
```

Cellwright writes into `cellwright-samples` under `Path.GetTempPath()` and prints that path at startup.

The existing examples run first, followed by these self-contained scenarios. Each new scenario creates its inputs, reloads or checks its output, and throws if the demonstrated result changes. All captions and messages use English.

| Example | Scenario | Output |
| --- | --- | --- |
| [CsvRoundTrip](Cellwright.Samples/CsvRoundTrip.cs) | Write fields containing commas and quotes, convert to XLSX, and read exported CSV incrementally. | `sales-input.csv; sales-from-csv.xlsx; sales-export.csv` |
| [ChartReport](Cellwright.Samples/ChartReport.cs) | Create a quarterly revenue chart, save it, then verify the saved series and title. | `chart-report.xlsx` |
| [NamedRangeBudget](Cellwright.Samples/NamedRangeBudget.cs) | Combine a named range, input validation and formula recalculation; save calculated values in the workbook. | `named-range-budget.xlsx` |

Each source file exposes `RunAsync(directory)` and can also be called independently. See the [getting-started guide](../docs/getting-started.md) for package installation. The same new method bodies are shown in the [website playground](https://wrightsuite.com/#code-playground).

CSV is a value exchange format and does not preserve workbook styling or formulas. The chart uses native worksheet references and cached plotted values. The budget example calculates before saving so readers can display cached totals immediately.
