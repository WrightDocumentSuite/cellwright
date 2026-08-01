# Benchmarks

Scenario: 100,000 rows × 20 columns (2 million cells) — numbers, repeated strings, dates, and
booleans. Each library uses its idiomatic per-cell API. BenchmarkDotNet 0.15.8,
RunStrategy.Monitoring (1 warmup plus 3 or 5 iterations depending on the class), MemoryDiagnoser.

Every table below comes from one measurement session, 2026-08-01 — mixing runs is what makes a
benchmark page disagree with itself. Environment: Intel Core i7-8700, Windows Server 2022,
.NET SDK 10.0.302 on .NET 10.0.10 (X64 RyuJIT x86-64-v3). Measurements were taken on a machine under
background load, and the confidence intervals show it: for several rows the reported error exceeds
the mean. Absolute numbers are therefore conservative and the relative ratios are the meaningful
part. Where a mean is dominated by one slow iteration the table says so and reports the median
instead. Versions: EPPlus 8.6.3, ClosedXML 0.105.0, ExcelDataReader 3.7.0.

## Write (2 million cells to MemoryStream)

| Method | Time (mean) | Allocations | Gen0/Gen1/Gen2 |
| --- | ---: | ---: | --- |
| **Cellwright streaming** | **2.49 s** | **32 MB** | 0 / 0 / 0 |
| **Cellwright DOM** | 2.51 s | 78 MB | 1k / 0 / 0 |
| ClosedXML | 8.91 s | 2,095 MB | 25k / 6k / 3k |
| EPPlus 8 | 10.73 s | 2,142 MB | 33k / 14k / 9k |

The streaming writer is 3.6× faster than ClosedXML and 4.3× faster than EPPlus, with **65×**
fewer allocations (64.7× and 66.2× respectively) — no Gen2 collections at all, and no Gen0 either.
The DOM path is almost as fast: data lives in column slices, and save runs through the same
streaming writer.

## Read (all cells, checksum)

| Method | Time (mean) | Allocations | Gen0/Gen1/Gen2 |
| --- | ---: | ---: | --- |
| **Cellwright streaming** | **2.99 s** | **607 MB** | 241k / 0 / 0 |
| **Cellwright DOM** | 3.08 s | 655 MB | 77k / 5k / 2k |
| EPPlus 8 | 5.20 s | 1,782 MB | 37k / 14k / 8k |
| ClosedXML | 11.86 s | 2,634 MB | 23k / 10k / 5k |

Read is 1.7× faster than EPPlus and 4.0× faster than ClosedXML. The gap to EPPlus is much narrower
than the write gap: reading is dominated by producing strings, which every engine has to do.
Read allocations are mostly text cell strings (unavoidable with a string API) plus `XmlReader`
buffers; further optimization (a custom UTF-8 XML parser instead of `XmlReader`) is a v1.1
candidate. Note how the allocation profile differs from the competition's: 607 MB in short-lived
Gen0 garbage that never survives a collection, against 1.8–2.6 GB that repeatedly reaches Gen2.

## Legacy .xls: read and write (50,000 rows × 4 columns)

A binary workbook written by this same package so both readers see exactly the same file. The old
grid stops at 65,536 rows, so the scenario is smaller than the xlsx one.

| Method | Time (median) | Allocations |
| --- | ---: | ---: |
| **Cellwright: read to DOM** | **28.9 ms** | **18.1 MB** |
| **Cellwright: row streaming** | 56.4 ms | 17.0 MB |
| ExcelDataReader 3.7: read | 193.5 ms | 96.1 MB |
| **Cellwright: write** | 116.9 ms | 26.4 MB |

Read allocations are 5.3× lower than ExcelDataReader: records are parsed as spans directly from
the stream buffer, and values land in column slices without a per-cell object. This is the noisiest
table on the page — the reported error is larger than the mean for every row, which is why it gives
medians; allocations are deterministic and can be compared directly. There is no third-party
baseline for write: ExcelDataReader is read-only.

Write costs more than it did before 0.1.0 — 26.4 MB against 10.6 MB, and 117 ms against 83 ms, on
the same workbook. The writer grew by some 1,900 lines over that period, gaining page setup,
property-set streams, data validation, conditional formatting and the Office Art layer. The
conditional-format, validation and drawing paths all return early on a sheet that has none of
those, so they are not the cause; the extra cost has not been attributed to a specific record
writer and is worth a profile before 1.0.

Streaming is not faster than DOM read and should not be: it saves not time but what stays in memory
afterward. Rows are yielded one at a time, but the container and row table are still read whole —
that is the nature of the format; see [xls.md](xls.md).

## Formulas: full recalculation (50,000 rows, 100,000 formulas)

A report workbook: two columns of constants, `C=A*B*VLOOKUP(...)` against a 100-row lookup table,
a nested `IF` in column D, and summary `SUM`/`COUNTIF`. Each engine opens the same file and
recalculates from scratch.

| Method | Time (mean) | Allocations |
| --- | ---: | ---: |
| **Cellwright** | **1.60 s** | 380 MB |
| EPPlus 8 | 4.54 s | not reported |
| ClosedXML 0.105 | 142.91 s | 1,067 MB |

2.8 times faster than EPPlus and 89 times faster than ClosedXML. Most of that time goes to reading
the package, not calculating: a pure recalc of the same workbook takes 122 ms (see the next
section), so the file read accounts for roughly nine tenths of the figure above.

Two caveats on this table. MemoryDiagnoser reported no allocation total for EPPlus in this session
even though it collected Gen0 45,000 times, so the cell says what was measured rather than a number
carried over from an earlier run. And ClosedXML's own spread here is enormous — ±75 s of error
around a 143 s mean — so the 89× is a floor, not a precise ratio: an earlier session on the same
machine put the same comparison at 125×.

## Formulas: pure recalculation (no file read)

The same report workbook, but only `Calculate` is measured: reading the package takes longer than
the calculation itself and hides in the combined number.

| Scenario | Time (mean) | Allocations |
| --- | ---: | ---: |
| Lookup table on a separate sheet | 121.6 ms | 13.73 MB |
| Lookup table on the same sheet as the results | 123.0 ms | 13.73 MB |

One hundred thousand formulas recalculate in 122 ms. The match between the two rows is the main
point — 1% apart, on identical allocations: in the second case every computed result lands on the
same sheet as the lookup table, and coarse sheet-version invalidation would flush the cache on
every row. Rectangle-based invalidation does not.

Before the range cache, each of 50,000 `VLOOKUP` calls copied the whole table: the same scenario
with load cost 557 MB of allocations versus 13.7 MB for recalc today.

## Formulas: range subscription lookup

Editing a cell must find formulas that read it through a range. The benchmark isolates the lookup
structure: 50,000 subscriptions of the form `SUM(A$1:A{n})`, none repeated.

| Method | Time (mean) |
| --- | ---: |
| Interval index | 57.5 ns |
| Linear scan of the same subscriptions | 105.2 µs |

An 1,835× difference. Neither variant allocates; only the number of subscriptions touched to
answer changes.

At workbook level it looks like this:

| Scenario (50,000 range formulas) | Time (mean) | Allocations |
| --- | ---: | ---: |
| Edit outside all subscribed ranges | 72.2 µs | 144 B |
| Edit found among 50,000 distinct ranges | 781 µs | 496 B |
| Edit in a table read by 50,000 formulas | 72.2 ms | 11.4 MB |

The last row is not about lookup but about work volume: editing the lookup table dirties all fifty
thousand dependent formulas, and they genuinely need recalculation.

## Formulas: incremental recalculation

The workbook is already calculated; one cell changes. The dependency graph exists so edit cost
follows the size of the affected subgraph, not the size of the workbook.

| Scenario | Time (median) | Allocations |
| --- | ---: | ---: |
| Edit with a chain of 10,000 dependent formulas behind it | 6.3 ms | 1.1 MB |
| Edit of one of 50,000 independent formulas | **96.8 µs** | 496 B |
| Forced full recalc of the same 50,000 formulas | 33.3 ms | 2.4 MB |

Medians again: the forced recalc row measured 54 ms on average with ±36 ms of error, one slow
iteration against two fast ones, so the mean says more about the machine than about the code.

A local edit costs roughly 340× less than a full recalc and almost no allocations. The first
version of traversal cost 5.5 ms for the same edit because it scanned the entire formula list;
reverse edges in the index and breadth-first traversal made cost proportional to the affected
subgraph. The 10,000-formula chain case cannot get cheaper: all 10,000 genuinely need
recalculation.

## Formulas: parsing

10,000 formulas of seven typical shapes, parsed into an RPN program.

| Scenario | Time (mean) | Allocations |
| --- | ---: | ---: |
| 10k distinct formulas | 17.6 ms | 16.9 MB |
| 10k formulas via `ProgramCache` (7 unique) | 0.44 ms | 12.6 KB |
| 10k formulas: parse plus reverse render to A1 | 31.4 ms | 24.5 MB |

That is about 570,000 formulas per second on a cold parse. Real sheets repeat one formula down a
column, and the program cache applies: the same 10,000 cells cost 40× less and three orders of
magnitude fewer allocations.

There is no third-party parsing baseline on purpose: the parser shipped with ClosedXML exposes a
visitor interface with two dozen members, and a comparison would mostly measure an adapter written
for it.

## How to reproduce

```bash
dotnet build Cellwright.slnx -c Release
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks --no-build -- --filter *ReadBenchmarks*
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks --no-build -- --filter *WriteBenchmarks*
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks --no-build -- --filter *FormulaParseBenchmarks*
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks --no-build -- --filter *FormulaCalcBenchmarks*
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks --no-build -- --filter *FormulaRecalcBenchmarks*
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks --no-build -- --filter *IncrementalCalcBenchmarks*
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks --no-build -- --filter *RangeEditBenchmarks*
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks --no-build -- --filter *SubscriptionLookupBenchmarks*
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks --no-build -- --filter *XlsBenchmarks*
```

`FormulaCalcBenchmarks` dominates the wall clock: ClosedXML spends two and a half minutes per
iteration there, so that one command alone runs for about ten minutes and the whole list for
something under half an hour.

Run the whole list before editing this page. Every table here is from one session on purpose — the
machine's noise is large enough that numbers taken weeks apart cannot be compared with each other,
and a page that mixes them ends up claiming ratios no single run ever produced.

Run each benchmark class with a separate command: BenchmarkSwitcher cleans generated artifacts
between sessions, and a single `--filter *` breaks the second session.
