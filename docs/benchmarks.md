# Benchmarks

Scenario: 100,000 rows × 20 columns (2 million cells) — numbers, repeated strings, dates, and
booleans. Each library uses its idiomatic per-cell API. BenchmarkDotNet 0.15.8,
RunStrategy.Monitoring (1 warmup + 3 iterations), MemoryDiagnoser.

Environment: Intel Core i7-8700, Windows Server, .NET 10.0.10 (X64 RyuJIT x86-64-v3). Measurements
were taken on a machine under background load — absolute numbers are conservative; relative ratios
are the meaningful part. Versions: EPPlus 8.6.3, ClosedXML 0.105.0, ExcelDataReader 3.7.0.

## Write (2 million cells to MemoryStream)

| Method | Time | Allocations | Gen0/Gen1/Gen2 |
| --- | ---: | ---: | --- |
| **Cellwright streaming** | **3.48 s** | **32 MB** | 0 / 0 / 0 |
| **Cellwright DOM** | 3.99 s | 78 MB | 1k / 0 / 0 |
| ClosedXML | 13.31 s | 2,095 MB | 41k / 6k / 4k |
| EPPlus 8 | 16.27 s | 2,142 MB | 41k / 19k / 13k |

The streaming writer is ~3.8× faster than ClosedXML and ~4.7× faster than EPPlus, with **65×**
fewer allocations — no Gen2 collections at all. The DOM path is almost as fast: data lives in
column slices, and save runs through the same streaming writer.

## Read (all cells, checksum)

| Method | Time | Allocations | Gen0/Gen1/Gen2 |
| --- | ---: | ---: | --- |
| **Cellwright streaming** | **3.80 s** | **607 MB** | 123k / 0 / 0 |
| **Cellwright DOM** | 5.59 s | 654 MB | 19k / 5k / 2k |
| EPPlus 8 | 17.09 s | 1,782 MB | 26k / 11k / 6k |
| ClosedXML | 25.49 s | 2,602 MB | 23k / 10k / 5k |

Read is ~4.5× faster than EPPlus and ~6.7× faster than ClosedXML. Read allocations are mostly
text cell strings (unavoidable with a string API) plus `XmlReader` buffers; further optimization (a
custom UTF-8 XML parser instead of `XmlReader`) is a v1.1 candidate.

## Legacy .xls: read and write (50,000 rows × 4 columns)

A binary workbook written by this same package so both readers see exactly the same file. The old
grid stops at 65,536 rows, so the scenario is smaller than the xlsx one.

| Method | Time (median) | Allocations |
| --- | ---: | ---: |
| **Cellwright: read to DOM** | **51 ms** | **17.7 MB** |
| **Cellwright: row streaming** | 128 ms | 16.6 MB |
| ExcelDataReader 3.7: read | 142 ms | 96.1 MB |
| **Cellwright: write** | 83 ms | 10.6 MB |

Read allocations are 5.4× lower than ExcelDataReader: records are parsed as spans directly from
the stream buffer, and values land in column slices without a per-cell object. Timings on this
machine are noisy (three iterations, spread up to ±100 ms); allocations are deterministic. There is
no third-party baseline for write: ExcelDataReader is read-only.

Streaming is not faster than DOM read and should not be: it saves not time but what stays in memory
afterward. Rows are yielded one at a time, but the container and row table are still read whole —
that is the nature of the format; see [xls.md](xls.md).

## Formulas: full recalculation (50,000 rows, 100,000 formulas)

A report workbook: two columns of constants, `C=A*B*VLOOKUP(...)` against a 100-row lookup table,
a nested `IF` in column D, and summary `SUM`/`COUNTIF`. Each engine opens the same file and
recalculates from scratch.

| Method | Time | Allocations |
| --- | ---: | ---: |
| **Cellwright** | **1.43 s** | 328 MB |
| EPPlus 8 | 4.32 s | 1,606 MB |
| ClosedXML 0.105 | 178.07 s | 1,068 MB |

Three times faster than EPPlus and 125 times faster than ClosedXML. More than half of that time
goes to reading the package, not calculation: a pure recalc of the same workbook takes 112 ms (see
the next section). The machine that produced these numbers is noisy — confidence intervals are wide
for all three — so ratios matter more than absolute values here.

## Formulas: pure recalculation (no file read)

The same report workbook, but only `Calculate` is measured: reading the package takes longer than
the calculation itself and hides in the combined number.

| Scenario | Time | Allocations |
| --- | ---: | ---: |
| Lookup table on a separate sheet | 112 ms | 13.7 MB |
| Lookup table on the same sheet as the results | 111 ms | 13.7 MB |

One hundred thousand formulas recalculate in 112 ms. The match between the two rows is the main
point: in the second case every computed result lands on the same sheet as the lookup table, and
coarse sheet-version invalidation would flush the cache on every row. Rectangle-based invalidation
does not.

Before the range cache, each of 50,000 `VLOOKUP` calls copied the whole table: the same scenario
with load cost 557 MB of allocations versus 13.7 MB for recalc today.

## Formulas: range subscription lookup

Editing a cell must find formulas that read it through a range. The benchmark isolates the lookup
structure: 50,000 subscriptions of the form `SUM(A$1:A{n})`, none repeated.

| Method | Time |
| --- | ---: |
| Interval index | 55 ns |
| Linear scan of the same subscriptions | 121 µs |

A 2,200× difference. Neither variant allocates; only the number of subscriptions touched to
answer changes.

At workbook level it looks like this:

| Scenario (50,000 range formulas) | Time | Allocations |
| --- | ---: | ---: |
| Edit outside all subscribed ranges | 68 µs | 144 B |
| Edit found among 50,000 distinct ranges | 820 µs | 496 B |
| Edit in a table read by 50,000 formulas | 57.6 ms | 11.4 MB |

The last row is not about lookup but about work volume: editing the lookup table dirties all fifty
thousand dependent formulas, and they genuinely need recalculation.

## Formulas: incremental recalculation

The workbook is already calculated; one cell changes. The dependency graph exists so edit cost
follows the size of the affected subgraph, not the size of the workbook.

| Scenario | Time | Allocations |
| --- | ---: | ---: |
| Edit with a chain of 10,000 dependent formulas behind it | 9.3 ms | 1.1 MB |
| Edit of one of 50,000 independent formulas | **114 µs** | 496 B |
| Forced full recalc of the same 50,000 formulas | 38.9 ms | 2.4 MB |

A local edit costs roughly 340× less than a full recalc and almost no allocations. The first
version of traversal cost 5.5 ms for the same edit because it scanned the entire formula list;
reverse edges in the index and breadth-first traversal made cost proportional to the affected
subgraph. The 10,000-formula chain case cannot get cheaper: all 10,000 genuinely need
recalculation.

## Formulas: parsing

10,000 formulas of seven typical shapes, parsed into an RPN program.

| Scenario | Time | Allocations |
| --- | ---: | ---: |
| 10k distinct formulas | 17.1 ms | 13.3 MB |
| 10k formulas via `ProgramCache` (7 unique) | 0.54 ms | 9.8 KB |
| 10k formulas: parse plus reverse render to A1 | 29.7 ms | 21.0 MB |

That is about 580,000 formulas per second on a cold parse. Real sheets repeat one formula down a
column, and the program cache applies: the same 10,000 cells cost 30× less and almost no
allocations.

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

The combined formula recalc run takes about twenty minutes: ClosedXML spends roughly three minutes
per iteration.

Run each benchmark class with a separate command: BenchmarkSwitcher cleans generated artifacts
between sessions, and a single `--filter *` breaks the second session.
