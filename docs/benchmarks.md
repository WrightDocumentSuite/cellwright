# Benchmarks

Measured on **2026-11-09**, Cellwright **0.1.4**, [public source `9540bfbf66a9`](https://github.com/WrightDocumentSuite/cellwright/tree/9540bfbf66a9b54f93504f016ccfee0616a78138).

Environment: Microsoft Windows Server 2025 Datacenter (build 26100), Intel Xeon Gold 5412U, 24 cores / 48 logical processors, 255.7 GiB RAM; .NET SDK 10.0.403, .NET 10.0.14 (10.0.14, 10.0.1426.52326), BenchmarkDotNet 0.15.8.

The server also ran other workloads. Suites and benchmark classes ran sequentially with their committed job parameters. These are shared-host measurements; the August i7-8700 results are a different hardware baseline and do not establish a code speedup.

Tables use the first complete pass throughout. Timings and allocations are per operation; KiB and MiB are binary units. Error is BenchmarkDotNet's reported confidence margin, N is its retained sample count, and CV is standard deviation divided by mean. High-CV rows and confidence margins larger than the mean should be treated as noisy observations.

The [machine-readable snapshot](benchmark-results/2026-11-09.json) records exact assembly versions, hashes, source commit, and every measured row. Runtime servicing and competitor versions are part of this snapshot, rather than labels copied from the previous tables.

| Assembly | Product version |
| --- | --- |
| `BenchmarkDotNet.Annotations.dll` | `0.15.8+1e8238ef6c08e18821edae83ac8d6b0a2bb1bb98` |
| `BenchmarkDotNet.dll` | `0.15.8+1e8238ef6c08e18821edae83ac8d6b0a2bb1bb98` |
| `Cellwright.Benchmarks.dll` | `0.1.4+9540bfbf66a9b54f93504f016ccfee0616a78138` |
| `Cellwright.dll` | `0.1.4+9540bfbf66a9b54f93504f016ccfee0616a78138` |
| `Cellwright.Formulas.dll` | `0.1.4+9540bfbf66a9b54f93504f016ccfee0616a78138` |
| `Cellwright.Pdf.dll` | `0.1.4+9540bfbf66a9b54f93504f016ccfee0616a78138` |
| `Cellwright.Xls.dll` | `0.1.4+9540bfbf66a9b54f93504f016ccfee0616a78138` |
| `ClosedXML.dll` | `0.105.0+d15f6690886801980c0d83b7eb1d1b5d2171ad31` |
| `ClosedXML.Parser.dll` | `1.0.0+658973aeedc2fe289e0ccba2bb9959f67692ded5` |
| `DocumentFormat.OpenXml.dll` | `3.1.1+Branch.main.Sha.f9e6ad7524e22b6b1f932deabc9055201b7870c0.f9e6ad7524e22b6b1f932deabc9055201b7870c0` |
| `DocumentFormat.OpenXml.Framework.dll` | `3.1.1+Branch.main.Sha.f9e6ad7524e22b6b1f932deabc9055201b7870c0.f9e6ad7524e22b6b1f932deabc9055201b7870c0` |
| `EPPlus.dll` | `8.6.3+21b0af6b3d004bcfe8cbfd5c03893fb9915bcb9f` |
| `EPPlus.Interfaces.dll` | `8.6.2+d8416ee1c78423f478a891b0b2112da6e0919e48` |
| `ExcelDataReader.dll` | `3.7.0+840007736467bb66be083764021a3a6b9afad5c2` |
| `Inkwright.dll` | `0.1.4+5028efe3b640193b92863f7b94f3452809d88107` |

The EPPlus write allocation total is unavailable for comparison: its raw export reports 0 B despite nonzero Gen0/Gen1/Gen2 collection counts. The raw report is retained, but the zero is not presented as an allocation-free operation. Timing statistics remain reported.

## AggregateBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| SumOneMillionCells | Cells=1000000 | 21.194 ms | 25.742 ms | 20.533 ms | 176 B | 5 | 31.5% |

## AutoFitBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Columns | Rows=20000 | 167.689 ms | 1.055 s | 197.895 ms | 22.209 MiB | 3 | 34.5% |
| ColumnsSampled | Rows=20000 | 16.998 ms | 47.983 ms | 16.318 ms | 2.278 MiB | 3 | 15.5% |
| ColumnsAndRows | Rows=20000 | 203.051 ms | 2.818 s | 134.190 ms | 44.522 MiB | 3 | 76.1% |

## CopiedFormulaBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| RelativeExpression | Rows=100000 | 21.833 ms | 5.553 ms | 21.993 ms | 4.578 MiB | 3 | 1.4% |

## CopiedRangeCacheBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| SlidingVlookup | Rows=70000 | 126.023 ms | 1.099 s | 99.261 ms | 129.776 MiB | 3 | 47.8% |

## FormulaCalcBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Cellwright | Rows=50000 | 1.016 s | 1.342 s | 982.680 ms | 397.700 MiB | 3 | 7.2% |
| ClosedXml | Rows=50000 | 80.545 s | 1.553 s | 80.573 s | 1,066.783 MiB | 3 | 0.1% |
| EpPlus | Rows=50000 | 2.027 s | 3.650 s | 1.952 s | 1,606.073 MiB | 3 | 9.9% |

## FormulaParseBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Distinct | Count=10000 | 9.110 ms | 1.127 ms | 9.136 ms | 17.351 MiB | 4 | 1.9% |
| Cached | Count=10000 | 162.274 µs | 5.297 µs | 162.169 µs | 12.922 KiB | 5 | 0.8% |
| Render | Count=10000 | 14.581 ms | 986.677 µs | 14.457 ms | 24.948 MiB | 5 | 1.8% |

## FormulaRecalcBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| SeparateSheet | Rows=50000 | 76.670 ms | 36.029 ms | 72.156 ms | 13.733 MiB | 5 | 12.2% |
| SameSheet | Rows=50000 | 71.560 ms | 32.488 ms | 68.733 ms | 13.733 MiB | 5 | 11.8% |

## IncrementalCalcBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Chain |  | 7.149 ms | 8.021 ms | 6.463 ms | 1.048 MiB | 5 | 29.1% |
| Independent |  | 175.140 µs | 65.708 µs | 176.700 µs | 496 B | 5 | 9.7% |
| Forced |  | 19.049 ms | 7.237 ms | 18.961 ms | 2.289 MiB | 5 | 9.9% |

## PdfBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Plain | Rows=20000 | 387.897 ms | 3.989 s | 307.994 ms | 109.398 MiB | 3 | 56.4% |
| Sparse | Rows=20000 | 14.082 ms | 54.665 ms | 12.755 ms | 6.395 MiB | 3 | 21.3% |
| Styled | Rows=20000 | 478.451 ms | 3.777 s | 401.938 ms | 145.306 MiB | 3 | 43.3% |
| Tagged | Rows=20000 | 692.607 ms | 3.408 s | 640.868 ms | 464.910 MiB | 3 | 27.0% |
| Conditional | Rows=20000 | 428.963 ms | 2.174 s | 363.500 ms | 206.115 MiB | 3 | 27.8% |
| Graded | Rows=20000 | 484.098 ms | 3.307 s | 400.555 ms | 216.395 MiB | 3 | 37.4% |

## RangeEditBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| EditOutsideRanges | Formulas=50000 | 85.060 µs | 55.356 µs | 79.000 µs | 144 B | 5 | 16.9% |
| EditInsideSharedTable | Formulas=50000 | 63.349 ms | 53.087 ms | 67.985 ms | 10.894 MiB | 5 | 21.8% |
| EditAmongDistinctRanges | Formulas=50000 | 405.420 µs | 191.694 µs | 382.100 µs | 496 B | 5 | 12.3% |

## ReadBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| CellwrightStreaming | Rows=100000 | 2.024 s | 320.361 ms | 2.027 s | 626.729 MiB | 3 | 0.9% |
| CellwrightDom | Rows=100000 | 2.097 s | 462.444 ms | 2.096 s | 674.461 MiB | 3 | 1.2% |
| ClosedXml | Rows=100000 | 5.283 s | 5.763 s | 5.128 s | 2,617.666 MiB | 3 | 6.0% |
| EpPlus | Rows=100000 | 2.410 s | 1.069 s | 2.392 s | 1,782.266 MiB | 3 | 2.4% |

## SubscriptionLookupBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Indexed | Subscriptions=50000 | 25.921 ns | 0.614 ns | 25.916 ns | 0 B | 4 | 0.4% |
| Linear | Subscriptions=50000 | 51.373 µs | 8.928 µs | 50.779 µs | 0 B | 5 | 4.5% |

## WriteBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| CellwrightStreaming | Rows=100000 | 1.373 s | 1.300 s | 1.378 s | 32.299 MiB | 3 | 5.2% |
| CellwrightDom | Rows=100000 | 1.293 s | 303.767 ms | 1.294 s | 78.412 MiB | 3 | 1.3% |
| ClosedXml | Rows=100000 | 4.159 s | 1.115 s | 4.126 s | 2,094.892 MiB | 3 | 1.5% |
| EpPlus | Rows=100000 | 5.981 s | 3.154 s | 5.983 s | unavailable | 3 | 2.9% |

## XlsBenchmarks

| Method | Parameters | Mean | Error | Median | Allocated | N | CV |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| CellwrightRead | Rows=50000 | 53.423 ms | 181.919 ms | 53.513 ms | 23.861 MiB | 3 | 18.7% |
| CellwrightStreamRows | Rows=50000 | 103.456 ms | 126.555 ms | 100.301 ms | 22.766 MiB | 3 | 6.7% |
| ExcelDataReaderRead | Rows=50000 | 190.709 ms | 1.662 s | 241.591 ms | 96.126 MiB | 3 | 47.8% |
| CellwrightWrite | Rows=50000 | 53.865 ms | 27.764 ms | 53.689 ms | 10.605 MiB | 3 | 2.8% |

## Reproduce

Use a clean standalone public checkout at the source commit above, the pinned SDK, and the required fonts. Run each class in a separate process with a distinct artifact directory; this avoids shared generated-build cleanup between benchmark sessions.

```sh
dotnet restore benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --locked-mode
dotnet build benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj -c Release --no-restore
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.AggregateBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/AggregateBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.AutoFitBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/AutoFitBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.CopiedFormulaBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/CopiedFormulaBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.CopiedRangeCacheBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/CopiedRangeCacheBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.FormulaCalcBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/FormulaCalcBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.FormulaParseBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/FormulaParseBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.FormulaRecalcBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/FormulaRecalcBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.IncrementalCalcBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/IncrementalCalcBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.PdfBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/PdfBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.RangeEditBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/RangeEditBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.ReadBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/ReadBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.SubscriptionLookupBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/SubscriptionLookupBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.WriteBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/WriteBenchmarks"
dotnet run -c Release --project benchmarks/Cellwright.Benchmarks/Cellwright.Benchmarks.csproj --no-build -- --filter "Cellwright.Benchmarks.XlsBenchmarks.*" --exporters json --artifacts "BenchmarkDotNet.Artifacts/XlsBenchmarks"
```

For a new comparison, rerun the complete suite and retain the environment and raw reports.

## Historical results and correctness

The [previous benchmark document](https://github.com/WrightDocumentSuite/cellwright/blob/9540bfbf66a9b54f93504f016ccfee0616a78138/docs/benchmarks.md) preserves the August measurements and their correctness/corpus notes. Corpus snapshots and build/test wall-clock telemetry are separate evidence from the measurements above.
