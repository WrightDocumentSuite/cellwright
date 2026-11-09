# Cellwright benchmarks

Latest measured run: **2026-11-09**, **Cellwright 0.1.4**.

- [Results, environment, limitations and reproduction commands](../docs/benchmarks.md)
- [Machine-readable snapshot and source/assembly hashes](../docs/benchmark-results/2026-11-09.json)
- [Raw BenchmarkDotNet reports](../docs/benchmark-results/2026-11-09/)
- [Website charts](https://wrightsuite.com/#benchmarks)

The measured source is [`9540bfbf66a9`](https://github.com/WrightDocumentSuite/cellwright/tree/9540bfbf66a9b54f93504f016ccfee0616a78138).
The benchmark project is [Cellwright.Benchmarks](Cellwright.Benchmarks/).

The run used a shared Windows Server 2025 host with an Intel Xeon Gold 5412U,
.NET SDK 10.0.403, runtime 10.0.14 and BenchmarkDotNet 0.15.8. Timing tables
include confidence margins; allocations use binary KiB/MiB per operation.
Results from the August i7-8700 run are a different hardware baseline, not
evidence of a code speedup. Follow the report's pinned source and per-class
commands when reproducing; a normal build alone does not refresh measurements.

EPPlus XLSX write allocation is unavailable: the raw zero conflicts with observed GC collections. Do not treat it as zero allocation.
