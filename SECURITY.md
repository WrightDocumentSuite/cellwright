# Security policy

Cellwright parses spreadsheets supplied by callers. Treat every workbook as untrusted input and
apply a finite [`DocumentLoadBudget`](docs/loading-untrusted-input.md), even when the file extension
looks familiar.

## Supported versions

Security fixes are made on `main` and released in the newest Cellwright package set. Older package
versions are not maintained as separate security branches; upgrade to the latest NuGet release
before reporting that a fix is missing.

## Reporting a vulnerability

Use a [private GitHub security advisory](https://github.com/WrightDocumentSuite/cellwright/security/advisories/new).
Include the affected package and version, the smallest reproducer you can safely share, the observed
impact, and any budget or password options used. Do not attach a confidential workbook: reduce it
to synthetic data or describe how the maintainers can generate an equivalent file.

Please do not open a public issue until a fix or mitigation is available. The maintainers will use
the advisory to coordinate validation, a release, and disclosure. Reports about memory exhaustion,
excessive CPU, container expansion, parser crashes, path traversal, or cryptographic integrity are
in scope even when they require a malformed workbook.

## Security verification

The regular CI build restores from committed lock files in an isolated NuGet package directory. A
monthly security workflow rejects direct or transitive NuGet advisories and runs a real
coverage-guided SharpFuzz/libFuzzer target over the format detector, CFB/BIFF, XLSB, OOXML and
encrypted-package paths. The native libFuzzer driver, SharpFuzz tool and public seed corpora are all
pinned; downloaded binaries and source archives are checked by SHA-256 before execution.

NPOI is a test-only interoperability oracle. It is not referenced by any shipped package and is not
loaded by production code. Tests use NPOI 2.7.6 with `System.Security.Cryptography.Xml` 10.0.10
explicitly pinned; the dependency audit covers that test graph without blanket `NU1902`/`NU1903`
suppressions. Corpus and fuzz jobs run on disposable GitHub-hosted runners with finite job and
document-load limits.

Fuzzing and load budgets reduce risk; they are not a sandbox. Applications that accept hostile
documents should also isolate parsing at the process or container boundary and enforce their own
time, memory and upload-size limits.
