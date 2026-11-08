# NuGet release

Publication is irreversible. A version can be unlisted, but it cannot be deleted or replaced: any
number that once reached nuget.org is taken forever. This page describes what to verify before
`push`, not only how to type the command.

## What gets published

All five packages ship together, under one version taken from
[Directory.Build.props](../Directory.Build.props).

| Package | Dependencies |
| --- | --- |
| `Cellwright` | — |
| `Cellwright.Formulas` | `Cellwright`, `System.Numerics.Tensors` |
| `Cellwright.Mapping` | `Cellwright` (analyzer bundled in the package) |
| `Cellwright.Xls` | `Cellwright`, `Cellwright.Formulas` |
| `Cellwright.Pdf` | `Cellwright`, [`Inkwright`](https://www.nuget.org/packages/Inkwright) |

`Inkwright` is pinned to an exact version in
[Directory.Packages.props](../Directory.Packages.props): the PDF package is released against the
Inkwright it was tested against, and a floating range would let it resolve against one it never saw.
Changing Inkwright for this library means releasing Inkwright first, waiting for it to index, and
only then building the package that depends on it.

## Before release

- **Cut from `main`.** SourceLink embeds the branch and commit SHA the package was built from — they
  appear in `<repository>` inside the `.nuspec`. A package built from a working branch names that
  branch aloud and sends the debugger to the wrong place.
- **Set the version.** `<Version>` in [Directory.Build.props](../Directory.Build.props) is shared
  across all packages, and pass the same value to the verifier below. While major is zero,
  breaking changes are allowed in minor releases.
- **Update [CHANGELOG.md](../CHANGELOG.md).** The "Unreleased" section becomes the version section
  with a date.
- **Verify build and tests.** Warnings in this repository are errors, so a clean build is part of the
  contract:

  ```bash
  dotnet build Cellwright.slnx -c Release
  dotnet test Cellwright.slnx -c Release --no-build
  ```

- **Build the packages from the pushed commit.** SourceLink records the commit of the tree `pack`
  ran in; a package built before the code was pushed names a commit the public repository does not
  have, and the debugger goes nowhere.

## Reproducible restore and local packages

Every project commits `packages.lock.json`. CI restores with `--locked-mode` into a repository-local
`NUGET_PACKAGES` directory, and its cache key includes the SDK selection, project files, central
props/targets and every lock file. There is deliberately no broad cache fallback: a package built
locally under a reused stable version must never satisfy a later restore.

Use the package verifier for development builds:

```powershell
pwsh ./eng/verify-packages.ps1
```

It assigns one unique `0.1.4-dev.<UTC>.g<commit>.<random>` version to the whole package set, uses an
isolated NuGet cache, validates the finished nuspecs and analyzer/SourceLink assets, installs the
packages into an empty consumer, then publishes and runs that consumer with trimming and Native AOT.
The output lives below `artifacts/package-verification/runs/<version>/`. A normal `dotnet pack` is
rejected even for prereleases, so local packages cannot reuse an ambiguous identity or bypass the
package/consumer checks.

For a release, start with a new package cache and keep it for that release invocation only:

```bash
release_nuget_packages="$(mktemp -d "${TMPDIR:-/tmp}/cellwright-release-nuget.XXXXXX")"
export NUGET_PACKAGES="$release_nuget_packages"
dotnet restore Cellwright.slnx --locked-mode
dotnet build Cellwright.slnx -c Release --no-restore
dotnet test Cellwright.slnx -c Release --no-build --no-restore
```

Review all lock-file changes before release; if locked restore fails, update and review the locks
rather than disabling locked mode. Never invoke the stable-pack MSBuild opt-in directly: the
verifier owns that escape hatch so a release cannot bypass package validation.

## Package verification

Build the stable package set through the same gate used by CI:

```powershell
$releaseVersion = 'X.Y.Z'
pwsh ./eng/verify-packages.ps1 `
  -PackageVersion $releaseVersion `
  -AllowStablePackageVersion
```

Stable verification requires a clean public checkout and performs the strict SourceLink content
check; `-SkipSourceLinkContentCheck` is deliberately rejected. The command prints the unique
`package-directory` for this run. Publish only the `.nupkg` files from that directory. Before
publishing, inspect one of them, for example with
`unzip -l "$package_directory/Cellwright.X.Y.Z.nupkg"`, where `package_directory` is the printed
path.

- `lib/net10.0/` contains both `.dll` and `.xml` — XML documentation is included
  (`GenerateDocumentationFile`); without it, the IDE shows no tooltips.
- `README.md` inside the package: it appears on the package page.
- Beside each `.nupkg` lies a `.snupkg` with symbols.
- `<repository>` inside the `.nuspec` shows the expected branch and commit.
- `Cellwright.Mapping` includes `analyzers/dotnet/cs/Cellwright.Mapping.Generator.dll`: without it
  the generator does not load and the package is useless.
- `<dependencies>` lists nothing that is not on nuget.org.
- Every Cellwright-to-Cellwright dependency and `Inkwright` uses an exact bracketed range such as
  `[0.1.3]`, not the open-ended minimum range represented by bare `0.1.3`.

The scheduled `Corpus integration` workflow downloads ClosedXML, ExcelDataReader, four pinned EPPlus
releases, and both the NPOI 1.2.2 and 2.7.6-rc1 source corpora at pinned commits and verifies their
archive SHA-256 values. It also rejects any
corpus-related skipped test through the produced TRX files. The monthly `Security` workflow audits
the complete dependency graph and runs the coverage-guided spreadsheet fuzz target. Update
`eng/corpus/manifest.json` deliberately when advancing an upstream corpus: verify both the commit
and the new archive digest, then run the corpus suite before merging.

For an exact local reproduction, extract only those manifest archives below one directory and set
`CELLWRIGHT_CORPUS_ROOT` to its absolute path. When the variable is present the tests fail on a
missing/malformed root and do not mix in fixtures from neighbouring checkouts. Without it, local
development retains the convenient historical sibling-repository discovery.

## Publication

Order matters: a dependency must be live before whatever references it. NuGet does not index a new
package instantly, so wait between layers and confirm the previous one is visible.

```bash
package_directory="/absolute/path/printed/by/the/verifier"

# 1. Core.
dotnet nuget push "$package_directory/Cellwright.X.Y.Z.nupkg" \
  --source https://api.nuget.org/v3/index.json --api-key "$NUGET_PUB_UNPUB"

# 2. Packages that depend only on core.
dotnet nuget push "$package_directory/Cellwright.Formulas.X.Y.Z.nupkg" ...
dotnet nuget push "$package_directory/Cellwright.Mapping.X.Y.Z.nupkg" ...

# 3. Packages that depend on Formulas, and on Inkwright.
dotnet nuget push "$package_directory/Cellwright.Xls.X.Y.Z.nupkg" ...
dotnet nuget push "$package_directory/Cellwright.Pdf.X.Y.Z.nupkg" ...
```

The `.snupkg` follows its `.nupkg` on the same `push` automatically.

Do not store the API key in the repository or pass it on a command line that shell history records:
use an environment variable or a CI secret only.

After publication — tag and GitHub release:

```bash
git tag -a vX.Y.Z -m "Cellwright X.Y.Z"
git push origin vX.Y.Z
```

## Post-release verification

Install what was published into an empty project; do not trust the website alone:

```bash
mkdir /tmp/verify && cd /tmp/verify && dotnet new console
dotnet add package Cellwright --version X.Y.Z
dotnet build
```

That catches exactly the two failures a paper check misses: an unresolved dependency and a package
missing required content.

## What is not here yet

Publication is manual: workflows under `.github/workflows/` build, test, and upload artifacts, but
do not push anywhere. While releases are rare, that is arguably a feature — an irreversible action
stays deliberate. When they become frequent, add a separate workflow on tag `v*` with the key in
secrets and the same checks listed above.

## SDK servicing and lock files

This release uses exactly .NET SDK 10.0.403 (`global.json`, roll-forward disabled).
The SDK supplies implicit linker dependencies, so changing its servicing patch can
invalidate a locked restore even when the explicit package references are unchanged.
To service the SDK, update `global.json`, restore with `--force-evaluate`, review the
lock-file diff, and repeat the locked restore, source tests, and package consumer gates.
Commit these changes together; do not disable locked restore in CI.

The test projects deliberately use `xunit.v3.mtp-off` 3.2.2 with the VSTest adapter.
This retains the existing `dotnet test`, TRX, filters, and coverage commands while
removing the unused, deprecated Microsoft.Testing.Platform.MSBuild 1.9.1 dependency.
Do not switch to the MTP v2 package without migrating those commands together.

Linux PDF test geometry assumes the Office-compatible fonts installed by
`bash eng/install-linux-fonts.sh`. Without Carlito, Calibri can fall back to a wider face and
move chart anchors across a page boundary. Font substitution remains a diagnostic; the tests
accept only the expected Calibri-to-Carlito message and still reject content diagnostics.
The Hebrew bidi assertion explicitly uses DejaVu Sans on Linux because Carlito has no Hebrew
glyphs. These font choices isolate layout and bidi checks from the runner's incidental inventory.
