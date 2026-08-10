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
  across all packages. While major is zero, breaking changes are allowed in minor releases.
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

## Package verification

Inspect the build before it goes out. Check:

```bash
dotnet pack Cellwright.slnx -c Release
unzip -l artifacts/package/release/Cellwright.0.1.3.nupkg
```

- `lib/net10.0/` contains both `.dll` and `.xml` — XML documentation is included
  (`GenerateDocumentationFile`); without it, the IDE shows no tooltips.
- `README.md` inside the package: it appears on the package page.
- Beside each `.nupkg` lies a `.snupkg` with symbols.
- `<repository>` inside the `.nuspec` shows the expected branch and commit.
- `Cellwright.Mapping` includes `analyzers/dotnet/cs/Cellwright.Mapping.Generator.dll`: without it
  the generator does not load and the package is useless.
- `<dependencies>` lists nothing that is not on nuget.org.

## Publication

Order matters: a dependency must be live before whatever references it. NuGet does not index a new
package instantly, so wait between layers and confirm the previous one is visible.

```bash
# 1. Core.
dotnet nuget push artifacts/package/release/Cellwright.0.1.3.nupkg \
  --source https://api.nuget.org/v3/index.json --api-key "$NUGET_PUB_UNPUB"

# 2. Packages that depend only on core.
dotnet nuget push artifacts/package/release/Cellwright.Formulas.0.1.3.nupkg ...
dotnet nuget push artifacts/package/release/Cellwright.Mapping.0.1.3.nupkg ...

# 3. Packages that depend on Formulas, and on Inkwright.
dotnet nuget push artifacts/package/release/Cellwright.Xls.0.1.3.nupkg ...
dotnet nuget push artifacts/package/release/Cellwright.Pdf.0.1.3.nupkg ...
```

The `.snupkg` follows its `.nupkg` on the same `push` automatically.

Do not store the API key in the repository or pass it on a command line that shell history records:
use an environment variable or a CI secret only.

After publication — tag and GitHub release:

```bash
git tag -a v0.1.3 -m "Cellwright 0.1.3"
git push origin v0.1.3
```

## Post-release verification

Install what was published into an empty project; do not trust the website alone:

```bash
mkdir /tmp/verify && cd /tmp/verify && dotnet new console
dotnet add package Cellwright --version 0.1.3
dotnet build
```

That catches exactly the two failures a paper check misses: an unresolved dependency and a package
missing required content.

## What is not here yet

Publication is manual: workflows under `.github/workflows/` build, test, and upload artifacts, but
do not push anywhere. While releases are rare, that is arguably a feature — an irreversible action
stays deliberate. When they become frequent, add a separate workflow on tag `v*` with the key in
secrets and the same checks listed above.
