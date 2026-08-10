# Macros (VBA)

The `Cellwright.Xls` package extracts macro source from a workbook — from both `.xlsm` and `.xls`.
Both formats store the project the same way, so one reader handles both.

```csharp
using Cellwright.Xls;

VbaProject? macros = await VbaProject.LoadAsync("report.xlsm");
foreach (VbaModule module in macros?.Modules ?? [])
{
    Console.WriteLine($"{module.Kind} {module.Name}");
    Console.WriteLine(module.Code);
}
```

Books without macros are common, so their absence is returned as `null`, not an exception.
The `.xlsm` extension guarantees nothing: a file with a button but no code is still `.xlsm`.

If the book is already loaded, there is no need to read the file again:

```csharp
Workbook book = await Workbook.LoadAsync("report.xlsm");
VbaProject? macros = book.GetVbaProject();
```

This works while preservation of unknown parts is enabled (`LoadOptions.PreserveUnknownParts`, on
by default): the project is not part of the model; it is one of the pieces the loader keeps as raw
bytes. A book read from `.xls` will not have a project here at all — that reader skips it; for a
binary file, call `VbaProject.LoadAsync`.

## What you get

| Property | Contents |
| --- | --- |
| `Name` | Project name in the editor tree, usually `VBAProject` |
| `Description` | Project description, if present |
| `CompilationConstants` | Conditional compilation constants as `Name = Value : Name = Value` |
| `CodePage` | Code page used for names and source |
| `IsPasswordProtected` | Whether the project is password-protected |
| `Protection` | Who locked the project and whether it is visible at all |
| `References` | Libraries the project references |
| `Modules` | Modules in declaration order |

Each module has `Name`, `StreamName`, `Kind`, `Code`, `Description`, `IsReadOnly`, and `IsPrivate`.
There are four kinds: `Procedural` — a standard module, `Document` — code behind a sheet or the
workbook, `Class` — a class, `Designer` — code behind a user form. The last three differ only in
the `PROJECT` stream: the project catalog collapses them into one type.

`Code` is the text as the editor shows it, lines separated by CR+LF. Hidden lines such as
`Attribute VB_Name` are left in place: they are part of what is stored in the file.

## What the macro connects to

For someone parsing an uploaded document, the first question is what the code can reach at all —
before reading a single line:

```csharp
foreach (VbaReference reference in macros.References)
    Console.WriteLine($"{reference.Name}: {reference.Description} ({reference.Path})");
// Scripting: Microsoft Scripting Runtime (C:\Windows\System32\scrrun.dll)
// MSXML2: Microsoft XML, v6.0 (C:\WINDOWS\System32\msxml6.dll)
```

The identifier in the file is a packed string of the form `*\G{GUID}#version#lcid#path#description`
(MS-OVBA 2.1.1.8); `TypeLibrary`, `Path`, and `Description` are its parsed pieces, and `Identifier`
stores it whole because the producer may use a form the parser does not cover.
`Kind` distinguishes a normal library, another VBA project, and an ActiveX control; for the last,
the identifier is conventionally all zeros and `TypeLibrary` is then empty.

## A password does not hide the source

`IsPasswordProtected` means the project is locked with a password — but the source is still
extracted. The password stops the editor from showing code, not the file from storing it; in the
file it sits exactly as in an unprotected project. For someone parsing an uploaded document, that
is what matters.

Locking the project is not a single flag in the file — there are three, and `Protection` exposes
them separately:

```csharp
if (macros.Protection.LockedByUser) { /* locked by whoever edited the project */ }
if (macros.Protection.LockedByHost) { /* locked by the host application */ }
if (!macros.Protection.IsVisible)   { /* project not visible in the editor at all */ }
```

Telling a protected project from an open one requires decrypting the fields: Excel writes `CMG`,
`DPB`, and `GC` into every project, protected or not. The obfuscation there is a rolling XOR with a
seed stored next to the data (MS-OVBA 2.4.3); the library unwinds it only to answer these
questions — the password itself is not recovered.

The password hash is not the only form: the spec (2.3.1.16) allows a plaintext password string in
`DPB`. The tell is the same in both cases — an unprotected project has exactly one zero byte there;
anything else means a password is set.

## How it works

The project is a compound file inside the book: in `.xlsm` it is part `xl/vbaProject.bin`, in `.xls`
it is storage `_VBA_PROJECT_CUR` next to the workbook stream. Inside is storage `VBA` with stream
`dir` and one stream per module.

- **Compression.** Both `dir` and sources are stored compressed (MS-OVBA 2.4.1): a container of
  chunks no larger than 4096 unpacked bytes, one flag byte per eight tokens, each token either a
  literal or a back-reference. The boundary between length and offset on a reference shifts as the
  chunk fills and resets at the chunk boundary — the place where readers like this usually break,
  often not immediately but after a few hundred bytes of plausible garbage.
  A back-reference must not leave its chunk — the spec requires it; violating that means the stream
  is not what it claims, and what follows would be a convincing imitation of text, not text.
- **Catalog.** `dir` records are id-size-data — all except `PROJECTVERSION`, which declares four
  bytes and writes six. A reader that trusts it loses alignment for the rest of the stream.
  Reference records look irregular in the spec but size honestly: `REFERENCECONTROL` continues in
  separate records, and `REFERENCEORIGINAL` nests `REFERENCECONTROL` in the next record, not inside
  itself.
- **Module.** The module stream starts with an editor cache, then compressed source; where the
  first ends is given by `MODULEOFFSET` from the catalog.
- **Names.** Module streams are named after modules, so in `.xls` a module can be named `Workbook` —
  like the workbook stream itself. Lookup goes by path through storages, not by name across the whole
  catalog; names compare case-insensitively as the spec requires.

## A file you do not trust

Extracting macros is usually parsing an uploaded document, so decompression has two layers of
limits: six bytes of compressed chunk can expand to four thousand, and a file built to abuse that
would turn megabytes into gigabytes. Directory and module output contributes to the caller's
`DocumentLoadBudget.MaxInflatedBytes` and observes `MaxPartBytes`; a policy breach is a
`DocumentLoadLimitException` with machine-readable fields. The independent 64 MiB codec guard — two
orders of magnitude above the largest real project — remains even for an unlimited outer budget and
reports an explanatory `XlsFormatException`. Decompression and directory parsing observe the load
cancellation token.

Reading the package does not buffer the whole file: one part is taken from the zip even when the
book weighs hundreds of megabytes. The MS-CFB container must stay in memory — its sectors are
scattered and the directory addresses them by offset. OPC and nested VBA CFB/codec work use one
shared state during `VbaProject.LoadAsync`.

## Limitations

- **Read-only.** You cannot write a changed macro back. That is deliberate: the project part is
  carried through save byte-for-byte (verified by test), so there is nothing to rewrite, and building
  a project anew would mean recompression, rebuilding the cache, and rebuilding the catalog.
- **Forms.** Code behind a form is extracted, but not its layout: control layout lives in the
  designer storage with stream `VBFrame` beside it, which the model does not touch.
- **Signatures.** Content hashes used to sign the project (MS-OVBA 2.4.2, three generations) are not
  computed: this package cannot verify or apply a signature.
- **Code pages.** Names and source are single-byte in the code page declared by the project.
  Most such pages in .NET are available only after
  `Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)`; without registration, text is read
  as Windows-1252 — Latin script is correct, Cyrillic is not.

These are deliberately not read because they add nothing: streams `__SRP_n` are a performance cache
the spec says to ignore on read (2.2.6); `PROJECTwm` — module names in two encodings, redundant
once the catalog gives `MODULENAMEUNICODE`; `PROJECTlk` — ActiveX control licenses.

## How it is verified

- **Specification**: algorithms checked line-by-line against MS-OVBA — chunk size arithmetic, copy
  token bit split, decryption state order, shape of every catalog record. That caught one mismatch:
  back-references were validated from the start of the buffer, not from the start of the chunk.
- **Corpus**: 114 books with projects (23 packages and 91 binary), 595 modules and 301 library
  references — all unpack without a single exception. The invariant checked: every module’s source
  starts with `Attribute VB_Name = `, and a broken decompressor breaks that immediately.
- **Long sources**: `VBADecompressBug.xlsm` — a file EPPlus keeps as a regression for exactly this
  decompressor bug.
- **Unit tests**: back-reference overlapping its own output; reference reaching into the previous
  chunk; reference longer than a chunk; chunk stored uncompressed; container that is not a container;
  stream that inflates without end.
- **Protection**: fields from real books — open decrypts to one zero byte, protected to twenty-nine
  bytes of key and hash. In `31749.xls` the project is locked by the user and hidden from the editor,
  and the source is still extracted.
