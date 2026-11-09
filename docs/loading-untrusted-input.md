# Loading untrusted spreadsheets

Every public spreadsheet reader applies a finite `DocumentLoadBudget` by default. One budget
instance supplies the policy and one internal state carries the counters through the complete
import: package/container discovery, worksheet parsing, shared strings, media, embedded objects,
and a nested CFB VBA project. Format detection does not open an XLSX twice: `Spreadsheet.LoadAsync`
hands the already-admitted OPC package to the XML or XLSB reader. An encrypted workbook likewise
keeps one state from its outer CFB container through the decrypted OPC package.

```csharp
using Cellwright;
using Cellwright.Xls;

DocumentLoadBudget budget = DocumentLoadBudget.Default with
{
    MaxInputBytes = 64 * 1024 * 1024,
    MaxInflatedBytes = 256 * 1024 * 1024,
    MaxPartBytes = 64 * 1024 * 1024,
    MaxCells = 2_000_000,
    MaxTotalCellTextCharacters = 32 * 1024 * 1024,
};

Workbook xlsx = await Workbook.LoadAsync(
    upload,
    new LoadOptions { Budget = budget });

Workbook anyExcel = await Spreadsheet.LoadAsync(
    upload,
    new XlsLoadOptions { Budget = budget });
```

The same property is available on `XlsxReaderOptions`, so row streaming and DOM loading enforce
the same policy. `Spreadsheet.LoadAsync`, `XlsWorkbook.LoadAsync`, `XlsReader.OpenAsync`, and
`VbaProject.LoadAsync` use `XlsLoadOptions.Budget`.

Limits are inclusive. A breach throws `DocumentLoadLimitException`; `LimitName`, `Limit`, and
`Observed` are stable machine-readable fields. Damage to the file remains an
`XlsxFormatException` or `XlsFormatException`, and recoverable repairs remain
`WorkbookWarning`s. A service can therefore distinguish a resource-policy refusal from malformed
content without parsing an English message.

## Default ceilings

| Property | Default | What it bounds |
| --- | ---: | --- |
| `MaxInputBytes` | 512 MiB | Compressed XLSX/XLSB input and the complete CFB `.xls` input |
| `MaxPackageParts` | 16,384 | OPC ZIP entries plus nested CFB directory entries |
| `MaxInflatedBytes` | 2 GiB | Total declared OPC/CFB uncompressed payload in one import |
| `MaxPartBytes` | 512 MiB | One OPC part or CFB stream; actual streamed bytes are checked too |
| `MaxXmlCharactersPerPart` | 512 Mi chars | Decoded XML characters in one part |
| `MaxXmlNodes` | 64,000,000 | XML nodes across parts actually opened by the import |
| `MaxXmlDepth` | 256 | XML nesting |
| `MaxCells` | 10,000,000 | Cell records across worksheets |
| `MaxSharedStrings` | 4,000,000 | Stored shared-string items |
| `MaxCellTextCharacters` | 1 Mi chars | One stored inline/shared text item; leaves room for producer-invalid legacy strings that the XLS reader repairs to Excel's 32,767-character model limit |
| `MaxTotalCellTextCharacters` | 256 Mi chars | Stored cell, formula, and shared-string source text |
| `MaxMediaBytes` | 64 MiB | One image/audio/video payload |
| `MaxTotalMediaBytes` | 512 MiB | Media payload across the import |
| `MaxEmbeddedObjectBytes` | 128 MiB | One OPC/CFB embedded object |
| `MaxEmbeddedObjects` | 4,096 | Embedded-object count |

Shared strings are charged once where they are stored, not again for every cell that references
the same item. Re-reading the same XLSX part through `XlsxReader` likewise does not consume its XML
or cell count twice.

## Where checks happen

- A seekable source length is rejected before a whole-file allocation. A non-seekable source is
  copied through a bounded, cancellation-aware loop and stops on the first byte past the limit.
- The ZIP end record is inspected for its entry count before `ZipArchive` materializes the central
  directory. Central-directory part and expansion metrics are checked before part buffers. Actual
  decompression is reconciled with the declared sizes; reopening one physical part records a
  maximum rather than charging it twice.
- XML validation uses DTD-prohibited readers, a decoded-character ceiling, shared node/depth
  counters, and the caller's cancellation token even inside one large text node. XML identity
  follows content types too, including `application/*+xml` parts whose names have no `.xml` suffix;
  a malformed part's valid prefix still consumes the shared counters before format diagnostics run.
- CFB directory, FAT, DIFAT, miniFAT, stream lengths, and nested object-pool payloads are checked
  before their arrays are allocated. BIFF/BIFF12 loops check cancellation between records. VBA
  directory and module expansion contributes to the same total-inflated counter, observes the
  per-part ceiling, and retains its independent 64 MiB codec guard.
- Agile `EncryptionInfo` XML shares the XML counters and has a 1 MiB descriptor ceiling checked
  from the CFB directory before the stream or any XML attribute is materialized. Each encoded salt,
  verifier, key, or HMAC attribute also has a fixed 4 Ki-character crypto-metadata ceiling before
  its Base64 output buffer is allocated; those fields are normally only tens or hundreds of characters.
- Media and embedded objects are classified from OPC content types and relationship roles, not
  only `/media/` and `/embeddings/` names. Conventional folders remain a damaged-metadata fallback.
  Transitional and Strict relationship URIs are equivalent, and duplicate relationships to one
  physical target do not duplicate its resource charge.
  Raw and safely decoded `%20`-style part names point to the same physical item; encoded separators
  and dot segments are never allowed to rewrite package identity.

`PackageLimits`, `LoadOptions.Limits`, `XlsxReaderOptions.Limits`, and `XlsLoadOptions.Limits`
remain source-compatible. Setting those legacy properties projects their two values onto input,
per-part, and inflated ceilings. The public `Budget` getter exposes that effective projection, so
copying it to another reader does not lose legacy settings; when both properties are assigned, the
last assignment wins. `PackageLimits.None` retains its original meaning and selects
`DocumentLoadBudget.Unlimited`; prefer a deliberately raised finite `Budget` for server input.

CSV and TSV input goes through this same budget: set it with `CsvOptions.LoadBudget`.
`MaxInputBytes` is measured on the bytes of the source stream, so the limit gives the same answer
for seekable and non-seekable streams and for any encoding. A `TextReader` has no bytes, so its
characters are counted against the limit instead. Cell text and cell counts use the same budget
as a workbook.

A quoted field with text after its closing quote is recovered by keeping that text as part of the
field, which is what Excel does. Set `CsvOptions.StrictQuotes` to raise `InvalidDataException` for
that and for a quote that is never closed. Use it when the file comes from a producer that should
have written valid RFC 4180.

Cellwright has no HTML importer: Markdown is export-only. Any future importer must accept this
same policy rather than creating format-local size options.

## Removing active and linked content

Reading a package, preserving what the model does not understand, and cleaning a file are separate
operations. A load or a plain save never removes anything. To produce a clean copy, name what to
leave out and read the report:

```csharp
using Cellwright.Sanitization;

SanitizeReport report = await workbook.SaveSanitizedAsync(
    "clean.xlsx",
    ContentCategories.Macros | ContentCategories.ExternalLinks);

foreach (SanitizedItem item in report.Removed)
    Console.WriteLine($"removed {item.Kind} {item.Location}");
```

| Category | Removed |
| --- | --- |
| `Macros` | The VBA project and Excel 4.0 macro sheets; the workbook stops declaring itself macro-enabled |
| `EmbeddedObjects` | OLE objects, embedded packages, ActiveX controls, and the preview pictures only they used |
| `ExternalLinks` | Links to other workbooks and every external relationship, hyperlinks included |
| `CustomXml` | Custom XML parts and their properties |
| `Signatures` | Digital signatures and the parts that list them |

The workbook object is not changed. Parts, relationships, content types and the XML references to
them are removed together. A formula that reads another workbook loses the formula and keeps its
last value, a defined name that points at one is removed, and so is a conditional format or
validation rule that does. A copy without macros is not an `.xlsm` any more and should not keep
that extension. XML maps in `xl/xmlMaps.xml` are not custom XML parts and are left in place.
