# Pictures and Theme Colors

Images on a sheet are read from all three formats and written to xlsx and `.xls`. A loaded
workbook exposes them through `sheet.Pictures` — an ordinary model collection, like `Comments`
or `Tables`.

```csharp
using Cellwright;
using Cellwright.Xls;

Workbook book = await Spreadsheet.LoadAsync("report.xls");
foreach (Picture picture in book.Sheets[0].Pictures)
    Console.WriteLine($"{picture.Name}: {picture.Format}, {picture.Content.Length} bytes, {picture.Anchor}");

// Converting from .xls to xlsx carries pictures over: the old format has no drawing parts,
// so the model owns them, and the xlsx writer assembles the part from scratch.
await book.SaveAsync("report.xlsx");
```

To add your own image:

```csharp
byte[] logo = await File.ReadAllBytesAsync("logo.png");
sheet.Pictures.Add(logo, "B2", "E8", name: "Logo");
```

Format is determined from the bytes themselves, not the file name: each supported format declares
itself in the first bytes, and the part extension is a formality.

## Bytes are untouched

An image is returned exactly as stored in the file and written back the same way. The library does
not decode, recompress, or scale pictures: an incoming progressive JPEG leaves as the same
progressive JPEG, byte for byte. The only exception is metafiles in `.xls`, which the format stores
deflate-compressed; decompression returns the original file, not a new one.

## Anchors

How a picture is attached to a sheet has three different answers, and they are not interchangeable.

| Kind | Behavior when the grid changes | What it carries |
| --- | --- | --- |
| `TwoCell` | Moves and resizes with cells | Both corner cells with offsets |
| `OneCell` | Moves; size preserved | Top-left cell and size |
| `Absolute` | Does not move at all | Position and size from the sheet corner |

Offsets and sizes are in EMU (English Metric Units), the unit used throughout any drawing part:
914,400 per inch, 12,700 per point. Converting them to something familiar would mean guessing
effective column widths, so the model stores what the file said and leaves arithmetic to the `Emu`
class.

On row insert and delete, pictures move with cells — except those whose `MoveBehavior` is `Fixed`.
The collection is marked changed **only if something actually moved**: this matters more than for
other collections, because a changed collection means rewriting the drawing part (see below).

## How formats differ

**xlsx and xlsb share one reader.** A binary workbook holds cells in BIFF12 records and drawing in
the same DrawingML as the modern format. Their drawing parts are identical, so one reader serves
both; only the byte supplier differs.

**`.xls` is different.** There the layer is Office Art (MS-ODRAW): all image bytes in the workbook
sit together in its `MsoDrawingGroup`, and sheet shapes refer to them by index in that store. Hence
two passes in different substreams — the store from global records, shapes from each sheet.

**What counts as a picture in `.xls`** is not a trivial question. A shape becomes a picture through
its `pib` property, not its type: an ordinary image is a picture frame, but an embedded object and
a form control show their "face" the same way, through a BLIP index, and for someone looking at the
sheet that is a picture in the same sense. The reader keys off the property.

**The drawing layer does not fit in the record format it rides in.** The Office Art tree is cut at
the byte where the record ends, so a container's declared length regularly extends beyond the record
carrying it. The workbook store continues in a second `MsoDrawingGroup`; sheet shapes continue in
following `MsoDrawing` records, and parsing a chunk alone reads exactly as much as that chunk
happened to capture. Chunks are glued before parsing. Failure here does not throw: the image prefix
matches, and the tail is lost.

**In-cell offsets are computed differently.** DrawingML writes absolute length in EMU; BIFF writes
fractions of a cell (1/1024 of column width, 1/256 of row height). Translation requires effective
widths and heights, so for `.xls` a picture's position inside a cell is reconstructed, not read
directly.

## Writing

Writing is supported to xlsx and `.xls`; `.xlsb` is not written at all. The BIFF8 writer emits PNG,
JPEG, BMP, EMF and WMF images into a shared workbook image store. GIF and TIFF cannot be represented
there and produce `PictureFormatNotWritten` through the save warning callback.

In xlsx, a sheet whose pictures were not touched preserves **the original drawing part byte for byte** —
together with charts, autoshapes, and text boxes that lived beside them. Once the collection is
changed, the part is rebuilt from the picture and chart models. Pictures and embedded charts remain;
autoshapes, text boxes, and other unmodelled anchors disappear. If the original part held any of
those, save reports `DrawingShapesNotWritten` through `SaveOptions.OnWarning` — such loss cannot be
silent. Copying a sheet uses the same boundary: modelled pictures and charts get independent anchors,
while source-only drawing content is reported as `SheetCopyContentNotCopied`.

## Theme colors

A theme color is not exotic; it is what almost any styled file uses: Excel's palette is built on a
theme. Previously a theme slot came back as "automatic" because there was nothing to resolve it
with; now `xl/theme/theme1.xml` is read, and `workbook.Theme` can turn an `ExcelColor` into a
concrete ARGB.

```csharp
ExcelColor themed = ExcelColor.FromTheme(4, tint: -0.25);
ExcelColor actual = book.Theme.Resolve(themed);   // #FF3B6091
```

Two pitfalls, both in the spec, not the implementation:

- **Slot order is not the order the theme declares.** The scheme opens with two dark colors and two
  light ones, but styles refer to them with swapped pairs: index 0 is light background, index 1 is
  dark text. The model keeps style numbering (§18.8.3); the theme reader and `BrtColor` reader swap
  pairs on input.
- **Tint is arithmetic on HSL luminance** (§18.8.3), not RGB mixing with white. Negative tint
  compresses luminance toward zero; positive tint walks a fraction of the remaining path to one.

A file without a theme and a workbook created here use the Office scheme Excel opens for a new
document.

The binary format now resolves colors too: `BrtColor` with type 1 resolves through the BIFF8
palette; with type 3 — to a theme slot with tint. Previously both read as "automatic".

## Deliberate limitations

- **Autoshapes, text boxes, and SmartArt are not modeled.** Embedded charts are read and written;
  the other drawing objects survive only with an untouched drawing part and are reported if a
  picture/chart edit or sheet copy must rebuild it.
- **GIF and TIFF are not written to `.xls`**; no pictures are written to xlsb because there is no
  xlsb writer.
- **WMF and EMF are not rendered** — bytes are returned as-is; nothing interprets them.
- **Crop, rotation, effects, and transparency** are not read: the model knows where a picture is and
  what is in it, not how it was styled.
- **In xlsx, one image on two sheets is written twice.** A shared part would be smaller, but then
  removing the picture from one sheet could zero it on the other. The `.xls` writer instead uses
  the format's workbook-wide image store.

## How this is verified

- **NPOI oracle, `.xls`**: 66 corpus files with images, 337 placements. NPOI answers a slightly
  different question — `GetAllPictures` returns the workbook *store*, that is all file images
  regardless of whether anything displays them, while this library returns *sheet placements*. So
  comparison is built around that difference: it is required that **every byte returned is a byte
  from the store**, and in 53 of 66 files the sets match exactly. That requirement caught a
  truncated store where the prefix matched and the tail was lost.
- **NPOI oracle, xlsx**: 42 packages, 151 placements, same requirement. The two sides read by
  completely different paths — DrawingML through XML, Office Art through records — and are checked
  against one library.
- **Anchors**: separate comparison with NPOI shapes on a named fixture. A byte-perfect picture in
  the wrong place is still an error, and store comparison does not see position.
- **Theme colors**: 119 corpus packages, 1,428 slots against NPOI `ThemesTable`, no mismatches. Tint
  is checked separately — on grays from the Excel palette, with tint values as Excel writes them and
  results as the user sees them.
- **Cross-format oracle**: picture count and position were added to `.xlsb` twin comparison.
- **Round-trip**: a picture added through the API survives save and read; all three anchor kinds
  return unchanged; the assembled package passes the OpenXml validator; a workbook without pictures
  gets no drawing part at all.
- **Synthetic MS-ODRAW coverage** for branches absent from the corpus: compressed and uncompressed
  metafile, variant with two digests, nested groups, reference beyond the store, and a tree declaring
  infinite nesting — the last used to blow the stack.
