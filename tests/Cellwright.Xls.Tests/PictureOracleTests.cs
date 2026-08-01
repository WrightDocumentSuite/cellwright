using System.Security.Cryptography;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Checks the images read out of the legacy corpus against NPOI, which reads the same Office Art
/// layer through its own implementation.
/// </summary>
/// <remarks>
/// The two libraries answer slightly different questions, and the comparison is built around
/// that rather than against it. NPOI's <c>GetAllPictures</c> returns the workbook's image
/// <em>store</em> — every image the file holds, whether or not anything displays it. This
/// library returns the images <em>placed on sheets</em>, which is what the model is for. So a
/// store entry no shape uses shows up in NPOI and not here, and one image shown six times shows
/// up once in NPOI and six times here.
/// <para>
/// What the two must agree on absolutely is the bytes: every image this library hands back has
/// to be one the store actually holds, byte for byte. That is the assertion below, and it is a
/// strong one — it caught a truncated store, where the prefix matched and the tail was missing.
/// </para>
/// </remarks>
public sealed class PictureOracleTests
{
    static PictureOracleTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Fact]
    public async Task ImagesAgreeWithTheOracle()
    {
        Assert.SkipWhen(XlsCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        int files = 0, matched = 0, placements = 0;
        var foreign = new List<string>();

        foreach (string path in XlsCorpus.Files)
        {
            if (ReadStore(path) is not { Count: > 0 } store)
                continue;

            Workbook mine;
            try
            {
                mine = await Spreadsheet.LoadAsync(path, options: null, TestContext.Current.CancellationToken);
            }
            catch (Exception exception) when (XlsCorpus.IsUnreadable(exception))
            {
                continue;
            }

            files++;
            HashSet<string> ours = [];
            foreach (Worksheet sheet in mine.Sheets)
            {
                foreach (Picture picture in sheet.Pictures)
                {
                    placements++;
                    ours.Add(Digest(picture.Content.Span));
                }
            }

            if (ours.SetEquals(store))
                matched++;

            foreach (string unknown in ours.Except(store, StringComparer.Ordinal))
                foreign.Add($"{Path.GetFileName(path)}: {unknown[..16]} is not in the file's image store");
        }

        Assert.True(files > 60, $"Only {files} files with images were compared; coverage has dropped.");
        Assert.True(placements > 300, $"Only {placements} pictures were placed; coverage has dropped.");

        // Not one byte may be invented. Anything here is a misread record, not a difference of
        // opinion about what counts as a picture.
        Assert.True(
            foreign.Count == 0,
            $"{foreign.Count} images do not match the oracle:{Environment.NewLine}" +
            string.Join(Environment.NewLine, foreign.Take(20)));

        // Where every store entry is placed on a sheet, the two answers coincide exactly. The
        // rest hold an image no shape displays — a header graphic, a control's face.
        Assert.True(matched > 45, $"Only {matched} of {files} files matched the store exactly.");
    }

    /// <summary>
    /// The same comparison for the modern format, where the store is the package's media parts.
    /// </summary>
    /// <remarks>
    /// Two formats, two entirely separate readers on this side — DrawingML through the XML
    /// reader, Office Art through the record reader — checked against the same library. A bug
    /// shared by both would have to be a misunderstanding of what a picture is, not a misread.
    /// </remarks>
    [Fact]
    public async Task ImagesOfPackagesAgreeWithTheOracle()
    {
        Assert.SkipWhen(XlsCorpus.Packages.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        int files = 0, placements = 0;
        var foreign = new List<string>();

        foreach (string path in XlsCorpus.Packages)
        {
            if (ReadPackageStore(path) is not { Count: > 0 } store)
                continue;

            Workbook mine;
            try
            {
                mine = await Workbook.LoadAsync(path, cancellationToken: TestContext.Current.CancellationToken);
            }
            catch (Exception)
            {
                continue;
            }

            files++;
            foreach (Worksheet sheet in mine.Sheets)
            {
                foreach (Picture picture in sheet.Pictures)
                {
                    placements++;
                    if (!store.Contains(Digest(picture.Content.Span)))
                        foreign.Add($"{Path.GetFileName(path)} [{sheet.Name}]: an image not in the package");
                }
            }
        }

        Assert.True(files > 30, $"Only {files} packages with images were compared; coverage has dropped.");
        Assert.True(placements > 60, $"Only {placements} pictures were placed; coverage has dropped.");
        Assert.True(
            foreign.Count == 0,
            $"{foreign.Count} images do not match the oracle:{Environment.NewLine}" +
            string.Join(Environment.NewLine, foreign.Take(20)));
    }

    /// <summary>
    /// The anchors of a named fixture, against the shapes NPOI reports. A byte-for-byte image in
    /// the wrong place is still wrong, and the store comparison cannot see position at all.
    /// </summary>
    [Fact]
    public async Task AnchorsAgreeWithTheOracle()
    {
        string? path = XlsCorpus.Files.FirstOrDefault(
            static p => Path.GetFileName(p).Equals("drawings.xls", StringComparison.OrdinalIgnoreCase));
        Assert.SkipWhen(path is null, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        List<string> oracle = [];
        using (FileStream input = File.OpenRead(path!))
        {
            var book = new HSSFWorkbook(input);
            for (int i = 0; i < book.NumberOfSheets; i++)
            {
                if (book.GetSheetAt(i).DrawingPatriarch is not HSSFPatriarch patriarch)
                    continue;

                foreach (HSSFShape shape in patriarch.Children)
                {
                    if (shape is HSSFPicture picture && picture.Anchor is HSSFClientAnchor anchor)
                        oracle.Add($"{anchor.Col1},{anchor.Row1}-{anchor.Col2},{anchor.Row2}");
                }
            }
        }

        Assert.NotEmpty(oracle);

        Workbook mine = await Spreadsheet.LoadAsync(path!, options: null, TestContext.Current.CancellationToken);
        List<string> ours = [];
        foreach (Worksheet sheet in mine.Sheets)
        {
            foreach (Picture picture in sheet.Pictures)
            {
                PictureAnchor anchor = picture.Anchor;
                ours.Add(
                    $"{anchor.From.Column - 1},{anchor.From.Row - 1}-{anchor.To.Column - 1},{anchor.To.Row - 1}");
            }
        }

        Assert.Equal(oracle.Order(StringComparer.Ordinal), ours.Order(StringComparer.Ordinal));
    }

    /// <summary>
    /// A legacy workbook converted to xlsx keeps its pictures. The old format has no drawing
    /// part to carry over, so the only way they survive is by the model owning them and the
    /// xlsx writer generating a drawing from scratch.
    /// </summary>
    [Fact]
    public async Task PicturesSurviveTheConversionToXlsx()
    {
        string? path = XlsCorpus.Files.FirstOrDefault(
            static p => Path.GetFileName(p).Equals("drawings.xls", StringComparison.OrdinalIgnoreCase));
        Assert.SkipWhen(path is null, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        Workbook legacy = await Spreadsheet.LoadAsync(path!, options: null, TestContext.Current.CancellationToken);
        List<string> before = Digests(legacy);
        Assert.NotEmpty(before);

        using var saved = new MemoryStream();
        await legacy.SaveAsync(saved, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        saved.Position = 0;

        Workbook converted = await Workbook.LoadAsync(saved, cancellationToken: TestContext.Current.CancellationToken);
        Assert.Equal(before, Digests(converted));
    }

    private static List<string> Digests(Workbook workbook) =>
        [.. workbook.Sheets
            .SelectMany(static sheet => sheet.Pictures)
            .Select(static picture => Digest(picture.Content.Span))
            .Order(StringComparer.Ordinal)];

    /// <summary>Every image the workbook stores, by digest, as the oracle sees them.</summary>
    private static HashSet<string>? ReadStore(string path)
    {
        try
        {
            using FileStream input = File.OpenRead(path);
            var book = new HSSFWorkbook(input);
            HashSet<string> store = [];
            foreach (IPictureData picture in book.GetAllPictures())
                store.Add(Digest(picture.Data));
            return store;
        }
        catch (Exception)
        {
            // A file the oracle cannot open is no oracle; the corpus tests cover reading it.
            return null;
        }
    }

    /// <summary>Every image a package holds, by digest, as the oracle sees them.</summary>
    private static HashSet<string>? ReadPackageStore(string path)
    {
        try
        {
            using FileStream input = File.OpenRead(path);
            using var book = new NPOI.XSSF.UserModel.XSSFWorkbook(input);
            HashSet<string> store = [];
            foreach (IPictureData picture in book.GetAllPictures())
                store.Add(Digest(picture.Data));
            return store;
        }
        catch (Exception)
        {
            return null;
        }
    }

    private static string Digest(ReadOnlySpan<byte> content) => Convert.ToHexString(SHA256.HashData(content));
}
