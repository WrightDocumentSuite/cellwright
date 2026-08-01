using Cellwright.Xls.Vba;
using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Reading the macros out of a workbook. Both formats keep a VBA project as a compound file of
/// its own, with the source of every module compressed behind a cache the editor keeps, so the
/// evidence that this works is source that reads like source.
/// </summary>
public sealed class VbaTests
{
    private static string? FindPackage(string name) => XlsCorpus.Packages
        .FirstOrDefault(path => Path.GetFileName(path).Equals(name, StringComparison.OrdinalIgnoreCase));

    private static string? FindLegacy(string name) => XlsCorpus.Files
        .FirstOrDefault(path => Path.GetFileName(path).Equals(name, StringComparison.OrdinalIgnoreCase));

    [Fact]
    public async Task ThePackageOfAMacroWorkbookYieldsItsModules()
    {
        string? path = FindPackage("Control.xlsm");
        Assert.SkipWhen(path is null, "Control.xlsm is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        VbaProject project = Assert.IsType<VbaProject>(
            await VbaProject.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken));

        Assert.Equal("VBAProject", project.Name);
        Assert.Equal(4, project.Modules.Count);
        Assert.False(project.IsPasswordProtected);

        VbaModule module = Assert.IsType<VbaModule>(project.Find("Module1"));
        Assert.Equal(VbaModuleKind.Procedural, module.Kind);
        Assert.Contains("Sub DropDown3_Change()", module.Code, StringComparison.Ordinal);

        // The code behind a sheet is a module of a different kind, and only the PROJECT stream
        // says so: the directory lumps it in with classes and forms.
        Assert.Equal(VbaModuleKind.Document, project.Find("ThisWorkbook")!.Kind);
        Assert.Equal(VbaModuleKind.Document, project.Find("Sheet1")!.Kind);
    }

    /// <summary>
    /// A project long enough to fill several compression chunks. The bit split of a back
    /// reference changes as a chunk fills and resets at every boundary, so a reader that gets
    /// it wrong produces plausible-looking rubbish only after the first few hundred bytes.
    /// </summary>
    [Fact]
    public async Task SourceLongerThanOneChunkComesOutIntact()
    {
        string? path = FindPackage("VBADecompressBug.xlsm");
        Assert.SkipWhen(path is null, "VBADecompressBug.xlsm is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        VbaProject project = Assert.IsType<VbaProject>(
            await VbaProject.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken));

        VbaModule utils = Assert.IsType<VbaModule>(project.Find("Utils"));
        Assert.True(utils.Code.Length > 3000, $"'Utils' came out {utils.Code.Length} characters long.");
        Assert.StartsWith("Attribute VB_Name = \"Utils\"", utils.Code, StringComparison.Ordinal);
        Assert.Contains("Option Explicit", utils.Code, StringComparison.Ordinal);

        // A wrong back reference corrupts bytes rather than dropping them, so the giveaway is
        // text that stops being text.
        Assert.DoesNotContain('\0', utils.Code);
    }

    [Fact]
    public async Task ALegacyWorkbookKeepsItsProjectBesideTheWorkbookStream()
    {
        string? path = FindLegacy("ExcelDataset.xls");
        Assert.SkipWhen(path is null, "ExcelDataset.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        VbaProject project = Assert.IsType<VbaProject>(
            await VbaProject.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken));

        Assert.NotEmpty(project.Modules);
        Assert.Contains(project.Modules, static module => module.Name == "ThisWorkbook");
    }

    [Fact]
    public async Task AWorkbookWithoutMacrosHasNoProject()
    {
        string? path = FindPackage("FileWithButton.xlsm");
        Assert.SkipWhen(path is null, "FileWithButton.xlsm is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        // The extension promises macros; the package has no project part at all. That is an
        // answer, not a failure.
        Assert.Null(await VbaProject.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken));
    }

    /// <summary>
    /// A password stops the editor from showing the code, not a reader from finding it. Telling
    /// a protected project from an open one means decoding the field, because every project
    /// Excel saves carries it either way.
    /// </summary>
    [Fact]
    public async Task APasswordIsReportedAndDoesNotHideTheSource()
    {
        string? path = FindLegacy("31749.xls");
        Assert.SkipWhen(path is null, "31749.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        VbaProject project = Assert.IsType<VbaProject>(
            await VbaProject.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken));

        Assert.True(project.IsPasswordProtected);
        Assert.NotEmpty(project.Modules);
    }

    /// <summary>
    /// Both fields are taken from corpus workbooks. An unprotected project decodes to a single
    /// zero; a protected one to the twenty-nine bytes that hold the salt and the hash.
    /// </summary>
    [Theory]
    [InlineData("CFCD24F63C0A8B0B8B0B8B", 1)]
    [InlineData("7270DE250A420A42F5BE0B4294A1C0C0DF1F8FBA4480966C87FCE8E7925C9910E27617B341", 29)]
    public void TheProtectionFieldDecodesToItsDocumentedShape(string field, int expected)
    {
        byte[] decoded = Assert.IsType<byte[]>(VbaProtection.Decode(field));
        Assert.Equal(expected, decoded.Length);
    }

    [Fact]
    public void AProtectionFieldThatIsNotOneIsRefused()
    {
        // The second byte has to decode to a version of two; random hex does not.
        Assert.Null(VbaProtection.Decode("00112233445566778899"));
        Assert.Null(VbaProtection.Decode("ABC"));
    }

    /// <summary>
    /// The copy token of the compression splits its sixteen bits between length and offset
    /// according to how full the chunk already is. This is the smallest case that exercises it:
    /// one literal, then a back reference that reaches into what it is itself producing.
    /// </summary>
    [Fact]
    public void ABackReferenceCanOverlapWhatItProduces()
    {
        // Signature, a compressed chunk of four bytes, a flag byte marking token two as a copy,
        // the literal 'a', then a token of offset one and length seven.
        byte[] container = [0x01, 0x03, 0xB0, 0x02, (byte)'a', 0x04, 0x00];

        Assert.Equal("aaaaaaaa", System.Text.Encoding.ASCII.GetString(VbaCompression.Decompress(container)));
    }

    [Fact]
    public void AChunkStoredWithoutCompressionIsCopiedThrough()
    {
        byte[] raw = [.. Enumerable.Range(0, 4096).Select(static i => (byte)(i % 251))];

        // A raw chunk declares the largest size there is and clears the compressed flag.
        byte[] container = [0x01, 0xFF, 0x3F, .. raw];

        Assert.Equal(raw, VbaCompression.Decompress(container));
    }

    [Fact]
    public void SomethingThatIsNotAContainerIsRefused()
    {
        Assert.Throws<XlsFormatException>(static () => VbaCompression.Decompress([0x02, 0x00, 0x00]));
    }

    /// <summary>
    /// Six bytes of chunk produce four thousand, so a file built for the purpose turns
    /// megabytes into gigabytes. Extracting macros is exactly what gets pointed at a document
    /// somebody was sent, which makes this the worst place to grow a buffer without a bound.
    /// </summary>
    [Fact]
    public void AStreamThatExpandsWithoutEndIsAbandoned()
    {
        // One literal and a copy token of the longest run a chunk may hold, reaching back a
        // single byte: every six bytes of input asks for another four thousand out.
        List<byte> container = [0x01];
        for (int chunk = 0; chunk < 64; chunk++)
            container.AddRange([0x03, 0xB0, 0x02, (byte)'a', 0xFC, 0x0F]);

        XlsFormatException failure = Assert.Throws<XlsFormatException>(
            () => VbaCompression.Decompress([.. container], sizeLimit: 4096));

        Assert.Contains("expands past", failure.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// A back reference belongs to its own chunk and may not reach past the chunk it started in
    /// (MS-OVBA 2.4.1.3.19). One that does means the stream is not what it claims, and going on
    /// would produce not text but a convincing imitation of it.
    /// </summary>
    [Fact]
    public void ABackReferenceIsConfinedToItsOwnChunk()
    {
        // A first chunk of eight bytes, then a second whose very first token reaches back into
        // it — legal arithmetic, illegal reference.
        byte[] container =
        [
            0x01,
            0x03, 0xB0, 0x02, (byte)'a', 0x04, 0x00,
            0x03, 0xB0, 0x02, (byte)'b', 0x04, 0x00,
        ];

        // Both chunks decode on their own; only reaching across is refused.
        Assert.Equal("aaaaaaaabbbbbbbb", System.Text.Encoding.ASCII.GetString(VbaCompression.Decompress(container)));

        // Offset five in a chunk holding one byte reaches four bytes into the chunk before it.
        byte[] crossing = [0x01, 0x03, 0xB0, 0x02, (byte)'a', 0x04, 0x00, 0x03, 0xB0, 0x02, (byte)'b', 0x04, 0x40];

        XlsFormatException failure = Assert.Throws<XlsFormatException>(() => VbaCompression.Decompress(crossing));
        Assert.Contains("before the start of its chunk", failure.Message, StringComparison.Ordinal);
    }

    /// <summary>A run longer than a chunk cannot be a run inside one.</summary>
    [Fact]
    public void ABackReferenceLongerThanAChunkIsRefused()
    {
        // Length 4098, three past what a chunk holds.
        byte[] container = [0x01, 0x03, 0xB0, 0x02, (byte)'a', 0xFF, 0x0F];

        XlsFormatException failure = Assert.Throws<XlsFormatException>(() => VbaCompression.Decompress(container));
        Assert.Contains("does not fit a chunk", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheDefaultLimitLeavesRealProjectsAlone()
    {
        // Two orders of magnitude above the largest project anybody writes.
        Assert.Equal(64 * 1024 * 1024, VbaCompression.DefaultSizeLimit);
    }

    /// <summary>
    /// What the code can reach, before reading a line of it. The identifier a reference carries
    /// packs the class id, the path and the display name into one string.
    /// </summary>
    [Fact]
    public async Task TheLibrariesAProjectBindsToAreNamed()
    {
        string? path = FindPackage("VBADecompressBug.xlsm");
        Assert.SkipWhen(path is null, "VBADecompressBug.xlsm is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        VbaProject project = Assert.IsType<VbaProject>(
            await VbaProject.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken));

        VbaReference xml = Assert.Single(project.References, static r => r.Name == "MSXML2");
        Assert.Equal(VbaReferenceKind.Library, xml.Kind);
        Assert.Equal(Guid.Parse("f5078f18-c551-11d3-89b9-0000f81fe221"), xml.TypeLibrary);
        Assert.Equal("Microsoft XML, v6.0", xml.Description);
        Assert.EndsWith("msxml6.dll", xml.Path!, StringComparison.OrdinalIgnoreCase);

        // The control reference of a form carries the placeholder identifier the specification
        // prescribes for a twiddled library, which is not an identifier at all.
        VbaReference forms = Assert.Single(project.References, static r => r.Kind == VbaReferenceKind.Control);
        Assert.Null(forms.TypeLibrary);
    }

    /// <summary>
    /// A file records three separate decisions about closing a project, not one flag, and
    /// whether the editor shows it at all. None of it hides the source.
    /// </summary>
    [Fact]
    public async Task WhoClosedTheProjectIsReportedSeparately()
    {
        string? path = FindLegacy("31749.xls");
        Assert.SkipWhen(path is null, "31749.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        VbaProject project = Assert.IsType<VbaProject>(
            await VbaProject.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken));

        Assert.True(project.Protection.LockedByUser);
        Assert.True(project.Protection.IsLocked);
        Assert.False(project.Protection.IsVisible);
        Assert.NotEmpty(project.Modules);
    }

    [Fact]
    public async Task AnUnprotectedProjectSaysSoOnEveryCount()
    {
        string? path = FindPackage("Control.xlsm");
        Assert.SkipWhen(path is null, "Control.xlsm is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        VbaProject project = Assert.IsType<VbaProject>(
            await VbaProject.LoadAsync(path!, cancellationToken: TestContext.Current.CancellationToken));

        Assert.Equal(VbaProtectionState.Open, project.Protection);
        Assert.False(project.IsPasswordProtected);
    }

    /// <summary>
    /// Every project in the corpus, both formats. A module's source always opens with the
    /// attribute naming it, which is the invariant that catches a decompression gone wrong
    /// wherever it happens.
    /// </summary>
    [Fact]
    public async Task EveryProjectInTheCorpusUnpacksIntoSource()
    {
        Assert.SkipWhen(XlsCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        int projects = 0;
        int modules = 0;
        int references = 0;
        foreach (string file in XlsCorpus.Packages.Concat(XlsCorpus.Files))
        {
            VbaProject? project;
            try
            {
                project = await VbaProject.LoadAsync(file, cancellationToken: TestContext.Current.CancellationToken);
            }
            catch (Exception exception) when (XlsCorpus.IsUnreadable(exception))
            {
                continue;
            }

            if (project is null)
                continue;

            projects++;
            references += project.References.Count;
            foreach (VbaModule module in project.Modules)
            {
                modules++;
                Assert.StartsWith("Attribute VB_Name = ", module.Code, StringComparison.Ordinal);
            }
        }

        Assert.SkipWhen(projects == 0, "No workbook in the corpus carries macros. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");
        Assert.True(modules > 400, $"Only {modules} module(s) came out of {projects} project(s).");

        // The reference records sit between the project information and the modules; reaching
        // the modules at all means the walk passed them, and this says it read them too.
        Assert.True(references > 250, $"Only {references} reference(s) came out of {projects} project(s).");
    }
}
