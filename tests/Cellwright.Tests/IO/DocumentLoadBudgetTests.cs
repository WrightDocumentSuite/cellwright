using System.IO.Compression;
using System.Text;
using Cellwright.IO;
using Xunit;

namespace Cellwright.Tests.IO;

public sealed class DocumentLoadBudgetTests
{
    [Fact]
    public void DefaultsAreFiniteValidatedAndLegacyNoneRemainsUnlimited()
    {
        DocumentLoadBudget.Default.Validate();
        Assert.Equal(512L * 1024 * 1024, DocumentLoadBudget.Default.MaxInputBytes);
        Assert.Equal(16_384, DocumentLoadBudget.Default.MaxPackageParts);
        Assert.Equal(10_000_000, DocumentLoadBudget.Default.MaxCells);
        Assert.Equal(1024 * 1024, DocumentLoadBudget.Default.MaxCellTextCharacters);

        var options = new LoadOptions { Limits = PackageLimits.None };
        Assert.Same(DocumentLoadBudget.Unlimited, options.EffectiveBudget);
        Assert.Equal(long.MaxValue, options.EffectiveBudget.MaxXmlNodes);
        Assert.Equal(int.MaxValue, options.EffectiveBudget.MaxEmbeddedObjects);
    }

    [Fact]
    public void OptionsExposeTheEffectiveBudgetAndTheLastAssignmentWins()
    {
        var legacy = new PackageLimits { MaxPartSize = 11, MaxTotalSize = 22 };
        var custom = DocumentLoadBudget.Default with { MaxCells = 7, MaxPartBytes = 33 };

        var loadLegacyLast = new LoadOptions { Budget = custom, Limits = legacy };
        Assert.Equal(11, loadLegacyLast.Budget.MaxPartBytes);
        Assert.Equal(22, loadLegacyLast.Budget.MaxInputBytes);
        Assert.Equal(22, loadLegacyLast.Budget.MaxInflatedBytes);
        Assert.Equal(7, loadLegacyLast.Budget.MaxCells);

        var loadBudgetLast = new LoadOptions { Limits = legacy, Budget = custom };
        Assert.Same(custom, loadBudgetLast.Budget);

        var streaming = new XlsxReaderOptions { Limits = PackageLimits.None };
        Assert.Same(DocumentLoadBudget.Unlimited, streaming.Budget);

        var asymmetric = DocumentLoadBudget.Default with
        {
            MaxInputBytes = 1 * 1024 * 1024,
            MaxInflatedBytes = 2L * 1024 * 1024 * 1024,
        };
        PackageLimits compatibilityView = new LoadOptions { Budget = asymmetric }.Limits;
        Assert.Equal(asymmetric.MaxInputBytes, compatibilityView.MaxTotalSize);

        var roundTripped = new LoadOptions { Limits = compatibilityView };
        Assert.Equal(asymmetric.MaxInputBytes, roundTripped.Budget.MaxInputBytes);
        Assert.Equal(asymmetric.MaxInputBytes, roundTripped.Budget.MaxInflatedBytes);
    }

    [Fact]
    public async Task InvalidBudgetIsRejectedBeforeInputIsRead()
    {
        var invalid = DocumentLoadBudget.Default with { MaxCells = 0 };
        Assert.Throws<ArgumentOutOfRangeException>(invalid.Validate);

        using var input = new MemoryStream([0x50, 0x4B], writable: false);
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () =>
            await Workbook.LoadAsync(
                input,
                new LoadOptions { Budget = invalid, LeaveOpen = true },
                TestContext.Current.CancellationToken));
        Assert.Equal(0, input.Position);
    }

    [Fact]
    public async Task OpcPackageCeilingsHaveStableMachineReadableIdentity()
    {
        byte[] package = await WorkbookPackageAsync("budget");
        (string Name, DocumentLoadBudget Budget)[] cases =
        [
            (nameof(DocumentLoadBudget.MaxInputBytes),
                DocumentLoadBudget.Default with { MaxInputBytes = package.LongLength - 1 }),
            (nameof(DocumentLoadBudget.MaxPackageParts),
                DocumentLoadBudget.Default with { MaxPackageParts = 1 }),
            (nameof(DocumentLoadBudget.MaxInflatedBytes),
                DocumentLoadBudget.Default with { MaxInflatedBytes = 1 }),
            (nameof(DocumentLoadBudget.MaxPartBytes),
                DocumentLoadBudget.Default with { MaxPartBytes = 1 }),
            (nameof(DocumentLoadBudget.MaxXmlCharactersPerPart),
                DocumentLoadBudget.Default with { MaxXmlCharactersPerPart = 1 }),
            (nameof(DocumentLoadBudget.MaxXmlNodes),
                DocumentLoadBudget.Default with { MaxXmlNodes = 1 }),
            (nameof(DocumentLoadBudget.MaxXmlDepth),
                DocumentLoadBudget.Default with { MaxXmlDepth = 1 }),
        ];

        foreach ((string name, DocumentLoadBudget budget) in cases)
        {
            DocumentLoadLimitException error = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
                await Workbook.LoadAsync(
                    new MemoryStream(package, writable: false),
                    new LoadOptions { Budget = budget },
                    TestContext.Current.CancellationToken));
            Assert.Equal(name, error.LimitName);
            Assert.True(error.Observed > error.Limit);
            Assert.Contains($"'{name}'", error.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task CellSharedStringAndTextCeilingsAreChargedFromStoredSourceOnce()
    {
        byte[] package = await WorkbookPackageAsync("first", "second");

        await AssertLimitAsync(
            package,
            DocumentLoadBudget.Default with { MaxCells = 1 },
            nameof(DocumentLoadBudget.MaxCells));
        await AssertLimitAsync(
            package,
            DocumentLoadBudget.Default with { MaxSharedStrings = 1 },
            nameof(DocumentLoadBudget.MaxSharedStrings));
        await AssertLimitAsync(
            package,
            DocumentLoadBudget.Default with { MaxCellTextCharacters = 5 },
            nameof(DocumentLoadBudget.MaxCellTextCharacters));
        await AssertLimitAsync(
            package,
            DocumentLoadBudget.Default with { MaxTotalCellTextCharacters = 10 },
            nameof(DocumentLoadBudget.MaxTotalCellTextCharacters));
    }

    [Fact]
    public async Task RepeatedSharedStringReferencesDoNotRechargeTextOrIndexDigits()
    {
        byte[] package = await WorkbookPackageAsync("same", "same");
        Workbook loaded = await Workbook.LoadAsync(
            new MemoryStream(package, writable: false),
            new LoadOptions
            {
                Budget = DocumentLoadBudget.Default with
                {
                    MaxCellTextCharacters = 4,
                    MaxTotalCellTextCharacters = 4,
                },
            },
            TestContext.Current.CancellationToken);

        Assert.Equal("same", loaded.Sheets[0]["A1"].GetText());
        Assert.Equal("same", loaded.Sheets[0]["B1"].GetText());
    }

    [Fact]
    public async Task SemanticMediaAndEmbeddedRelationshipsWorkOutsideConventionalFolders()
    {
        byte[] original = await WorkbookPackageAsync("resources");
        byte[] media = Enumerable.Repeat((byte)0x42, 31).ToArray();
        byte[] embedded = Enumerable.Repeat((byte)0x51, 37).ToArray();
        byte[] package = AddRelatedParts(
            original,
            new RelatedPart(
                "resources/My%20Preview.dat",
                "image/png",
                "https://cellwright.example/relationships/resource",
                media),
            new RelatedPart(
                "objects/payload.dat",
                "application/octet-stream",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                embedded));

        await AssertLimitAsync(
            package,
            DocumentLoadBudget.Default with { MaxMediaBytes = media.Length - 1 },
            nameof(DocumentLoadBudget.MaxMediaBytes));
        await AssertLimitAsync(
            package,
            DocumentLoadBudget.Default with { MaxEmbeddedObjectBytes = embedded.Length - 1 },
            nameof(DocumentLoadBudget.MaxEmbeddedObjectBytes));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RoleOnlyMediaIsClassifiedInTransitionalAndStrictPackagesWithoutDuplicateCharging(
        bool strict)
    {
        byte[] media = Enumerable.Repeat((byte)0x62, 43).ToArray();
        byte[] package = AddRelatedParts(
            await WorkbookPackageAsync("role-only"),
            new RelatedPart(
                "custom/assets/preview.bin",
                "application/octet-stream",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                media,
                RelationshipCopies: 2));
        if (strict)
            package = MakeStrict(package);

        await AssertLimitAsync(
            package,
            DocumentLoadBudget.Default with { MaxMediaBytes = media.Length - 1 },
            nameof(DocumentLoadBudget.MaxMediaBytes));

        Workbook loaded = await Workbook.LoadAsync(
            new MemoryStream(package, writable: false),
            new LoadOptions
            {
                Budget = DocumentLoadBudget.Default with { MaxTotalMediaBytes = media.Length },
            },
            TestContext.Current.CancellationToken);
        Assert.Equal("role-only", loaded.Sheets[0]["A1"].GetText());
    }

    [Fact]
    public async Task LogicalAliasCollisionClassifiesEveryPhysicalMediaEntry()
    {
        byte[] small = Enumerable.Repeat((byte)0x31, 11).ToArray();
        byte[] oversized = Enumerable.Repeat((byte)0x32, 47).ToArray();
        byte[] package = AddRelatedParts(
            await WorkbookPackageAsync("alias collision"),
            new RelatedPart(
                "custom/assets/My%20Pic.bin",
                "application/octet-stream",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                small));
        package = AddPhysicalPart(
            package,
            "custom/assets/My Pic.bin",
            "application/octet-stream",
            oversized);

        await AssertLimitAsync(
            package,
            DocumentLoadBudget.Default with { MaxMediaBytes = small.Length },
            nameof(DocumentLoadBudget.MaxMediaBytes));
    }

    [Fact]
    public async Task ExactDuplicateTargetsChargeEveryPhysicalMediaEntryOnce()
    {
        byte[] first = Enumerable.Repeat((byte)0x41, 13).ToArray();
        byte[] second = Enumerable.Repeat((byte)0x42, 37).ToArray();
        const string entryName = "custom/assets/duplicate.bin";
        byte[] package = AddRelatedParts(
            await WorkbookPackageAsync("exact duplicate"),
            new RelatedPart(
                entryName,
                "application/octet-stream",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                first));
        package = AddPhysicalPart(
            package,
            entryName,
            "application/octet-stream",
            second);

        await AssertLimitAsync(
            package,
            DocumentLoadBudget.Default with
            {
                MaxMediaBytes = second.Length,
                MaxTotalMediaBytes = second.Length,
            },
            nameof(DocumentLoadBudget.MaxTotalMediaBytes));
    }

    [Fact]
    public async Task ArbitraryPlusXmlPartsUseCharacterNodeAndDepthBudgets()
    {
        string xml = "<root payload=\"" + new string('x', 4_096) + "\">" +
                     string.Concat(Enumerable.Repeat("<n>", 32)) +
                     string.Concat(Enumerable.Repeat("</n>", 32)) +
                     string.Concat(Enumerable.Repeat("<item/>", 100)) +
                     "</root>";
        byte[] package = CustomXmlPackage(("custom/payload.data", xml));

        await AssertPartLimitAsync(
            package,
            DocumentLoadBudget.Default with { MaxXmlCharactersPerPart = 1_024 },
            nameof(DocumentLoadBudget.MaxXmlCharactersPerPart),
            "/custom/payload.data");
        await AssertPartLimitAsync(
            package,
            DocumentLoadBudget.Default with { MaxXmlNodes = 80 },
            nameof(DocumentLoadBudget.MaxXmlNodes),
            "/custom/payload.data");
        await AssertPartLimitAsync(
            package,
            DocumentLoadBudget.Default with { MaxXmlDepth = 16 },
            nameof(DocumentLoadBudget.MaxXmlDepth),
            "/custom/payload.data");
    }

    [Fact]
    public async Task MalformedXmlValidPrefixStillConsumesSharedNodeBudget()
    {
        string prefix = "<root>" + string.Concat(Enumerable.Repeat("<item/>", 30));
        string valid = prefix + "</root>";
        byte[] package = CustomXmlPackage(
            ("custom/bad.data", prefix),
            ("custom/good.data", valid));
        var budget = DocumentLoadBudget.Default with { MaxXmlNodes = 50 };

        await using OpcPackage opc = await OpcPackage.OpenReadAsync(
            new MemoryStream(package, writable: false),
            leaveOpen: false,
            TestContext.Current.CancellationToken,
            budget);
        await opc.ReadPartBytesAsync("/custom/bad.data", TestContext.Current.CancellationToken);
        DocumentLoadLimitException error = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
            await opc.ReadPartBytesAsync("/custom/good.data", TestContext.Current.CancellationToken));

        Assert.Equal(nameof(DocumentLoadBudget.MaxXmlNodes), error.LimitName);
    }

    [Fact]
    public async Task DuplicatePhysicalRelationshipEntriesCannotShareXmlValidationState()
    {
        const string relationships =
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"urn:test\" Target=\"https://example.com\" TargetMode=\"External\"/>" +
            "</Relationships>";
        byte[] package = CustomXmlPackage(
            ("_rels/.rels", relationships),
            ("_rels/.rels", relationships.Replace("rId1", "rId2", StringComparison.Ordinal)));

        DocumentLoadLimitException error = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
            await OpcPackage.OpenReadAsync(
                new MemoryStream(package, writable: false),
                leaveOpen: false,
                TestContext.Current.CancellationToken,
                DocumentLoadBudget.Default with { MaxXmlNodes = 7 }));

        Assert.Equal(nameof(DocumentLoadBudget.MaxXmlNodes), error.LimitName);
    }

    [Fact]
    public async Task CancellationInterruptsOneGiantXmlAttribute()
    {
        string xml = "<root payload=\"" + new string('x', 1024 * 1024) + "\"/>";
        byte[] package = CustomXmlPackage(("custom/giant.data", xml));
        using var cancellation = new CancellationTokenSource();
        await using var source = new CancellingSeekableStream(package);
        await using OpcPackage opc = await OpcPackage.OpenReadAsync(
            source,
            leaveOpen: true,
            CancellationToken.None,
            DocumentLoadBudget.Default);

        source.Arm(cancellation);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await opc.ReadPartBytesAsync("/custom/giant.data", cancellation.Token));
    }

    [Fact]
    public async Task EncodedSeparatorsAndDotSegmentsDoNotChangePartIdentity()
    {
        Assert.Equal("/safe/My File.xml", OpcPath.Unescape("/safe/My%20File.xml"));
        Assert.Equal("/safe%2Fsecret.xml", OpcPath.Unescape("/safe%2Fsecret.xml"));
        Assert.Equal("/safe/%2E%2E/secret.xml", OpcPath.Unescape("/safe/%2E%2E/secret.xml"));

        byte[] package = CustomXmlPackage(
            ("safe/My%20File.data", "<root/>"),
            ("safe%2Fsecret.data", "<root/>"),
            ("safe/%2E%2E/secret.data", "<root/>"));
        await using OpcPackage opc = await OpcPackage.OpenReadAsync(
            new MemoryStream(package, writable: false),
            leaveOpen: false,
            TestContext.Current.CancellationToken,
            DocumentLoadBudget.Default);

        Assert.True(opc.PartExists("/safe/My File.data"));
        Assert.True(opc.PartExists("/safe%2Fsecret.data"));
        Assert.True(opc.PartExists("/safe/%2E%2E/secret.data"));
        Assert.False(opc.PartExists("/safe/secret.data"));
        Assert.False(opc.PartExists("/secret.data"));
    }

    [Fact]
    public async Task NonSeekableInputStopsAtFirstBytePastLimitAndHonorsCancellation()
    {
        byte[] package = await WorkbookPackageAsync("stream");
        await using var input = new NonSeekableReadStream(package);
        DocumentLoadLimitException error = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
            await Workbook.LoadAsync(
                input,
                new LoadOptions
                {
                    LeaveOpen = true,
                    Budget = DocumentLoadBudget.Default with { MaxInputBytes = package.Length - 1 },
                },
                TestContext.Current.CancellationToken));
        Assert.Equal(nameof(DocumentLoadBudget.MaxInputBytes), error.LimitName);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await Workbook.LoadAsync(
                new NonSeekableReadStream(package),
                cancellationToken: cancellation.Token));
    }

    private static async Task AssertLimitAsync(
        byte[] package, DocumentLoadBudget budget, string expected)
    {
        DocumentLoadLimitException error = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
            await Workbook.LoadAsync(
                new MemoryStream(package, writable: false),
                new LoadOptions { Budget = budget },
                TestContext.Current.CancellationToken));
        Assert.Equal(expected, error.LimitName);
    }

    private static async Task AssertPartLimitAsync(
        byte[] package,
        DocumentLoadBudget budget,
        string expected,
        string part)
    {
        await using OpcPackage opc = await OpcPackage.OpenReadAsync(
            new MemoryStream(package, writable: false),
            leaveOpen: false,
            TestContext.Current.CancellationToken,
            budget);
        DocumentLoadLimitException error = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
            await opc.ReadPartBytesAsync(part, TestContext.Current.CancellationToken));
        Assert.Equal(expected, error.LimitName);
    }

    private static async Task<byte[]> WorkbookPackageAsync(params string[] values)
    {
        Workbook workbook = Workbook.Create();
        Worksheet sheet = workbook.Sheets.Add("Data");
        for (int i = 0; i < values.Length; i++)
            sheet[1, i + 1] = values[i];

        using var stream = new MemoryStream();
        await workbook.SaveAsync(stream, cancellationToken: TestContext.Current.CancellationToken);
        return stream.ToArray();
    }

    private static byte[] AddRelatedParts(byte[] package, params RelatedPart[] parts)
    {
        using var output = new MemoryStream();
        using (var destination = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        using (var source = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read))
        {
            foreach (ZipArchiveEntry entry in source.Entries)
            {
                ZipArchiveEntry copy = destination.CreateEntry(entry.FullName, CompressionLevel.SmallestSize);
                using Stream target = copy.Open();
                using Stream input = entry.Open();
                if (entry.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                {
                    using var reader = new StreamReader(input, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                    string declarations = string.Concat(parts.Select(static part =>
                        $"<Override PartName=\"/{part.EntryName}\" ContentType=\"{part.ContentType}\"/>"));
                    string xml = reader.ReadToEnd().Replace(
                        "</Types>", declarations + "</Types>", StringComparison.Ordinal);
                    using var writer = new StreamWriter(target, new UTF8Encoding(false), leaveOpen: true);
                    writer.Write(xml);
                    continue;
                }

                if (entry.FullName.Equals("xl/_rels/workbook.xml.rels", StringComparison.OrdinalIgnoreCase))
                {
                    using var reader = new StreamReader(input, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                    int relationshipIndex = 0;
                    string declarations = string.Concat(parts.SelectMany(part =>
                        Enumerable.Range(0, part.RelationshipCopies).Select(_ =>
                            $"<Relationship Id=\"rIdBudget{++relationshipIndex}\" Type=\"{part.RelationshipType}\" Target=\"../{part.EntryName}\"/>")));
                    string xml = reader.ReadToEnd().Replace(
                        "</Relationships>", declarations + "</Relationships>", StringComparison.Ordinal);
                    using var writer = new StreamWriter(target, new UTF8Encoding(false), leaveOpen: true);
                    writer.Write(xml);
                    continue;
                }

                input.CopyTo(target);
            }

            foreach (RelatedPart part in parts)
            {
                ZipArchiveEntry added = destination.CreateEntry(part.EntryName, CompressionLevel.SmallestSize);
                using Stream content = added.Open();
                content.Write(part.Content);
            }
        }

        return output.ToArray();
    }

    private static byte[] CustomXmlPackage(params (string EntryName, string Xml)[] parts)
    {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            ZipArchiveEntry contentTypes = archive.CreateEntry("[Content_Types].xml");
            using (var writer = new StreamWriter(contentTypes.Open(), new UTF8Encoding(false)))
            {
                writer.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
                foreach ((string name, _) in parts)
                {
                    writer.Write($"<Override PartName=\"/{name}\" ContentType=\"application/vnd.cellwright.test+xml\"/>");
                }
                writer.Write("</Types>");
            }

            foreach ((string name, string xml) in parts)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.SmallestSize);
                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.Write(xml);
            }
        }

        return output.ToArray();
    }

    private static byte[] AddPhysicalPart(
        byte[] package,
        string entryName,
        string contentType,
        byte[] content)
    {
        using var output = new MemoryStream();
        using (var destination = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        using (var source = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read))
        {
            foreach (ZipArchiveEntry entry in source.Entries)
            {
                ZipArchiveEntry copy = destination.CreateEntry(entry.FullName, CompressionLevel.SmallestSize);
                using Stream target = copy.Open();
                using Stream input = entry.Open();
                if (entry.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                {
                    using var reader = new StreamReader(input, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                    string xml = reader.ReadToEnd().Replace(
                        "</Types>",
                        $"<Override PartName=\"/{entryName}\" ContentType=\"{contentType}\"/></Types>",
                        StringComparison.Ordinal);
                    using var writer = new StreamWriter(target, new UTF8Encoding(false), leaveOpen: true);
                    writer.Write(xml);
                    continue;
                }

                input.CopyTo(target);
            }

            ZipArchiveEntry added = destination.CreateEntry(entryName, CompressionLevel.SmallestSize);
            using Stream addedContent = added.Open();
            addedContent.Write(content);
        }

        return output.ToArray();
    }

    private static byte[] MakeStrict(byte[] package)
    {
        using var output = new MemoryStream();
        using (var destination = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        using (var source = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read))
        {
            foreach (ZipArchiveEntry entry in source.Entries)
            {
                ZipArchiveEntry copy = destination.CreateEntry(entry.FullName, CompressionLevel.SmallestSize);
                using Stream target = copy.Open();
                using Stream input = entry.Open();
                if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                    !entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                {
                    input.CopyTo(target);
                    continue;
                }

                using var reader = new StreamReader(input, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                string xml = reader.ReadToEnd()
                    .Replace(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        "http://purl.oclc.org/ooxml/officeDocument/relationships",
                        StringComparison.Ordinal)
                    .Replace(
                        "http://schemas.openxmlformats.org/package/2006/relationships",
                        "http://purl.oclc.org/ooxml/package/relationships",
                        StringComparison.Ordinal)
                    .Replace(
                        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                        "http://purl.oclc.org/ooxml/spreadsheetml/main",
                        StringComparison.Ordinal);
                using var writer = new StreamWriter(target, new UTF8Encoding(false), leaveOpen: true);
                writer.Write(xml);
            }
        }

        return output.ToArray();
    }

    private sealed record RelatedPart(
        string EntryName,
        string ContentType,
        string RelationshipType,
        byte[] Content,
        int RelationshipCopies = 1);

    private sealed class CancellingSeekableStream(byte[] content) : Stream
    {
        private readonly MemoryStream _inner = new(content, writable: false);
        private CancellationTokenSource? _cancellation;

        public void Arm(CancellationTokenSource cancellation) => _cancellation = cancellation;

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position
        {
            get => _inner.Position;
            set => _inner.Position = value;
        }

        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count)
        {
            int read = _inner.Read(buffer, offset, count);
            _cancellation?.Cancel();
            return read;
        }

        public override int Read(Span<byte> buffer)
        {
            int read = _inner.Read(buffer);
            _cancellation?.Cancel();
            return read;
        }

        public override async ValueTask<int> ReadAsync(
            Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            int read = await _inner.ReadAsync(buffer, cancellationToken);
            _cancellation?.Cancel();
            return read;
        }

        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class NonSeekableReadStream(byte[] content) : Stream
    {
        private readonly MemoryStream _inner = new(content, writable: false);

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override ValueTask<int> ReadAsync(
            Memory<byte> buffer, CancellationToken cancellationToken = default) =>
            _inner.ReadAsync(buffer, cancellationToken);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
