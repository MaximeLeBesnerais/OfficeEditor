using System.IO.Compression;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxEditor.Core.Builders;
using XlsxEditor.Core.Exceptions;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Deterministic malformed-input tests for the public <see cref="WorkbookBuilder"/> Open
/// boundary only. All fixtures are generated in memory — no binary assets on disk.
/// Contract: any payload that is not a structurally valid XLSX workbook must fail during
/// Open with a typed <see cref="XlsxException"/> (original SDK/parse failure preserved as
/// the InnerException where one exists). No NullReferenceException, no opaque raw SDK
/// exception escapes the boundary; the caller's own stream is never disposed.
/// </summary>
public class WorkbookBuilderMalformedInputTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private const string ContentTypesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
        """;

    private const string PackageRelsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private const string WorkbookXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
          </sheets>
        </workbook>
        """;

    private const string WorkbookRelsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        </Relationships>
        """;

    private const string WorkbookWithoutSheetsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
        """;

    private const string WorksheetXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
        </worksheet>
        """;

    private const string CorruptWorkbookXml = """
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets>
        """;

    private const string WrongWorkbookXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
        """;

    private const string CorruptWorksheetXml = """
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
        """;

    private const string WrongWorksheetXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
        """;

    private const string WorkbookWithSheetWithoutIdXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheets>
            <sheet name="Sheet1" sheetId="1"/>
          </sheets>
        </workbook>
        """;

    private const string WorkbookWithUndeclaredSheetIdXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="Sheet1" sheetId="1" r:id="rId99"/>
          </sheets>
        </workbook>
        """;

    private const string StylesRelationshipRelsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
        """;

    private const string StylesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
        """;

    /// <summary>
    /// Every malformed byte payload class the boundary must reject, plus the SDK exception
    /// family preserved as the <see cref="XlsxException.InnerException"/>. A null expected
    /// inner type means the failure is our own structural validation (no SDK exception to
    /// preserve). These cover the full failure spectrum deterministically.
    /// </summary>
    public static IEnumerable<object?[]> MalformedBytePayloads()
    {
        yield return new object?[] { Array.Empty<byte>(), "empty package", null };
        yield return new object?[] { CreateRandomBytes(), "random non-archive bytes", typeof(System.IO.FileFormatException) };

        var valid = CreateValidWorkbook();
        yield return new object?[] { Truncate(valid, valid.Length / 2), "zip truncated in half", typeof(System.IO.FileFormatException) };
        yield return new object?[] { Truncate(valid, valid.Length - 24), "zip with end-of-central-directory removed", typeof(System.IO.FileFormatException) };

        yield return new object?[] { CreateValidDocx(), "valid DOCX opened as a spreadsheet (wrong main-part content type)", typeof(InvalidDataException) };
        yield return new object?[] { CreateOpcWithoutWorkbookPart(), "OPC package missing the workbook main part", typeof(InvalidOperationException) };
        yield return new object?[] { CreateOpcWithoutWorksheetPart(), "OPC package missing a referenced worksheet part", typeof(InvalidOperationException) };
    }

    [Theory]
    [MemberData(nameof(MalformedBytePayloads))]
    public void Open_ByteArray_MalformedPackage_ThrowsXlsxExceptionWithOriginalInner(
        byte[] payload, string scenario, Type? expectedInnerType)
    {
        var ex = Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));
        AssertInnerPreserved(ex, expectedInnerType, $"Open(byte[]) — {scenario}");
    }

    [Theory]
    [MemberData(nameof(MalformedBytePayloads))]
    public void Open_Stream_MalformedPackage_ThrowsXlsxExceptionWithOriginalInner(
        byte[] payload, string scenario, Type? expectedInnerType)
    {
        using var stream = new MemoryStream(payload);
        var ex = Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(stream));
        AssertInnerPreserved(ex, expectedInnerType, $"Open(Stream) — {scenario}");
    }

    [Fact]
    public void Open_Stream_MalformedBytes_ThrowsAndLeavesCallerOwnedStreamUsable()
    {
        var garbage = CreateRandomBytes();
        using var stream = new MemoryStream(garbage);

        Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(stream));

        // Caller retains ownership: the stream is not disposed, its length is unchanged,
        // and the failed Open consumed it exactly once (position parked at the end).
        Assert.True(stream.CanRead, "Open failure must not dispose the caller-owned stream.");
        Assert.Equal(garbage.Length, stream.Length);
        Assert.Equal(garbage.Length, stream.Position);

        // A subsequent valid XLSX can still be opened from the same caller-owned stream.
        var valid = CreateValidWorkbook();
        stream.SetLength(0);
        stream.Write(valid, 0, valid.Length);
        stream.Position = 0;

        using var builder = WorkbookBuilder.Open(stream);
        Assert.Equal(new List<string> { "Sheet1" }, builder.GetWorksheetNames());
    }

    [Fact]
    public void Open_ByteArray_StructuralValidationFailure_PropagatesXlsxExceptionUnchanged()
    {
        // The package is structurally valid OPC with a readable workbook part, but its
        // sheet is missing the r:id the SDK needs to resolve the worksheet — our own
        // structural validation rejects it. The typed exception must propagate unchanged:
        // not re-wrapped, no invented SDK inner exception, message intact.
        var payload = BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", WorkbookWithSheetWithoutIdXml),
            ("xl/_rels/workbook.xml.rels", WorkbookRelsXml),
            ("xl/worksheets/sheet1.xml", WorksheetXml));

        var ex = Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));

        Assert.Null(ex.InnerException);
        Assert.Contains("relationship id", ex.Message);
        Assert.Contains("Sheet1", ex.Message);
    }

    [Fact]
    public void Open_Stream_StructuralValidationFailure_LeavesCallerStreamUntouched()
    {
        var payload = BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", WorkbookWithSheetWithoutIdXml),
            ("xl/_rels/workbook.xml.rels", WorkbookRelsXml),
            ("xl/worksheets/sheet1.xml", WorksheetXml));
        using var stream = new MemoryStream(payload);

        var ex = Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(stream));

        Assert.Null(ex.InnerException);
        Assert.Contains("relationship id", ex.Message);

        // The caller-owned stream survives the failed open unchanged, and the boundary
        // stays usable afterwards.
        Assert.True(stream.CanRead, "Open failure must not dispose the caller-owned stream.");
        Assert.Equal(payload.Length, stream.Length);

        using var builder = WorkbookBuilder.Open(CreateValidWorkbook());
        Assert.Equal(new List<string> { "Sheet1" }, builder.GetWorksheetNames());
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("\t")]
    public void Open_Path_EmptyOrWhitespace_ThrowsArgumentException(string path)
    {
        var ex = Assert.Throws<ArgumentException>(() => WorkbookBuilder.Open(path));
        Assert.Equal("path", ex.ParamName);

        // The boundary remains usable after the failed calls.
        using var builder = WorkbookBuilder.Open(CreateValidWorkbook());
        Assert.Equal(new List<string> { "Sheet1" }, builder.GetWorksheetNames());
    }

    [Fact]
    public void Open_Stream_SourceReadFailure_PropagatesUnexpectedExceptionUnchanged()
    {
        // A source stream that throws mid-copy is neither a malformed package nor our
        // structural validation — it is an unexpected I/O failure and must propagate as
        // its original type, never be normalized into an XlsxException.
        using var broken = new ThrowingReadStream();

        var ex = Assert.Throws<IOException>(() => WorkbookBuilder.Open(broken));
        Assert.Contains("Simulated I/O failure", ex.Message);

        using var builder = WorkbookBuilder.Open(CreateValidWorkbook());
        Assert.Equal(new List<string> { "Sheet1" }, builder.GetWorksheetNames());
    }

    [Fact]
    public void Open_Path_PermissionOrPathError_PropagatesUnchanged()
    {
        // A directory is not a readable package: the SDK fails with a filesystem-level
        // exception. That is a path/permission error, not a malformed package, so the
        // boundary must clean up and rethrow the original type rather than normalizing it.
        var directory = Path.GetTempPath();

        var ex = Assert.ThrowsAny<Exception>(() => WorkbookBuilder.Open(directory));

        Assert.IsNotType<XlsxException>(ex);
        Assert.True(
            ex is UnauthorizedAccessException or IOException or DirectoryNotFoundException,
            $"expected a filesystem error, got {ex.GetType().Name}: '{ex.Message}'.");
    }

    [Fact]
    public void Open_Path_MalformedFile_ThrowsAndLeavesSourceFileIntact()
    {
        var path = CreateTempFile(CreateRandomBytes());
        Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(path));
        Assert.True(File.Exists(path), "failed Open(path) must not delete the source file.");
    }

    [Fact]
    public void Open_Path_MissingFile_ThrowsFileNotFoundException()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"missing_{Guid.NewGuid():N}.xlsx");
        var ex = Assert.Throws<FileNotFoundException>(() => WorkbookBuilder.Open(missing));
        Assert.Contains(Path.GetFileName(missing), ex.Message);
    }

    [Fact]
    public void Open_Path_ValidWorkbookStillOpens_AfterAMalformedFileFailure()
    {
        var garbagePath = CreateTempFile(CreateRandomBytes());
        Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(garbagePath));

        var validPath = CreateTempFile(CreateValidWorkbook());
        using var builder = WorkbookBuilder.Open(validPath);
        Assert.Equal(new List<string> { "Sheet1" }, builder.GetWorksheetNames());
    }

    [Fact]
    public void Open_ByteArray_ValidWorkbookStillOpens_AfterMalformedPayloadFailures()
    {
        foreach (var payload in new[]
                 {
                     Array.Empty<byte>(), CreateRandomBytes(), CreateValidDocx(),
                     CreateOpcWithoutWorkbookPart(), CreateOpcWithoutWorksheetPart()
                 })
        {
            Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));
        }

        using var builder = WorkbookBuilder.Open(CreateValidWorkbook());
        Assert.Equal(new List<string> { "Sheet1" }, builder.GetWorksheetNames());
    }

    [Fact]
    public void Open_OpcPackage_WorkbookWithoutSheets_OpensGracefullyWithZeroWorksheets()
    {
        // Positive control: a structurally valid workbook that merely has no <sheets>
        // must NOT be treated as an error.
        using var builder = WorkbookBuilder.Open(CreateOpcWorkbookWithoutSheets());
        Assert.Empty(builder.GetWorksheetNames());
    }

    [Fact]
    public void Open_NullInputs_ThrowArgumentNullExceptionWithParameterName()
    {
        var bytesEx = Assert.Throws<ArgumentNullException>(() => WorkbookBuilder.Open((byte[])null!));
        Assert.Equal("bytes", bytesEx.ParamName);

        var streamEx = Assert.Throws<ArgumentNullException>(() => WorkbookBuilder.Open((Stream)null!));
        Assert.Equal("stream", streamEx.ParamName);

        var pathEx = Assert.Throws<ArgumentNullException>(() => WorkbookBuilder.Open((string)null!));
        Assert.Equal("path", pathEx.ParamName);

        // The boundary remains usable after the failed calls.
        using var builder = WorkbookBuilder.Open(CreateValidWorkbook());
        Assert.Equal(new List<string> { "Sheet1" }, builder.GetWorksheetNames());
    }

    [Fact]
    public void Open_OpcPackage_WorkbookRootMissing_ThrowsXlsxException()
    {
        // The workbook part exists but its root element is absent (empty part bytes).
        var payload = BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", string.Empty));

        Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));
    }

    [Fact]
    public void Open_OpcPackage_WorkbookRootCorrupt_ThrowsXlsxExceptionWithXmlException()
    {
        var payload = BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", CorruptWorkbookXml));

        var ex = Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));
        Assert.True(ex.InnerException is XmlException, $"expected XmlException inner, got {ex.InnerException?.GetType().Name}");
    }

    [Fact]
    public void Open_OpcPackage_WorkbookRootWrongElement_ThrowsXlsxExceptionWithInvalidDataException()
    {
        var payload = BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", WrongWorkbookXml));

        var ex = Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));
        Assert.True(ex.InnerException is InvalidDataException, $"expected InvalidDataException inner, got {ex.InnerException?.GetType().Name}");
    }

    [Fact]
    public void Open_OpcPackage_WorksheetRootCorrupt_ThrowsXlsxExceptionWithXmlException()
    {
        var payload = BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", WorkbookXml),
            ("xl/_rels/workbook.xml.rels", WorkbookRelsXml),
            ("xl/worksheets/sheet1.xml", CorruptWorksheetXml));

        var ex = Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));
        Assert.True(ex.InnerException is XmlException, $"expected XmlException inner, got {ex.InnerException?.GetType().Name}");
    }

    [Fact]
    public void Open_OpcPackage_WorksheetRootWrongElement_ThrowsXlsxExceptionWithInvalidDataException()
    {
        var payload = BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", WorkbookXml),
            ("xl/_rels/workbook.xml.rels", WorkbookRelsXml),
            ("xl/worksheets/sheet1.xml", WrongWorksheetXml));

        var ex = Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));
        Assert.True(ex.InnerException is InvalidDataException, $"expected InvalidDataException inner, got {ex.InnerException?.GetType().Name}");
    }

    [Fact]
    public void Open_OpcPackage_WorksheetRootMissing_ThrowsXlsxException()
    {
        // The worksheet part exists but has no root element (empty part bytes).
        var payload = BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", WorkbookXml),
            ("xl/_rels/workbook.xml.rels", WorkbookRelsXml),
            ("xl/worksheets/sheet1.xml", string.Empty));

        Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));
    }

    [Fact]
    public void Open_OpcPackage_SheetWithoutRelationshipId_ThrowsXlsxException()
    {
        var payload = BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", WorkbookWithSheetWithoutIdXml),
            ("xl/_rels/workbook.xml.rels", WorkbookRelsXml),
            ("xl/worksheets/sheet1.xml", WorksheetXml));

        Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));
    }

    [Fact]
    public void Open_OpcPackage_UndeclaredRelationshipId_ThrowsXlsxExceptionWithOriginalInner()
    {
        var payload = BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", WorkbookWithUndeclaredSheetIdXml),
            ("xl/_rels/workbook.xml.rels", WorkbookRelsXml),
            ("xl/worksheets/sheet1.xml", WorksheetXml));

        var ex = Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));
        Assert.True(ex.InnerException is ArgumentOutOfRangeException,
            $"expected ArgumentOutOfRangeException inner, got {ex.InnerException?.GetType().Name}");
    }

    [Fact]
    public void Open_OpcPackage_RelationshipResolvesToNonWorksheetPart_ThrowsXlsxExceptionWithOriginalInner()
    {
        // Sheet r:id points at a styles part instead of a worksheet.
        var payload = BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", WorkbookXml),
            ("xl/_rels/workbook.xml.rels", StylesRelationshipRelsXml),
            ("xl/styles.xml", StylesXml));

        var ex = Assert.Throws<XlsxException>(() => WorkbookBuilder.Open(payload));
        Assert.NotNull(ex.InnerException);
    }

    /// <summary>
    /// The typed contract: an <see cref="XlsxException"/>, with the SDK's original failure
    /// preserved as the inner exception whenever the SDK actually raised it (a null expected
    /// inner means our own structural validation rejected the payload directly).
    /// </summary>
    private static void AssertInnerPreserved(XlsxException ex, Type? expectedInnerType, string scenario)
    {
        if (expectedInnerType is null)
        {
            Assert.Null(ex.InnerException);
            return;
        }

        Assert.NotNull(ex.InnerException);
        Assert.True(expectedInnerType.IsAssignableFrom(ex.InnerException.GetType()),
            $"[{scenario}] expected inner {expectedInnerType.Name}, got " +
            $"{ex.InnerException!.GetType().FullName}: '{ex.InnerException.Message}'.");
    }

    private static byte[] CreateValidWorkbook()
    {
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            workbookPart.Workbook = new Workbook(new Sheets(new Sheet
            {
                Name = "Sheet1",
                SheetId = 1u,
                Id = workbookPart.GetIdOfPart(worksheetPart)
            }));
            doc.Save();
        }
        return ms.ToArray();
    }

    private static byte[] CreateValidDocx()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            doc.Save();
        }
        return ms.ToArray();
    }

    private static byte[] CreateRandomBytes()
    {
        var random = new Random(0x5EED);
        var bytes = new byte[4096];
        random.NextBytes(bytes);
        // Force the signature away from the ZIP 'PK' magic so detection is unambiguous,
        // keeping this a deterministic "not an archive" payload regardless of PRNG output.
        bytes[0] = 0x00;
        bytes[1] = 0x01;
        return bytes;
    }

    /// <summary>
    /// A stream that advertises readability but throws on every read, simulating an
    /// unexpected I/O failure while the boundary copies the caller's content into its
    /// internal buffer.
    /// </summary>
    private sealed class ThrowingReadStream : Stream
    {
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => true;
        public override long Length => 0;
        public override long Position { get; set; }

        public override void Flush() { }

        public override int Read(byte[] buffer, int offset, int count)
            => throw new IOException("Simulated I/O failure during copy.");

        public override long Seek(long offset, SeekOrigin origin) => 0;

        public override void SetLength(long value) { }

        public override void Write(byte[] buffer, int offset, int count) { }
    }

    private static byte[] Truncate(byte[] bytes, int length) => bytes.Take(length).ToArray();

    private static byte[] CreateOpcWithoutWorkbookPart()
        => BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml));

    private static byte[] CreateOpcWithoutWorksheetPart()
        => BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", WorkbookXml),
            ("xl/_rels/workbook.xml.rels", WorkbookRelsXml));

    private static byte[] CreateOpcWorkbookWithoutSheets()
        => BuildOpcPackage(
            ("[Content_Types].xml", ContentTypesXml),
            ("_rels/.rels", PackageRelsXml),
            ("xl/workbook.xml", WorkbookWithoutSheetsXml));

    private static byte[] BuildOpcPackage(params (string Name, string Content)[] entries)
    {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, content) in entries)
            {
                var entry = archive.CreateEntry(name);
                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.Write(content);
            }
        }
        return ms.ToArray();
    }

    private string CreateTempFile(byte[] content)
    {
        var path = Path.Combine(Path.GetTempPath(), $"malformed_{Guid.NewGuid():N}.xlsx");
        File.WriteAllBytes(path, content);
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var file in _tempFiles)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }
    }
}
