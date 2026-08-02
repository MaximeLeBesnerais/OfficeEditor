using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxEditor.Core.Builders;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Deterministic malformed-input tests for the public <see cref="WorkbookBuilder"/> Open
/// boundary only. All fixtures are generated in memory — no binary assets on disk.
/// The assertions prove a controlled, managed failure (a well-known SDK exception family)
/// without depending on OS-specific exception message text.
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

    /// <summary>
    /// Every malformed byte payload class the boundary must reject. The DOCX and the two
    /// hand-built OPC packages are structurally valid ZIP/OPC containers that are still not
    /// an openable workbook, so this suite covers the full failure spectrum deterministically.
    /// </summary>
    public static IEnumerable<object[]> MalformedBytePayloads()
    {
        yield return new object[] { Array.Empty<byte>(), "empty package" };
        yield return new object[] { CreateRandomBytes(), "random non-archive bytes" };

        var valid = CreateValidWorkbook();
        yield return new object[] { Truncate(valid, valid.Length / 2), "zip truncated in half" };
        yield return new object[] { Truncate(valid, valid.Length - 24), "zip with end-of-central-directory removed" };

        yield return new object[] { CreateValidDocx(), "valid DOCX opened as a spreadsheet (wrong main-part content type)" };
        yield return new object[] { CreateOpcWithoutWorkbookPart(), "OPC package missing the workbook main part" };
        yield return new object[] { CreateOpcWithoutWorksheetPart(), "OPC package missing a referenced worksheet part" };
    }

    [Theory]
    [MemberData(nameof(MalformedBytePayloads))]
    public void Open_ByteArray_MalformedPackage_ThrowsControlledManagedException(byte[] payload, string scenario)
    {
        var ex = Record.Exception(() => WorkbookBuilder.Open(payload));
        AssertControlledManagedFailure(ex, $"Open(byte[]) — {scenario}");
    }

    [Theory]
    [MemberData(nameof(MalformedBytePayloads))]
    public void Open_Stream_MalformedPackage_ThrowsControlledManagedException(byte[] payload, string scenario)
    {
        using var stream = new MemoryStream(payload);
        var ex = Record.Exception(() => WorkbookBuilder.Open(stream));
        AssertControlledManagedFailure(ex, $"Open(Stream) — {scenario}");
    }

    [Fact]
    public void Open_Stream_MalformedBytes_ThrowsAndLeavesCallerOwnedStreamUsable()
    {
        var garbage = CreateRandomBytes();
        using var stream = new MemoryStream(garbage);

        var ex = Record.Exception(() => WorkbookBuilder.Open(stream));
        AssertControlledManagedFailure(ex, "Open(Stream) caller-owned garbage stream");

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
    public void Open_Path_MalformedFile_ThrowsAndLeavesSourceFileIntact()
    {
        var path = CreateTempFile(CreateRandomBytes());
        var ex = Record.Exception(() => WorkbookBuilder.Open(path));
        AssertControlledManagedFailure(ex, "Open(path) garbage file");
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
        var ex = Record.Exception(() => WorkbookBuilder.Open(garbagePath));
        AssertControlledManagedFailure(ex, "garbage file on disk");

        var validPath = CreateTempFile(CreateValidWorkbook());
        using var builder = WorkbookBuilder.Open(validPath);
        Assert.Equal(new List<string> { "Sheet1" }, builder.GetWorksheetNames());
    }

    [Fact]
    public void Open_OpcPackage_WorkbookWithoutSheets_OpensGracefullyWithZeroWorksheets()
    {
        // Positive control: a structurally valid workbook that merely has no <sheets>
        // must NOT be treated as an error — LoadExistingWorksheets returns early.
        using var builder = WorkbookBuilder.Open(CreateOpcWorkbookWithoutSheets());
        Assert.Empty(builder.GetWorksheetNames());
    }

    [Fact]
    public void Open_NullInputs_ThrowWithoutCorruptingFutureOpens()
    {
        // Gap: the public boundary does not validate null args. Open(byte[]) and
        // Open(Stream) leak the raw NullReferenceException from the MemoryStream ctor /
        // CopyTo; only Open(path) is validated (ArgumentNullException) by the SDK.
        Assert.Throws<NullReferenceException>(() => WorkbookBuilder.Open((byte[])null!));
        Assert.Throws<NullReferenceException>(() => WorkbookBuilder.Open((Stream)null!));
        Assert.Throws<ArgumentNullException>(() => WorkbookBuilder.Open((string)null!));

        // The boundary remains usable after the failed calls.
        using var builder = WorkbookBuilder.Open(CreateValidWorkbook());
        Assert.Equal(new List<string> { "Sheet1" }, builder.GetWorksheetNames());
    }

    /// <summary>
    /// Proves the open failed with a controlled managed exception from the well-known
    /// SDK/BCL family — never an unmanaged/fatal condition, and without overfitting to a
    /// single exception type or to OS-specific message text.
    /// </summary>
    private static void AssertControlledManagedFailure(Exception? ex, string scenario)
    {
        Assert.True(ex is not null, $"[{scenario}] Open returned a builder instead of failing on malformed input.");
        Assert.True(
            ex is FileFormatException
                or InvalidDataException
                or IOException
                or OpenXmlPackageException
                or ArgumentException
                or ArgumentOutOfRangeException
                or InvalidOperationException
                or KeyNotFoundException
                or NullReferenceException,
            $"[{scenario}] Open failed with {ex?.GetType().FullName} ('{ex?.Message}') — not a controlled managed " +
            "failure (allowed: FileFormatException, InvalidDataException, IOException, OpenXmlPackageException, " +
            "ArgumentException, ArgumentOutOfRangeException, InvalidOperationException, KeyNotFoundException, " +
            "NullReferenceException).");
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
