using System.IO;
using System.IO.Packaging;
using System.Xml;
using DocxEditor.Core.Builders;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeEditor.Core.Exceptions;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Deterministic malformed-input suite for the public <see cref="DocumentBuilder"/> open boundary.
/// The open boundary is the single place malformed DOCX input is rejected: corrupt,
/// wrong-format, sparse, or structurally-invalid packages fail during <c>Open</c> with a
/// <see cref="OfficeEditorException"/> (the original SDK/package failure preserved as the inner
/// exception) — never an uncontrolled NullReferenceException and never deferred to later builder
/// use. Null and empty arguments are rejected up front as ArgumentNullException/ArgumentException,
/// a failed open never claims the caller's stream, and ordinary path/permission errors
/// (<see cref="FileNotFoundException"/>, <see cref="DirectoryNotFoundException"/>, ...) propagate
/// unchanged instead of being normalized.
/// </summary>
public class DocumentBuilderMalformedInputTests : IDisposable
{
    private const string WordMainContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    private const string OfficeDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    private readonly List<string> _tempFiles = new();

    // ---------------------------------------------------------------- fixture helpers

    private static byte[] CreateValidDocxBytes()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text("Hello from valid docx")))));
            doc.Save();
        }

        return stream.ToArray();
    }

    private static byte[] CreateValidXlsxBytes()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var part = doc.AddWorkbookPart();
            part.Workbook = new Workbook();
            doc.Save();
        }

        return stream.ToArray();
    }

    private static byte[] CreateValidPptxBytes()
    {
        using var stream = new MemoryStream();
        using (var doc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var part = doc.AddPresentationPart();
            part.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation();
            doc.Save();
        }

        return stream.ToArray();
    }

    private static byte[] CreateRandomBytes(int length)
    {
        var buffer = new byte[length];
        new Random(0x5EED).NextBytes(buffer);
        return buffer;
    }

    private static byte[] Truncate(byte[] bytes, int length) => bytes[..length];

    // Structurally valid OPC package with a main document part whose document.xml has no <w:body>.
    private static byte[] BuildWordPackageMissingBody()
    {
        using var stream = new MemoryStream();
        using (var package = Package.Open(stream, FileMode.Create))
        {
            var documentPart = package.CreatePart(new Uri("/word/document.xml", UriKind.Relative), WordMainContentType);
            using (var writer = new StreamWriter(documentPart.GetStream(), new System.Text.UTF8Encoding(false)))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" +
                             @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""/>");
            }

            package.CreateRelationship(documentPart.Uri, TargetMode.Internal, OfficeDocumentRelationship);
        }

        return stream.ToArray();
    }

    // Structurally valid OPC package with no officeDocument relationship (no main document part).
    private static byte[] BuildWordPackageMissingMainPart()
    {
        using var stream = new MemoryStream();
        using (var package = Package.Open(stream, FileMode.Create))
        {
            package.CreatePart(new Uri("/custom.xml", UriKind.Relative), "application/xml");
        }

        return stream.ToArray();
    }

    // OPC package whose officeDocument relationship targets a part that does not exist.
    private static byte[] BuildWordPackageWithBrokenMainRelationship()
    {
        using var stream = new MemoryStream();
        using (var package = Package.Open(stream, FileMode.Create))
        {
            package.CreatePart(new Uri("/custom.xml", UriKind.Relative), "application/xml");
            package.CreateRelationship(new Uri("/word/document.xml", UriKind.Relative), TargetMode.Internal, OfficeDocumentRelationship);
        }

        return stream.ToArray();
    }

    // OPC package whose main document part contains malformed (non-well-formed) XML.
    private static byte[] BuildWordPackageWithMalformedXml()
    {
        using var stream = new MemoryStream();
        using (var package = Package.Open(stream, FileMode.Create))
        {
            var documentPart = package.CreatePart(new Uri("/word/document.xml", UriKind.Relative), WordMainContentType);
            using (var writer = new StreamWriter(documentPart.GetStream(), new System.Text.UTF8Encoding(false)))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" +
                             @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body></w:document>");
            }

            package.CreateRelationship(documentPart.Uri, TargetMode.Internal, OfficeDocumentRelationship);
        }

        return stream.ToArray();
    }

    /// <summary>
    /// The public open boundary must reject every malformed input with a domain exception carrying
    /// a diagnostic message, regardless of which overload or SDK failure mode is hit.
    /// </summary>
    private static void AssertOpenFailure(Action action, string scenario)
    {
        var ex = Assert.Throws<OfficeEditorException>(() => action());
        Assert.False(string.IsNullOrWhiteSpace(ex.Message), $"[{scenario}] malformed input must fail with a diagnostic message.");
    }

    // ---------------------------------------------------------------- byte[] overload

    [Fact]
    public void Open_EmptyBytes_ThrowsOfficeEditorException()
    {
        AssertOpenFailure(() => DocumentBuilder.Open(Array.Empty<byte>()), "empty byte array");
    }

    [Fact]
    public void Open_RandomBytes_ThrowsOfficeEditorException()
    {
        AssertOpenFailure(() => DocumentBuilder.Open(CreateRandomBytes(1024)), "random non-archive bytes");
    }

    [Fact]
    public void Open_GarbageBytes_ThrowsDomainExceptionWrappingFileFormatException()
    {
        // Non-archive input surfaces as System.IO.FileFormatException from the package reader and
        // must be normalized to the domain exception with the original preserved as the inner.
        var ex = Assert.Throws<OfficeEditorException>(() => DocumentBuilder.Open(CreateRandomBytes(1024)));
        Assert.IsType<System.IO.FileFormatException>(ex.InnerException);
    }

    [Theory]
    [MemberData(nameof(TruncatedDocxCases))]
    public void Open_TruncatedDocxBytes_ThrowsOfficeEditorException(byte[] truncated)
    {
        AssertOpenFailure(() => DocumentBuilder.Open(truncated), "truncated DOCX");
    }

    [Theory]
    [MemberData(nameof(WrongFormatOfficeCases))]
    public void Open_WrongFormatOfficePackage_ThrowsDomainExceptionWrappingInvalidData(byte[] wrongFormat)
    {
        // The OpenXml SDK validates the main part content type against the requested document
        // type, so a valid XLSX/PPTX package must be rejected by the DOCX opener. The SDK's
        // InvalidDataException is normalized to the domain exception with the original preserved.
        var ex = Assert.Throws<OfficeEditorException>(() => DocumentBuilder.Open(wrongFormat));
        Assert.IsType<InvalidDataException>(ex.InnerException);
    }

    [Fact]
    public void Open_WordPackageMissingBody_ThrowsDuringOpen()
    {
        // Gap fixed: the constructor no longer dereferences Document.Body with a null-forgiving
        // operator. A word package whose document.xml has no <w:body> is rejected at Open time as
        // a domain exception — before any builder operation can run.
        var bytes = BuildWordPackageMissingBody();

        var ex = Assert.Throws<OfficeEditorException>(() => DocumentBuilder.Open(bytes));
        Assert.Contains("body", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Open_WordPackageMissingMainPart_ThrowsDuringOpen()
    {
        // Gap fixed: without an officeDocument relationship the package opens but MainDocumentPart
        // is null; the constructor now rejects it with a domain exception instead of an NRE.
        var bytes = BuildWordPackageMissingMainPart();
        AssertOpenFailure(() => DocumentBuilder.Open(bytes), "OPC package without an officeDocument relationship");
    }

    [Fact]
    public void Open_WordPackageWithBrokenMainRelationship_ThrowsDuringOpen()
    {
        var bytes = BuildWordPackageWithBrokenMainRelationship();
        AssertOpenFailure(() => DocumentBuilder.Open(bytes), "OPC package whose main relationship targets a missing part");
    }

    [Fact]
    public void Open_WordPackageWithMalformedXml_ThrowsDomainExceptionWrappingXmlException()
    {
        // A well-formed OPC package whose main part is not well-formed XML surfaces as
        // XmlException when the root element is loaded; it must be normalized to the domain
        // exception with the original preserved as the inner.
        var bytes = BuildWordPackageWithMalformedXml();
        var ex = Assert.Throws<OfficeEditorException>(() => DocumentBuilder.Open(bytes));
        Assert.IsType<XmlException>(ex.InnerException);
    }

    [Fact]
    public void Open_NullBytes_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => DocumentBuilder.Open((byte[])null!));
    }

    // ---------------------------------------------------------------- Stream overload

    [Fact]
    public void Open_EmptyStream_ThrowsOfficeEditorException()
    {
        using var stream = new MemoryStream();
        AssertOpenFailure(() => DocumentBuilder.Open(stream), "empty stream");
    }

    [Fact]
    public void Open_InvalidStream_Failure_LeavesCallerStreamOwned()
    {
        var bytes = CreateRandomBytes(512);
        using var stream = new MemoryStream(bytes);

        AssertOpenFailure(() => DocumentBuilder.Open(stream), "garbage stream");

        // DocumentBuilder.Open(Stream) copies the source to an internal buffer before opening;
        // the caller retains ownership, so the stream must stay open, readable, and reusable
        // even after the open fails.
        Assert.True(stream.CanRead, "Caller stream must not be disposed after a failed Open.");
        Assert.True(stream.CanSeek, "Caller stream must remain seekable after a failed Open.");
        Assert.Equal(bytes.Length, stream.Position);

        stream.Position = 0;
        var buffer = new byte[bytes.Length];
        var read = stream.Read(buffer, 0, buffer.Length);
        Assert.Equal(bytes.Length, read);
    }

    [Fact]
    public void Open_ValidStream_Success_LeavesCallerStreamOwned()
    {
        using var stream = new MemoryStream(CreateValidDocxBytes());

        using (var builder = DocumentBuilder.Open(stream))
        {
            builder.AddParagraph("caller owns the stream");
            Assert.True(stream.CanRead, "Caller stream must stay open while the builder is alive.");
            Assert.True(stream.CanSeek, "Caller stream must stay seekable while the builder is alive.");
        }

        // Disposing the builder must not dispose the caller's stream: it can still be read,
        // seeked, and handed to a fresh open.
        Assert.True(stream.CanRead, "Disposing the builder must not dispose the caller stream.");
        Assert.True(stream.CanSeek, "Disposing the builder must not seal the caller stream.");

        stream.Position = 0;
        using var reopened = WordprocessingDocument.Open(stream, false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        Assert.Contains(body.Elements<Paragraph>(), p => p.InnerText == "Hello from valid docx");
    }

    [Fact]
    public void Open_InvalidStream_ThenValidStream_StillOpens()
    {
        using var badStream = new MemoryStream(CreateRandomBytes(256));
        AssertOpenFailure(() => DocumentBuilder.Open(badStream), "garbage stream");

        using var goodStream = new MemoryStream(CreateValidDocxBytes());
        using var builder = DocumentBuilder.Open(goodStream);
        builder.AddParagraph("recovered");
        var roundTripped = builder.SaveToBytes();

        using var reopened = WordprocessingDocument.Open(new MemoryStream(roundTripped), false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        Assert.Contains(body.Elements<Paragraph>(), p => p.InnerText == "Hello from valid docx");
        Assert.Contains(body.Elements<Paragraph>(), p => p.InnerText == "recovered");
    }

    [Fact]
    public void Open_NullStream_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => DocumentBuilder.Open((Stream)null!));
    }

    // ---------------------------------------------------------------- path overload

    [Fact]
    public void Open_InvalidFile_ThrowsOfficeEditorException()
    {
        var path = Path.Combine(Path.GetTempPath(), $"malformed_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);
        File.WriteAllBytes(path, CreateRandomBytes(1024));

        AssertOpenFailure(() => DocumentBuilder.Open(path), "garbage file on disk");
    }

    [Fact]
    public void Open_NullPath_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => DocumentBuilder.Open((string)null!));
    }

    [Fact]
    public void Open_EmptyPath_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() => DocumentBuilder.Open(string.Empty));
    }

    [Fact]
    public void Open_WhitespacePath_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() => DocumentBuilder.Open("   "));
    }

    [Fact]
    public void Open_NonexistentPath_ThrowsFileNotFoundException()
    {
        // Ordinary path errors are not malformed-package failures and must propagate unchanged.
        var path = Path.Combine(Path.GetTempPath(), $"missing_{Guid.NewGuid():N}.docx");
        Assert.Throws<FileNotFoundException>(() => DocumentBuilder.Open(path));
    }

    [Fact]
    public void Open_MissingDirectoryPath_ThrowsDirectoryNotFoundException()
    {
        var path = Path.Combine(Path.GetTempPath(), $"no_such_dir_{Guid.NewGuid():N}", "doc.docx");
        Assert.Throws<DirectoryNotFoundException>(() => DocumentBuilder.Open(path));
    }

    // ---------------------------------------------------------------- create overload

    [Fact]
    public void Create_NullPath_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => DocumentBuilder.Create((string)null!));
    }

    [Fact]
    public void Create_EmptyPath_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() => DocumentBuilder.Create(string.Empty));
    }

    [Fact]
    public void Create_WhitespacePath_ThrowsArgumentExceptionAndCreatesNoFile()
    {
        // Whitespace-only names are legal files on some platforms; Create must reject them up
        // front so a stray whitespace path never materializes a junk document file.
        Assert.Throws<ArgumentException>(() => DocumentBuilder.Create("   "));
        Assert.False(File.Exists("   "), "A rejected whitespace path must not create a file.");
    }

    [Fact]
    public void Create_PathIsDirectory_ThrowsWithoutNormalizing()
    {
        // Create is not a malformed-package boundary: ordinary path/permission errors must
        // propagate unchanged instead of being wrapped in the domain exception.
        var dir = Path.Combine(Path.GetTempPath(), $"create_dir_{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        try
        {
            var ex = Assert.ThrowsAny<Exception>(() => DocumentBuilder.Create(dir));
            Assert.IsNotType<OfficeEditorException>(ex);
        }
        finally
        {
            Directory.Delete(dir);
        }
    }

    [Fact]
    public void Create_MissingParentDirectory_ThrowsDirectoryNotFoundException()
    {
        var path = Path.Combine(Path.GetTempPath(), $"no_such_dir_{Guid.NewGuid():N}", "doc.docx");
        Assert.Throws<DirectoryNotFoundException>(() => DocumentBuilder.Create(path));
    }

    [Fact]
    public void Create_AfterFailedCreate_CanStillCreateValidDocx()
    {
        var badPath = Path.Combine(Path.GetTempPath(), $"no_such_dir_{Guid.NewGuid():N}", "doc.docx");
        Assert.Throws<DirectoryNotFoundException>(() => DocumentBuilder.Create(badPath));

        var goodPath = Path.Combine(Path.GetTempPath(), $"create_ok_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(goodPath);
        using (var builder = DocumentBuilder.Create(goodPath))
        {
            builder.AddParagraph("after failed create");
            builder.Save();
        }

        Assert.True(File.Exists(goodPath));
        using var doc = WordprocessingDocument.Open(goodPath, false);
        Assert.Contains("after failed create", doc.MainDocumentPart!.Document!.Body!.InnerText);
    }

    // ---------------------------------------------------------------- recovery

    [Fact]
    public void Open_AfterFailedOpen_CanStillOpenValidDocx()
    {
        AssertOpenFailure(() => DocumentBuilder.Open(CreateRandomBytes(512)), "random bytes");
        AssertOpenFailure(() => DocumentBuilder.Open(Truncate(CreateValidDocxBytes(), 10)), "truncated DOCX");

        using var builder = DocumentBuilder.Open(CreateValidDocxBytes());
        builder.AddParagraph("after failure");
        var roundTripped = builder.SaveToBytes();

        using var reopened = WordprocessingDocument.Open(new MemoryStream(roundTripped), false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        Assert.Contains(body.Elements<Paragraph>(), p => p.InnerText == "Hello from valid docx");
        Assert.Contains(body.Elements<Paragraph>(), p => p.InnerText == "after failure");
    }

    // ---------------------------------------------------------------- theory data

    public static IEnumerable<object[]> TruncatedDocxCases()
    {
        var valid = CreateValidDocxBytes();
        yield return new object[] { Truncate(valid, valid.Length / 2) };
        yield return new object[] { Truncate(valid, valid.Length / 3) };
        yield return new object[] { Truncate(valid, Math.Max(0, valid.Length - 22)) };
    }

    public static IEnumerable<object[]> WrongFormatOfficeCases()
    {
        yield return new object[] { CreateValidXlsxBytes() };
        yield return new object[] { CreateValidPptxBytes() };
    }

    public void Dispose()
    {
        foreach (var file in _tempFiles)
        {
            try
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
            catch (Exception)
            {
                // best-effort temp cleanup
            }
        }
    }
}
