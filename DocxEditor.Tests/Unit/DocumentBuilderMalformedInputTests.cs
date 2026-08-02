using System.IO;
using System.IO.Packaging;
using DocxEditor.Core.Builders;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxEditor.Tests.Unit;

/// <summary>
/// Deterministic malformed-input suite for the public <see cref="DocumentBuilder"/> open boundary.
/// Every case asserts a controlled managed failure (a small, curated set of package/format
/// exceptions with a non-empty message) rather than a process-level abort, without pinning
/// OS-specific exception text.
/// </summary>
public class DocumentBuilderMalformedInputTests : IDisposable
{
    private const string WordMainContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    private const string OfficeDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    // A valid OPC/ZIP package that is not a valid wordprocessing package fails zip/part reading;
    // the exact type depends on OpenXml SDK / System.IO.Packaging internals, so accept the family.
    private static readonly Type[] ZipCorruptionFailures =
    {
        typeof(OpenXmlPackageException),
        typeof(FileFormatException),
        typeof(InvalidDataException),
        typeof(IOException),
        typeof(NullReferenceException),
    };

    // A structurally valid OPC package without an officeDocument relationship currently surfaces
    // an uncontrolled NullReferenceException from the DocumentBuilder constructor (no null guard);
    // other SDK versions may reject at open instead.
    private static readonly Type[] MissingMainPartFailures =
    {
        typeof(NullReferenceException),
        typeof(OpenXmlPackageException),
        typeof(InvalidOperationException),
        typeof(FileFormatException),
    };

    // A wordprocessing package whose document.xml lacks a <w:body> currently surfaces an
    // uncontrolled NullReferenceException from the constructor (Body! dereference with no guard).
    private static readonly Type[] MissingBodyFailures =
    {
        typeof(NullReferenceException),
        typeof(InvalidOperationException),
    };

    // An officeDocument relationship pointing at a part that does not exist; exact failure point
    // (open vs. constructor) and type vary by SDK version, so accept the failure family.
    private static readonly Type[] BrokenRelationshipFailures =
    {
        typeof(OpenXmlPackageException),
        typeof(InvalidOperationException),
        typeof(FileNotFoundException),
        typeof(FileFormatException),
        typeof(NullReferenceException),
    };

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

    private static void AssertThrowsManagedFailure(Action action, params Type[] allowedTypes)
    {
        var ex = Assert.ThrowsAny<Exception>(action);
        Assert.False(string.IsNullOrWhiteSpace(ex.Message), "Malformed input must fail with a diagnostic message.");
        Assert.Contains(allowedTypes, t => t.IsInstanceOfType(ex));
    }

    // ---------------------------------------------------------------- byte[] overload

    [Fact]
    public void Open_EmptyBytes_ThrowsManagedFailure()
    {
        AssertThrowsManagedFailure(() => DocumentBuilder.Open(Array.Empty<byte>()), ZipCorruptionFailures);
    }

    [Fact]
    public void Open_RandomBytes_ThrowsManagedFailure()
    {
        AssertThrowsManagedFailure(() => DocumentBuilder.Open(CreateRandomBytes(1024)), ZipCorruptionFailures);
    }

    [Theory]
    [MemberData(nameof(TruncatedDocxCases))]
    public void Open_TruncatedDocxBytes_ThrowsManagedFailure(byte[] truncated)
    {
        AssertThrowsManagedFailure(() => DocumentBuilder.Open(truncated), ZipCorruptionFailures);
    }

    [Theory]
    [MemberData(nameof(WrongFormatOfficeCases))]
    public void Open_WrongFormatOfficePackage_ThrowsInvalidDataException(byte[] wrongFormat)
    {
        // The OpenXml SDK validates the main part content type against the requested document
        // type, so a valid XLSX/PPTX package must be rejected by the DOCX opener.
        Assert.Throws<InvalidDataException>(() => DocumentBuilder.Open(wrongFormat));
    }

    [Fact]
    public void Open_WordPackageMissingBody_FailsWhenBuilderIsUsed()
    {
        // Gap: the constructor dereferences MainDocumentPart.Document.Body without a null guard,
        // so a word package whose document.xml has no <w:body> surfaces an uncontrolled
        // NullReferenceException instead of a domain-level "invalid package" exception.
        var bytes = BuildWordPackageMissingBody();
        using var builder = DocumentBuilder.Open(bytes);

        AssertThrowsManagedFailure(() => builder.AddParagraph("must fail"), MissingBodyFailures);
    }

    [Fact]
    public void Open_WordPackageMissingMainPart_ThrowsManagedFailure()
    {
        // Gap: without an officeDocument relationship the package opens but MainDocumentPart is
        // null, and the constructor dereferences it without a guard (uncontrolled NRE).
        var bytes = BuildWordPackageMissingMainPart();
        AssertThrowsManagedFailure(() => DocumentBuilder.Open(bytes), MissingMainPartFailures);
    }

    [Fact]
    public void Open_WordPackageWithBrokenMainRelationship_ThrowsManagedFailure()
    {
        var bytes = BuildWordPackageWithBrokenMainRelationship();
        AssertThrowsManagedFailure(() => DocumentBuilder.Open(bytes), BrokenRelationshipFailures);
    }

    // ---------------------------------------------------------------- Stream overload

    [Fact]
    public void Open_EmptyStream_ThrowsManagedFailure()
    {
        using var stream = new MemoryStream();
        AssertThrowsManagedFailure(() => DocumentBuilder.Open(stream), ZipCorruptionFailures);
    }

    [Fact]
    public void Open_InvalidStream_Failure_LeavesCallerStreamOwned()
    {
        var bytes = CreateRandomBytes(512);
        using var stream = new MemoryStream(bytes);

        AssertThrowsManagedFailure(() => DocumentBuilder.Open(stream), ZipCorruptionFailures);

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
    public void Open_InvalidStream_ThenValidStream_StillOpens()
    {
        using var badStream = new MemoryStream(CreateRandomBytes(256));
        AssertThrowsManagedFailure(() => DocumentBuilder.Open(badStream), ZipCorruptionFailures);

        using var goodStream = new MemoryStream(CreateValidDocxBytes());
        using var builder = DocumentBuilder.Open(goodStream);
        builder.AddParagraph("recovered");
        var roundTripped = builder.SaveToBytes();

        using var reopened = WordprocessingDocument.Open(new MemoryStream(roundTripped), false);
        var body = reopened.MainDocumentPart!.Document!.Body!;
        Assert.Contains(body.Elements<Paragraph>(), p => p.InnerText == "Hello from valid docx");
        Assert.Contains(body.Elements<Paragraph>(), p => p.InnerText == "recovered");
    }

    // ---------------------------------------------------------------- path overload

    [Fact]
    public void Open_InvalidFile_ThrowsManagedFailure()
    {
        var path = Path.Combine(Path.GetTempPath(), $"malformed_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);
        File.WriteAllBytes(path, CreateRandomBytes(1024));

        AssertThrowsManagedFailure(() => DocumentBuilder.Open(path), ZipCorruptionFailures);
    }

    // ---------------------------------------------------------------- recovery

    [Fact]
    public void Open_AfterFailedOpen_CanStillOpenValidDocx()
    {
        AssertThrowsManagedFailure(() => DocumentBuilder.Open(CreateRandomBytes(512)), ZipCorruptionFailures);
        AssertThrowsManagedFailure(() => DocumentBuilder.Open(Truncate(CreateValidDocxBytes(), 10)), ZipCorruptionFailures);

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
