using PptxEditor.Core.Builders;
using PptxEditor.Core.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace DocxEditor.Tests.Unit;

public class PptxAnatomizerTests : IDisposable
{
    private readonly string _testFilePath = Path.Combine(Path.GetTempPath(), $"test_pptx_anatomy_{Guid.NewGuid()}.pptx");
    private readonly string _testImagePath = Path.Combine(Path.GetTempPath(), $"test_image_{Guid.NewGuid()}.jpg");

    public PptxAnatomizerTests()
    {
        // Create a minimal test image (1x1 pixel JPEG)
        CreateMinimalJpeg(_testImagePath);
    }

    private void CreateMinimalJpeg(string path)
    {
        // Minimal valid JPEG: 1x1 pixel, gray
        var jpegBytes = new byte[]
        {
            0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01,
            0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0xFF, 0xDB, 0x00, 0x43,
            0x00, 0x08, 0x06, 0x06, 0x07, 0x06, 0x05, 0x08, 0x07, 0x07, 0x07, 0x09,
            0x09, 0x08, 0x0A, 0x0C, 0x14, 0x0D, 0x0C, 0x0B, 0x0B, 0x0C, 0x19, 0x12,
            0x13, 0x0F, 0x14, 0x1D, 0x1A, 0x1F, 0x1E, 0x1D, 0x1A, 0x1C, 0x1C, 0x20,
            0x24, 0x2E, 0x27, 0x20, 0x22, 0x2C, 0x23, 0x1C, 0x1C, 0x28, 0x37, 0x29,
            0x2C, 0x30, 0x31, 0x34, 0x34, 0x34, 0x1F, 0x27, 0x39, 0x3D, 0x38, 0x32,
            0x3C, 0x2E, 0x33, 0x34, 0x32, 0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00, 0x01,
            0x00, 0x01, 0x01, 0x01, 0x11, 0x00, 0xFF, 0xC4, 0x00, 0x1F, 0x00, 0x00,
            0x01, 0x05, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
            0x09, 0x0A, 0x0B, 0xFF, 0xC4, 0x00, 0xB5, 0x10, 0x00, 0x02, 0x01, 0x03,
            0x03, 0x02, 0x04, 0x03, 0x05, 0x05, 0x04, 0x04, 0x00, 0x00, 0x01, 0x7D,
            0x01, 0x02, 0x03, 0x00, 0x04, 0x11, 0x05, 0x12, 0x21, 0x31, 0x41, 0x06,
            0x13, 0x51, 0x61, 0x07, 0x22, 0x71, 0x14, 0x32, 0x81, 0x91, 0xA1, 0x08,
            0x23, 0x42, 0xB1, 0xC1, 0x15, 0x52, 0xD1, 0xF0, 0x24, 0x33, 0x62, 0x72,
            0x82, 0x09, 0x0A, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x25, 0x26, 0x27, 0x28,
            0x29, 0x2A, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3A, 0x43, 0x44, 0x45,
            0x46, 0x47, 0x48, 0x49, 0x4A, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59,
            0x5A, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69, 0x6A, 0x73, 0x74, 0x75,
            0x76, 0x77, 0x78, 0x79, 0x7A, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89,
            0x8A, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9A, 0xA2, 0xA3,
            0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6,
            0xB7, 0xB8, 0xB9, 0xBA, 0xC2, 0xC3, 0xC4, 0xC5, 0xC6, 0xC7, 0xC8, 0xC9,
            0xCA, 0xD2, 0xD3, 0xD4, 0xD5, 0xD6, 0xD7, 0xD8, 0xD9, 0xDA, 0xE1, 0xE2,
            0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8, 0xE9, 0xEA, 0xF1, 0xF2, 0xF3, 0xF4,
            0xF5, 0xF6, 0xF7, 0xF8, 0xF9, 0xFA, 0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01,
            0x00, 0x00, 0x3F, 0x00, 0xFB, 0xD5, 0xDB, 0x20, 0xB8, 0xF7, 0xFF, 0xD9
        };
        File.WriteAllBytes(path, jpegBytes);
    }

    [Fact]
    public void Analyze_ShouldDetectTextShapes()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Test Title");
            builder.CurrentSlide.AddText("Test Body");
            builder.Save();
        }

        // Act
        List<SlideAnatomy> anatomy;
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            anatomy = builder.Analyze();
        }

        // Assert
        Assert.Single(anatomy);
        Assert.True(anatomy[0].Elements.Count >= 2);
        
        var textElements = anatomy[0].Elements.Where(e => e.Type == "Text").ToList();
        Assert.True(textElements.Count >= 2);
        Assert.Contains(textElements, e => e.Text == "Test Title");
        Assert.Contains(textElements, e => e.Text == "Test Body");
    }

    [Fact]
    public void Analyze_ShouldDetectTables()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTable(new List<List<string>>
            {
                new() { "Header 1", "Header 2" },
                new() { "Cell 1", "Cell 2" }
            });
            builder.Save();
        }

        // Act
        List<SlideAnatomy> anatomy;
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            anatomy = builder.Analyze();
        }

        // Assert
        Assert.Single(anatomy);
        var tableElement = anatomy[0].Elements.FirstOrDefault(e => e.Type == "Table");
        Assert.NotNull(tableElement);
        Assert.NotNull(tableElement.TableData);
        Assert.Equal(2, tableElement.TableData.Count);
        Assert.Equal("Header 1", tableElement.TableData[0][0]);
        Assert.Equal("Cell 2", tableElement.TableData[1][1]);
    }

    [Fact]
    public void Analyze_ShouldDetectImages()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddImage(_testImagePath);
            builder.Save();
        }

        // Act
        List<SlideAnatomy> anatomy;
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            anatomy = builder.Analyze();
        }

        // Assert
        Assert.Single(anatomy);
        var imageElement = anatomy[0].Elements.FirstOrDefault(e => e.Type == "Image");
        Assert.NotNull(imageElement);
        Assert.True(imageElement.Id > 0);
        // Image name is auto-generated as "Image {id}" by SlideBuilder
        Assert.StartsWith("Image", imageElement.Name);
    }

    [Fact]
    public void Analyze_ShouldAssignUniqueIds()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Title");
            builder.CurrentSlide.AddText("Body");
            builder.CurrentSlide.AddTable(new List<List<string>>
            {
                new() { "A", "B" }
            });
            builder.Save();
        }

        // Act
        List<SlideAnatomy> anatomy;
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            anatomy = builder.Analyze();
        }

        // Assert - should have at least 3 elements (title, text, table)
        Assert.True(anatomy[0].Elements.Count >= 3, $"Expected at least 3 elements, found {anatomy[0].Elements.Count}");
        var ids = anatomy[0].Elements.Select(e => e.Id).ToList();
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Fact]
    public void ReplaceElement_ShouldUpdateText()
    {
        // Arrange
        uint titleId;
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Original Title");
            builder.Save();
        }

        // Get the ID of the title element
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            var anatomy = builder.Analyze();
            titleId = anatomy[0].Elements.First(e => e.Text == "Original Title").Id;
            
            // Act
            builder.ReplaceElement(titleId, "New Title");
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("New Title", text);
        Assert.DoesNotContain("Original Title", text);
    }

    [Fact]
    public void ReplaceTable_ShouldUpdateTableData()
    {
        // Arrange
        uint tableId;
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTable(new List<List<string>>
            {
                new() { "Old 1", "Old 2" },
                new() { "Old 3", "Old 4" }
            });
            builder.Save();
        }

        // Get the ID of the table element
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            var anatomy = builder.Analyze();
            tableId = anatomy[0].Elements.First(e => e.Type == "Table").Id;
            
            // Act
            builder.ReplaceTable(tableId, new List<List<string>>
            {
                new() { "New 1", "New 2" },
                new() { "New 3", "New 4" }
            });
            builder.Save();
        }

        // Assert
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var text = slidePart.Slide!.InnerText;
        Assert.Contains("New 1", text);
        Assert.Contains("New 4", text);
        Assert.DoesNotContain("Old 1", text);
    }

    [Fact]
    public void ReplaceImage_ShouldUpdateImage()
    {
        // Arrange
        uint imageId;
        var newImagePath = Path.Combine(Path.GetTempPath(), $"new_image_{Guid.NewGuid()}.jpg");
        CreateMinimalJpeg(newImagePath);

        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddImage(_testImagePath);
            builder.Save();
        }

        // Get the ID of the image element
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            var anatomy = builder.Analyze();
            imageId = anatomy[0].Elements.First(e => e.Type == "Image").Id;
            
            // Act
            builder.ReplaceImage(imageId, newImagePath);
            builder.Save();
        }

        // Assert - verify the image part was updated by checking it's still valid
        using var doc = PresentationDocument.Open(_testFilePath, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var imageParts = slidePart.ImageParts.ToList();
        Assert.Single(imageParts);

        // Cleanup
        File.Delete(newImagePath);
    }

    [Fact]
    public void Analyze_ShouldReturnCorrectLocationFormat()
    {
        // Arrange
        using (var builder = PresentationBuilder.Create(_testFilePath))
        {
            builder.AddSlide();
            builder.CurrentSlide.AddTitle("Title");
            builder.Save();
        }

        // Act
        List<SlideAnatomy> anatomy;
        using (var builder = PresentationBuilder.Open(_testFilePath))
        {
            anatomy = builder.Analyze();
        }

        // Assert
        var titleElement = anatomy[0].Elements.First(e => e.Text == "Title");
        Assert.Equal("slide:1:shape:Title", titleElement.Location);
    }

    public void Dispose()
    {
        if (File.Exists(_testFilePath))
        {
            File.Delete(_testFilePath);
        }
        if (File.Exists(_testImagePath))
        {
            File.Delete(_testImagePath);
        }
    }
}
