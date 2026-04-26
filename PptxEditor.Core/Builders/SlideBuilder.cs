using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxEditor.Core.Builders;

public class SlideBuilder : ISlideBuilder
{
    private readonly SlidePart _slidePart;
    private readonly Slide _slide;
    private readonly ShapeTree _shapeTree;

    public SlidePart SlidePart => _slidePart;

    public SlideBuilder(SlidePart slidePart, Slide slide)
    {
        _slidePart = slidePart;
        _slide = slide;
        _shapeTree = slide.CommonSlideData!.ShapeTree!;
    }

    public ISlideBuilder AddTitle(string text)
    {
        var shape = CreateTextShape("Title", text, 0, 0, 7200000, 720000);
        _shapeTree.Append(shape);
        return this;
    }

    public ISlideBuilder AddSubtitle(string text)
    {
        var shape = CreateTextShape("Subtitle", text, 0, 720000, 7200000, 720000);
        _shapeTree.Append(shape);
        return this;
    }

    public ISlideBuilder AddText(string text)
    {
        var shape = CreateTextShape("Text", text, 0, 1440000, 7200000, 3600000);
        _shapeTree.Append(shape);
        return this;
    }

    public ISlideBuilder AddBulletList(IEnumerable<string> items)
    {
        var shape = CreateListShape("Bullet List", items, true);
        _shapeTree.Append(shape);
        return this;
    }

    public ISlideBuilder AddNumberedList(IEnumerable<string> items)
    {
        var shape = CreateListShape("Numbered List", items, false);
        _shapeTree.Append(shape);
        return this;
    }

    public ISlideBuilder AddImage(string imagePath)
    {
        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException($"Image not found: {imagePath}");
        }

        var ext = Path.GetExtension(imagePath).ToLowerInvariant();
        var imagePartType = ext switch
        {
            ".png" => ImagePartType.Png,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tiff" or ".tif" => ImagePartType.Tiff,
            ".svg" => ImagePartType.Svg,
            _ => ImagePartType.Jpeg
        };
        var imagePart = _slidePart.AddImagePart(imagePartType);
        using (var stream = new FileStream(imagePath, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        var imageId = GetNextShapeId();
        var picture = new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = imageId, Name = $"Image {imageId}" },
                new P.NonVisualPictureDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            new P.BlipFill(
                new A.Blip { Embed = _slidePart.GetIdOfPart(imagePart) },
                new A.Stretch(new A.FillRectangle())
            ),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0, Y = 1440000 },
                    new A.Extents { Cx = 7200000, Cy = 3600000 }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            )
        );

        _shapeTree.Append(picture);
        return this;
    }

    public ISlideBuilder AddTable(List<List<string>> rows)
    {
        if (rows.Count == 0 || rows[0].Count == 0)
        {
            throw new ArgumentException("Table must have at least one row and one column.");
        }

        var numRows = rows.Count;
        var numCols = rows[0].Count;

        var table = new A.Table(
            new A.TableProperties { FirstRow = true },
            new A.TableGrid(Enumerable.Range(0, numCols).Select(_ => new A.GridColumn { Width = 7200000 / numCols }))
        );

        foreach (var row in rows)
        {
            var tableRow = new A.TableRow { Height = 370840 };
            foreach (var cell in row)
            {
                var tableCell = new A.TableCell(
                    new A.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(cell)))
                    ),
                    new A.TableCellProperties()
                );
                tableRow.Append(tableCell);
            }
            table.Append(tableRow);
        }

        var graphicFrame = new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = GetNextShapeId(), Name = "Table" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            new P.Transform(
                new A.Offset { X = 0, Y = 1440000 },
                new A.Extents { Cx = 7200000, Cy = 3600000 }
            ),
            new A.Graphic(
                new A.GraphicData(table)
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }
            )
        );

        _shapeTree.Append(graphicFrame);
        return this;
    }

    public ISlideBuilder AddChart(ChartType type, Dictionary<string, int> data)
    {
        // For V1, we'll create a simple placeholder
        // Full chart implementation will be in Phase 3.3
        var shape = CreateTextShape("Chart", $"[{type} Chart: {string.Join(", ", data.Select(kvp => $"{kvp.Key}={kvp.Value}"))}]", 0, 1440000, 7200000, 3600000);
        _shapeTree.Append(shape);
        return this;
    }

    private P.Shape CreateTextShape(string name, string text, long x, long y, long cx, long cy)
    {
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = GetNextShapeId(), Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            ),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(text)))
            )
        );
    }

    private P.Shape CreateListShape(string name, IEnumerable<string> items, bool isBullet)
    {
        var textBody = new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle()
        );

        int index = 0;
        foreach (var item in items)
        {
            var paragraph = new A.Paragraph();
            
            if (isBullet)
            {
                paragraph.ParagraphProperties = new A.ParagraphProperties(
                    new A.BulletFont { Typeface = "Arial" },
                    new A.CharacterBullet { Char = "\u2022" }
                );
            }
            else
            {
                paragraph.ParagraphProperties = new A.ParagraphProperties(
                    new A.AutoNumberedBullet { Type = A.TextAutoNumberSchemeValues.ArabicPeriod }
                );
            }
            
            paragraph.Append(new A.Run(new A.Text(item)));
            textBody.Append(paragraph);
            index++;
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = GetNextShapeId(), Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0, Y = 1440000 },
                    new A.Extents { Cx = 7200000, Cy = 3600000 }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            ),
            textBody
        );
    }

    private uint _nextShapeId = 10; // Start after placeholder IDs
    private uint GetNextShapeId()
    {
        return _nextShapeId++;
    }
}
