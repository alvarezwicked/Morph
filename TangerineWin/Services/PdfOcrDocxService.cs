using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ImageMagick;
using UglyToad.PdfPig;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace TangerineWin.Services;

public sealed class PdfOcrDocxService
{
    private readonly PdfOcrService _ocrService;

    private readonly PdfOcrTableDetectionService
        _tableDetectionService;

    public PdfOcrDocxService()
    {
        _ocrService =
            new PdfOcrService();

        _tableDetectionService =
            new PdfOcrTableDetectionService();
    }

    public async Task<string> ConvertAsync(
        string sourcePath)
    {
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException(
                "PDF file was not found.",
                sourcePath);
        }

        var outputPath =
            CreateUniqueOutputPath(
                sourcePath);

        using var pdf =
            PdfDocument.Open(
                sourcePath);

        using var wordDocument =
            WordprocessingDocument.Create(
                outputPath,
                WordprocessingDocumentType.Document);

        var mainPart =
            wordDocument.AddMainDocumentPart();

        mainPart.Document =
            new Document();

        var body =
            mainPart.Document.AppendChild(
                new Body());

        var imageNumber = 1;

        for (var pageIndex = 0;
             pageIndex < pdf.NumberOfPages;
             pageIndex++)
        {
            var pngBytes =
                await _ocrService
                    .RenderPageToPngAsync(
                        sourcePath,
                        pageIndex);

            //
            // Сначала сохраняем визуальный оригинал.
            //
            AddPageImage(
                mainPart,
                body,
                pngBytes,
                imageNumber);

            imageNumber++;

            //
            // OCR.
            //
            var ocrResult =
                await _ocrService
                    .RecognizePngAsync(
                        pngBytes);

            var tables =
                _tableDetectionService
                    .Detect(
                        ocrResult.Words);

            AddOcrHeading(
                body);

            if (ocrResult.Words.Count > 0)
            {
                AddStructuredOcrContent(
                    body,
                    ocrResult.Words,
                    tables);
            }
            else if (!string.IsNullOrWhiteSpace(
                         ocrResult.Text))
            {
                // Fallback, если TSV по какой-то
                // причине оказался пустым.
                AddPlainOcrText(
                    body,
                    ocrResult.Text);
            }

            if (pageIndex <
                pdf.NumberOfPages - 1)
            {
                body.Append(
                    new Paragraph(
                        new Run(
                            new Break
                            {
                                Type =
                                    BreakValues.Page
                            })));
            }
        }

        mainPart.Document.Save();

        return outputPath;
    }

    private static void AddStructuredOcrContent(
        Body body,
        IReadOnlyList<PdfOcrService.OcrWord> words,
        IReadOnlyList<PdfOcrTableDetectionService.OcrTable> tables)
    {
        var tableWords =
            tables
                .SelectMany(table =>
                    table.Words)
                .ToHashSet();

        var normalLines =
            words
                .Where(word =>
                    !tableWords.Contains(word))
                .GroupBy(word =>
                    new
                    {
                        word.BlockNumber,
                        word.ParagraphNumber,
                        word.LineNumber
                    })
                .Select(group =>
                {
                    var ordered =
                        group
                            .OrderBy(word =>
                                word.Left)
                            .ToArray();

                    return new OcrTextLine(
                        ordered.Min(word =>
                            word.Top),
                        string.Join(
                            " ",
                            ordered.Select(word =>
                                word.Text)));
                })
                .Where(line =>
                    !string.IsNullOrWhiteSpace(
                        line.Text))
                .ToArray();

        var blocks =
            new List<OcrContentBlock>();

        foreach (var line in normalLines)
        {
            blocks.Add(
                new OcrContentBlock(
                    line.Top,
                    line.Text,
                    null));
        }

        foreach (var table in tables)
        {
            blocks.Add(
                new OcrContentBlock(
                    table.Top,
                    null,
                    table));
        }

        // У Tesseract начало координат сверху слева,
        // поэтому меньший Top идёт раньше.
        foreach (var block in
                 blocks.OrderBy(block =>
                     block.Top))
        {
            if (block.Table is not null)
            {
                AddWordTable(
                    body,
                    block.Table.Rows);
            }
            else if (!string.IsNullOrWhiteSpace(
                         block.Text))
            {
                AddOcrParagraph(
                    body,
                    block.Text);
            }
        }
    }

    private static void AddOcrHeading(
        Body body)
    {
        var run =
            new Run(
                new Text(
                    "Распознанный текст"));

        run.RunProperties =
            new RunProperties(
                new Bold(),
                new FontSize
                {
                    Val = "20"
                });

        var paragraph =
            new Paragraph(
                run);

        paragraph.ParagraphProperties =
            new ParagraphProperties(
                new SpacingBetweenLines
                {
                    Before = "160",
                    After = "100"
                });

        body.Append(
            paragraph);
    }

    private static void AddOcrParagraph(
        Body body,
        string text)
    {
        var run =
            new Run(
                new Text(
                    text)
                {
                    Space =
                        SpaceProcessingModeValues.Preserve
                });

        run.RunProperties =
            new RunProperties(
                new FontSize
                {
                    Val = "20"
                });

        var paragraph =
            new Paragraph(
                run);

        paragraph.ParagraphProperties =
            new ParagraphProperties(
                new SpacingBetweenLines
                {
                    Before = "0",
                    After = "80"
                });

        body.Append(
            paragraph);
    }

    private static void AddPlainOcrText(
        Body body,
        string text)
    {
        var lines =
            text.Split(
                new[]
                {
                    "\r\n",
                    "\n"
                },
                StringSplitOptions.RemoveEmptyEntries);

        foreach (var line in lines)
        {
            AddOcrParagraph(
                body,
                line);
        }
    }

    private static void AddWordTable(
        Body body,
        IReadOnlyList<IReadOnlyList<string>> rows)
    {
        if (rows.Count < 2)
            return;

        var columnCount =
            rows
                .Select(row =>
                    row.Count)
                .DefaultIfEmpty(0)
                .Max();

        if (columnCount < 2)
            return;

        var table =
            new DocumentFormat.OpenXml.Wordprocessing.Table();

        table.AppendChild(
            new TableProperties(

                new TableWidth
                {
                    Width = "0",
                    Type =
                        TableWidthUnitValues.Auto
                },

                new TableBorders(

                    new TopBorder
                    {
                        Val = BorderValues.Single,
                        Size = 4U
                    },

                    new LeftBorder
                    {
                        Val = BorderValues.Single,
                        Size = 4U
                    },

                    new BottomBorder
                    {
                        Val = BorderValues.Single,
                        Size = 4U
                    },

                    new RightBorder
                    {
                        Val = BorderValues.Single,
                        Size = 4U
                    },

                    new InsideHorizontalBorder
                    {
                        Val = BorderValues.Single,
                        Size = 4U
                    },

                    new InsideVerticalBorder
                    {
                        Val = BorderValues.Single,
                        Size = 4U
                    })));

        foreach (var sourceRow in rows)
        {
            var tableRow =
                new TableRow();

            for (var columnIndex = 0;
                 columnIndex < columnCount;
                 columnIndex++)
            {
                var value =
                    columnIndex <
                    sourceRow.Count
                        ? sourceRow[columnIndex]
                        : string.Empty;

                var cell =
                    new TableCell(
                        new Paragraph(
                            new Run(
                                new Text(
                                    value)
                                {
                                    Space =
                                        SpaceProcessingModeValues
                                            .Preserve
                                })));

                cell.Append(
                    new TableCellProperties(
                        new TableCellWidth
                        {
                            Type =
                                TableWidthUnitValues.Auto
                        }));

                tableRow.Append(
                    cell);
            }

            table.Append(
                tableRow);
        }

        body.Append(
            table);

        body.Append(
            new Paragraph());
    }

    private static void AddPageImage(
        MainDocumentPart mainPart,
        Body body,
        byte[] pngBytes,
        int imageNumber)
    {
        var imagePart =
            mainPart.AddImagePart(
                ImagePartType.Png);

        using (
            var stream =
                new MemoryStream(
                    pngBytes))
        {
            imagePart.FeedData(
                stream);
        }

        var relationshipId =
            mainPart.GetIdOfPart(
                imagePart);

        long widthEmu;
        long heightEmu;

        using (
            var image =
                new MagickImage(
                    pngBytes))
        {
            var width =
                (double)image.Width;

            var height =
                (double)image.Height;

            const long emusPerPixel =
                9525;

            widthEmu =
                (long)(
                    width *
                    emusPerPixel);

            heightEmu =
                (long)(
                    height *
                    emusPerPixel);

            const long maxWidthEmu =
                5_669_280;

            const long maxHeightEmu =
                6_500_000;

            var scale =
                Math.Min(
                    1.0,
                    Math.Min(
                        (double)maxWidthEmu /
                        widthEmu,

                        (double)maxHeightEmu /
                        heightEmu));

            widthEmu =
                (long)(
                    widthEmu *
                    scale);

            heightEmu =
                (long)(
                    heightEmu *
                    scale);
        }

        var drawing =
            CreateDrawing(
                relationshipId,
                widthEmu,
                heightEmu,
                imageNumber);

        var paragraph =
            new Paragraph(
                new Run(
                    drawing));

        paragraph.ParagraphProperties =
            new ParagraphProperties(
                new Justification
                {
                    Val =
                        JustificationValues.Center
                });

        body.Append(
            paragraph);
    }

    private static Drawing CreateDrawing(
        string relationshipId,
        long widthEmu,
        long heightEmu,
        int imageNumber)
    {
        return new Drawing(
            new DW.Inline(

                new DW.Extent
                {
                    Cx = widthEmu,
                    Cy = heightEmu
                },

                new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },

                new DW.DocProperties
                {
                    Id =
                        (UInt32Value)(
                            (uint)imageNumber),

                    Name =
                        $"OCR page {imageNumber}"
                },

                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks
                    {
                        NoChangeAspect = true
                    }),

                new A.Graphic(
                    new A.GraphicData(

                        new PIC.Picture(

                            new PIC.NonVisualPictureProperties(

                                new PIC.NonVisualDrawingProperties
                                {
                                    Id = 0U,

                                    Name =
                                        $"OCR page {imageNumber}"
                                },

                                new PIC.NonVisualPictureDrawingProperties()),

                            new PIC.BlipFill(

                                new A.Blip
                                {
                                    Embed =
                                        relationshipId
                                },

                                new A.Stretch(
                                    new A.FillRectangle())),

                            new PIC.ShapeProperties(

                                new A.Transform2D(

                                    new A.Offset
                                    {
                                        X = 0L,
                                        Y = 0L
                                    },

                                    new A.Extents
                                    {
                                        Cx = widthEmu,
                                        Cy = heightEmu
                                    }),

                                new A.PresetGeometry(
                                    new A.AdjustValueList())
                                {
                                    Preset =
                                        A.ShapeTypeValues.Rectangle
                                })
                        ))
                    {
                        Uri =
                            "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });
    }

    private static string CreateUniqueOutputPath(
        string sourcePath)
    {
        var directory =
            Path.GetDirectoryName(
                sourcePath)
            ?? throw new InvalidOperationException(
                "Unable to determine source directory.");

        var name =
            Path.GetFileNameWithoutExtension(
                sourcePath);

        var candidate =
            Path.Combine(
                directory,
                $"{name}-ocr.docx");

        if (!File.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name}-ocr ({index}).docx");

            if (!File.Exists(candidate))
                return candidate;

            index++;
        }
    }

    private sealed record OcrTextLine(
        int Top,
        string Text);

    private sealed record OcrContentBlock(
        int Top,
        string? Text,
        PdfOcrTableDetectionService.OcrTable? Table);
}