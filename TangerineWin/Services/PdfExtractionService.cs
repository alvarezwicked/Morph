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

public sealed class PdfExtractionService
{
    private readonly PdfTableExtractionService _tableExtractionService;

    public PdfExtractionService()
    {
        _tableExtractionService =
            new PdfTableExtractionService();
    }

    public async Task<string> ExtractToTxtAsync(
        string sourcePath)
    {
        return await Task.Run(() =>
        {
            var outputPath =
                CreateUniqueOutputPath(
                    sourcePath,
                    ".txt");

            using var document =
                PdfDocument.Open(sourcePath);

            using var writer =
                new StreamWriter(
                    outputPath,
                    false);

            foreach (var page in document.GetPages())
            {
                writer.WriteLine(
                    page.Text);

                writer.WriteLine();
            }

            return outputPath;
        });
    }

    public async Task<string> ExtractToDocxAsync(
        string sourcePath)
    {
        var extractedTables =
            await _tableExtractionService
                .ExtractAsync(
                    sourcePath);

        return await Task.Run(() =>
        {
            var outputPath =
                CreateUniqueOutputPath(
                    sourcePath,
                    ".docx");

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
                mainPart.Document
                    .AppendChild(
                        new Body());

            var imageNumber = 1;

            foreach (var page in pdf.GetPages())
            {
                var pageTables =
                    extractedTables
                        .Where(table =>
                            table.PageNumber ==
                            page.Number)
                        .ToArray();

                var textLines =
                    ExtractTextOutsideTables(
                        page,
                        pageTables);

                var pageItems =
                    new List<PageItem>();

                foreach (var line in textLines)
                {
                    pageItems.Add(
                        new PageItem(
                            line.Top,
                            line,
                            null));
                }

                foreach (var table in pageTables)
                {
                    pageItems.Add(
                        new PageItem(
                            table.Top,
                            null,
                            table));
                }

                // В PDF координата Y растёт снизу вверх.
                // Поэтому сортируем сверху страницы вниз.
                foreach (var item in
                         pageItems.OrderByDescending(
                             item => item.Top))
                {
                    if (item.TextLine is not null)
                    {
                        AddTextParagraph(
                            body,
                            item.TextLine.Text);
                    }
                    else if (item.Table is not null)
                    {
                        AddWordTable(
                            body,
                            item.Table.Rows);
                    }
                }

                // Картинки пока добавляем после
                // текстовых/табличных элементов страницы.
                foreach (var pdfImage in page.GetImages())
                {
                    try
                    {
                        var imageData =
                            GetImageData(
                                pdfImage);

                        if (imageData is null)
                            continue;

                        AddImageToDocument(
                            mainPart,
                            body,
                            imageData.Value.Bytes,
                            imageData.Value.PartType,
                            imageNumber);

                        imageNumber++;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"Unable to extract PDF image: {ex}");
                    }
                }

                if (page.Number <
                    pdf.NumberOfPages)
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
        });
    }

    private static IReadOnlyList<TextLineInfo>
        ExtractTextOutsideTables(
            UglyToad.PdfPig.Content.Page page,
            IReadOnlyList<PdfTableExtractionService.ExtractedTable> tables)
    {
        var words =
            page.GetWords()
                .Where(word =>
                    !string.IsNullOrWhiteSpace(
                        word.Text))
                .Select(word =>
                    new WordInfo(
                        word.Text,
                        word.BoundingBox.Left,
                        word.BoundingBox.Bottom,
                        word.BoundingBox.Right,
                        word.BoundingBox.Top))
                .Where(word =>
                    !IsWordInsideAnyTable(
                        word,
                        tables))
                .OrderByDescending(
                    word => word.CenterY)
                .ThenBy(
                    word => word.Left)
                .ToArray();

        if (words.Length == 0)
        {
            return Array.Empty<TextLineInfo>();
        }

        var lines =
            new List<TextLineBuilder>();

        foreach (var word in words)
        {
            TextLineBuilder? targetLine =
                null;

            foreach (var line in lines)
            {
                var tolerance =
                    Math.Max(
                        3.0,
                        Math.Max(
                            word.Height,
                            line.AverageHeight) *
                        0.55);

                if (Math.Abs(
                        line.CenterY -
                        word.CenterY)
                    <= tolerance)
                {
                    targetLine = line;
                    break;
                }
            }

            if (targetLine is null)
            {
                targetLine =
                    new TextLineBuilder();

                lines.Add(
                    targetLine);
            }

            targetLine.Add(
                word);
        }

        return lines
            .Select(line =>
                line.Build())
            .Where(line =>
                !string.IsNullOrWhiteSpace(
                    line.Text))
            .OrderByDescending(
                line => line.Top)
            .ToArray();
    }

    private static bool IsWordInsideAnyTable(
        WordInfo word,
        IReadOnlyList<PdfTableExtractionService.ExtractedTable> tables)
    {
        // Небольшой запас вокруг границы таблицы.
        const double margin = 2.0;

        foreach (var table in tables)
        {
            var inside =
                word.CenterX >=
                    table.Left - margin &&
                word.CenterX <=
                    table.Right + margin &&
                word.CenterY >=
                    table.Bottom - margin &&
                word.CenterY <=
                    table.Top + margin;

            if (inside)
                return true;
        }

        return false;
    }

    private static void AddTextParagraph(
        Body body,
        string text)
    {
        if (string.IsNullOrWhiteSpace(
            text))
        {
            return;
        }

        body.Append(
            new Paragraph(
                new Run(
                    new Text(
                        text)
                    {
                        Space =
                            SpaceProcessingModeValues
                                .Preserve
                    })));
    }

    private static void AddWordTable(
        Body body,
        IReadOnlyList<IReadOnlyList<string>> rows)
    {
        if (rows.Count == 0)
            return;

        var columnCount =
            rows
                .Select(row =>
                    row.Count)
                .DefaultIfEmpty(0)
                .Max();

        if (columnCount == 0)
            return;

        var table =
            new DocumentFormat.OpenXml.Wordprocessing.Table();

        var tableProperties =
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
                        Val =
                            BorderValues.Single,
                        Size = 4U
                    },

                    new LeftBorder
                    {
                        Val =
                            BorderValues.Single,
                        Size = 4U
                    },

                    new BottomBorder
                    {
                        Val =
                            BorderValues.Single,
                        Size = 4U
                    },

                    new RightBorder
                    {
                        Val =
                            BorderValues.Single,
                        Size = 4U
                    },

                    new InsideHorizontalBorder
                    {
                        Val =
                            BorderValues.Single,
                        Size = 4U
                    },

                    new InsideVerticalBorder
                    {
                        Val =
                            BorderValues.Single,
                        Size = 4U
                    }));

        table.AppendChild(
            tableProperties);

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

                var paragraph =
                    new Paragraph(
                        new Run(
                            new Text(
                                value)
                            {
                                Space =
                                    SpaceProcessingModeValues
                                        .Preserve
                            }));

                var cell =
                    new TableCell(
                        paragraph);

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

    private static (
        byte[] Bytes,
        PartTypeInfo PartType)?
        GetImageData(
            UglyToad.PdfPig.Content.IPdfImage image)
    {
        if (image.TryGetPng(
            out var pngBytes))
        {
            if (pngBytes is not null &&
                pngBytes.Length > 0)
            {
                return (
                    pngBytes,
                    ImagePartType.Png);
            }
        }

        var rawBytes =
            image.RawBytes
                .ToArray();

        if (IsJpeg(
            rawBytes))
        {
            return (
                rawBytes,
                ImagePartType.Jpeg);
        }

        return null;
    }

    private static void AddImageToDocument(
        MainDocumentPart mainPart,
        Body body,
        byte[] imageBytes,
        PartTypeInfo imagePartType,
        int imageNumber)
    {
        var imagePart =
            mainPart.AddImagePart(
                imagePartType);

        using (
            var stream =
                new MemoryStream(
                    imageBytes))
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
                    imageBytes))
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

            if (widthEmu >
                maxWidthEmu)
            {
                var scale =
                    (double)maxWidthEmu /
                    widthEmu;

                widthEmu =
                    maxWidthEmu;

                heightEmu =
                    (long)(
                        heightEmu *
                        scale);
            }
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
                        $"Image {imageNumber}"
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
                                        $"Image {imageNumber}"
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
                                        A.ShapeTypeValues
                                            .Rectangle
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

    private static bool IsJpeg(
        byte[] bytes)
    {
        return bytes.Length >= 3
            && bytes[0] == 0xFF
            && bytes[1] == 0xD8
            && bytes[2] == 0xFF;
    }

    private static string CreateUniqueOutputPath(
        string sourcePath,
        string targetExtension)
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
                name +
                targetExtension);

        if (!File.Exists(
            candidate))
        {
            return candidate;
        }

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name} ({index}){targetExtension}");

            if (!File.Exists(
                candidate))
            {
                return candidate;
            }

            index++;
        }
    }

    private sealed record WordInfo(
        string Text,
        double Left,
        double Bottom,
        double Right,
        double Top)
    {
        public double CenterX =>
            (Left + Right) / 2.0;

        public double CenterY =>
            (Bottom + Top) / 2.0;

        public double Height =>
            Math.Abs(
                Top - Bottom);
    }

    private sealed record TextLineInfo(
        string Text,
        double Top);

    private sealed class TextLineBuilder
    {
        private readonly List<WordInfo> _words =
            new();

        public double CenterY =>
            _words.Count == 0
                ? 0
                : _words.Average(
                    word =>
                        word.CenterY);

        public double AverageHeight =>
            _words.Count == 0
                ? 0
                : _words.Average(
                    word =>
                        word.Height);

        public void Add(
            WordInfo word)
        {
            _words.Add(
                word);
        }

        public TextLineInfo Build()
        {
            var orderedWords =
                _words
                    .OrderBy(
                        word =>
                            word.Left)
                    .ToArray();

            var text =
                string.Join(
                    " ",
                    orderedWords
                        .Select(
                            word =>
                                word.Text));

            var top =
                orderedWords.Length == 0
                    ? 0
                    : orderedWords.Max(
                        word =>
                            word.Top);

            return new TextLineInfo(
                text,
                top);
        }
    }

    private sealed record PageItem(
        double Top,
        TextLineInfo? TextLine,
        PdfTableExtractionService.ExtractedTable? Table);
}