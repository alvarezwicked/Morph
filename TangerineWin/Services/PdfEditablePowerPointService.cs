using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

using PdfPage = UglyToad.PdfPig.Content.Page;
using PdfWord = UglyToad.PdfPig.Content.Word;
using PdfLetter = UglyToad.PdfPig.Content.Letter;
using PdfImage = UglyToad.PdfPig.Content.IPdfImage;

namespace TangerineWin.Services;

public sealed class PdfEditablePowerPointService
{
    private readonly PdfTableExtractionService _tableExtractionService =
        new();

    private readonly PdfRenderService _pdfRenderService =
        new();

    public async Task<string> ConvertAsync(
        string sourcePath)
    {
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException(
                "PDF file was not found.",
                sourcePath);
        }

        var templatePath =
            Path.Combine(
                AppContext.BaseDirectory,
                "Assets",
                "PdfTemplate.pptx");

        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException(
                "PowerPoint template was not found.",
                templatePath);
        }

        var outputPath =
            CreateUniqueOutputPath(
                sourcePath);

        File.Copy(
            templatePath,
            outputPath,
            false);

        var tempRoot =
            Path.Combine(
                Path.GetTempPath(),
                "TangerineWin",
                "PdfEditablePptx",
                Guid.NewGuid().ToString("N"));

        Directory.CreateDirectory(
            tempRoot);

        try
        {
            var allTables =
                await _tableExtractionService
                    .ExtractAsync(
                        sourcePath);

            var tempPdfPath =
                Path.Combine(
                    tempRoot,
                    "source.pdf");

            File.Copy(
                sourcePath,
                tempPdfPath,
                true);

            var renderedDirectory =
                await _pdfRenderService
                    .RenderToImagesAsync(
                        tempPdfPath,
                        "pdf_to_png");

            var renderedPages =
                Directory
                    .GetFiles(
                        renderedDirectory,
                        "*.png")
                    .OrderBy(
                        NaturalPageOrder)
                    .ToArray();

            await Task.Run(() =>
            {
                FillPresentation(
                    sourcePath,
                    outputPath,
                    allTables,
                    renderedPages,
                    tempRoot);
            });

            return outputPath;
        }
        catch
        {
            try
            {
                if (File.Exists(
                        outputPath))
                {
                    File.Delete(
                        outputPath);
                }
            }
            catch
            {
            }

            throw;
        }
        finally
        {
            try
            {
                if (Directory.Exists(
                        tempRoot))
                {
                    Directory.Delete(
                        tempRoot,
                        true);
                }
            }
            catch
            {
            }
        }
    }

    // ============================================================
    // PRESENTATION
    // ============================================================

    private static void FillPresentation(
        string pdfPath,
        string pptxPath,
        IReadOnlyList<PdfTableExtractionService.ExtractedTable> allTables,
        IReadOnlyList<string> renderedPages,
        string tempRoot)
    {
        using var pdf =
            PdfDocument.Open(
                pdfPath);

        using var presentationDocument =
            PresentationDocument.Open(
                pptxPath,
                true);

        var presentationPart =
            presentationDocument.PresentationPart
            ?? throw new InvalidOperationException(
                "PresentationPart is missing.");

        var presentation =
            presentationPart.Presentation
            ?? throw new InvalidOperationException(
                "Presentation is missing.");

        var slideIdList =
            presentation.SlideIdList
            ?? throw new InvalidOperationException(
                "SlideIdList is missing.");

        if (pdf.NumberOfPages > 0)
        {
            var firstPage =
                pdf.GetPage(1);

            ApplyPdfAspectRatio(
                presentation,
                firstPage.Width,
                firstPage.Height);
        }

        var templateSlideId =
            slideIdList
                .Elements<P.SlideId>()
                .FirstOrDefault()
            ?? throw new InvalidOperationException(
                "Template must contain one slide.");

        var templateRelationshipId =
            templateSlideId
                .RelationshipId?
                .Value
            ?? throw new InvalidOperationException(
                "Template slide relationship is missing.");

        var templateSlidePart =
            (SlidePart)
            presentationPart.GetPartById(
                templateRelationshipId);

        var slideLayoutPart =
            templateSlidePart.SlideLayoutPart
            ?? throw new InvalidOperationException(
                "SlideLayoutPart is missing.");

        templateSlideId.Remove();

        presentationPart.DeletePart(
            templateSlidePart);

        uint slideId =
            256;

        for (var pageNumber = 1;
             pageNumber <= pdf.NumberOfPages;
             pageNumber++)
        {
            var page =
                pdf.GetPage(
                    pageNumber);

            var renderedPagePath =
                pageNumber - 1 <
                renderedPages.Count
                    ? renderedPages[
                        pageNumber - 1]
                    : null;

            var pageTables =
                PrepareTables(
                    page,
                    allTables
                        .Where(table =>
                            table.PageNumber ==
                            pageNumber)
                        .ToArray());

            var editableTables =
                pageTables
                    .Select(item =>
                        item.Table)
                    .ToArray();

            var styledLines =
                ExtractStyledTextLines(
                    page,
                    editableTables,
                    renderedPagePath);

            var pageImages =
                page.GetImages()
                    .ToArray();

            var editableImages =
                pageImages
                    .Where(image =>
                        !IsImagePartOfTable(
                            image,
                            pageTables))
                    .ToArray();

            string? backgroundPath =
                null;

            if (!string.IsNullOrWhiteSpace(
                    renderedPagePath))
            {
                backgroundPath =
                    CreateMaskedBackground(
                        renderedPagePath,
                        page,
                        pageTables,
                        pageImages,
                        tempRoot,
                        pageNumber);
            }

            var slidePart =
                presentationPart
                    .AddNewPart<SlidePart>();

            slidePart.AddPart(
                slideLayoutPart);

            CreateEditableSlide(
                slidePart,
                page,
                presentation,
                styledLines,
                editableTables,
                editableImages,
                backgroundPath,
                renderedPagePath);

            slideIdList.Append(
                new P.SlideId
                {
                    Id =
                        slideId++,

                    RelationshipId =
                        presentationPart
                            .GetIdOfPart(
                                slidePart)
                });
        }

        presentation.Save();
    }

    // ============================================================
    // TABLE DETECTION
    // ============================================================

    private static IReadOnlyList<PreparedTable>
        PrepareTables(
            PdfPage page,
            IReadOnlyList<PdfTableExtractionService.ExtractedTable>
                source)
    {
        var candidates =
            new List<PreparedTable>();

        foreach (var sourceTable in source)
        {
            var normalized =
                NormalizeDetectedTable(
                    sourceTable);

            if (normalized is null)
            {
                continue;
            }

            if (!IsLikelyRealTable(
                    page,
                    normalized))
            {
                continue;
            }

            var refined =
                RefineTableBounds(
                    page,
                    normalized);

            if (refined is null)
            {
                continue;
            }

            refined =
                TrimTableBoundsToContent(
                    page,
                    refined);

            if (!IsLikelyRealTable(
                    page,
                    refined))
            {
                continue;
            }

            candidates.Add(
                new PreparedTable(
                    refined,
                    sourceTable.Left,
                    sourceTable.Bottom,
                    sourceTable.Right,
                    sourceTable.Top));
        }

        candidates =
            candidates
                .OrderByDescending(item =>
                    GetTableQuality(
                        item.Table))
                .ToList();

        var result =
            new List<PreparedTable>();

        foreach (var candidate in candidates)
        {
            if (result.Any(existing =>
                    AreSameTableArea(
                        existing.Table,
                        candidate.Table)))
            {
                continue;
            }

            result.Add(
                candidate);
        }

        return result;
    }

    private static bool IsLikelyRealTable(
        PdfPage page,
        PdfTableExtractionService.ExtractedTable table)
    {
        var rowCount =
            table.Rows.Count;

        if (rowCount < 2)
        {
            return false;
        }

        var maxColumns =
            table.Rows
                .Select(row =>
                    row.Count)
                .DefaultIfEmpty(0)
                .Max();

        if (maxColumns < 2 ||
            maxColumns > 14)
        {
            return false;
        }

        var width =
            Math.Max(
                0,
                table.Right -
                table.Left);

        var height =
            Math.Max(
                0,
                table.Top -
                table.Bottom);

        if (width <= 0 ||
            height <= 0)
        {
            return false;
        }

        var widthRatio =
            width /
            Math.Max(
                1,
                page.Width);

        var heightRatio =
            height /
            Math.Max(
                1,
                page.Height);

        var areaRatio =
            widthRatio *
            heightRatio;

        var nonEmptyPerRow =
            table.Rows
                .Select(
                    CountNonEmptyCells)
                .ToArray();

        var nonEmptyCells =
            nonEmptyPerRow.Sum();

        var totalSlots =
            Math.Max(
                1,
                rowCount *
                maxColumns);

        var fillRatio =
            (double)nonEmptyCells /
            totalSlots;

        var multiColumnRows =
            nonEmptyPerRow.Count(count =>
                count >= 2);

        var multiColumnRatio =
            (double)multiColumnRows /
            rowCount;

        if (nonEmptyCells < 4 ||
            multiColumnRatio < 0.55 ||
            fillRatio < 0.28)
        {
            return false;
        }

        if (areaRatio > 0.48 &&
            rowCount <= 5)
        {
            return false;
        }

        if (widthRatio > 0.86 &&
            heightRatio > 0.34 &&
            rowCount <= 6)
        {
            return false;
        }

        if (areaRatio > 0.62 &&
            fillRatio < 0.58)
        {
            return false;
        }

        if (heightRatio > 0.72 &&
            rowCount < 8)
        {
            return false;
        }

        var bulletRowCount =
            table.Rows.Count(row =>
                row.Any(
                    ContainsBulletMarker));

        if (bulletRowCount >= 2)
        {
            var bulletRatio =
                (double)bulletRowCount /
                rowCount;

            if (bulletRatio >= 0.25)
            {
                return false;
            }
        }

        var firstColumnValues =
            table.Rows
                .Where(row =>
                    row.Count > 0)
                .Select(row =>
                    row[0]?.Trim() ??
                    string.Empty)
                .Where(value =>
                    !string.IsNullOrWhiteSpace(
                        value))
                .ToArray();

        if (firstColumnValues.Length > 0)
        {
            var listMarkerCount =
                firstColumnValues
                    .Count(
                        IsListMarker);

            var ratio =
                (double)listMarkerCount /
                firstColumnValues.Length;

            if (ratio >= 0.55)
            {
                return false;
            }
        }

        var populatedCells =
            table.Rows
                .SelectMany(row =>
                    row)
                .Where(cell =>
                    !string.IsNullOrWhiteSpace(
                        cell))
                .ToArray();

        var alphaCells =
            populatedCells.Count(cell =>
                cell.Any(
                    char.IsLetter));

        var alphaRatio =
            populatedCells.Length == 0
                ? 0
                : (double)alphaCells /
                  populatedCells.Length;

        if (maxColumns >= 5 &&
            rowCount >= 5 &&
            fillRatio < 0.42 &&
            alphaRatio < 0.30)
        {
            return false;
        }

        return true;
    }

    private static
        PdfTableExtractionService.ExtractedTable?
        NormalizeDetectedTable(
            PdfTableExtractionService.ExtractedTable table)
    {
        var rows =
            table.Rows
                .Select(row =>
                    (IReadOnlyList<string>)
                    row.ToArray())
                .ToList();

        while (rows.Count >= 3)
        {
            var firstNonEmpty =
                CountNonEmptyCells(
                    rows[0]);

            var followingMultiColumnRows =
                rows
                    .Skip(1)
                    .Take(
                        Math.Min(
                            4,
                            rows.Count - 1))
                    .Count(row =>
                        CountNonEmptyCells(
                            row) >= 2);

            if (firstNonEmpty <= 1 &&
                followingMultiColumnRows >= 2)
            {
                rows.RemoveAt(0);
                continue;
            }

            break;
        }

        if (rows.Count < 2)
        {
            return null;
        }

        return new PdfTableExtractionService
            .ExtractedTable(
                table.PageNumber,
                table.Left,
                table.Bottom,
                table.Right,
                table.Top,
                rows);
    }

    private static
        PdfTableExtractionService.ExtractedTable?
        RefineTableBounds(
            PdfPage page,
            PdfTableExtractionService.ExtractedTable table)
    {
        var words =
            page.GetWords()
                .Where(word =>
                    !string.IsNullOrWhiteSpace(
                        word.Text))
                .ToArray();

        if (words.Length == 0)
        {
            return table;
        }

        var originalWidth =
            Math.Max(
                1,
                table.Right -
                table.Left);

        var originalHeight =
            Math.Max(
                1,
                table.Top -
                table.Bottom);

        var horizontalTolerance =
            Math.Max(
                10,
                originalWidth *
                0.08);

        var verticalTolerance =
            Math.Max(
                10,
                originalHeight *
                0.10);

        var matchedWords =
            new List<PdfWord>();

        foreach (var word in words)
        {
            var centerX =
                (
                    word.BoundingBox.Left +
                    word.BoundingBox.Right
                ) /
                2.0;

            var centerY =
                (
                    word.BoundingBox.Bottom +
                    word.BoundingBox.Top
                ) /
                2.0;

            if (centerX <
                    table.Left -
                    horizontalTolerance ||
                centerX >
                    table.Right +
                    horizontalTolerance ||
                centerY <
                    table.Bottom -
                    verticalTolerance ||
                centerY >
                    table.Top +
                    verticalTolerance)
            {
                continue;
            }

            if (!TableContainsWord(
                    table,
                    word.Text))
            {
                continue;
            }

            matchedWords.Add(
                word);
        }

        if (matchedWords.Count < 4)
        {
            return table;
        }

        var left =
            matchedWords.Min(word =>
                word.BoundingBox.Left);

        var right =
            matchedWords.Max(word =>
                word.BoundingBox.Right);

        var bottom =
            matchedWords.Min(word =>
                word.BoundingBox.Bottom);

        var top =
            matchedWords.Max(word =>
                word.BoundingBox.Top);

        var heights =
            matchedWords
                .Select(word =>
                    word.BoundingBox.Height)
                .Where(height =>
                    height > 0)
                .OrderBy(height =>
                    height)
                .ToArray();

        var typicalHeight =
            heights.Length == 0
                ? 8
                : heights[
                    heights.Length / 2];

        left -=
            Math.Max(
                3,
                typicalHeight * 0.55);

        right +=
            Math.Max(
                3,
                typicalHeight * 0.55);

        bottom -=
            Math.Max(
                1.5,
                typicalHeight * 0.30);

        top +=
            Math.Max(
                1.5,
                typicalHeight * 0.30);

        return new PdfTableExtractionService
            .ExtractedTable(
                table.PageNumber,
                Math.Max(
                    0,
                    left),
                Math.Max(
                    0,
                    bottom),
                Math.Min(
                    page.Width,
                    right),
                Math.Min(
                    page.Height,
                    top),
                table.Rows);
    }

    private static PdfTableExtractionService.ExtractedTable
        TrimTableBoundsToContent(
            PdfPage page,
            PdfTableExtractionService.ExtractedTable table)
    {
        var words =
            page.GetWords()
                .Where(word =>
                    !string.IsNullOrWhiteSpace(
                        word.Text))
                .Where(word =>
                {
                    var centerX =
                        (
                            word.BoundingBox.Left +
                            word.BoundingBox.Right
                        ) /
                        2.0;

                    var centerY =
                        (
                            word.BoundingBox.Bottom +
                            word.BoundingBox.Top
                        ) /
                        2.0;

                    return
                        centerX >= table.Left &&
                        centerX <= table.Right &&
                        centerY >= table.Bottom &&
                        centerY <= table.Top;
                })
                .ToArray();

        if (words.Length < 4)
        {
            return table;
        }

        var relevant =
            words
                .Where(word =>
                    TableContainsWord(
                        table,
                        word.Text))
                .ToArray();

        if (relevant.Length < 4)
        {
            return table;
        }

        var heights =
            relevant
                .Select(word =>
                    word.BoundingBox.Height)
                .Where(value =>
                    value > 0)
                .OrderBy(value =>
                    value)
                .ToArray();

        var typicalHeight =
            heights.Length > 0
                ? heights[
                    heights.Length / 2]
                : 8;

        var top =
            relevant.Max(word =>
                word.BoundingBox.Top);

        var bottom =
            relevant.Min(word =>
                word.BoundingBox.Bottom);

        var left =
            relevant.Min(word =>
                word.BoundingBox.Left);

        var right =
            relevant.Max(word =>
                word.BoundingBox.Right);

        //
        // В отличие от старой версии не даём
        // таблице "растягиваться" до footer.
        //
        var safeBottom =
            bottom -
            Math.Max(
                2,
                typicalHeight * 0.45);

        var safeTop =
            top +
            Math.Max(
                2,
                typicalHeight * 0.45);

        var safeLeft =
            left -
            Math.Max(
                3,
                typicalHeight * 0.55);

        var safeRight =
            right +
            Math.Max(
                3,
                typicalHeight * 0.55);

        return new PdfTableExtractionService
            .ExtractedTable(
                table.PageNumber,
                Math.Max(
                    table.Left,
                    safeLeft),
                Math.Max(
                    table.Bottom,
                    safeBottom),
                Math.Min(
                    table.Right,
                    safeRight),
                Math.Min(
                    table.Top,
                    safeTop),
                table.Rows);
    }

    // ============================================================
    // TABLE GEOMETRY
    // ============================================================

    private static IReadOnlyList<long>
        EstimateColumnWidths(
            PdfPage page,
            PdfTableExtractionService.ExtractedTable table,
            long totalWidth)
    {
        var columnCount =
            Math.Max(
                1,
                table.Rows
                    .Select(row =>
                        row.Count)
                    .DefaultIfEmpty(1)
                    .Max());

        if (columnCount == 1)
        {
            return new[]
            {
                totalWidth
            };
        }

        var words =
            GetWordsInsideTable(
                page,
                table);

        if (words.Length == 0)
        {
            return CreateEqualSizes(
                totalWidth,
                columnCount);
        }

        var rowClusters =
            ClusterWordsIntoPhysicalRows(
                words);

        var gapCandidates =
            new List<ColumnGapCandidate>();

        foreach (var row in rowClusters)
        {
            var ordered =
                row
                    .OrderBy(word =>
                        word.BoundingBox.Left)
                    .ToArray();

            for (var index = 1;
                 index < ordered.Length;
                 index++)
            {
                var previous =
                    ordered[index - 1];

                var current =
                    ordered[index];

                var gap =
                    current.BoundingBox.Left -
                    previous.BoundingBox.Right;

                if (gap <= 0)
                {
                    continue;
                }

                gapCandidates.Add(
                    new ColumnGapCandidate(
                        (
                            previous.BoundingBox.Right +
                            current.BoundingBox.Left
                        ) /
                        2.0,
                        gap));
            }
        }

        if (gapCandidates.Count == 0)
        {
            return CreateEqualSizes(
                totalWidth,
                columnCount);
        }

        var averageColumnWidth =
            (
                table.Right -
                table.Left
            ) /
            columnCount;

        var mergeTolerance =
            Math.Max(
                3,
                averageColumnWidth *
                0.055);

        var clusters =
            new List<
                List<ColumnGapCandidate>>();

        foreach (var candidate in
                 gapCandidates
                     .OrderBy(item =>
                         item.Position))
        {
            var existing =
                clusters
                    .FirstOrDefault(cluster =>
                        Math.Abs(
                            cluster.Average(item =>
                                item.Position) -
                            candidate.Position) <=
                        mergeTolerance);

            if (existing is null)
            {
                existing =
                    new List<ColumnGapCandidate>();

                clusters.Add(
                    existing);
            }

            existing.Add(
                candidate);
        }

        var minimumRowSupport =
            Math.Max(
                2,
                (int)Math.Ceiling(
                    rowClusters.Count *
                    0.35));

        var separatorCandidates =
            clusters
                .Select(cluster =>
                    new
                    {
                        Position =
                            cluster.Average(item =>
                                item.Position),

                        AverageGap =
                            cluster.Average(item =>
                                item.Gap),

                        Support =
                            cluster.Count
                    })
                .Where(item =>
                    item.Support >=
                    minimumRowSupport)
                .Where(item =>
                    item.AverageGap >
                    averageColumnWidth *
                    0.08)
                .OrderByDescending(item =>
                    item.Support * 1000 +
                    item.AverageGap)
                .Take(
                    columnCount - 1)
                .OrderBy(item =>
                    item.Position)
                .ToArray();

        if (separatorCandidates.Length !=
            columnCount - 1)
        {
            //
            // Fallback: берём наиболее сильные
            // промежутки даже если support меньше.
            //
            separatorCandidates =
                clusters
                    .Select(cluster =>
                        new
                        {
                            Position =
                                cluster.Average(item =>
                                    item.Position),

                            AverageGap =
                                cluster.Average(item =>
                                    item.Gap),

                            Support =
                                cluster.Count
                        })
                    .OrderByDescending(item =>
                        item.Support * 1000 +
                        item.AverageGap)
                    .Take(
                        columnCount - 1)
                    .OrderBy(item =>
                        item.Position)
                    .ToArray();
        }

        if (separatorCandidates.Length !=
            columnCount - 1)
        {
            return CreateEqualSizes(
                totalWidth,
                columnCount);
        }

        var boundaries =
            new List<double>
            {
                table.Left
            };

        boundaries.AddRange(
            separatorCandidates.Select(item =>
                item.Position));

        boundaries.Add(
            table.Right);

        var minimumWidth =
            Math.Max(
                2,
                averageColumnWidth *
                0.10);

        for (var index = 0;
             index < columnCount;
             index++)
        {
            if (boundaries[index + 1] -
                boundaries[index] <
                minimumWidth)
            {
                return CreateEqualSizes(
                    totalWidth,
                    columnCount);
            }
        }

        var result =
            new List<long>();

        for (var index = 0;
             index < columnCount;
             index++)
        {
            var fraction =
                (
                    boundaries[index + 1] -
                    boundaries[index]
                ) /
                Math.Max(
                    1,
                    table.Right -
                    table.Left);

            result.Add(
                Math.Max(
                    1,
                    (long)Math.Round(
                        fraction *
                        totalWidth)));
        }

        NormalizeSizesToTotal(
            result,
            totalWidth);

        return result;
    }

    private static IReadOnlyList<List<PdfWord>>
        ClusterWordsIntoPhysicalRows(
            IReadOnlyList<PdfWord> words)
    {
        var rows =
            new List<List<PdfWord>>();

        foreach (var word in
                 words
                     .OrderByDescending(word =>
                         word.BoundingBox.Top))
        {
            var centerY =
                (
                    word.BoundingBox.Top +
                    word.BoundingBox.Bottom
                ) /
                2.0;

            List<PdfWord>? row =
                null;

            foreach (var candidate in rows)
            {
                var averageY =
                    candidate.Average(item =>
                        (
                            item.BoundingBox.Top +
                            item.BoundingBox.Bottom
                        ) /
                        2.0);

                var averageHeight =
                    candidate.Average(item =>
                        item.BoundingBox.Height);

                var tolerance =
                    Math.Max(
                        2,
                        Math.Max(
                            averageHeight,
                            word.BoundingBox.Height) *
                        0.50);

                if (Math.Abs(
                        averageY -
                        centerY) <=
                    tolerance)
                {
                    row =
                        candidate;

                    break;
                }
            }

            if (row is null)
            {
                row =
                    new List<PdfWord>();

                rows.Add(
                    row);
            }

            row.Add(
                word);
        }

        return rows;
    }

    private static IReadOnlyList<long>
        EstimateRowHeights(
            PdfPage page,
            PdfTableExtractionService.ExtractedTable table,
            long totalHeight)
    {
        var rowCount =
            Math.Max(
                1,
                table.Rows.Count);

        if (rowCount == 1)
        {
            return new[]
            {
                totalHeight
            };
        }

        var words =
            GetWordsInsideTable(
                page,
                table);

        if (words.Length == 0)
        {
            return CreateEqualSizes(
                totalHeight,
                rowCount);
        }

        var physicalRows =
            ClusterWordsIntoPhysicalRows(
                words)
            .OrderByDescending(row =>
                row.Average(word =>
                    (
                        word.BoundingBox.Top +
                        word.BoundingBox.Bottom
                    ) /
                    2.0))
            .ToArray();

        if (physicalRows.Length <
            rowCount)
        {
            return CreateEqualSizes(
                totalHeight,
                rowCount);
        }

        var centers =
            physicalRows
                .Select(row =>
                    row.Average(word =>
                        (
                            word.BoundingBox.Top +
                            word.BoundingBox.Bottom
                        ) /
                        2.0))
                .ToArray();

        var gaps =
            new List<
                (double Position, double Gap)>();

        for (var index = 1;
             index < centers.Length;
             index++)
        {
            gaps.Add(
                (
                    (
                        centers[index - 1] +
                        centers[index]
                    ) /
                    2.0,

                    centers[index - 1] -
                    centers[index]
                ));
        }

        var separators =
            gaps
                .OrderByDescending(item =>
                    item.Gap)
                .Take(
                    rowCount - 1)
                .Select(item =>
                    item.Position)
                .OrderByDescending(value =>
                    value)
                .ToArray();

        if (separators.Length !=
            rowCount - 1)
        {
            return CreateEqualSizes(
                totalHeight,
                rowCount);
        }

        var boundaries =
            new List<double>
            {
                table.Top
            };

        boundaries.AddRange(
            separators);

        boundaries.Add(
            table.Bottom);

        var result =
            new List<long>();

        for (var index = 0;
             index < rowCount;
             index++)
        {
            var height =
                boundaries[index] -
                boundaries[index + 1];

            if (height <= 0)
            {
                return CreateEqualSizes(
                    totalHeight,
                    rowCount);
            }

            result.Add(
                Math.Max(
                    1,
                    (long)Math.Round(
                        height /
                        Math.Max(
                            1,
                            table.Top -
                            table.Bottom) *
                        totalHeight)));
        }

        NormalizeSizesToTotal(
            result,
            totalHeight);

        return result;
    }

    private static PdfWord[]
        GetWordsInsideTable(
            PdfPage page,
            PdfTableExtractionService.ExtractedTable table)
    {
        return page
            .GetWords()
            .Where(word =>
                !string.IsNullOrWhiteSpace(
                    word.Text))
            .Where(word =>
            {
                var centerX =
                    (
                        word.BoundingBox.Left +
                        word.BoundingBox.Right
                    ) /
                    2.0;

                var centerY =
                    (
                        word.BoundingBox.Bottom +
                        word.BoundingBox.Top
                    ) /
                    2.0;

                return
                    centerX >= table.Left &&
                    centerX <= table.Right &&
                    centerY >= table.Bottom &&
                    centerY <= table.Top;
            })
            .ToArray();
    }

    private static IReadOnlyList<long>
        CreateEqualSizes(
            long total,
            int count)
    {
        count =
            Math.Max(
                1,
                count);

        var result =
            Enumerable
                .Repeat(
                    Math.Max(
                        1,
                        total /
                        count),
                    count)
                .ToList();

        NormalizeSizesToTotal(
            result,
            total);

        return result;
    }

    private static void NormalizeSizesToTotal(
        IList<long> sizes,
        long total)
    {
        if (sizes.Count == 0)
        {
            return;
        }

        var difference =
            total -
            sizes.Sum();

        sizes[
            sizes.Count - 1] =
                Math.Max(
                    1,
                    sizes[
                        sizes.Count - 1] +
                    difference);
    }

    private static TableGeometry
        BuildTableGeometry(
            PdfTableExtractionService.ExtractedTable table,
            IReadOnlyList<long> columnWidths,
            IReadOnlyList<long> rowHeights)
    {
        var columnBoundaries =
            new double[
                columnWidths.Count + 1];

        var rowBoundaries =
            new double[
                rowHeights.Count + 1];

        columnBoundaries[0] =
            table.Left;

        rowBoundaries[0] =
            table.Top;

        var totalColumnWidth =
            Math.Max(
                1L,
                columnWidths.Sum());

        var totalRowHeight =
            Math.Max(
                1L,
                rowHeights.Sum());

        long current =
            0;

        for (var index = 0;
             index < columnWidths.Count;
             index++)
        {
            current +=
                columnWidths[index];

            columnBoundaries[index + 1] =
                table.Left +
                (
                    table.Right -
                    table.Left
                ) *
                current /
                totalColumnWidth;
        }

        columnBoundaries[^1] =
            table.Right;

        current =
            0;

        for (var index = 0;
             index < rowHeights.Count;
             index++)
        {
            current +=
                rowHeights[index];

            rowBoundaries[index + 1] =
                table.Top -
                (
                    table.Top -
                    table.Bottom
                ) *
                current /
                totalRowHeight;
        }

        rowBoundaries[^1] =
            table.Bottom;

        return new TableGeometry(
            columnBoundaries,
            rowBoundaries);
    }

    // ============================================================
    // SLIDE
    // ============================================================

    private static void CreateEditableSlide(
        SlidePart slidePart,
        PdfPage page,
        P.Presentation presentation,
        IReadOnlyList<StyledTextLine> lines,
        IReadOnlyList<PdfTableExtractionService.ExtractedTable> tables,
        IReadOnlyList<PdfImage> images,
        string? backgroundPath,
        string? renderedPagePath)
    {
        var slideSize =
            presentation.SlideSize
            ?? throw new InvalidOperationException(
                "SlideSize is missing.");

        var slideWidth =
            slideSize.Cx?.Value
            ?? throw new InvalidOperationException(
                "Slide width is missing.");

        var slideHeight =
            slideSize.Cy?.Value
            ?? throw new InvalidOperationException(
                "Slide height is missing.");

        var shapeTree =
            CreateEmptyShapeTree();

        uint shapeId =
            2;

        if (!string.IsNullOrWhiteSpace(
                backgroundPath) &&
            File.Exists(
                backgroundPath))
        {
            shapeTree.Append(
                CreateBackgroundPicture(
                    slidePart,
                    shapeId++,
                    backgroundPath,
                    slideWidth,
                    slideHeight));
        }

        foreach (var line in lines)
        {
            shapeTree.Append(
                CreateStyledTextShape(
                    shapeId++,
                    line,
                    page.Width,
                    page.Height,
                    slideWidth,
                    slideHeight));
        }

        foreach (var image in images)
        {
            var picture =
                TryCreateImageShape(
                    slidePart,
                    shapeId,
                    image,
                    page.Width,
                    page.Height,
                    slideWidth,
                    slideHeight);

            if (picture is null)
            {
                continue;
            }

            shapeTree.Append(
                picture);

            shapeId++;
        }

        foreach (var table in tables)
        {
            shapeTree.Append(
                CreateTableFrame(
                    shapeId++,
                    page,
                    table,
                    page.Width,
                    page.Height,
                    slideWidth,
                    slideHeight,
                    renderedPagePath));
        }

        slidePart.Slide =
            new P.Slide(
                new P.CommonSlideData(
                    shapeTree),

                new P.ColorMapOverride(
                    new A.MasterColorMapping()));

        slidePart.Slide.Save();
    }

    // ============================================================
    // BACKGROUND
    // ============================================================

    private static string?
        CreateMaskedBackground(
            string renderedPagePath,
            PdfPage page,
            IReadOnlyList<PreparedTable> tables,
            IReadOnlyList<PdfImage> images,
            string tempRoot,
            int pageNumber)
    {
        var outputPath =
            Path.Combine(
                tempRoot,
                $"masked-page-{pageNumber:D4}.png");

        using var sourceBitmap =
            new System.Drawing.Bitmap(
                renderedPagePath);

        using var bitmap =
            new System.Drawing.Bitmap(
                sourceBitmap);

        using var graphics =
            System.Drawing.Graphics
                .FromImage(
                    bitmap);

        var acceptedTables =
            tables
                .Select(item =>
                    item.Table)
                .ToArray();

        foreach (var word in
                 page.GetWords())
        {
            if (string.IsNullOrWhiteSpace(
                    word.Text))
            {
                continue;
            }

            if (IsInsideAnyTable(
                    word,
                    acceptedTables))
            {
                continue;
            }

            if (ShouldKeepWordAsRaster(
                    page,
                    word,
                    sourceBitmap))
            {
                continue;
            }

            var box =
                word.BoundingBox;

            var rectangle =
                PdfRectangleToPixels(
                    box.Left,
                    box.Bottom,
                    box.Right,
                    box.Top,
                    page.Width,
                    page.Height,
                    sourceBitmap.Width,
                    sourceBitmap.Height);

            rectangle.Inflate(
                Math.Max(
                    1,
                    sourceBitmap.Width /
                    1800),

                Math.Max(
                    1,
                    sourceBitmap.Height /
                    1400));

            FillRectangleWithLocalBackground(
                sourceBitmap,
                graphics,
                rectangle);
        }

        foreach (var preparedTable in
                 tables)
        {
            var rectangle =
                PdfRectangleToPixels(
                    preparedTable.MaskLeft,
                    preparedTable.MaskBottom,
                    preparedTable.MaskRight,
                    preparedTable.MaskTop,
                    page.Width,
                    page.Height,
                    sourceBitmap.Width,
                    sourceBitmap.Height);

            rectangle.Inflate(
                Math.Max(
                    4,
                    sourceBitmap.Width /
                    350),

                Math.Max(
                    3,
                    sourceBitmap.Height /
                    450));

            FillRectangleWithLocalBackground(
                sourceBitmap,
                graphics,
                rectangle);
        }

        foreach (var image in
                 images)
        {
            if (IsImagePartOfTable(
                    image,
                    tables))
            {
                continue;
            }

            var box =
                image.BoundingBox;

            var rectangle =
                PdfRectangleToPixels(
                    box.Left,
                    box.Bottom,
                    box.Right,
                    box.Top,
                    page.Width,
                    page.Height,
                    sourceBitmap.Width,
                    sourceBitmap.Height);

            FillRectangleWithLocalBackground(
                sourceBitmap,
                graphics,
                rectangle);
        }

        if (!HasMeaningfulGraphics(
                bitmap))
        {
            return null;
        }

        bitmap.Save(
            outputPath,
            System.Drawing.Imaging.ImageFormat.Png);

        return outputPath;
    }

    private static void FillRectangleWithLocalBackground(
        System.Drawing.Bitmap sourceBitmap,
        System.Drawing.Graphics graphics,
        System.Drawing.Rectangle rectangle)
    {
        rectangle =
            ClampRectangle(
                rectangle,
                sourceBitmap.Width,
                sourceBitmap.Height);

        if (rectangle.Width <= 0 ||
            rectangle.Height <= 0)
        {
            return;
        }

        var info =
            AnalyzeLocalBackground(
                sourceBitmap,
                rectangle);

        using var brush =
            new System.Drawing.SolidBrush(
                info.Color);

        graphics.FillRectangle(
            brush,
            rectangle);
    }

    private static bool ShouldKeepWordAsRaster(
        PdfPage page,
        PdfWord word,
        System.Drawing.Bitmap bitmap)
    {
        var rectangle =
            PdfRectangleToPixels(
                word.BoundingBox.Left,
                word.BoundingBox.Bottom,
                word.BoundingBox.Right,
                word.BoundingBox.Top,
                page.Width,
                page.Height,
                bitmap.Width,
                bitmap.Height);

        rectangle =
            ClampRectangle(
                rectangle,
                bitmap.Width,
                bitmap.Height);

        var info =
            AnalyzeLocalBackground(
                bitmap,
                rectangle);

        //
        // Однородная плашка/фон:
        // текст можно безопасно удалить и
        // создать editable.
        //
        if (info.Dominance >= 0.66 &&
            info.ColorSpread <= 60)
        {
            return false;
        }

        //
        // На белом фоне тоже обычно безопасно.
        //
        if (GetBrightness(
                info.Color) > 235 &&
            info.ColorSpread < 48)
        {
            return false;
        }

        return true;
    }

    private static LocalBackgroundInfo
        AnalyzeLocalBackground(
            System.Drawing.Bitmap bitmap,
            System.Drawing.Rectangle rectangle)
    {
        rectangle =
            ClampRectangle(
                rectangle,
                bitmap.Width,
                bitmap.Height);

        if (rectangle.Width <= 0 ||
            rectangle.Height <= 0)
        {
            return new LocalBackgroundInfo(
                System.Drawing.Color.White,
                1,
                0);
        }

        var marginX =
            Math.Max(
                4,
                Math.Min(
                    18,
                    rectangle.Width /
                    2 +
                    3));

        var marginY =
            Math.Max(
                4,
                Math.Min(
                    14,
                    rectangle.Height +
                    2));

        var outer =
            System.Drawing.Rectangle.FromLTRB(
                Math.Max(
                    0,
                    rectangle.Left -
                    marginX),

                Math.Max(
                    0,
                    rectangle.Top -
                    marginY),

                Math.Min(
                    bitmap.Width,
                    rectangle.Right +
                    marginX),

                Math.Min(
                    bitmap.Height,
                    rectangle.Bottom +
                    marginY));

        var samples =
            new List<System.Drawing.Color>();

        var stepX =
            Math.Max(
                1,
                outer.Width /
                24);

        var stepY =
            Math.Max(
                1,
                outer.Height /
                16);

        for (var y = outer.Top;
             y < outer.Bottom;
             y += stepY)
        {
            for (var x = outer.Left;
                 x < outer.Right;
                 x += stepX)
            {
                //
                // Не берём сами пиксели текста.
                //
                if (x >= rectangle.Left &&
                    x < rectangle.Right &&
                    y >= rectangle.Top &&
                    y < rectangle.Bottom)
                {
                    continue;
                }

                samples.Add(
                    bitmap.GetPixel(
                        x,
                        y));
            }
        }

        if (samples.Count == 0)
        {
            return new LocalBackgroundInfo(
                System.Drawing.Color.White,
                1,
                0);
        }

        var buckets =
            samples
                .GroupBy(color =>
                    (
                        R: color.R / 12,
                        G: color.G / 12,
                        B: color.B / 12
                    ))
                .OrderByDescending(group =>
                    group.Count())
                .ToArray();

        var dominant =
            buckets[0]
                .ToArray();

        var dominantColor =
            System.Drawing.Color
                .FromArgb(
                    Median(
                        dominant.Select(color =>
                            (int)color.R)),

                    Median(
                        dominant.Select(color =>
                            (int)color.G)),

                    Median(
                        dominant.Select(color =>
                            (int)color.B)));

        var dominance =
            (double)dominant.Length /
            samples.Count;

        var distances =
            samples
                .Select(color =>
                    ColorDistance(
                        color,
                        dominantColor))
                .ToArray();

        var spread =
            distances.Length == 0
                ? 0
                : MedianDouble(
                    distances);

        return new LocalBackgroundInfo(
            dominantColor,
            dominance,
            spread);
    }

    // ============================================================
    // NORMAL TEXT
    // ============================================================

    private static IReadOnlyList<StyledTextLine>
        ExtractStyledTextLines(
            PdfPage page,
            IReadOnlyList<PdfTableExtractionService.ExtractedTable>
                tables,
            string? renderedPagePath)
    {
        System.Drawing.Bitmap? bitmap =
            null;

        try
        {
            if (!string.IsNullOrWhiteSpace(
                    renderedPagePath) &&
                File.Exists(
                    renderedPagePath))
            {
                bitmap =
                    new System.Drawing.Bitmap(
                        renderedPagePath);
            }

            var letters =
                page.Letters;

            var words =
                page.GetWords()
                    .Where(word =>
                        !string.IsNullOrWhiteSpace(
                            word.Text))
                    .Where(word =>
                        !IsInsideAnyTable(
                            word,
                            tables))
                    .Where(word =>
                        bitmap is null ||
                        !ShouldKeepWordAsRaster(
                            page,
                            word,
                            bitmap))
                    .OrderByDescending(word =>
                        word.BoundingBox.Top)
                    .ThenBy(word =>
                        word.BoundingBox.Left)
                    .Select(word =>
                        new StyledWord(
                            word,
                            GetWordStyle(
                                word,
                                letters,
                                page,
                                bitmap)))
                    .ToArray();

            var physicalLines =
                new List<List<StyledWord>>();

            foreach (var styledWord in
                     words)
            {
                var centerY =
                    (
                        styledWord.Word.BoundingBox.Top +
                        styledWord.Word.BoundingBox.Bottom
                    ) /
                    2.0;

                List<StyledWord>? line =
                    null;

                foreach (var candidate in
                         physicalLines)
                {
                    var y =
                        candidate.Average(item =>
                            (
                                item.Word.BoundingBox.Top +
                                item.Word.BoundingBox.Bottom
                            ) /
                            2.0);

                    var height =
                        candidate.Average(item =>
                            item.Word.BoundingBox.Height);

                    var tolerance =
                        Math.Max(
                            2.2,
                            Math.Max(
                                height,
                                styledWord.Word.BoundingBox.Height) *
                            0.40);

                    if (Math.Abs(
                            y -
                            centerY) <=
                        tolerance)
                    {
                        line =
                            candidate;

                        break;
                    }
                }

                if (line is null)
                {
                    line =
                        new List<StyledWord>();

                    physicalLines.Add(
                        line);
                }

                line.Add(
                    styledWord);
            }

            var result =
                new List<StyledTextLine>();

            foreach (var physicalLine in
                     physicalLines)
            {
                var ordered =
                    physicalLine
                        .OrderBy(item =>
                            item.Word.BoundingBox.Left)
                        .ToArray();

                if (ordered.Length == 0)
                {
                    continue;
                }

                var averageHeight =
                    ordered.Average(item =>
                        item.Word.BoundingBox.Height);

                var gapThreshold =
                    Math.Max(
                        page.Width *
                        0.035,

                        averageHeight *
                        2.8);

                var current =
                    new List<StyledWord>
                    {
                        ordered[0]
                    };

                for (var index = 1;
                     index < ordered.Length;
                     index++)
                {
                    var gap =
                        ordered[index]
                            .Word
                            .BoundingBox
                            .Left -
                        ordered[index - 1]
                            .Word
                            .BoundingBox
                            .Right;

                    if (gap >
                        gapThreshold)
                    {
                        AddTextSegment(
                            result,
                            current);

                        current =
                            new List<StyledWord>();
                    }

                    current.Add(
                        ordered[index]);
                }

                AddTextSegment(
                    result,
                    current);
            }

            return result
                .OrderByDescending(line =>
                    line.Top)
                .ThenBy(line =>
                    line.Left)
                .ToArray();
        }
        finally
        {
            bitmap?.Dispose();
        }
    }

    private static TextStyle GetWordStyle(
        PdfWord word,
        IReadOnlyList<PdfLetter> letters,
        PdfPage page,
        System.Drawing.Bitmap? bitmap)
    {
        var wordLetters =
            letters
                .Where(letter =>
                    IsLetterInsideWord(
                        letter,
                        word))
                .ToArray();

        var pointSize =
            Math.Max(
                6,
                word.BoundingBox.Height *
                0.75);

        string? fontName =
            null;

        var bold =
            false;

        var italic =
            false;

        if (wordLetters.Length > 0)
        {
            var sizes =
                wordLetters
                    .Select(letter =>
                        Convert.ToDouble(
                            letter.PointSize))
                    .Where(size =>
                        size > 0 &&
                        size < 200)
                    .ToArray();

            if (sizes.Length > 0)
            {
                pointSize =
                    sizes.Average();
            }

            var rawFontName =
                wordLetters
                    .Select(letter =>
                        letter.FontName)
                    .FirstOrDefault(name =>
                        !string.IsNullOrWhiteSpace(
                            name));

            fontName =
                NormalizeFontName(
                    rawFontName);

            var source =
                rawFontName ??
                string.Empty;

            bold =
                ContainsAny(
                    source,
                    "bold",
                    "semibold",
                    "semi-bold",
                    "demibold",
                    "black",
                    "heavy");

            italic =
                ContainsAny(
                    source,
                    "italic",
                    "oblique",
                    "slanted");
        }

        return new TextStyle(
            pointSize,
            fontName,
            bold,
            italic,
            EstimateWordTextColor(
                page,
                word,
                bitmap));
    }

    private static string EstimateWordTextColor(
        PdfPage page,
        PdfWord word,
        System.Drawing.Bitmap? bitmap)
    {
        if (bitmap is null)
        {
            return "000000";
        }

        var rectangle =
            PdfRectangleToPixels(
                word.BoundingBox.Left,
                word.BoundingBox.Bottom,
                word.BoundingBox.Right,
                word.BoundingBox.Top,
                page.Width,
                page.Height,
                bitmap.Width,
                bitmap.Height);

        rectangle =
            ClampRectangle(
                rectangle,
                bitmap.Width,
                bitmap.Height);

        var backgroundInfo =
            AnalyzeLocalBackground(
                bitmap,
                rectangle);

        var background =
            backgroundInfo.Color;

        var backgroundBrightness =
            GetBrightness(
                background);

        //
        // Самое важное изменение:
        // для явно тёмного однородного фона
        // не доверяем anti-aliasing вообще.
        //
        if (backgroundInfo.Dominance >
                0.55 &&
            backgroundBrightness <
                105)
        {
            return "FFFFFF";
        }

        var pixels =
            new List<System.Drawing.Color>();

        var stepX =
            Math.Max(
                1,
                rectangle.Width /
                40);

        var stepY =
            Math.Max(
                1,
                rectangle.Height /
                18);

        for (var y = rectangle.Top;
             y < rectangle.Bottom;
             y += stepY)
        {
            for (var x = rectangle.Left;
                 x < rectangle.Right;
                 x += stepX)
            {
                var color =
                    bitmap.GetPixel(
                        Math.Clamp(
                            x,
                            0,
                            bitmap.Width - 1),

                        Math.Clamp(
                            y,
                            0,
                            bitmap.Height - 1));

                if (ColorDistance(
                        color,
                        background) <
                    26)
                {
                    continue;
                }

                pixels.Add(
                    color);
            }
        }

        if (pixels.Count == 0)
        {
            return backgroundBrightness <
                   128
                ? "FFFFFF"
                : "000000";
        }

        IEnumerable<System.Drawing.Color>
            candidates;

        if (backgroundBrightness <
            125)
        {
            candidates =
                pixels
                    .Where(color =>
                        GetBrightness(
                            color) >
                        backgroundBrightness +
                        50)
                    .OrderByDescending(color =>
                        GetBrightness(
                            color));
        }
        else if (backgroundBrightness >
                 175)
        {
            candidates =
                pixels
                    .Where(color =>
                        GetBrightness(
                            color) <
                        backgroundBrightness -
                        45)
                    .OrderBy(color =>
                        GetBrightness(
                            color));
        }
        else
        {
            candidates =
                pixels
                    .OrderByDescending(color =>
                        ColorDistance(
                            color,
                            background));
        }

        var selected =
            candidates
                .Take(
                    Math.Max(
                        4,
                        pixels.Count /
                        3))
                .ToArray();

        if (selected.Length == 0)
        {
            return backgroundBrightness <
                   128
                ? "FFFFFF"
                : "000000";
        }

        var sampled =
            System.Drawing.Color
                .FromArgb(
                    Median(
                        selected.Select(color =>
                            (int)color.R)),

                    Median(
                        selected.Select(color =>
                            (int)color.G)),

                    Median(
                        selected.Select(color =>
                            (int)color.B)));

        //
        // Ещё одна защита:
        // тёмная плашка + слишком тёмный
        // результат = почти наверняка
        // anti-aliasing.
        //
        if (backgroundBrightness <
                125 &&
            GetBrightness(
                sampled) <
                155)
        {
            return "FFFFFF";
        }

        if (backgroundBrightness >
                225 &&
            GetBrightness(
                sampled) >
                170)
        {
            return "000000";
        }

        return
            $"{sampled.R:X2}" +
            $"{sampled.G:X2}" +
            $"{sampled.B:X2}";
    }

    private static void AddTextSegment(
        ICollection<StyledTextLine> result,
        IReadOnlyList<StyledWord> words)
    {
        if (words.Count == 0)
        {
            return;
        }

        var ordered =
            words
                .OrderBy(item =>
                    item.Word.BoundingBox.Left)
                .ToArray();

        result.Add(
            new StyledTextLine(
                ordered.Min(item =>
                    item.Word.BoundingBox.Left),

                ordered.Min(item =>
                    item.Word.BoundingBox.Bottom),

                ordered.Max(item =>
                    item.Word.BoundingBox.Right),

                ordered.Max(item =>
                    item.Word.BoundingBox.Top),

                ordered.Max(item =>
                    item.Word.BoundingBox.Right) -
                ordered.Min(item =>
                    item.Word.BoundingBox.Left),

                ordered.Max(item =>
                    item.Word.BoundingBox.Top) -
                ordered.Min(item =>
                    item.Word.BoundingBox.Bottom),

                ordered
                    .Select(item =>
                        new StyledRun(
                            item.Word.Text,
                            item.Style))
                    .ToArray()));
    }

    private static bool IsInsideAnyTable(
        PdfWord word,
        IReadOnlyList<PdfTableExtractionService.ExtractedTable>
            tables)
    {
        var centerX =
            (
                word.BoundingBox.Left +
                word.BoundingBox.Right
            ) /
            2.0;

        var centerY =
            (
                word.BoundingBox.Bottom +
                word.BoundingBox.Top
            ) /
            2.0;

        foreach (var table in tables)
        {
            if (centerX >= table.Left &&
                centerX <= table.Right &&
                centerY >= table.Bottom &&
                centerY <= table.Top)
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsLetterInsideWord(
        PdfLetter letter,
        PdfWord word)
    {
        var centerX =
            (
                letter.BoundingBox.Left +
                letter.BoundingBox.Right
            ) /
            2.0;

        var centerY =
            (
                letter.BoundingBox.Bottom +
                letter.BoundingBox.Top
            ) /
            2.0;

        const double tolerance =
            1.5;

        return
            centerX >=
                word.BoundingBox.Left -
                tolerance &&
            centerX <=
                word.BoundingBox.Right +
                tolerance &&
            centerY >=
                word.BoundingBox.Bottom -
                tolerance &&
            centerY <=
                word.BoundingBox.Top +
                tolerance;
    }

    private static P.Shape CreateStyledTextShape(
        uint shapeId,
        StyledTextLine line,
        double pdfWidth,
        double pdfHeight,
        long slideWidth,
        long slideHeight)
    {
        var x =
            ConvertX(
                line.Left,
                pdfWidth,
                slideWidth);

        var y =
            ConvertY(
                line.Top,
                pdfHeight,
                slideHeight);

        var width =
            ConvertLengthX(
                Math.Max(
                    line.Width *
                    1.06,
                    20),
                pdfWidth,
                slideWidth);

        var maxPointSize =
            line.Runs
                .Select(run =>
                    run.Style.PointSize)
                .DefaultIfEmpty(10)
                .Max();

        var height =
            ConvertLengthY(
                Math.Max(
                    line.Height *
                    1.35,
                    maxPointSize *
                    1.28),
                pdfHeight,
                slideHeight);

        var paragraph =
            new A.Paragraph(
                new A.ParagraphProperties
                {
                    Alignment =
                        A.TextAlignmentTypeValues.Left
                });

        for (var index = 0;
             index < line.Runs.Count;
             index++)
        {
            var run =
                line.Runs[index];

            var properties =
                new A.RunProperties
                {
                    FontSize =
                        Math.Clamp(
                            (int)Math.Round(
                                run.Style.PointSize *
                                100),
                            500,
                            7200),

                    Bold =
                        run.Style.Bold,

                    Italic =
                        run.Style.Italic,

                    Language =
                        "ru-RU"
                };

            if (!string.IsNullOrWhiteSpace(
                    run.Style.FontName))
            {
                properties.Append(
                    new A.LatinFont
                    {
                        Typeface =
                            run.Style.FontName
                    });
            }

            properties.Append(
                new A.SolidFill(
                    new A.RgbColorModelHex
                    {
                        Val =
                            run.Style.TextColor
                    }));

            paragraph.Append(
                new A.Run(
                    properties,
                    new A.Text(
                        index == 0
                            ? run.Text
                            : " " +
                              run.Text)));
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(

                new P.NonVisualDrawingProperties
                {
                    Id =
                        shapeId,

                    Name =
                        $"Text {shapeId}"
                },

                new P.NonVisualShapeDrawingProperties(
                    new A.ShapeLocks
                    {
                        NoGrouping =
                            true
                    }),

                new P.ApplicationNonVisualDrawingProperties()),

            new P.ShapeProperties(

                new A.Transform2D(

                    new A.Offset
                    {
                        X = x,
                        Y = y
                    },

                    new A.Extents
                    {
                        Cx = width,
                        Cy = height
                    }),

                new A.PresetGeometry(
                    new A.AdjustValueList())
                {
                    Preset =
                        A.ShapeTypeValues.Rectangle
                }),

            new P.TextBody(

                new A.BodyProperties
                {
                    Wrap =
                        A.TextWrappingValues.None,

                    LeftInset = 0,
                    RightInset = 0,
                    TopInset = 0,
                    BottomInset = 0
                },

                new A.ListStyle(),

                paragraph));
    }

    // ============================================================
    // TABLE CONTENT
    // ============================================================

    private static P.GraphicFrame CreateTableFrame(
        uint shapeId,
        PdfPage page,
        PdfTableExtractionService.ExtractedTable table,
        double pdfWidth,
        double pdfHeight,
        long slideWidth,
        long slideHeight,
        string? renderedPagePath)
    {
        var x =
            ConvertX(
                table.Left,
                pdfWidth,
                slideWidth);

        var y =
            ConvertY(
                table.Top,
                pdfHeight,
                slideHeight);

        var width =
            ConvertLengthX(
                table.Right -
                table.Left,
                pdfWidth,
                slideWidth);

        var height =
            ConvertLengthY(
                table.Top -
                table.Bottom,
                pdfHeight,
                slideHeight);

        var columnWidths =
            EstimateColumnWidths(
                page,
                table,
                width);

        var rowHeights =
            EstimateRowHeights(
                page,
                table,
                height);

        var geometry =
            BuildTableGeometry(
                table,
                columnWidths,
                rowHeights);

        var reconstructedRows =
            ReconstructTableRows(
                page,
                geometry);

        var fills =
            EstimateTableCellFills(
                page,
                table,
                renderedPagePath,
                geometry);

        var styles =
            EstimateTableCellStyles(
                page,
                reconstructedRows,
                renderedPagePath,
                geometry);

        var borders =
            EstimateTableCellBorders(
                page,
                table,
                renderedPagePath,
                geometry);

        var drawingTable =
            CreateDrawingTable(
                reconstructedRows,
                columnWidths,
                rowHeights,
                fills,
                styles,
                borders);

        return new P.GraphicFrame(

            new P.NonVisualGraphicFrameProperties(

                new P.NonVisualDrawingProperties
                {
                    Id =
                        shapeId,

                    Name =
                        $"Table {shapeId}"
                },

                new P.NonVisualGraphicFrameDrawingProperties(),

                new P.ApplicationNonVisualDrawingProperties()),

            new P.Transform(

                new A.Offset
                {
                    X = x,
                    Y = y
                },

                new A.Extents
                {
                    Cx = width,
                    Cy = height
                }),

            new A.Graphic(

                new A.GraphicData(
                    drawingTable)
                {
                    Uri =
                        "http://schemas.openxmlformats.org/drawingml/2006/table"
                }));
    }

    private static IReadOnlyList<IReadOnlyList<string>>
        ReconstructTableRows(
            PdfPage page,
            TableGeometry geometry)
    {
        var rows =
            new List<IReadOnlyList<string>>();

        var rowCount =
            geometry.RowBoundaries.Count -
            1;

        var columnCount =
            geometry.ColumnBoundaries.Count -
            1;

        var words =
            page.GetWords()
                .Where(word =>
                    !string.IsNullOrWhiteSpace(
                        word.Text))
                .ToArray();

        for (var rowIndex = 0;
             rowIndex < rowCount;
             rowIndex++)
        {
            var cells =
                new List<string>();

            for (var columnIndex = 0;
                 columnIndex < columnCount;
                 columnIndex++)
            {
                var bounds =
                    GetCellBounds(
                        geometry,
                        rowIndex,
                        columnIndex);

                var cellWords =
                    words
                        .Where(word =>
                            IsWordInsideBounds(
                                word,
                                bounds))
                        .ToArray();

                cells.Add(
                    BuildCellText(
                        cellWords));
            }

            rows.Add(
                cells);
        }

        return rows;
    }

    private static string BuildCellText(
        IReadOnlyList<PdfWord> words)
    {
        if (words.Count == 0)
        {
            return string.Empty;
        }

        var lines =
            ClusterWordsIntoPhysicalRows(
                words)
            .OrderByDescending(line =>
                line.Average(word =>
                    word.BoundingBox.Top))
            .Select(line =>
                string.Join(
                    " ",
                    line
                        .OrderBy(word =>
                            word.BoundingBox.Left)
                        .Select(word =>
                            word.Text.Trim())
                        .Where(text =>
                            text.Length > 0)))
            .Where(text =>
                text.Length > 0)
            .ToArray();

        return string.Join(
            "\n",
            lines);
    }

    private static A.Table CreateDrawingTable(
        IReadOnlyList<IReadOnlyList<string>> rows,
        IReadOnlyList<long> columnWidths,
        IReadOnlyList<long> rowHeights,
        IReadOnlyList<IReadOnlyList<string?>> cellFills,
        IReadOnlyList<IReadOnlyList<TableCellStyle>> cellStyles,
        IReadOnlyList<IReadOnlyList<TableCellBorderStyle>> cellBorders)
    {
        var table =
            new A.Table();

        table.Append(
            new A.TableProperties
            {
                FirstRow =
                    true
            });

        var columnCount =
            Math.Max(
                1,
                columnWidths.Count);

        var grid =
            new A.TableGrid();

        foreach (var width in
                 columnWidths)
        {
            grid.Append(
                new A.GridColumn
                {
                    Width =
                        Math.Max(
                            1,
                            width)
                });
        }

        table.Append(
            grid);

        for (var rowIndex = 0;
             rowIndex < rows.Count;
             rowIndex++)
        {
            var row =
                new A.TableRow
                {
                    Height =
                        Math.Max(
                            1,
                            rowIndex <
                            rowHeights.Count
                                ? rowHeights[
                                    rowIndex]
                                : 400000)
                };

            for (var columnIndex = 0;
                 columnIndex < columnCount;
                 columnIndex++)
            {
                var value =
                    columnIndex <
                    rows[rowIndex].Count
                        ? rows[rowIndex][
                            columnIndex]
                        : string.Empty;

                var style =
                    rowIndex <
                        cellStyles.Count &&
                    columnIndex <
                        cellStyles[
                            rowIndex].Count
                        ? cellStyles[
                            rowIndex][
                            columnIndex]
                        : new TableCellStyle(
                            10,
                            null,
                            false,
                            false,
                            "000000",
                            GuessCellAlignment(
                                value));

                var body =
                    CreateTableTextBody(
                        value,
                        style);

                var fill =
                    rowIndex <
                        cellFills.Count &&
                    columnIndex <
                        cellFills[
                            rowIndex].Count
                        ? cellFills[
                            rowIndex][
                            columnIndex]
                        : null;

                var border =
                    rowIndex <
                        cellBorders.Count &&
                    columnIndex <
                        cellBorders[
                            rowIndex].Count
                        ? cellBorders[
                            rowIndex][
                            columnIndex]
                        : CreateDefaultCellBorderStyle();

                row.Append(
                    new A.TableCell(
                        body,
                        CreateTableCellProperties(
                            rowIndex == 0,
                            fill,
                            border)));
            }

            table.Append(
                row);
        }

        return table;
    }

    private static A.TextBody CreateTableTextBody(
        string value,
        TableCellStyle style)
    {
        var body =
            new A.TextBody(
                new A.BodyProperties
                {
                    Wrap =
                        A.TextWrappingValues.Square
                },
                new A.ListStyle());

        var lines =
            (
                value ??
                string.Empty
            )
            .Replace(
                "\r\n",
                "\n")
            .Split(
                '\n');

        if (lines.Length == 0)
        {
            lines =
                new[]
                {
                    string.Empty
                };
        }

        foreach (var line in lines)
        {
            var paragraphProperties =
                new A.ParagraphProperties
                {
                    Alignment =
                        style.Alignment
                };

            var runProperties =
                new A.RunProperties
                {
                    FontSize =
                        Math.Clamp(
                            (int)Math.Round(
                                style.FontSize *
                                100),
                            500,
                            7200),

                    Bold =
                        style.Bold,

                    Italic =
                        style.Italic,

                    Language =
                        "ru-RU"
                };

            if (!string.IsNullOrWhiteSpace(
                    style.FontName))
            {
                runProperties.Append(
                    new A.LatinFont
                    {
                        Typeface =
                            style.FontName
                    });
            }

            runProperties.Append(
                new A.SolidFill(
                    new A.RgbColorModelHex
                    {
                        Val =
                            style.TextColor
                    }));

            body.Append(
                new A.Paragraph(
                    paragraphProperties,

                    new A.Run(
                        runProperties,

                        new A.Text(
                            line))));
        }

        return body;
    }

    // ============================================================
    // TABLE CELL STYLE
    // ============================================================

    private static IReadOnlyList<IReadOnlyList<TableCellStyle>>
        EstimateTableCellStyles(
            PdfPage page,
            IReadOnlyList<IReadOnlyList<string>> reconstructedRows,
            string? renderedPagePath,
            TableGeometry geometry)
    {
        var result =
            new List<IReadOnlyList<TableCellStyle>>();

        var words =
            page.GetWords()
                .Where(word =>
                    !string.IsNullOrWhiteSpace(
                        word.Text))
                .ToArray();

        using var bitmap =
            !string.IsNullOrWhiteSpace(
                renderedPagePath) &&
            File.Exists(
                renderedPagePath)
                ? new System.Drawing.Bitmap(
                    renderedPagePath)
                : null;

        for (var rowIndex = 0;
             rowIndex <
             geometry.RowBoundaries.Count -
             1;
             rowIndex++)
        {
            var row =
                new List<TableCellStyle>();

            for (var columnIndex = 0;
                 columnIndex <
                 geometry.ColumnBoundaries.Count -
                 1;
                 columnIndex++)
            {
                var bounds =
                    GetCellBounds(
                        geometry,
                        rowIndex,
                        columnIndex);

                var cellWords =
                    words
                        .Where(word =>
                            IsWordInsideBounds(
                                word,
                                bounds))
                        .ToArray();

                var value =
                    rowIndex <
                        reconstructedRows.Count &&
                    columnIndex <
                        reconstructedRows[
                            rowIndex].Count
                        ? reconstructedRows[
                            rowIndex][
                            columnIndex]
                        : string.Empty;

                if (cellWords.Length == 0)
                {
                    row.Add(
                        new TableCellStyle(
                            rowIndex == 0
                                ? 11
                                : 10,
                            null,
                            rowIndex == 0,
                            false,
                            "000000",
                            GuessCellAlignment(
                                value)));

                    continue;
                }

                var styles =
                    cellWords
                        .Select(word =>
                            GetWordStyle(
                                word,
                                page.Letters,
                                page,
                                bitmap))
                        .ToArray();

                row.Add(
                    new TableCellStyle(
                        styles
                            .Average(style =>
                                style.PointSize),

                        styles
                            .Where(style =>
                                !string.IsNullOrWhiteSpace(
                                    style.FontName))
                            .GroupBy(style =>
                                style.FontName)
                            .OrderByDescending(group =>
                                group.Count())
                            .Select(group =>
                                group.Key)
                            .FirstOrDefault(),

                        styles.Count(style =>
                            style.Bold) >=
                        Math.Ceiling(
                            styles.Length /
                            2.0),

                        styles.Count(style =>
                            style.Italic) >=
                        Math.Ceiling(
                            styles.Length /
                            2.0),

                        styles
                            .GroupBy(style =>
                                style.TextColor)
                            .OrderByDescending(group =>
                                group.Count())
                            .Select(group =>
                                group.Key)
                            .FirstOrDefault()
                        ??
                        "000000",

                        DetectCellAlignment(
                            cellWords,
                            bounds.Left,
                            bounds.Right,
                            value)));
            }

            result.Add(
                row);
        }

        return result;
    }

    private static IReadOnlyList<IReadOnlyList<string?>>
        EstimateTableCellFills(
            PdfPage page,
            PdfTableExtractionService.ExtractedTable table,
            string? renderedPagePath,
            TableGeometry geometry)
    {
        var result =
            new List<IReadOnlyList<string?>>();

        if (string.IsNullOrWhiteSpace(
                renderedPagePath) ||
            !File.Exists(
                renderedPagePath))
        {
            for (var row = 0;
                 row <
                 geometry.RowBoundaries.Count -
                 1;
                 row++)
            {
                result.Add(
                    Enumerable
                        .Repeat<string?>(
                            null,
                            geometry
                                .ColumnBoundaries
                                .Count -
                            1)
                        .ToArray());
            }

            return result;
        }

        using var bitmap =
            new System.Drawing.Bitmap(
                renderedPagePath);

        for (var rowIndex = 0;
             rowIndex <
             geometry.RowBoundaries.Count -
             1;
             rowIndex++)
        {
            var row =
                new List<string?>();

            for (var columnIndex = 0;
                 columnIndex <
                 geometry.ColumnBoundaries.Count -
                 1;
                 columnIndex++)
            {
                var bounds =
                    GetCellBounds(
                        geometry,
                        rowIndex,
                        columnIndex);

                row.Add(
                    SampleCellBackgroundColor(
                        bitmap,
                        page,
                        bounds.Left,
                        bounds.Bottom,
                        bounds.Right,
                        bounds.Top));
            }

            result.Add(
                row);
        }

        return result;
    }

    private static string?
        SampleCellBackgroundColor(
            System.Drawing.Bitmap bitmap,
            PdfPage page,
            double left,
            double bottom,
            double right,
            double top)
    {
        var rectangle =
            PdfRectangleToPixels(
                left,
                bottom,
                right,
                top,
                page.Width,
                page.Height,
                bitmap.Width,
                bitmap.Height);

        var insetX =
            Math.Max(
                2,
                rectangle.Width /
                5);

        var insetY =
            Math.Max(
                2,
                rectangle.Height /
                5);

        rectangle =
            System.Drawing.Rectangle
                .FromLTRB(
                    rectangle.Left +
                    insetX,

                    rectangle.Top +
                    insetY,

                    rectangle.Right -
                    insetX,

                    rectangle.Bottom -
                    insetY);

        rectangle =
            ClampRectangle(
                rectangle,
                bitmap.Width,
                bitmap.Height);

        if (rectangle.Width <= 1 ||
            rectangle.Height <= 1)
        {
            return null;
        }

        var colors =
            new List<System.Drawing.Color>();

        for (var iy = 0;
             iy < 4;
             iy++)
        {
            for (var ix = 0;
                 ix < 4;
                 ix++)
            {
                var x =
                    rectangle.Left +
                    rectangle.Width *
                    ix /
                    3;

                var y =
                    rectangle.Top +
                    rectangle.Height *
                    iy /
                    3;

                colors.Add(
                    bitmap.GetPixel(
                        Math.Clamp(
                            x,
                            0,
                            bitmap.Width - 1),
                        Math.Clamp(
                            y,
                            0,
                            bitmap.Height - 1)));
            }
        }

        var red =
            Median(
                colors.Select(color =>
                    (int)color.R));

        var green =
            Median(
                colors.Select(color =>
                    (int)color.G));

        var blue =
            Median(
                colors.Select(color =>
                    (int)color.B));

        if (red > 245 &&
            green > 245 &&
            blue > 245)
        {
            return null;
        }

        return
            $"{red:X2}{green:X2}{blue:X2}";
    }

    private static IReadOnlyList<IReadOnlyList<TableCellBorderStyle>>
        EstimateTableCellBorders(
            PdfPage page,
            PdfTableExtractionService.ExtractedTable table,
            string? renderedPagePath,
            TableGeometry geometry)
    {
        var result =
            new List<IReadOnlyList<TableCellBorderStyle>>();

        if (string.IsNullOrWhiteSpace(
                renderedPagePath) ||
            !File.Exists(
                renderedPagePath))
        {
            for (var row = 0;
                 row <
                 geometry.RowBoundaries.Count -
                 1;
                 row++)
            {
                result.Add(
                    Enumerable
                        .Range(
                            0,
                            geometry
                                .ColumnBoundaries
                                .Count -
                            1)
                        .Select(_ =>
                            CreateDefaultCellBorderStyle())
                        .ToArray());
            }

            return result;
        }

        using var bitmap =
            new System.Drawing.Bitmap(
                renderedPagePath);

        for (var rowIndex = 0;
             rowIndex <
             geometry.RowBoundaries.Count -
             1;
             rowIndex++)
        {
            var row =
                new List<TableCellBorderStyle>();

            for (var columnIndex = 0;
                 columnIndex <
                 geometry.ColumnBoundaries.Count -
                 1;
                 columnIndex++)
            {
                var bounds =
                    GetCellBounds(
                        geometry,
                        rowIndex,
                        columnIndex);

                row.Add(
                    new TableCellBorderStyle(

                        SampleVerticalBorder(
                            bitmap,
                            page,
                            bounds.Left,
                            bounds.Bottom,
                            bounds.Top)
                        ??
                        CreateDefaultBorderSide(),

                        SampleVerticalBorder(
                            bitmap,
                            page,
                            bounds.Right,
                            bounds.Bottom,
                            bounds.Top)
                        ??
                        CreateDefaultBorderSide(),

                        SampleHorizontalBorder(
                            bitmap,
                            page,
                            bounds.Left,
                            bounds.Right,
                            bounds.Top)
                        ??
                        CreateDefaultBorderSide(),

                        SampleHorizontalBorder(
                            bitmap,
                            page,
                            bounds.Left,
                            bounds.Right,
                            bounds.Bottom)
                        ??
                        CreateDefaultBorderSide()));
            }

            result.Add(
                row);
        }

        return result;
    }

    private static PdfBounds GetCellBounds(
        TableGeometry geometry,
        int rowIndex,
        int columnIndex)
    {
        return new PdfBounds(
            geometry.ColumnBoundaries[
                columnIndex],

            geometry.RowBoundaries[
                rowIndex + 1],

            geometry.ColumnBoundaries[
                columnIndex + 1],

            geometry.RowBoundaries[
                rowIndex]);
    }

    private static bool IsWordInsideBounds(
        PdfWord word,
        PdfBounds bounds)
    {
        var centerX =
            (
                word.BoundingBox.Left +
                word.BoundingBox.Right
            ) /
            2.0;

        var centerY =
            (
                word.BoundingBox.Bottom +
                word.BoundingBox.Top
            ) /
            2.0;

        return
            centerX >= bounds.Left &&
            centerX <= bounds.Right &&
            centerY >= bounds.Bottom &&
            centerY <= bounds.Top;
    }

    private static A.TextAlignmentTypeValues
        GuessCellAlignment(
            string value)
    {
        var normalized =
            (
                value ??
                string.Empty
            )
            .Trim()
            .Replace(
                ",",
                ".",
                StringComparison.Ordinal)
            .Replace(
                "%",
                string.Empty,
                StringComparison.Ordinal)
            .Replace(
                "$",
                string.Empty,
                StringComparison.Ordinal);

        if (double.TryParse(
                normalized,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out _))
        {
            return
                A.TextAlignmentTypeValues.Right;
        }

        return
            A.TextAlignmentTypeValues.Left;
    }

    private static A.TextAlignmentTypeValues
        DetectCellAlignment(
            IReadOnlyList<PdfWord> words,
            double cellLeft,
            double cellRight,
            string value)
    {
        if (words.Count == 0)
        {
            return GuessCellAlignment(
                value);
        }

        var textLeft =
            words.Min(word =>
                word.BoundingBox.Left);

        var textRight =
            words.Max(word =>
                word.BoundingBox.Right);

        var leftGap =
            textLeft -
            cellLeft;

        var rightGap =
            cellRight -
            textRight;

        var width =
            Math.Max(
                1,
                cellRight -
                cellLeft);

        if (Math.Abs(
                leftGap -
                rightGap) <
            width *
            0.08)
        {
            return
                A.TextAlignmentTypeValues.Center;
        }

        if (rightGap <
            leftGap *
            0.55)
        {
            return
                A.TextAlignmentTypeValues.Right;
        }

        return GuessCellAlignment(
            value);
    }

    // ============================================================
    // BORDERS
    // ============================================================

    private static TableBorderSide?
        SampleVerticalBorder(
            System.Drawing.Bitmap bitmap,
            PdfPage page,
            double pdfX,
            double pdfBottom,
            double pdfTop)
    {
        var pixelX =
            (int)Math.Round(
                pdfX /
                page.Width *
                bitmap.Width);

        var top =
            (int)Math.Round(
                (
                    page.Height -
                    pdfTop
                ) /
                page.Height *
                bitmap.Height);

        var bottom =
            (int)Math.Round(
                (
                    page.Height -
                    pdfBottom
                ) /
                page.Height *
                bitmap.Height);

        pixelX =
            Math.Clamp(
                pixelX,
                0,
                bitmap.Width - 1);

        top =
            Math.Clamp(
                top,
                0,
                bitmap.Height - 1);

        bottom =
            Math.Clamp(
                bottom,
                0,
                bitmap.Height);

        if (bottom <= top)
        {
            return null;
        }

        var values =
            new List<(int Pos, double Brightness)>();

        for (var x =
                 Math.Max(
                     0,
                     pixelX - 4);
             x <=
             Math.Min(
                 bitmap.Width - 1,
                 pixelX + 4);
             x++)
        {
            var samples =
                new List<double>();

            var step =
                Math.Max(
                    1,
                    (
                        bottom -
                        top
                    ) /
                    30);

            for (var y = top;
                 y < bottom;
                 y += step)
            {
                samples.Add(
                    GetBrightness(
                        bitmap.GetPixel(
                            x,
                            y)));
            }

            values.Add(
                (
                    x,
                    samples.Average()
                ));
        }

        var darkest =
            values
                .OrderBy(value =>
                    value.Brightness)
                .First();

        var median =
            MedianDouble(
                values.Select(value =>
                    value.Brightness));

        if (median -
            darkest.Brightness <
            14)
        {
            return null;
        }

        var positions =
            values
                .Where(value =>
                    value.Brightness <=
                    darkest.Brightness +
                    12)
                .Select(value =>
                    value.Pos)
                .ToArray();

        var colors =
            new List<System.Drawing.Color>();

        foreach (var x in positions)
        {
            var step =
                Math.Max(
                    1,
                    (
                        bottom -
                        top
                    ) /
                    30);

            for (var y = top;
                 y < bottom;
                 y += step)
            {
                colors.Add(
                    bitmap.GetPixel(
                        x,
                        y));
            }
        }

        return new TableBorderSide(
            GetMedianColorHex(
                colors)
            ??
            "BFBFBF",

            ConvertPixelThicknessToEmu(
                positions.Length));
    }

    private static TableBorderSide?
        SampleHorizontalBorder(
            System.Drawing.Bitmap bitmap,
            PdfPage page,
            double pdfLeft,
            double pdfRight,
            double pdfY)
    {
        var left =
            (int)Math.Round(
                pdfLeft /
                page.Width *
                bitmap.Width);

        var right =
            (int)Math.Round(
                pdfRight /
                page.Width *
                bitmap.Width);

        var pixelY =
            (int)Math.Round(
                (
                    page.Height -
                    pdfY
                ) /
                page.Height *
                bitmap.Height);

        left =
            Math.Clamp(
                left,
                0,
                bitmap.Width - 1);

        right =
            Math.Clamp(
                right,
                0,
                bitmap.Width);

        pixelY =
            Math.Clamp(
                pixelY,
                0,
                bitmap.Height - 1);

        if (right <= left)
        {
            return null;
        }

        var values =
            new List<(int Pos, double Brightness)>();

        for (var y =
                 Math.Max(
                     0,
                     pixelY - 4);
             y <=
             Math.Min(
                 bitmap.Height - 1,
                 pixelY + 4);
             y++)
        {
            var samples =
                new List<double>();

            var step =
                Math.Max(
                    1,
                    (
                        right -
                        left
                    ) /
                    30);

            for (var x = left;
                 x < right;
                 x += step)
            {
                samples.Add(
                    GetBrightness(
                        bitmap.GetPixel(
                            x,
                            y)));
            }

            values.Add(
                (
                    y,
                    samples.Average()
                ));
        }

        var darkest =
            values
                .OrderBy(value =>
                    value.Brightness)
                .First();

        var median =
            MedianDouble(
                values.Select(value =>
                    value.Brightness));

        if (median -
            darkest.Brightness <
            14)
        {
            return null;
        }

        var positions =
            values
                .Where(value =>
                    value.Brightness <=
                    darkest.Brightness +
                    12)
                .Select(value =>
                    value.Pos)
                .ToArray();

        var colors =
            new List<System.Drawing.Color>();

        foreach (var y in positions)
        {
            var step =
                Math.Max(
                    1,
                    (
                        right -
                        left
                    ) /
                    30);

            for (var x = left;
                 x < right;
                 x += step)
            {
                colors.Add(
                    bitmap.GetPixel(
                        x,
                        y));
            }
        }

        return new TableBorderSide(
            GetMedianColorHex(
                colors)
            ??
            "BFBFBF",

            ConvertPixelThicknessToEmu(
                positions.Length));
    }

    private static A.TableCellProperties
        CreateTableCellProperties(
            bool isHeader,
            string? fillColor,
            TableCellBorderStyle borderStyle)
    {
        var properties =
            new A.TableCellProperties
            {
                LeftMargin =
                    50000,

                RightMargin =
                    50000,

                TopMargin =
                    25000,

                BottomMargin =
                    25000,

                Anchor =
                    A.TextAnchoringTypeValues.Center
            };

        properties.Append(
            CreateLeftBorder(
                borderStyle.Left));

        properties.Append(
            CreateRightBorder(
                borderStyle.Right));

        properties.Append(
            CreateTopBorder(
                borderStyle.Top));

        properties.Append(
            CreateBottomBorder(
                borderStyle.Bottom));

        if (!string.IsNullOrWhiteSpace(
                fillColor))
        {
            properties.Append(
                new A.SolidFill(
                    new A.RgbColorModelHex
                    {
                        Val =
                            fillColor
                    }));
        }
        else if (isHeader)
        {
            properties.Append(
                new A.SolidFill(
                    new A.RgbColorModelHex
                    {
                        Val =
                            "EDEDED"
                    }));
        }

        return properties;
    }

    private static A.LeftBorderLineProperties
        CreateLeftBorder(
            TableBorderSide style)
    {
        var result =
            new A.LeftBorderLineProperties
            {
                Width =
                    style.Width
            };

        ApplyBorderStyle(
            result,
            style);

        return result;
    }

    private static A.RightBorderLineProperties
        CreateRightBorder(
            TableBorderSide style)
    {
        var result =
            new A.RightBorderLineProperties
            {
                Width =
                    style.Width
            };

        ApplyBorderStyle(
            result,
            style);

        return result;
    }

    private static A.TopBorderLineProperties
        CreateTopBorder(
            TableBorderSide style)
    {
        var result =
            new A.TopBorderLineProperties
            {
                Width =
                    style.Width
            };

        ApplyBorderStyle(
            result,
            style);

        return result;
    }

    private static A.BottomBorderLineProperties
        CreateBottomBorder(
            TableBorderSide style)
    {
        var result =
            new A.BottomBorderLineProperties
            {
                Width =
                    style.Width
            };

        ApplyBorderStyle(
            result,
            style);

        return result;
    }

    private static void ApplyBorderStyle(
        OpenXmlCompositeElement border,
        TableBorderSide style)
    {
        border.Append(
            new A.SolidFill(
                new A.RgbColorModelHex
                {
                    Val =
                        style.Color
                }));

        border.Append(
            new A.PresetDash
            {
                Val =
                    A.PresetLineDashValues.Solid
            });
    }

    private static TableBorderSide
        CreateDefaultBorderSide()
    {
        return new TableBorderSide(
            "BFBFBF",
            12700);
    }

    private static TableCellBorderStyle
        CreateDefaultCellBorderStyle()
    {
        var side =
            CreateDefaultBorderSide();

        return new TableCellBorderStyle(
            side,
            side,
            side,
            side);
    }

    // ============================================================
    // TABLE HELPERS
    // ============================================================

    private static int CountNonEmptyCells(
        IReadOnlyList<string> row)
    {
        return row.Count(cell =>
            !string.IsNullOrWhiteSpace(
                cell));
    }

    private static double GetTableQuality(
        PdfTableExtractionService.ExtractedTable table)
    {
        var counts =
            table.Rows
                .Select(
                    CountNonEmptyCells)
                .ToArray();

        return
            counts.Count(value =>
                value >= 2) *
            10 +
            counts.Sum() -
            counts.Count(value =>
                value <= 1) *
            8;
    }

    private static bool AreSameTableArea(
        PdfTableExtractionService.ExtractedTable first,
        PdfTableExtractionService.ExtractedTable second)
    {
        var left =
            Math.Max(
                first.Left,
                second.Left);

        var right =
            Math.Min(
                first.Right,
                second.Right);

        var bottom =
            Math.Max(
                first.Bottom,
                second.Bottom);

        var top =
            Math.Min(
                first.Top,
                second.Top);

        if (right <= left ||
            top <= bottom)
        {
            return false;
        }

        var intersection =
            (
                right -
                left
            ) *
            (
                top -
                bottom
            );

        var firstArea =
            (
                first.Right -
                first.Left
            ) *
            (
                first.Top -
                first.Bottom
            );

        var secondArea =
            (
                second.Right -
                second.Left
            ) *
            (
                second.Top -
                second.Bottom
            );

        return
            intersection /
            Math.Max(
                1,
                Math.Min(
                    firstArea,
                    secondArea)) >=
            0.82;
    }

    private static bool TableContainsWord(
        PdfTableExtractionService.ExtractedTable table,
        string word)
    {
        var target =
            NormalizeComparisonText(
                word);

        if (target.Length == 0)
        {
            return false;
        }

        foreach (var cell in
                 table.Rows.SelectMany(row =>
                     row))
        {
            var normalized =
                NormalizeComparisonText(
                    cell);

            if (normalized
                .Split(
                    ' ',
                    StringSplitOptions.RemoveEmptyEntries)
                .Contains(
                    target,
                    StringComparer.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static string NormalizeComparisonText(
        string? value)
    {
        if (string.IsNullOrWhiteSpace(
                value))
        {
            return string.Empty;
        }

        return string.Join(
            " ",
            new string(
                value
                    .ToLowerInvariant()
                    .Select(character =>
                        char.IsLetterOrDigit(
                            character)
                            ? character
                            : ' ')
                    .ToArray())
                .Split(
                    ' ',
                    StringSplitOptions.RemoveEmptyEntries));
    }

    private static bool ContainsBulletMarker(
        string? value)
    {
        if (string.IsNullOrWhiteSpace(
                value))
        {
            return false;
        }

        return
            value.Contains('•') ||
            value.Contains('▪') ||
            value.Contains('◦') ||
            value.Contains('●') ||
            value.Contains('■');
    }

    private static bool IsListMarker(
        string value)
    {
        if (string.IsNullOrWhiteSpace(
                value))
        {
            return false;
        }

        var text =
            value.Trim();

        if (text is
            "•" or
            "▪" or
            "◦" or
            "●" or
            "-" or
            "*" or
            "■")
        {
            return true;
        }

        return false;
    }

    // ============================================================
    // IMAGES
    // ============================================================

    private static bool IsImagePartOfTable(
        PdfImage image,
        IReadOnlyList<PreparedTable> tables)
    {
        var bounds =
            image.BoundingBox;

        var imageArea =
            Math.Max(
                0,
                bounds.Right -
                bounds.Left) *
            Math.Max(
                0,
                bounds.Top -
                bounds.Bottom);

        if (imageArea <= 0)
        {
            return false;
        }

        foreach (var table in tables)
        {
            var left =
                Math.Max(
                    bounds.Left,
                    table.MaskLeft);

            var right =
                Math.Min(
                    bounds.Right,
                    table.MaskRight);

            var bottom =
                Math.Max(
                    bounds.Bottom,
                    table.MaskBottom);

            var top =
                Math.Min(
                    bounds.Top,
                    table.MaskTop);

            if (right <= left ||
                top <= bottom)
            {
                continue;
            }

            var area =
                (
                    right -
                    left
                ) *
                (
                    top -
                    bottom
                );

            if (area /
                imageArea >=
                0.75)
            {
                return true;
            }
        }

        return false;
    }

    private static P.Picture?
        TryCreateImageShape(
            SlidePart slidePart,
            uint shapeId,
            PdfImage image,
            double pdfWidth,
            double pdfHeight,
            long slideWidth,
            long slideHeight)
    {
        if (!TryGetImageData(
                image,
                out var bytes,
                out var type))
        {
            return null;
        }

        var bounds =
            image.BoundingBox;

        var imagePart =
            slidePart.AddImagePart(
                type);

        using (var stream =
               new MemoryStream(
                   bytes,
                   false))
        {
            imagePart.FeedData(
                stream);
        }

        var relationshipId =
            slidePart.GetIdOfPart(
                imagePart);

        return new P.Picture(

            new P.NonVisualPictureProperties(

                new P.NonVisualDrawingProperties
                {
                    Id =
                        shapeId,

                    Name =
                        $"Image {shapeId}"
                },

                new P.NonVisualPictureDrawingProperties(
                    new A.PictureLocks
                    {
                        NoChangeAspect =
                            true
                    }),

                new P.ApplicationNonVisualDrawingProperties()),

            new P.BlipFill(

                new A.Blip
                {
                    Embed =
                        relationshipId
                },

                new A.Stretch(
                    new A.FillRectangle())),

            new P.ShapeProperties(

                new A.Transform2D(

                    new A.Offset
                    {
                        X =
                            ConvertX(
                                bounds.Left,
                                pdfWidth,
                                slideWidth),

                        Y =
                            ConvertY(
                                bounds.Top,
                                pdfHeight,
                                slideHeight)
                    },

                    new A.Extents
                    {
                        Cx =
                            ConvertLengthX(
                                bounds.Width,
                                pdfWidth,
                                slideWidth),

                        Cy =
                            ConvertLengthY(
                                bounds.Height,
                                pdfHeight,
                                slideHeight)
                    }),

                new A.PresetGeometry(
                    new A.AdjustValueList())
                {
                    Preset =
                        A.ShapeTypeValues.Rectangle
                }));
    }

    private static bool TryGetImageData(
        PdfImage image,
        out byte[] bytes,
        out PartTypeInfo type)
    {
        if (image.TryGetPng(
                out var png) &&
            png is not null &&
            png.Length > 0)
        {
            bytes =
                png;

            type =
                ImagePartType.Png;

            return true;
        }

        var raw =
            image.RawMemory
                .ToArray();

        if (raw.Length >= 3 &&
            raw[0] == 0xFF &&
            raw[1] == 0xD8 &&
            raw[2] == 0xFF)
        {
            bytes =
                raw;

            type =
                ImagePartType.Jpeg;

            return true;
        }

        bytes =
            Array.Empty<byte>();

        type =
            ImagePartType.Png;

        return false;
    }

    // ============================================================
    // COMMON HELPERS
    // ============================================================

    private static System.Drawing.Rectangle
        PdfRectangleToPixels(
            double left,
            double bottom,
            double right,
            double top,
            double pdfWidth,
            double pdfHeight,
            int pixelWidth,
            int pixelHeight)
    {
        var x =
            (int)Math.Round(
                left /
                pdfWidth *
                pixelWidth);

        var y =
            (int)Math.Round(
                (
                    pdfHeight -
                    top
                ) /
                pdfHeight *
                pixelHeight);

        var width =
            (int)Math.Round(
                (
                    right -
                    left
                ) /
                pdfWidth *
                pixelWidth);

        var height =
            (int)Math.Round(
                (
                    top -
                    bottom
                ) /
                pdfHeight *
                pixelHeight);

        return ClampRectangle(
            new System.Drawing.Rectangle(
                x,
                y,
                Math.Max(
                    1,
                    width),
                Math.Max(
                    1,
                    height)),
            pixelWidth,
            pixelHeight);
    }

    private static System.Drawing.Rectangle
        ClampRectangle(
            System.Drawing.Rectangle rectangle,
            int width,
            int height)
    {
        var left =
            Math.Clamp(
                rectangle.Left,
                0,
                width);

        var top =
            Math.Clamp(
                rectangle.Top,
                0,
                height);

        var right =
            Math.Clamp(
                rectangle.Right,
                0,
                width);

        var bottom =
            Math.Clamp(
                rectangle.Bottom,
                0,
                height);

        return System.Drawing.Rectangle
            .FromLTRB(
                left,
                top,
                Math.Max(
                    left,
                    right),
                Math.Max(
                    top,
                    bottom));
    }

    private static double GetBrightness(
        System.Drawing.Color color)
    {
        return
            color.R *
            0.299 +
            color.G *
            0.587 +
            color.B *
            0.114;
    }

    private static double ColorDistance(
        System.Drawing.Color first,
        System.Drawing.Color second)
    {
        return Math.Sqrt(
            Math.Pow(
                first.R -
                second.R,
                2) +
            Math.Pow(
                first.G -
                second.G,
                2) +
            Math.Pow(
                first.B -
                second.B,
                2));
    }

    private static double MedianDouble(
        IEnumerable<double> values)
    {
        var array =
            values
                .OrderBy(value =>
                    value)
                .ToArray();

        if (array.Length == 0)
        {
            return 0;
        }

        return array.Length %
               2 == 1
            ? array[
                array.Length /
                2]
            : (
                  array[
                      array.Length /
                      2 -
                      1] +
                  array[
                      array.Length /
                      2]
              ) /
              2.0;
    }

    private static int Median(
        IEnumerable<int> values)
    {
        var array =
            values
                .OrderBy(value =>
                    value)
                .ToArray();

        return array.Length == 0
            ? 0
            : array[
                array.Length /
                2];
    }

    private static string?
        GetMedianColorHex(
            IReadOnlyList<System.Drawing.Color> colors)
    {
        if (colors.Count == 0)
        {
            return null;
        }

        return
            $"{Median(colors.Select(c => (int)c.R)):X2}" +
            $"{Median(colors.Select(c => (int)c.G)):X2}" +
            $"{Median(colors.Select(c => (int)c.B)):X2}";
    }

    private static int ConvertPixelThicknessToEmu(
        int pixels)
    {
        return pixels switch
        {
            <= 1 => 12700,
            2 => 19050,
            3 => 25400,
            4 => 31750,
            _ => 38100
        };
    }

    private static bool HasMeaningfulGraphics(
        System.Drawing.Bitmap bitmap)
    {
        var count =
            0;

        for (var y = 0;
             y < bitmap.Height;
             y += 4)
        {
            for (var x = 0;
                 x < bitmap.Width;
                 x += 4)
            {
                var pixel =
                    bitmap.GetPixel(
                        x,
                        y);

                if (pixel.R < 242 ||
                    pixel.G < 242 ||
                    pixel.B < 242)
                {
                    count++;
                }
            }
        }

        return count >= 250;
    }

    private static P.Picture CreateBackgroundPicture(
        SlidePart slidePart,
        uint shapeId,
        string imagePath,
        long slideWidth,
        long slideHeight)
    {
        var imagePart =
            slidePart.AddImagePart(
                ImagePartType.Png);

        using (var stream =
               File.OpenRead(
                   imagePath))
        {
            imagePart.FeedData(
                stream);
        }

        return new P.Picture(

            new P.NonVisualPictureProperties(

                new P.NonVisualDrawingProperties
                {
                    Id =
                        shapeId,

                    Name =
                        "PDF background"
                },

                new P.NonVisualPictureDrawingProperties(
                    new A.PictureLocks
                    {
                        NoChangeAspect =
                            true
                    }),

                new P.ApplicationNonVisualDrawingProperties()),

            new P.BlipFill(

                new A.Blip
                {
                    Embed =
                        slidePart.GetIdOfPart(
                            imagePart)
                },

                new A.Stretch(
                    new A.FillRectangle())),

            new P.ShapeProperties(

                new A.Transform2D(

                    new A.Offset
                    {
                        X = 0,
                        Y = 0
                    },

                    new A.Extents
                    {
                        Cx =
                            slideWidth,

                        Cy =
                            slideHeight
                    }),

                new A.PresetGeometry(
                    new A.AdjustValueList())
                {
                    Preset =
                        A.ShapeTypeValues.Rectangle
                }));
    }

    private static string? NormalizeFontName(
        string? fontName)
    {
        if (string.IsNullOrWhiteSpace(
                fontName))
        {
            return null;
        }

        var value =
            fontName;

        var plus =
            value.IndexOf('+');

        if (plus >= 0 &&
            plus <
            value.Length - 1)
        {
            value =
                value[
                    (plus + 1)..];
        }

        foreach (var suffix in new[]
                 {
                     "-BoldItalicMT",
                     "-BoldMT",
                     "-ItalicMT",
                     "-Regular",
                     "PSMT"
                 })
        {
            value =
                value.Replace(
                    suffix,
                    string.Empty,
                    StringComparison.OrdinalIgnoreCase);
        }

        return value.Trim();
    }

    private static bool ContainsAny(
        string value,
        params string[] values)
    {
        return values.Any(candidate =>
            value.Contains(
                candidate,
                StringComparison.OrdinalIgnoreCase));
    }

    private static void ApplyPdfAspectRatio(
        P.Presentation presentation,
        double pdfWidth,
        double pdfHeight)
    {
        var slideSize =
            presentation.SlideSize;

        if (slideSize is null ||
            pdfWidth <= 0 ||
            pdfHeight <= 0)
        {
            return;
        }

        var height =
            slideSize.Cy?.Value
            ??
            6858000;

        slideSize.Cx =
            (int)Math.Round(
                height *
                pdfWidth /
                pdfHeight);
    }

    private static long ConvertX(
        double value,
        double width,
        long slideWidth)
    {
        return
            (long)Math.Round(
                value /
                width *
                slideWidth);
    }

    private static long ConvertY(
        double top,
        double height,
        long slideHeight)
    {
        return
            (long)Math.Round(
                (
                    height -
                    top
                ) /
                height *
                slideHeight);
    }

    private static long ConvertLengthX(
        double value,
        double width,
        long slideWidth)
    {
        return
            Math.Max(
                1,
                (long)Math.Round(
                    value /
                    width *
                    slideWidth));
    }

    private static long ConvertLengthY(
        double value,
        double height,
        long slideHeight)
    {
        return
            Math.Max(
                1,
                (long)Math.Round(
                    value /
                    height *
                    slideHeight));
    }

    private static P.ShapeTree CreateEmptyShapeTree()
    {
        return new P.ShapeTree(

            new P.NonVisualGroupShapeProperties(

                new P.NonVisualDrawingProperties
                {
                    Id = 1,
                    Name = string.Empty
                },

                new P.NonVisualGroupShapeDrawingProperties(),

                new P.ApplicationNonVisualDrawingProperties()),

            new P.GroupShapeProperties(

                new A.TransformGroup(

                    new A.Offset
                    {
                        X = 0,
                        Y = 0
                    },

                    new A.Extents
                    {
                        Cx = 0,
                        Cy = 0
                    },

                    new A.ChildOffset
                    {
                        X = 0,
                        Y = 0
                    },

                    new A.ChildExtents
                    {
                        Cx = 0,
                        Cy = 0
                    })));
    }

    private static string CreateUniqueOutputPath(
        string sourcePath)
    {
        var directory =
            Path.GetDirectoryName(
                sourcePath)
            ??
            throw new InvalidOperationException(
                "Unable to determine source directory.");

        var name =
            Path.GetFileNameWithoutExtension(
                sourcePath);

        var candidate =
            Path.Combine(
                directory,
                $"{name}-editable.pptx");

        var index =
            1;

        while (File.Exists(
                   candidate))
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name}-editable ({index++}).pptx");
        }

        return candidate;
    }

    private static int NaturalPageOrder(
        string path)
    {
        var digits =
            new string(
                Path
                    .GetFileNameWithoutExtension(
                        path)
                    .Where(
                        char.IsDigit)
                    .ToArray());

        return int.TryParse(
            digits,
            out var result)
                ? result
                : int.MaxValue;
    }

    // ============================================================
    // MODELS
    // ============================================================

    private sealed record TextStyle(
        double PointSize,
        string? FontName,
        bool Bold,
        bool Italic,
        string TextColor);

    private sealed record StyledWord(
        PdfWord Word,
        TextStyle Style);

    private sealed record StyledRun(
        string Text,
        TextStyle Style);

    private sealed record StyledTextLine(
        double Left,
        double Bottom,
        double Right,
        double Top,
        double Width,
        double Height,
        IReadOnlyList<StyledRun> Runs);

    private sealed record PreparedTable(
        PdfTableExtractionService.ExtractedTable Table,
        double MaskLeft,
        double MaskBottom,
        double MaskRight,
        double MaskTop);

    private sealed record LocalBackgroundInfo(
        System.Drawing.Color Color,
        double Dominance,
        double ColorSpread);

    private sealed record TableGeometry(
        IReadOnlyList<double> ColumnBoundaries,
        IReadOnlyList<double> RowBoundaries);

    private sealed record PdfBounds(
        double Left,
        double Bottom,
        double Right,
        double Top);

    private sealed record TableCellStyle(
        double FontSize,
        string? FontName,
        bool Bold,
        bool Italic,
        string TextColor,
        A.TextAlignmentTypeValues Alignment);

    private sealed record TableBorderSide(
        string Color,
        int Width);

    private sealed record TableCellBorderStyle(
        TableBorderSide Left,
        TableBorderSide Right,
        TableBorderSide Top,
        TableBorderSide Bottom);

    private sealed record ColumnGapCandidate(
        double Position,
        double Gap);
}