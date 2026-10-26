using UglyToad.PdfPig;
using UglyToad.PdfPig.Parser;
using Tabula;
using Tabula.Extractors;

namespace TangerineWin.Services;

public sealed class PdfTableExtractionService
{
    public sealed record ExtractedTable(
        int PageNumber,
        double Left,
        double Bottom,
        double Right,
        double Top,
        IReadOnlyList<IReadOnlyList<string>> Rows);

    public async Task<IReadOnlyList<ExtractedTable>> ExtractAsync(
        string sourcePath)
    {
        return await Task.Run(() =>
        {
            var result =
                new List<ExtractedTable>();

            using var document =
                PdfDocument.Open(
                    sourcePath,
                    new ParsingOptions
                    {
                        ClipPaths = true
                    });

            for (var pageNumber = 1;
                 pageNumber <= document.NumberOfPages;
                 pageNumber++)
            {
                var pageArea =
                    ObjectExtractor.Extract(
                        document,
                        pageNumber);

                var tables =
                    ExtractTables(
                        pageArea);

                foreach (var table in tables)
                {
                    var rows =
                        new List<IReadOnlyList<string>>();

                    foreach (var row in table.Rows)
                    {
                        var cells =
                            row
                                .Select(cell =>
                                    NormalizeCellText(
                                        cell.GetText()))
                                .ToArray();

                        if (cells.All(
                            string.IsNullOrWhiteSpace))
                        {
                            continue;
                        }

                        rows.Add(
                            cells);
                    }

                    if (!IsUsefulTable(rows))
                        continue;

                    var cellsWithBounds =
                        table.Rows
                            .SelectMany(row => row)
                            .Where(cell =>
                                cell.BoundingBox.Width > 0 ||
                                cell.BoundingBox.Height > 0)
                            .ToArray();

                    if (cellsWithBounds.Length == 0)
                        continue;

                    var left =
                        cellsWithBounds
                            .Min(cell =>
                                cell.BoundingBox.Left);

                    var bottom =
                        cellsWithBounds
                            .Min(cell =>
                                cell.BoundingBox.Bottom);

                    var right =
                        cellsWithBounds
                            .Max(cell =>
                                cell.BoundingBox.Right);

                    var top =
                        cellsWithBounds
                            .Max(cell =>
                                cell.BoundingBox.Top);

                    result.Add(
                        new ExtractedTable(
                            pageNumber,
                            left,
                            bottom,
                            right,
                            top,
                            rows));
                }
            }

            return result;
        });
    }

    private static IReadOnlyList<Table> ExtractTables(
        PageArea page)
    {
        var spreadsheet =
            new SpreadsheetExtractionAlgorithm();

        var spreadsheetTables =
            spreadsheet
                .Extract(page)
                .Where(IsUsefulTable)
                .ToArray();

        if (spreadsheetTables.Length > 0)
        {
            return spreadsheetTables;
        }

        var basic =
            new BasicExtractionAlgorithm();

        return basic
            .Extract(page)
            .Where(IsUsefulTable)
            .ToArray();
    }

    private static bool IsUsefulTable(
        Table table)
    {
        if (table.Rows.Count < 2)
            return false;

        var maximumColumnCount =
            table.Rows
                .Select(row =>
                    row.Count)
                .DefaultIfEmpty(0)
                .Max();

        return maximumColumnCount >= 2;
    }

    private static bool IsUsefulTable(
        IReadOnlyList<IReadOnlyList<string>> rows)
    {
        if (rows.Count < 2)
            return false;

        var maximumColumnCount =
            rows
                .Select(row =>
                    row.Count)
                .DefaultIfEmpty(0)
                .Max();

        return maximumColumnCount >= 2;
    }

    private static string NormalizeCellText(
        string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        return string.Join(
            " ",
            value
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Split(
                    ' ',
                    StringSplitOptions.RemoveEmptyEntries |
                    StringSplitOptions.TrimEntries));
    }
}