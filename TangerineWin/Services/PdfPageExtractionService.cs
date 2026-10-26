using System.IO;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Writer;

namespace TangerineWin.Services;

public sealed class PdfPageExtractionService
{
    public async Task<string> ExtractAsync(
        string sourcePath,
        IReadOnlyList<int> pageNumbers)
    {
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("PDF file was not found.", sourcePath);

        if (pageNumbers.Count == 0)
            throw new InvalidOperationException("Не выбраны страницы.");

        return await Task.Run(() =>
        {
            using var document = PdfDocument.Open(sourcePath);

            var validPages = pageNumbers
                .Distinct()
                .Where(page => page >= 1 && page <= document.NumberOfPages)
                .ToArray();

            if (validPages.Length == 0)
                throw new InvalidOperationException("Нет корректных страниц для извлечения.");

            using var builder = new PdfDocumentBuilder();

            foreach (var pageNumber in validPages)
            {
                builder.AddPage(document, pageNumber);
            }

            var outputPath = CreateUniqueOutputPath(sourcePath);

            File.WriteAllBytes(
                outputPath,
                builder.Build());

            return outputPath;
        });
    }

    public IReadOnlyList<int> ParsePageRange(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return Array.Empty<int>();

        var result = new List<int>();

        var parts = value.Split(
            ',',
            StringSplitOptions.RemoveEmptyEntries |
            StringSplitOptions.TrimEntries);

        foreach (var part in parts)
        {
            if (part.Contains('-'))
            {
                var rangeParts = part.Split(
                    '-',
                    StringSplitOptions.RemoveEmptyEntries |
                    StringSplitOptions.TrimEntries);

                if (rangeParts.Length != 2)
                    throw new FormatException($"Некорректный диапазон: {part}");

                if (!int.TryParse(rangeParts[0], out var start) ||
                    !int.TryParse(rangeParts[1], out var end))
                {
                    throw new FormatException($"Некорректный диапазон: {part}");
                }

                if (start > end)
                    throw new FormatException($"Начало диапазона больше конца: {part}");

                for (var page = start; page <= end; page++)
                {
                    result.Add(page);
                }
            }
            else
            {
                if (!int.TryParse(part, out var page))
                    throw new FormatException($"Некорректный номер страницы: {part}");

                result.Add(page);
            }
        }

        return result;
    }

    private static string CreateUniqueOutputPath(
        string sourcePath)
    {
        var directory =
            Path.GetDirectoryName(sourcePath)
            ?? throw new InvalidOperationException(
                "Unable to determine source directory.");

        var name =
            Path.GetFileNameWithoutExtension(sourcePath);

        var candidate =
            Path.Combine(
                directory,
                $"{name}-pages.pdf");

        if (!File.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name}-pages ({index}).pdf");

            if (!File.Exists(candidate))
                return candidate;

            index++;
        }
    }
}