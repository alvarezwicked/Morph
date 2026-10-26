using System.IO;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Writer;

namespace TangerineWin.Services;

public sealed class PdfPageReorderService
{
    public async Task<string> ReorderAsync(
        string sourcePath,
        IReadOnlyList<int> pageOrder)
    {
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException(
                "PDF file was not found.",
                sourcePath);
        }

        return await Task.Run(() =>
        {
            using var document =
                PdfDocument.Open(sourcePath);

            var pageCount =
                document.NumberOfPages;

            if (pageOrder.Count != pageCount)
            {
                throw new InvalidOperationException(
                    $"Нужно указать все {pageCount} страниц.");
            }

            if (pageOrder.Any(
                page =>
                    page < 1 ||
                    page > pageCount))
            {
                throw new InvalidOperationException(
                    $"Номера страниц должны быть от 1 до {pageCount}.");
            }

            if (pageOrder.Distinct().Count() != pageCount)
            {
                throw new InvalidOperationException(
                    "Каждая страница должна быть указана ровно один раз.");
            }

            using var builder =
                new PdfDocumentBuilder();

            foreach (var pageNumber in pageOrder)
            {
                builder.AddPage(
                    document,
                    pageNumber);
            }

            var outputPath =
                CreateUniqueOutputPath(
                    sourcePath);

            File.WriteAllBytes(
                outputPath,
                builder.Build());

            return outputPath;
        });
    }

    private static string CreateUniqueOutputPath(
        string sourcePath)
    {
        var directory =
            Path.GetDirectoryName(sourcePath)
            ?? throw new InvalidOperationException(
                "Unable to determine source directory.");

        var name =
            Path.GetFileNameWithoutExtension(
                sourcePath);

        var candidate =
            Path.Combine(
                directory,
                $"{name}-reordered.pdf");

        if (!File.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name}-reordered ({index}).pdf");

            if (!File.Exists(candidate))
                return candidate;

            index++;
        }
    }
}