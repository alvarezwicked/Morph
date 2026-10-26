using System.IO;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Writer;

namespace TangerineWin.Services;

public sealed class PdfPageDeletionService
{
    public async Task<string> DeleteAsync(
        string sourcePath,
        IReadOnlyList<int> pagesToDelete)
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

            var deleteSet =
                pagesToDelete
                    .Where(page =>
                        page >= 1 &&
                        page <= document.NumberOfPages)
                    .ToHashSet();

            if (deleteSet.Count == 0)
            {
                throw new InvalidOperationException(
                    "Не указаны корректные страницы для удаления.");
            }

            if (deleteSet.Count >= document.NumberOfPages)
            {
                throw new InvalidOperationException(
                    "Нельзя удалить все страницы PDF.");
            }

            using var builder =
                new PdfDocumentBuilder();

            for (var pageNumber = 1;
                 pageNumber <= document.NumberOfPages;
                 pageNumber++)
            {
                if (deleteSet.Contains(pageNumber))
                    continue;

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
                $"{name}-deleted-pages.pdf");

        if (!File.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name}-deleted-pages ({index}).pdf");

            if (!File.Exists(candidate))
                return candidate;

            index++;
        }
    }
}