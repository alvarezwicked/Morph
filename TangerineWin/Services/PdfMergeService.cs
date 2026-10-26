using System.IO;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Writer;

namespace TangerineWin.Services;

public sealed class PdfMergeService
{
    public async Task<string> MergeAsync(
        IReadOnlyList<string> files)
    {
        if (files.Count < 2)
        {
            throw new InvalidOperationException(
                "Для объединения нужно выбрать минимум два PDF-файла.");
        }

        return await Task.Run(() =>
        {
            var pdfFiles = files
                .Where(File.Exists)
                .Where(path =>
                    Path.GetExtension(path)
                        .Equals(
                            ".pdf",
                            StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (pdfFiles.Length < 2)
            {
                throw new InvalidOperationException(
                    "Для объединения нужно минимум два PDF-файла.");
            }

            var outputPath =
                CreateUniqueOutputPath(pdfFiles[0]);

            using var builder =
                new PdfDocumentBuilder();

            foreach (var pdfPath in pdfFiles)
            {
                using var document =
                    PdfDocument.Open(pdfPath);

                for (var pageNumber = 1;
                     pageNumber <= document.NumberOfPages;
                     pageNumber++)
                {
                    builder.AddPage(
                        document,
                        pageNumber);
                }
            }

            var bytes =
                builder.Build();

            File.WriteAllBytes(
                outputPath,
                bytes);

            return outputPath;
        });
    }

    private static string CreateUniqueOutputPath(
        string firstPdfPath)
    {
        var directory =
            Path.GetDirectoryName(firstPdfPath)
            ?? throw new InvalidOperationException(
                "Unable to determine output directory.");

        var candidate =
            Path.Combine(
                directory,
                "merged.pdf");

        if (!File.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"merged ({index}).pdf");

            if (!File.Exists(candidate))
                return candidate;

            index++;
        }
    }
}