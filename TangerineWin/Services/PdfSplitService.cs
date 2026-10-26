using System.IO;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Writer;

namespace TangerineWin.Services;

public sealed class PdfSplitService
{
    public async Task<string> SplitAsync(
        string sourcePath)
    {
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException(
                "PDF file was not found.",
                sourcePath);
        }

        return await Task.Run(() =>
        {
            var outputDirectory =
                CreateUniqueOutputDirectory(
                    sourcePath);

            Directory.CreateDirectory(
                outputDirectory);

            using var document =
                PdfDocument.Open(sourcePath);

            for (var pageNumber = 1;
                 pageNumber <= document.NumberOfPages;
                 pageNumber++)
            {
                using var builder =
                    new PdfDocumentBuilder();

                builder.AddPage(
                    document,
                    pageNumber);

                var bytes =
                    builder.Build();

                var outputPath =
                    Path.Combine(
                        outputDirectory,
                        $"page-{pageNumber}.pdf");

                File.WriteAllBytes(
                    outputPath,
                    bytes);
            }

            return outputDirectory;
        });
    }

    private static string CreateUniqueOutputDirectory(
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
                $"{name}-split");

        if (!Directory.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name}-split ({index})");

            if (!Directory.Exists(candidate))
                return candidate;

            index++;
        }
    }
}