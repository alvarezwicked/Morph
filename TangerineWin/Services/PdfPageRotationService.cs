using System.IO;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Writer;

namespace TangerineWin.Services;

public sealed class PdfPageRotationService
{
    public async Task<string> RotateAsync(
        string sourcePath,
        IReadOnlyList<int> pagesToRotate,
        int rotationDegrees)
    {
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException(
                "PDF file was not found.",
                sourcePath);
        }

        if (rotationDegrees is not (90 or 180 or 270))
        {
            throw new ArgumentException(
                "Rotation must be 90, 180 or 270 degrees.",
                nameof(rotationDegrees));
        }

        return await Task.Run(() =>
        {
            using var document =
                PdfDocument.Open(sourcePath);

            var rotateSet =
                pagesToRotate
                    .Where(page =>
                        page >= 1 &&
                        page <= document.NumberOfPages)
                    .ToHashSet();

            if (rotateSet.Count == 0)
            {
                throw new InvalidOperationException(
                    "Не указаны корректные страницы для поворота.");
            }

            using var builder =
                new PdfDocumentBuilder();

            for (var pageNumber = 1;
                 pageNumber <= document.NumberOfPages;
                 pageNumber++)
            {
                var originalPage =
                    document.GetPage(pageNumber);

                var newPage =
                    builder.AddPage(
                        document,
                        pageNumber);

                if (!rotateSet.Contains(pageNumber))
                    continue;

                var currentRotation =
                    originalPage.Rotation.Value;

                var newRotation =
                    (currentRotation + rotationDegrees) % 360;

                newPage.SetRotation(
                    new PageRotationDegrees(
                        newRotation));
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
                $"{name}-rotated.pdf");

        if (!File.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name}-rotated ({index}).pdf");

            if (!File.Exists(candidate))
                return candidate;

            index++;
        }
    }
}