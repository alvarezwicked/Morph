using System.IO;
using UglyToad.PdfPig;

namespace TangerineWin.Services;

public sealed class PdfImageExtractionService
{
    public async Task<string> ExtractAsync(
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
                CreateUniqueOutputDirectory(sourcePath);

            Directory.CreateDirectory(
                outputDirectory);

            using var document =
                PdfDocument.Open(sourcePath);

            var extractedCount = 0;

            foreach (var page in document.GetPages())
            {
                var imageIndex = 1;

                foreach (var image in page.GetImages())
                {
                    extractedCount++;

                    var baseName =
                        $"page-{page.Number}-image-{imageIndex}";

                    if (image.TryGetPng(out var pngBytes)
                        && pngBytes is not null
                        && pngBytes.Length > 0)
                    {
                        var outputPath =
                            Path.Combine(
                                outputDirectory,
                                baseName + ".png");

                        File.WriteAllBytes(
                            outputPath,
                            pngBytes);
                    }
                    else
                    {
                        var rawBytes =
                            image.RawBytes.ToArray();

                        if (IsJpeg(rawBytes))
                        {
                            var outputPath =
                                Path.Combine(
                                    outputDirectory,
                                    baseName + ".jpg");

                            File.WriteAllBytes(
                                outputPath,
                                rawBytes);
                        }
                        else
                        {
                            // Некоторые PDF используют внутренние
                            // форматы/цветовые пространства, которые
                            // PdfPig не умеет экспортировать напрямую.
                            //
                            // Такие изображения пока пропускаем.
                            extractedCount--;
                        }
                    }

                    imageIndex++;
                }
            }

            if (extractedCount == 0)
            {
                Directory.Delete(
                    outputDirectory,
                    recursive: true);

                throw new InvalidOperationException(
                    "В PDF не найдено изображений, которые можно извлечь.");
            }

            return outputDirectory;
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
                $"{name}-images");

        if (!Directory.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name}-images ({index})");

            if (!Directory.Exists(candidate))
                return candidate;

            index++;
        }
    }
}