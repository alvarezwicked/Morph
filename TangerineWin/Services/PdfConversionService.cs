using System.IO;
using ImageMagick;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Writer;

namespace TangerineWin.Services;

public sealed class PdfConversionService
{
    public async Task<string> ConvertImageToPdfAsync(
        string sourcePath)
    {
        return await Task.Run(() =>
        {
            var outputPath =
                CreateUniqueOutputPath(sourcePath);

            using var image =
                new MagickImage(sourcePath);

            byte[] imageBytes;

            var extension =
                Path.GetExtension(sourcePath)
                    .ToLowerInvariant();

            var isJpeg =
                extension is ".jpg" or ".jpeg";

            if (isJpeg)
            {
                image.Format = MagickFormat.Jpeg;
                imageBytes = image.ToByteArray();
            }
            else
            {
                image.Format = MagickFormat.Png;
                imageBytes = image.ToByteArray();
            }

            var builder =
                new PdfDocumentBuilder();

            // Страница A4.
            var page =
                builder.AddPage(PageSize.A4);

            // Размер A4 в PDF points.
            const double pageWidth = 595.0;
            const double pageHeight = 842.0;

            // Небольшие поля.
            const double margin = 24.0;

            var availableWidth =
                pageWidth - margin * 2;

            var availableHeight =
                pageHeight - margin * 2;

            var imageWidth =
                (double)image.Width;

            var imageHeight =
                (double)image.Height;

            var scale = Math.Min(
                availableWidth / imageWidth,
                availableHeight / imageHeight);

            // Маленькие картинки не увеличиваем.
            if (scale > 1.0)
                scale = 1.0;

            var width =
                imageWidth * scale;

            var height =
                imageHeight * scale;

            // Центрируем.
            var x =
                (pageWidth - width) / 2.0;

            var y =
                (pageHeight - height) / 2.0;

            var rectangle =
                new PdfRectangle(
                    x,
                    y,
                    x + width,
                    y + height);

            if (isJpeg)
            {
                page.AddJpeg(
                    imageBytes,
                    rectangle);
            }
            else
            {
                page.AddPng(
                    imageBytes,
                    rectangle);
            }

            var pdfBytes =
                builder.Build();

            File.WriteAllBytes(
                outputPath,
                pdfBytes);

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
                name + ".pdf");

        if (!File.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name} ({index}).pdf");

            if (!File.Exists(candidate))
                return candidate;

            index++;
        }
    }
}