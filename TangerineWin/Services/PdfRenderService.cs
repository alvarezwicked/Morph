using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using Docnet.Core;
using Docnet.Core.Models;

namespace TangerineWin.Services;

public sealed class PdfRenderService
{
    public async Task<string> RenderToImagesAsync(
        string sourcePath,
        string actionId)
    {
        return await Task.Run(() =>
        {
            var extension = actionId switch
            {
                "pdf_to_png" => ".png",
                "pdf_to_jpg" => ".jpg",

                _ => throw new NotSupportedException(
                    $"Unsupported PDF render action: {actionId}")
            };

            var outputDirectory =
                CreateUniqueOutputDirectory(sourcePath);

            Directory.CreateDirectory(outputDirectory);

            using var docReader =
                DocLib.Instance.GetDocReader(
                    sourcePath,
                    new PageDimensions(1600, 2200));

            var pageCount =
                docReader.GetPageCount();

            for (var pageIndex = 0;
                 pageIndex < pageCount;
                 pageIndex++)
            {
                using var pageReader =
                    docReader.GetPageReader(pageIndex);

                var rawBytes =
                    pageReader.GetImage();

                var width =
                    pageReader.GetPageWidth();

                var height =
                    pageReader.GetPageHeight();

                // Docnet отдаёт BGRA.
                // Принудительно делаем каждый пиксель
                // полностью непрозрачным.
                CompositeOnWhite(rawBytes);

                using var bitmap =
                    new Bitmap(
                        width,
                        height,
                        PixelFormat.Format32bppArgb);

                CopyPixels(
                    bitmap,
                    rawBytes,
                    width,
                    height);

                var outputPath =
                    Path.Combine(
                        outputDirectory,
                        $"page-{pageIndex + 1}{extension}");

                if (extension == ".png")
                {
                    bitmap.Save(
                        outputPath,
                        ImageFormat.Png);
                }
                else
                {
                    SaveJpeg(
                        bitmap,
                        outputPath);
                }
            }

            return outputDirectory;
        });
    }

    private static void CompositeOnWhite(
        byte[] pixels)
    {
        // Docnet/PDFium возвращает BGRA.
        //
        // Прозрачные области страницы обычно:
        // B=0, G=0, R=0, A=0
        //
        // Нельзя просто поставить A=255 —
        // тогда прозрачный фон станет черным.
        //
        // Здесь смешиваем каждый пиксель
        // с белым фоном.

        for (var i = 0;
             i + 3 < pixels.Length;
             i += 4)
        {
            var alpha = pixels[i + 3];

            // PDFium использует premultiplied BGRA,
            // поэтому для белого фона:
            //
            // result = channel + (255 - alpha)

            var background = 255 - alpha;

            pixels[i] =
                (byte)Math.Min(
                    255,
                    pixels[i] + background);

            pixels[i + 1] =
                (byte)Math.Min(
                    255,
                    pixels[i + 1] + background);

            pixels[i + 2] =
                (byte)Math.Min(
                    255,
                    pixels[i + 2] + background);

            pixels[i + 3] = 255;
        }
    }

    private static void CopyPixels(
        Bitmap bitmap,
        byte[] rawBytes,
        int width,
        int height)
    {
        var rect =
            new Rectangle(
                0,
                0,
                width,
                height);

        var bitmapData =
            bitmap.LockBits(
                rect,
                ImageLockMode.WriteOnly,
                PixelFormat.Format32bppArgb);

        try
        {
            var sourceStride =
                width * 4;

            var destinationStride =
                bitmapData.Stride;

            // Обычно stride совпадает, но лучше
            // обрабатывать корректно в любом случае.
            if (sourceStride == destinationStride)
            {
                Marshal.Copy(
                    rawBytes,
                    0,
                    bitmapData.Scan0,
                    rawBytes.Length);
            }
            else
            {
                for (var y = 0;
                     y < height;
                     y++)
                {
                    var sourceOffset =
                        y * sourceStride;

                    var destinationPointer =
                        IntPtr.Add(
                            bitmapData.Scan0,
                            y * destinationStride);

                    Marshal.Copy(
                        rawBytes,
                        sourceOffset,
                        destinationPointer,
                        sourceStride);
                }
            }
        }
        finally
        {
            bitmap.UnlockBits(bitmapData);
        }
    }

    private static void SaveJpeg(
        Bitmap bitmap,
        string outputPath)
    {
        var jpegCodec =
            ImageCodecInfo
                .GetImageEncoders()
                .First(codec =>
                    codec.FormatID ==
                    ImageFormat.Jpeg.Guid);

        using var parameters =
            new EncoderParameters(1);

        parameters.Param[0] =
            new EncoderParameter(
                Encoder.Quality,
                92L);

        bitmap.Save(
            outputPath,
            jpegCodec,
            parameters);
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
                $"{name}-pages");

        if (!Directory.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name}-pages ({index})");

            if (!Directory.Exists(candidate))
                return candidate;

            index++;
        }
    }
}