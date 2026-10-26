using Docnet.Core;
using Docnet.Core.Models;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Tesseract;

namespace TangerineWin.Services;

public sealed class PdfOcrService
{
    public sealed record OcrWord(
        string Text,
        int Left,
        int Top,
        int Width,
        int Height,
        float Confidence,
        int BlockNumber,
        int ParagraphNumber,
        int LineNumber)
    {
        public int Right => Left + Width;
        public int Bottom => Top + Height;
    }

    public sealed record OcrResult(
        string Text,
        float Confidence,
        IReadOnlyList<OcrWord> Words);

    public async Task<OcrResult> RecognizePageAsync(
        string pdfPath,
        int pageIndex)
    {
        var pngBytes =
            await RenderPageToPngAsync(
                pdfPath,
                pageIndex);

        return await RecognizePngAsync(
            pngBytes);
    }

    public async Task<OcrResult> RecognizePngAsync(
        byte[] pngBytes)
    {
        return await Task.Run(() =>
        {
            var tessdataPath =
                Path.Combine(
                    AppContext.BaseDirectory,
                    "tessdata");

            if (!Directory.Exists(tessdataPath))
            {
                throw new DirectoryNotFoundException(
                    $"Tessdata directory not found: {tessdataPath}");
            }

            using var engine =
                new TesseractEngine(
                    tessdataPath,
                    "rus+eng",
                    EngineMode.LstmOnly);

            using var pix =
                Pix.LoadFromMemory(
                    pngBytes);

            using var page =
                engine.Process(
                    pix,
                    PageSegMode.Auto);

            var text =
                page.GetText()
                    ?.Trim()
                ?? string.Empty;

            var confidence =
                page.GetMeanConfidence();

            var tsv =
                page.GetTsvText(0);

            var words =
                ParseTsv(tsv);

            return new OcrResult(
                text,
                confidence,
                words);
        });
    }

    private static IReadOnlyList<OcrWord> ParseTsv(
        string? tsv)
    {
        if (string.IsNullOrWhiteSpace(tsv))
            return Array.Empty<OcrWord>();

        var result =
            new List<OcrWord>();

        var lines =
            tsv.Split(
                new[]
                {
                    "\r\n",
                    "\n"
                },
                StringSplitOptions.RemoveEmptyEntries);

        // Первая строка TSV — заголовок.
        foreach (var line in lines.Skip(1))
        {
            var columns =
                line.Split('\t');

            if (columns.Length < 12)
                continue;

            if (!int.TryParse(columns[0], out var level))
                continue;

            // Level 5 = word.
            if (level != 5)
                continue;

            if (!int.TryParse(
                    columns[2],
                    out var blockNumber))
                continue;

            if (!int.TryParse(
                    columns[3],
                    out var paragraphNumber))
                continue;

            if (!int.TryParse(
                    columns[4],
                    out var lineNumber))
                continue;

            if (!int.TryParse(
                    columns[6],
                    out var left))
                continue;

            if (!int.TryParse(
                    columns[7],
                    out var top))
                continue;

            if (!int.TryParse(
                    columns[8],
                    out var width))
                continue;

            if (!int.TryParse(
                    columns[9],
                    out var height))
                continue;

            if (!float.TryParse(
                    columns[10],
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out var wordConfidence))
            {
                wordConfidence = 0;
            }

            var wordText =
                string.Join(
                    "\t",
                    columns.Skip(11))
                .Trim();

            if (string.IsNullOrWhiteSpace(wordText))
                continue;

            result.Add(
                new OcrWord(
                    wordText,
                    left,
                    top,
                    width,
                    height,
                    wordConfidence,
                    blockNumber,
                    paragraphNumber,
                    lineNumber));
        }

        return result;
    }

    public async Task<byte[]> RenderPageToPngAsync(
        string pdfPath,
        int pageIndex)
    {
        return await Task.Run(() =>
        {
            using var docReader =
                DocLib.Instance.GetDocReader(
                    pdfPath,
                    new PageDimensions(
                        2200,
                        3000));

            using var pageReader =
                docReader.GetPageReader(
                    pageIndex);

            var rawBytes =
                pageReader.GetImage();

            var width =
                pageReader.GetPageWidth();

            var height =
                pageReader.GetPageHeight();

            CompositeOnWhite(
                rawBytes);

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

            using var stream =
                new MemoryStream();

            bitmap.Save(
                stream,
                System.Drawing.Imaging.ImageFormat.Png);

            return stream.ToArray();
        });
    }

    private static void CompositeOnWhite(
        byte[] pixels)
    {
        for (var i = 0;
             i + 3 < pixels.Length;
             i += 4)
        {
            var alpha =
                pixels[i + 3];

            var background =
                255 - alpha;

            pixels[i] =
                (byte)Math.Min(
                    255,
                    pixels[i] +
                    background);

            pixels[i + 1] =
                (byte)Math.Min(
                    255,
                    pixels[i + 1] +
                    background);

            pixels[i + 2] =
                (byte)Math.Min(
                    255,
                    pixels[i + 2] +
                    background);

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

            if (sourceStride ==
                destinationStride)
            {
                Marshal.Copy(
                    rawBytes,
                    0,
                    bitmapData.Scan0,
                    rawBytes.Length);

                return;
            }

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
        finally
        {
            bitmap.UnlockBits(
                bitmapData);
        }
    }
}