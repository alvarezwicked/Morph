using System.IO;
using ImageMagick;

namespace TangerineWin.Services;

public sealed class ImageConversionService
{
    public async Task<string> ConvertAsync(
        string sourcePath,
        string actionId)
    {
        return await Task.Run(() =>
        {
            using var image = new MagickImage(sourcePath);

            switch (actionId)
            {
                case "to_png":
                    {
                        image.Format = MagickFormat.Png;

                        var outputPath =
                            CreateUniqueOutputPath(
                                sourcePath,
                                ".png");

                        image.Write(outputPath);

                        return outputPath;
                    }

                case "to_jpg":
                    {
                        image.Format = MagickFormat.Jpeg;
                        image.Quality = 92;

                        var outputPath =
                            CreateUniqueOutputPath(
                                sourcePath,
                                ".jpg");

                        image.Write(outputPath);

                        return outputPath;
                    }

                case "to_webp":
                    {
                        image.Format = MagickFormat.WebP;
                        image.Quality = 85;

                        var outputPath =
                            CreateUniqueOutputPath(
                                sourcePath,
                                ".webp");

                        image.Write(outputPath);

                        return outputPath;
                    }

                default:
                    throw new NotSupportedException(
                        $"Unsupported image action: {actionId}");
            }
        });
    }

    private static string CreateUniqueOutputPath(
        string sourcePath,
        string targetExtension)
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
                name + targetExtension);

        if (!File.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name} ({index}){targetExtension}");

            if (!File.Exists(candidate))
                return candidate;

            index++;
        }
    }
}