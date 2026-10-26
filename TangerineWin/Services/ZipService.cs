using System.IO;
using System.IO.Compression;

namespace TangerineWin.Services;

public sealed class ZipService
{
    public async Task<string> CreateAsync(
        IReadOnlyList<string> files)
    {
        if (files.Count == 0)
            throw new ArgumentException(
                "No files provided.",
                nameof(files));

        return await Task.Run(() =>
        {
            var existingFiles = files
                .Where(File.Exists)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            if (existingFiles.Length == 0)
                throw new InvalidOperationException(
                    "No valid files were found.");

            var outputPath =
                CreateOutputPath(existingFiles);

            using var archive = ZipFile.Open(
                outputPath,
                ZipArchiveMode.Create);

            foreach (var file in existingFiles)
            {
                archive.CreateEntryFromFile(
                    file,
                    Path.GetFileName(file),
                    CompressionLevel.Optimal);
            }

            return outputPath;
        });
    }

    private static string CreateOutputPath(
        IReadOnlyList<string> files)
    {
        var firstFile = files[0];

        var directory =
            Path.GetDirectoryName(firstFile)
            ?? throw new InvalidOperationException(
                "Unable to determine output directory.");

        string baseName;

        if (files.Count == 1)
        {
            baseName =
                Path.GetFileNameWithoutExtension(firstFile);
        }
        else
        {
            baseName = "Archive";
        }

        var candidate =
            Path.Combine(
                directory,
                baseName + ".zip");

        if (!File.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{baseName} ({index}).zip");

            if (!File.Exists(candidate))
                return candidate;

            index++;
        }
    }
}