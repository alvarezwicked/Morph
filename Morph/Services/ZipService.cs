using System.IO;
using System.IO.Compression;

namespace Morph.Services;

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
                .Distinct(
                    StringComparer.OrdinalIgnoreCase)
                .ToArray();

            if (existingFiles.Length == 0)
                throw new InvalidOperationException(
                    "No valid files were found.");

            var outputPath =
                CreateOutputPath(existingFiles);

            try
            {
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
            }
            catch
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);

                throw;
            }

            return outputPath;
        });
    }

    /// <summary>
    /// Распаковывает ZIP-архивы в отдельные папки.
    /// </summary>
    public Task<IReadOnlyList<string>> ExtractAsync(
        IReadOnlyList<string> zipFiles)
    {
        if (zipFiles.Count == 0)
            throw new ArgumentException(
                "No ZIP files provided.",
                nameof(zipFiles));

        return Task.Run<IReadOnlyList<string>>(() =>
        {
            var sources = zipFiles
                .Distinct(
                    StringComparer.OrdinalIgnoreCase)
                .ToArray();

            foreach (var file in sources)
            {
                if (!File.Exists(file) ||
                    !Path.GetExtension(file).Equals(
                        ".zip",
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        $"Not a ZIP archive: {file}");
                }
            }

            var results =
                new List<string>();

            foreach (var source in sources)
            {
                results.Add(
                    ExtractOne(source));
            }

            return results;
        });
    }

    private static string ExtractOne(
        string source)
    {
        var parent =
            Path.GetDirectoryName(source)
            ?? throw new InvalidOperationException(
                "Unable to determine archive directory.");

        var baseName =
            Path.GetFileNameWithoutExtension(source);

        if (string.IsNullOrWhiteSpace(baseName))
            baseName = "Extracted";

        // Используем временную папку,
        // чтобы не оставлять неполную распаковку.
        var staging = Path.Combine(
            parent,
            ".morph-extract-" +
            Guid.NewGuid().ToString("N"));

        Directory.CreateDirectory(staging);

        try
        {
            using (var archive =
                   ZipFile.OpenRead(source))
            {
                var stagingRoot =
                    Path.GetFullPath(staging)
                        .TrimEnd(
                            Path.DirectorySeparatorChar)
                    + Path.DirectorySeparatorChar;

                foreach (var entry in archive.Entries)
                {
                    // Не разрешаем символические ссылки.
                    var unixType =
                        (entry.ExternalAttributes >> 16)
                        & 0xF000;

                    if (unixType == 0xA000)
                    {
                        throw new InvalidDataException(
                            "ZIP contains a symbolic link.");
                    }

                    var name = entry.FullName.Replace(
                        '/',
                        Path.DirectorySeparatorChar);

                    if (Path.IsPathRooted(name) ||
                        name.Contains(':'))
                    {
                        throw new InvalidDataException(
                            "ZIP contains an unsafe path.");
                    }

                    var target = Path.GetFullPath(
                        Path.Combine(staging, name));

                    if (!target.StartsWith(
                            stagingRoot,
                            StringComparison.OrdinalIgnoreCase)
                        &&
                        !string.Equals(
                            target,
                            stagingRoot.TrimEnd(
                                Path.DirectorySeparatorChar),
                            StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidDataException(
                            "ZIP contains a path outside " +
                            "the extraction folder.");
                    }

                    var isDirectory =
                        entry.FullName.EndsWith('/')
                        ||
                        entry.FullName.EndsWith('\\');

                    if (isDirectory)
                    {
                        Directory.CreateDirectory(target);
                        continue;
                    }

                    var directory =
                        Path.GetDirectoryName(target);

                    if (directory is not null)
                        Directory.CreateDirectory(directory);

                    // Не перезаписываем файлы.
                    using var input =
                        entry.Open();

                    using var output =
                        new FileStream(
                            target,
                            FileMode.CreateNew,
                            FileAccess.Write);

                    input.CopyTo(output);
                }
            }

            // Папка с именем архива.
            // При совпадении используем (1), (2) и т.д.
            for (var index = 0; ; index++)
            {
                var folderName =
                    index == 0
                        ? baseName
                        : $"{baseName} ({index})";

                var destination =
                    Path.Combine(
                        parent,
                        folderName);

                if (Directory.Exists(destination) ||
                    File.Exists(destination))
                {
                    continue;
                }

                try
                {
                    Directory.Move(
                        staging,
                        destination);

                    return destination;
                }
                catch (IOException) when (
                    Directory.Exists(destination) ||
                    File.Exists(destination))
                {
                    // Папка успела появиться.
                    // Пробуем следующее имя.
                }
            }
        }
        finally
        {
            if (Directory.Exists(staging))
            {
                Directory.Delete(
                    staging,
                    recursive: true);
            }
        }
    }

    private static string CreateOutputPath(
        IReadOnlyList<string> files)
    {
        var firstFile =
            files[0];

        var directory =
            Path.GetDirectoryName(firstFile)
            ?? throw new InvalidOperationException(
                "Unable to determine output directory.");

        var baseName =
            files.Count == 1
                ? Path.GetFileNameWithoutExtension(firstFile)
                : "Archive";

        var candidate = Path.Combine(
            directory,
            baseName + ".zip");

        if (!File.Exists(candidate) &&
            !Directory.Exists(candidate))
        {
            return candidate;
        }

        var index = 1;

        while (true)
        {
            candidate = Path.Combine(
                directory,
                $"{baseName} ({index}).zip");

            if (!File.Exists(candidate) &&
                !Directory.Exists(candidate))
            {
                return candidate;
            }

            index++;
        }
    }
}