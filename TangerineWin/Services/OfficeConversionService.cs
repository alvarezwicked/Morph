using System.Diagnostics;
using System.IO;

namespace TangerineWin.Services;

public sealed class OfficeConversionService
{
    private static readonly string[] LibreOfficeCandidates =
    {
        @"C:\Program Files\LibreOffice\program\soffice.exe",
        @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    };

    public async Task<string> ConvertToPdfAsync(
        string sourcePath)
    {
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException(
                "Source file was not found.",
                sourcePath);

        var sofficePath = FindLibreOffice();

        if (sofficePath is null)
        {
            throw new InvalidOperationException(
                "LibreOffice не найден. Установите LibreOffice для конвертации Office-файлов в PDF.");
        }

        var sourceDirectory =
            Path.GetDirectoryName(sourcePath)
            ?? throw new InvalidOperationException(
                "Unable to determine source directory.");

        var sourceName =
            Path.GetFileNameWithoutExtension(sourcePath);

        // LibreOffice сам задаёт имя выходного файла,
        // поэтому используем временную папку.
        var tempDirectory =
            Path.Combine(
                Path.GetTempPath(),
                "TangerineWin",
                Guid.NewGuid().ToString("N"));

        Directory.CreateDirectory(tempDirectory);

        try
        {
            var startInfo =
                new ProcessStartInfo
                {
                    FileName = sofficePath,

                    UseShellExecute = false,
                    CreateNoWindow = true,

                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

            startInfo.ArgumentList.Add("--headless");
            startInfo.ArgumentList.Add("--convert-to");
            startInfo.ArgumentList.Add("pdf");
            startInfo.ArgumentList.Add("--outdir");
            startInfo.ArgumentList.Add(tempDirectory);
            startInfo.ArgumentList.Add(sourcePath);

            using var process =
                new Process
                {
                    StartInfo = startInfo
                };

            process.Start();

            var stdoutTask =
                process.StandardOutput.ReadToEndAsync();

            var stderrTask =
                process.StandardError.ReadToEndAsync();

            await process.WaitForExitAsync();

            var stdout = await stdoutTask;
            var stderr = await stderrTask;

            System.Diagnostics.Debug.WriteLine(
                $"LibreOffice stdout: {stdout}");

            System.Diagnostics.Debug.WriteLine(
                $"LibreOffice stderr: {stderr}");

            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    $"LibreOffice conversion failed. Exit code: {process.ExitCode}. {stderr}");
            }

            var generatedPdf =
                Path.Combine(
                    tempDirectory,
                    sourceName + ".pdf");

            if (!File.Exists(generatedPdf))
            {
                throw new InvalidOperationException(
                    "LibreOffice завершил работу, но PDF-файл не был создан.");
            }

            var finalPath =
                CreateUniqueOutputPath(
                    sourcePath);

            File.Move(
                generatedPdf,
                finalPath);

            return finalPath;
        }
        finally
        {
            try
            {
                if (Directory.Exists(tempDirectory))
                {
                    Directory.Delete(
                        tempDirectory,
                        recursive: true);
                }
            }
            catch
            {
                // Очистка временной папки не должна
                // ломать успешную конвертацию.
            }
        }
    }

    private static string? FindLibreOffice()
    {
        foreach (var path in LibreOfficeCandidates)
        {
            if (File.Exists(path))
                return path;
        }

        return null;
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