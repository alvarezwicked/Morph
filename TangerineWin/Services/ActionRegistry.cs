using System.IO;
using TangerineWin.Models;

namespace TangerineWin.Services;

public sealed class ActionRegistry
{
    public IReadOnlyList<FileAction> GetActions(
        IReadOnlyList<string> files)
    {
        if (files.Count == 0)
            return Array.Empty<FileAction>();

        // Пока для нескольких файлов предлагаем только ZIP.
        if (files.Count > 1)
        {
            var allPdf = files.All(path =>
                Path.GetExtension(path)
                    .Equals(
                        ".pdf",
                        StringComparison.OrdinalIgnoreCase));

            if (allPdf)
            {
                return new[]
                {
            new FileAction("merge_pdf", "MERGE"),
            new FileAction("zip", "ZIP")
        };
            }

            return new[]
            {
        new FileAction("zip", "ZIP")
    };
        }

        var extension = Path
            .GetExtension(files[0])
            .ToLowerInvariant();

        return extension switch
        {
            ".jpg" or ".jpeg" => new[]
            {
                new FileAction("to_png", "PNG"),
                new FileAction("to_webp", "WEBP"),
                new FileAction("to_pdf", "PDF"),
                new FileAction("edit_image", "EDIT"),
                new FileAction("zip", "ZIP")
            },

            ".png" => new[]
            {
                new FileAction("to_jpg", "JPG"),
                new FileAction("to_webp", "WEBP"),
                new FileAction("to_pdf", "PDF"),
                new FileAction("edit_image", "EDIT"),
                new FileAction("zip", "ZIP")
            },

            ".webp" => new[]
            {
                new FileAction("to_jpg", "JPG"),
                new FileAction("to_png", "PNG"),
                new FileAction("to_pdf", "PDF"),
                new FileAction("edit_image", "EDIT"),
                new FileAction("zip", "ZIP")
            },

            ".docx" => new[]
            {
                new FileAction("to_pdf", "PDF"),
                new FileAction("zip", "ZIP")
            },

            ".xlsx" => new[]
            {
                new FileAction("to_pdf", "PDF"),
                new FileAction("zip", "ZIP")
            },

            ".pptx" => new[]
            {
                new FileAction("to_pdf", "PDF"),
                new FileAction("zip", "ZIP")
            },

            ".pdf" => new[]
            {
                new FileAction("pdf_to_docx", "DOCX"),
                new FileAction("pdf_to_ocr_docx", "OCR DOCX"),
                new FileAction("pdf_to_xlsx", "XLSX"),
                new FileAction("pdf_to_pptx", "PPTX"),
                new FileAction("pdf_to_txt", "TXT"),
                new FileAction("pdf_to_png", "PNG"),
                new FileAction("pdf_to_jpg", "JPG"),
                new FileAction("extract_pdf_images", "IMAGES"),
                new FileAction("extract_pdf_pages", "PAGES"),
                new FileAction("split_pdf", "SPLIT"),
                new FileAction("delete_pdf_pages", "DELETE"),
                new FileAction("rotate_pdf_pages", "ROTATE"),
                new FileAction("reorder_pdf_pages", "REORDER"),
                new FileAction("zip", "ZIP")
            },

            _ => new[]
            {
                new FileAction("zip", "ZIP")
            }
        };
    }
}