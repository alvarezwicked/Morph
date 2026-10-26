using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.IO;
using TangerineWin.Views;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using ExcelCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using ExcelFont = DocumentFormat.OpenXml.Spreadsheet.Font;
using PdfPage = UglyToad.PdfPig.Content.Page;
using PdfWord = UglyToad.PdfPig.Content.Word;

namespace TangerineWin.Services;

public sealed class PdfExcelExtractionService
{
    private readonly PdfTableExtractionService
        _tableExtractionService =
            new PdfTableExtractionService();
    private readonly PdfOcrService
    _ocrService =
        new PdfOcrService();

    private readonly PdfOcrTableDetectionService
        _ocrTableDetectionService =
            new PdfOcrTableDetectionService();

    public async Task<string> ConvertAsync(
        string sourcePath,
        PdfExcelExportMode exportMode)
    {
        var stage = "start";

        try
        {
            if (!File.Exists(sourcePath))
            {
                throw new FileNotFoundException(
                    "PDF file was not found.",
                    sourcePath);
            }

            stage = "extract tables";

            IReadOnlyList<PdfTableExtractionService.ExtractedTable>
                allTables;

            try
            {
                allTables =
                    await _tableExtractionService.ExtractAsync(
                        sourcePath);
            }
            catch
            {
                allTables =
                    Array.Empty<
                        PdfTableExtractionService.ExtractedTable>();
            }

            stage = "create xlsx";

            return await Task.Run(async () =>
            {
                stage = "create output path";

                var outputPath =
                    CreateUniqueOutputPath(
                        sourcePath);

                stage = "open pdf";

                using var pdf =
                    PdfDocument.Open(
                        sourcePath,
                        new ParsingOptions
                        {
                            ClipPaths = true
                        });

                stage = "create spreadsheet";

                using var spreadsheet =
                    SpreadsheetDocument.Create(
                        outputPath,
                        SpreadsheetDocumentType.Workbook);

                stage = "create workbook";

                var workbookPart =
                    spreadsheet.AddWorkbookPart();

                workbookPart.Workbook =
                    new Workbook();

                stage = "create worksheet";

                var worksheetPart =
                    workbookPart.AddNewPart<WorksheetPart>();

                var sheetData =
                    new SheetData();

                worksheetPart.Worksheet =
                    new Worksheet(
                        sheetData);

                stage = "create styles";

                var stylesPart =
                    workbookPart.AddNewPart<WorkbookStylesPart>();

                stylesPart.Stylesheet =
                    CreateStylesheet();

                stylesPart.Stylesheet.Save();

                stage = "create sheet";

                var sheets =
                    workbookPart.Workbook.AppendChild(
                        new Sheets());

                sheets.Append(
                    new Sheet
                    {
                        Id =
                            workbookPart.GetIdOfPart(
                                worksheetPart),

                        SheetId = 1,

                        Name = "Sheet1"
                    });

                uint rowNumber = 1;

                stage = "process pages";

                for (var pageNumber = 1;
                     pageNumber <= pdf.NumberOfPages;
                     pageNumber++)
                {
                    var page =
                        pdf.GetPage(
                            pageNumber);

                    //
                    // Определяем, есть ли на странице нормальный
                    // текстовый слой.
                    //
                    var nativeWords =
                        page.GetWords()
                            .Where(word =>
                                !string.IsNullOrWhiteSpace(
                                    word.Text))
                            .ToArray();

                    var meaningfulWordCount =
                        nativeWords.Count(word =>
                            word.Text.Any(
                                char.IsLetterOrDigit));

                    //
                    // Если PdfPig видит меньше 5 содержательных слов,
                    // считаем страницу сканом и запускаем OCR.
                    //
                    var useOcr =
                        meaningfulWordCount < 5;

                    if (!useOcr)
                    {
                        //
                        // Обычный PDF:
                        // используем уже работающий путь.
                        //
                        var pageTables =
                            allTables
                                .Where(table =>
                                    table.PageNumber ==
                                    pageNumber)
                                .OrderByDescending(table =>
                                    table.Top)
                                .ToArray();

                        if (exportMode ==
                            PdfExcelExportMode.TablesOnly)
                        {
                            WriteTablesOnly(
                                sheetData,
                                pageTables,
                                ref rowNumber);
                        }
                        else
                        {
                            WriteTextAndTables(
                                sheetData,
                                page,
                                pageTables,
                                ref rowNumber);
                        }
                    }
                    else
                    {
                        //
                        // Сканированная страница.
                        //
                        // PdfOcrService использует индекс страницы,
                        // поэтому pageNumber - 1.
                        //
                        var ocrResult =
                            await _ocrService
                                .RecognizePageAsync(
                                    sourcePath,
                                    pageNumber - 1);

                        var ocrTables =
                            _ocrTableDetectionService
                                .Detect(
                                    ocrResult.Words)
                                .OrderBy(table =>
                                    table.Top)
                                .ToArray();

                        if (exportMode ==
                            PdfExcelExportMode.TablesOnly)
                        {
                            //
                            // Только найденные OCR-таблицы.
                            //
                            foreach (var table in ocrTables)
                            {
                                AddRowsTable(
                                    sheetData,
                                    table.Rows,
                                    ref rowNumber);

                                rowNumber++;
                            }
                        }
                        else
                        {
                            //
                            // Таблицы + обычный OCR-текст.
                            //

                            //
                            // Слова, которые уже принадлежат таблицам,
                            // второй раз в обычный текст не добавляем.
                            //
                            var tableWords =
                                ocrTables
                                    .SelectMany(table =>
                                        table.Words)
                                    .ToHashSet();

                            //
                            // Здесь объединяем текстовые строки и таблицы
                            // в один список, чтобы сохранить порядок сверху вниз.
                            //
                            var blocks =
                                new List<(
                                    double Top,
                                    string? Text,
                                    IReadOnlyList<IReadOnlyList<string>>? Rows
                                )>();

                            var normalWords =
                                ocrResult.Words
                                    .Where(word =>
                                        !tableWords.Contains(word) &&
                                        !string.IsNullOrWhiteSpace(
                                            word.Text))
                                    .ToArray();

                            //
                            // Tesseract уже отдаёт номера Block /
                            // Paragraph / Line, поэтому строки можно
                            // собирать значительно точнее, чем вручную.
                            //
                            var ocrLineBuilders =
                                new List<OcrTextLineBuilder>();

                            foreach (var word in normalWords)
                            {
                                OcrTextLineBuilder? matchingLine =
                                    null;

                                foreach (var line in ocrLineBuilders)
                                {
                                    var tolerance =
                                        Math.Max(
                                            4.0,
                                            Math.Max(
                                                word.Height,
                                                line.AverageHeight) *
                                            0.6);

                                    if (Math.Abs(
                                            word.Top -
                                            line.Top) <=
                                        tolerance)
                                    {
                                        matchingLine =
                                            line;

                                        break;
                                    }
                                }

                                if (matchingLine is null)
                                {
                                    matchingLine =
                                        new OcrTextLineBuilder();

                                    ocrLineBuilders.Add(
                                        matchingLine);
                                }

                                matchingLine.Add(
                                    word);
                            }

                            var textLines =
                                ocrLineBuilders
                                    .Select(line =>
                                        line.Build())
                                    .Where(line =>
                                        !string.IsNullOrWhiteSpace(
                                            line.Text))
                                    .OrderBy(line =>
                                        line.Top)
                                    .ToArray();

                            foreach (var line in textLines)
                            {
                                blocks.Add(
                                    (
                                        line.Top,
                                        line.Text,
                                        null
                                    ));
                            }

                            foreach (var table in ocrTables)
                            {
                                blocks.Add(
                                    (
                                        table.Top,
                                        null,
                                        table.Rows
                                    ));
                            }

                            //
                            // В OCR координата Top идёт от верхнего края
                            // изображения, поэтому меньший Top находится выше.
                            //
                            foreach (var block in
                                     blocks.OrderBy(block =>
                                         block.Top))
                            {
                                if (block.Rows is not null)
                                {
                                    AddRowsTable(
                                        sheetData,
                                        block.Rows,
                                        ref rowNumber);
                                }
                                else if (!string.IsNullOrWhiteSpace(
                                             block.Text))
                                {
                                    AddTextRow(
                                        sheetData,
                                        block.Text,
                                        ref rowNumber);
                                }
                            }
                        }
                    }

                    //
                    // Пустая строка между страницами.
                    //
                    if (pageNumber <
                        pdf.NumberOfPages)
                    {
                        rowNumber++;
                    }
                }
                stage = "column widths";

                AddColumnWidths(
                    worksheetPart.Worksheet,
                    sheetData);

                stage = "save worksheet";

                worksheetPart.Worksheet.Save();

                stage = "save workbook";

                workbookPart.Workbook.Save();

                stage = "finished";

                return outputPath;
            });
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                $"PDF → XLSX failed at stage: {stage}\n\n" +
                $"{ex.GetType().Name}: {ex.Message}\n\n" +
                ex.StackTrace,
                "PDF → XLSX error",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);

            throw;
        }
    }

    // ====================================================
    // TABLES ONLY
    // ====================================================
    private static void AddRowsTable(
        SheetData sheetData,
        IReadOnlyList<IReadOnlyList<string>> rows,
        ref uint rowNumber)
    {
        for (var rowIndex = 0;
             rowIndex < rows.Count;
             rowIndex++)
        {
            var sourceRow =
                rows[rowIndex];

            var row =
                new Row
                {
                    RowIndex =
                        rowNumber
                };

            var styleIndex =
                rowIndex == 0
                    ? 2u
                    : 1u;

            foreach (var value
                     in sourceRow)
            {
                row.Append(
                    CreateCell(
                        value,
                        styleIndex));
            }

            sheetData.Append(
                row);

            rowNumber++;
        }
    }
    // ====================================================
    // TEXT + TABLES
    // ====================================================
    private static void WriteTablesOnly(
    SheetData sheetData,
    IReadOnlyList<PdfTableExtractionService.ExtractedTable> tables,
    ref uint rowNumber)
    {
        foreach (var table in tables)
        {
            AddTable(
                sheetData,
                table,
                ref rowNumber);

            // Пустая строка между таблицами.
            rowNumber++;
        }
    }
    private static void WriteTextAndTables(
        SheetData sheetData,
        PdfPage page,
        IReadOnlyList<PdfTableExtractionService.ExtractedTable>
            tables,
        ref uint rowNumber)
    {
        var textLines =
            ExtractTextOutsideTables(
                page,
                tables);

        var blocks =
            new List<PageBlock>();

        foreach (var line in textLines)
        {
            blocks.Add(
                new PageBlock(
                    line.Top,
                    line.Text,
                    null));
        }

        foreach (var table in tables)
        {
            blocks.Add(
                new PageBlock(
                    table.Top,
                    null,
                    table));
        }

        //
        // В PdfPig Y растёт снизу вверх,
        // поэтому верх страницы имеет больший Top.
        //
        foreach (var block in
                 blocks.OrderByDescending(
                     block => block.Top))
        {
            if (block.Table is not null)
            {
                AddTable(
                    sheetData,
                    block.Table,
                    ref rowNumber);
            }
            else if (!string.IsNullOrWhiteSpace(
                         block.Text))
            {
                AddTextRow(
                    sheetData,
                    block.Text,
                    ref rowNumber);
            }
        }
    }

    // ====================================================
    // TEXT EXTRACTION
    // ====================================================

    private static IReadOnlyList<TextLine>
        ExtractTextOutsideTables(
            PdfPage page,
            IReadOnlyList<PdfTableExtractionService.ExtractedTable>
                tables)
    {
        var words =
            page.GetWords()
                .Where(word =>
                    !IsInsideAnyTable(
                        word,
                        tables))
                .OrderByDescending(word =>
                    word.BoundingBox.Top)
                .ThenBy(word =>
                    word.BoundingBox.Left)
                .ToArray();

        if (words.Length == 0)
        {
            return Array.Empty<TextLine>();
        }

        var lines =
            new List<TextLineBuilder>();

        foreach (var word in words)
        {
            var centerY =
                (
                    word.BoundingBox.Top +
                    word.BoundingBox.Bottom
                ) / 2.0;

            TextLineBuilder? matchingLine =
                null;

            //
            // Ищем уже существующую строку
            // примерно на той же высоте.
            //
            foreach (var line in lines)
            {
                var tolerance =
                    Math.Max(
                        3.0,
                        Math.Max(
                            word.BoundingBox.Height,
                            line.AverageHeight) *
                        0.55);

                if (Math.Abs(
                        centerY -
                        line.CenterY) <=
                    tolerance)
                {
                    matchingLine =
                        line;

                    break;
                }
            }

            if (matchingLine is null)
            {
                matchingLine =
                    new TextLineBuilder();

                lines.Add(
                    matchingLine);
            }

            matchingLine.Add(
                word);
        }

        return lines
            .Select(line =>
                line.Build())
            .Where(line =>
                !string.IsNullOrWhiteSpace(
                    line.Text))
            .OrderByDescending(line =>
                line.Top)
            .ToArray();
    }

    private static bool IsInsideAnyTable(
        PdfWord word,
        IReadOnlyList<PdfTableExtractionService.ExtractedTable>
            tables)
    {
        var centerX =
            (
                word.BoundingBox.Left +
                word.BoundingBox.Right
            ) / 2.0;

        var centerY =
            (
                word.BoundingBox.Bottom +
                word.BoundingBox.Top
            ) / 2.0;

        //
        // Небольшой допуск нужен потому,
        // что координаты PdfPig и Tabula
        // могут отличаться на несколько единиц.
        //
        const double tolerance = 2.0;

        foreach (var table in tables)
        {
            if (centerX >=
                    table.Left -
                    tolerance &&
                centerX <=
                    table.Right +
                    tolerance &&
                centerY >=
                    table.Bottom -
                    tolerance &&
                centerY <=
                    table.Top +
                    tolerance)
            {
                return true;
            }
        }

        return false;
    }

    // ====================================================
    // WRITE TEXT
    // ====================================================

    private static void AddTextRow(
        SheetData sheetData,
        string text,
        ref uint rowNumber)
    {
        var row =
            new Row
            {
                RowIndex =
                    rowNumber
            };

        row.Append(
            CreateCell(
                text,
                0));

        sheetData.Append(
            row);

        rowNumber++;
    }

    // ====================================================
    // WRITE TABLE
    // ====================================================

    private static void AddTable(
        SheetData sheetData,
        PdfTableExtractionService.ExtractedTable table,
        ref uint rowNumber)
    {
        for (var rowIndex = 0;
             rowIndex < table.Rows.Count;
             rowIndex++)
        {
            var sourceRow =
                table.Rows[rowIndex];

            var row =
                new Row
                {
                    RowIndex =
                        rowNumber
                };

            var styleIndex =
                rowIndex == 0
                    ? 2u
                    : 1u;

            foreach (var value
                     in sourceRow)
            {
                row.Append(
                    CreateCell(
                        value,
                        styleIndex));
            }

            sheetData.Append(
                row);

            rowNumber++;
        }
    }
    // ====================================================
    // CELL CREATION
    // ====================================================

    private static ExcelCell CreateCell(
        string value,
        uint styleIndex)
    {
        value =
            value?.Trim()
            ?? string.Empty;

        if (TryParseNumber(
                value,
                out var number))
        {
            return new ExcelCell
            {
                DataType =
                    CellValues.Number,

                CellValue =
                    new CellValue(
                        number.ToString(
                            CultureInfo.InvariantCulture)),

                StyleIndex =
                    styleIndex
            };
        }

        return new ExcelCell
        {
            DataType =
                CellValues.InlineString,

            InlineString =
                new InlineString(
                    new Text(value)
                    {
                        Space =
                            SpaceProcessingModeValues
                                .Preserve
                    }),

            StyleIndex =
                styleIndex
        };
    }

    // ====================================================
    // NUMBER DETECTION
    // ====================================================

    private static bool TryParseNumber(
        string value,
        out decimal result)
    {
        result = 0;

        if (string.IsNullOrWhiteSpace(
                value))
        {
            return false;
        }

        var normalized =
            value
                .Replace(
                    "\u00A0",
                    "")
                .Replace(
                    " ",
                    "")
                .Replace(
                    ",",
                    ".");

        //
        // Не превращаем очевидные даты,
        // время и составные номера в числа.
        //
        if (normalized.Contains('/') ||
            normalized.Contains(':') ||
            (
                normalized.Contains('-') &&
                !normalized.StartsWith("-")
            ))
        {
            return false;
        }

        return decimal.TryParse(
            normalized,
            NumberStyles.Number |
            NumberStyles.AllowLeadingSign |
            NumberStyles.AllowDecimalPoint,
            CultureInfo.InvariantCulture,
            out result);
    }

    // ====================================================
    // COLUMN WIDTHS
    // ====================================================

    private static void AddColumnWidths(
        Worksheet worksheet,
        SheetData sheetData)
    {
        var maximumWidths =
            new Dictionary<int, double>();

        foreach (var row in sheetData.Elements<Row>())
        {
            var cells =
                row.Elements<ExcelCell>()
                    .ToArray();

            for (var columnIndex = 0;
                 columnIndex < cells.Length;
                 columnIndex++)
            {
                var cell =
                    cells[columnIndex];

                var text =
                    GetCellDisplayText(
                        cell);

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                //
                // Берём самую длинную строку,
                // если внутри ячейки есть переносы.
                //
                var longestLineLength =
                    text
                        .Split(
                            new[]
                            {
                            '\r',
                            '\n'
                            },
                            StringSplitOptions.RemoveEmptyEntries)
                        .Select(line =>
                            line.Length)
                        .DefaultIfEmpty(0)
                        .Max();

                //
                // Небольшой запас по ширине.
                //
                var desiredWidth =
                    longestLineLength + 2;

                //
                // Не даём колонке стать
                // слишком узкой или огромной.
                //
                desiredWidth =
                    Math.Clamp(
                        desiredWidth,
                        10,
                        45);

                if (!maximumWidths.TryGetValue(
                        columnIndex,
                        out var currentWidth) ||
                    desiredWidth > currentWidth)
                {
                    maximumWidths[columnIndex] =
                        desiredWidth;
                }
            }
        }


        //
        // Если данных вообще нет —
        // создаём хотя бы стандартную колонку A.
        //
        if (maximumWidths.Count == 0)
        {
            maximumWidths[0] = 15;
        }

        var columns =
            new Columns();

        foreach (var item in
                 maximumWidths.OrderBy(
                     pair => pair.Key))
        {
            var excelColumn =
                (uint)item.Key + 1;

            columns.Append(
                new Column
                {
                    Min = excelColumn,
                    Max = excelColumn,
                    Width = item.Value,
                    CustomWidth = true
                });
        }

        worksheet.InsertAt(
            columns,
            0);
    }
    private static string GetCellDisplayText(
    ExcelCell cell)
    {
        if (cell.DataType?.Value ==
            CellValues.InlineString)
        {
            return cell.InlineString?.InnerText
                   ?? string.Empty;
        }

        return cell.CellValue?.Text
               ?? cell.InnerText
               ?? string.Empty;
    }
    // ====================================================
    // EXCEL STYLES
    // ====================================================

    private static Stylesheet CreateStylesheet()
    {
        var fonts =
            new Fonts(

                //
                // Font 0 — обычный.
                //
                new ExcelFont(
                    new FontName
                    {
                        Val = "Calibri"
                    },
                    new FontSize
                    {
                        Val = 11
                    }),

                //
                // Font 1 — жирный заголовок таблицы.
                //
                new ExcelFont(
                    new Bold(),
                    new FontName
                    {
                        Val = "Calibri"
                    },
                    new FontSize
                    {
                        Val = 11
                    }));

        var fills =
            new Fills(

                //
                // Fill 0 — стандартный.
                //
                new Fill(
                    new PatternFill
                    {
                        PatternType =
                            PatternValues.None
                    }),

                //
                // Fill 1 — обязательный системный Gray125.
                //
                new Fill(
                    new PatternFill
                    {
                        PatternType =
                            PatternValues.Gray125
                    }),

                //
                // Fill 2 — лёгкая заливка заголовка таблицы.
                //
                new Fill(
                    new PatternFill(
                        new ForegroundColor
                        {
                            Rgb = "FFEDEDED"
                        },
                        new BackgroundColor
                        {
                            Indexed = 64
                        })
                    {
                        PatternType =
                            PatternValues.Solid
                    }));

        var borders =
            new Borders(

                //
                // Border 0 — без границ.
                //
                new Border(),

                //
                // Border 1 — тонкая сетка таблицы.
                //
                new Border(
                    new LeftBorder
                    {
                        Style =
                            BorderStyleValues.Thin
                    },
                    new RightBorder
                    {
                        Style =
                            BorderStyleValues.Thin
                    },
                    new TopBorder
                    {
                        Style =
                            BorderStyleValues.Thin
                    },
                    new BottomBorder
                    {
                        Style =
                            BorderStyleValues.Thin
                    },
                    new DiagonalBorder()));

        var cellFormats =
            new CellFormats(

                //
                // Style 0 — обычный текст вне таблицы.
                //
                new CellFormat
                {
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,

                    Alignment =
                        new Alignment
                        {
                            Vertical =
                                VerticalAlignmentValues.Top,

                            WrapText =
                                true
                        }
                },

                //
                // Style 1 — обычная строка таблицы.
                //
                new CellFormat
                {
                    FontId = 0,
                    FillId = 0,
                    BorderId = 1,

                    ApplyBorder = true,

                    Alignment =
                        new Alignment
                        {
                            Vertical =
                                VerticalAlignmentValues.Top,

                            WrapText =
                                true
                        }
                },

                //
                // Style 2 — заголовок таблицы.
                //
                new CellFormat
                {
                    FontId = 1,
                    FillId = 2,
                    BorderId = 1,

                    ApplyFont = true,
                    ApplyFill = true,
                    ApplyBorder = true,

                    Alignment =
                        new Alignment
                        {
                            Vertical =
                                VerticalAlignmentValues.Top,

                            WrapText =
                                true
                        }
                });

        return new Stylesheet(
            fonts,
            fills,
            borders,
            cellFormats);
    }
    // ====================================================
    // OUTPUT FILE
    // ====================================================

    private static string CreateUniqueOutputPath(
        string sourcePath)
    {
        var directory =
            Path.GetDirectoryName(
                sourcePath)
            ?? throw new InvalidOperationException(
                "Unable to determine source directory.");

        var name =
            Path.GetFileNameWithoutExtension(
                sourcePath);

        var candidate =
            Path.Combine(
                directory,
                $"{name}.xlsx");

        if (!File.Exists(candidate))
        {
            return candidate;
        }

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name} ({index}).xlsx");

            if (!File.Exists(candidate))
            {
                return candidate;
            }

            index++;
        }
    }

    // ====================================================
    // INTERNAL MODELS
    // ====================================================
    private sealed record OcrTextLine(
    double Top,
    string Text);
    private sealed record TextLine(
        double Top,
        string Text);

    private sealed record PageBlock(
        double Top,
        string? Text,
        PdfTableExtractionService.ExtractedTable? Table);

    private sealed class OcrTextLineBuilder
    {
        private readonly List<PdfOcrService.OcrWord>
            _words = new();

        public double Top
        {
            get
            {
                if (_words.Count == 0)
                    return 0;

                return _words.Average(
                    word =>
                        word.Top);
            }
        }

        public double AverageHeight
        {
            get
            {
                if (_words.Count == 0)
                    return 0;

                return _words.Average(
                    word =>
                        word.Height);
            }
        }

        public void Add(
            PdfOcrService.OcrWord word)
        {
            _words.Add(
                word);
        }

        public OcrTextLine Build()
        {
            var ordered =
                _words
                    .OrderBy(word =>
                        word.Left)
                    .ToArray();

            var text =
                string.Join(
                    " ",
                    ordered.Select(
                        word =>
                            word.Text));

            var top =
                ordered
                    .Select(word =>
                        (double)word.Top)
                    .DefaultIfEmpty(0)
                    .Min();

            return new OcrTextLine(
                top,
                text);
        }
    }
    private sealed class TextLineBuilder
    {
        private readonly List<PdfWord>
            _words = new();

        public double CenterY
        {
            get
            {
                if (_words.Count == 0)
                {
                    return 0;
                }

                return _words.Average(
                    word =>
                        (
                            word.BoundingBox.Top +
                            word.BoundingBox.Bottom
                        ) / 2.0);
            }
        }

        public double AverageHeight
        {
            get
            {
                if (_words.Count == 0)
                {
                    return 0;
                }

                return _words.Average(
                    word =>
                        word.BoundingBox.Height);
            }
        }

        public void Add(
            PdfWord word)
        {
            _words.Add(
                word);
        }

        public TextLine Build()
        {
            var ordered =
                _words
                    .OrderBy(word =>
                        word.BoundingBox.Left)
                    .ToArray();

            var text =
                string.Join(
                    " ",
                    ordered.Select(
                        word =>
                            word.Text));

            var top =
                ordered
                    .Select(word =>
                        word.BoundingBox.Top)
                    .DefaultIfEmpty(0)
                    .Max();

            return new TextLine(
                top,
                text);
        }
    }
}