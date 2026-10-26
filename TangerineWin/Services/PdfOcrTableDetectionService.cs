namespace TangerineWin.Services;

public sealed class PdfOcrTableDetectionService
{
    public sealed record OcrTable(
        int Top,
        int Bottom,
        IReadOnlyList<IReadOnlyList<string>> Rows,
        IReadOnlyList<PdfOcrService.OcrWord> Words);

    private sealed record CandidateRow(
        int Top,
        int Bottom,
        IReadOnlyList<CellInfo> Cells,
        IReadOnlyList<PdfOcrService.OcrWord> Words);

    private sealed record CellInfo(
        int Left,
        string Text);

    public IReadOnlyList<OcrTable> Detect(
        IReadOnlyList<PdfOcrService.OcrWord> words)
    {
        if (words.Count == 0)
            return Array.Empty<OcrTable>();

        var candidateRows =
            BuildCandidateRows(words)
                .OrderBy(row => row.Top)
                .ToArray();

        var tables =
            new List<OcrTable>();

        var index = 0;

        while (index < candidateRows.Length)
        {
            var group =
                new List<CandidateRow>
                {
                    candidateRows[index]
                };

            var currentIndex =
                index + 1;

            while (currentIndex <
                   candidateRows.Length)
            {
                var previous =
                    group[^1];

                var current =
                    candidateRows[currentIndex];

                if (!RowsAreCompatible(
                        previous,
                        current))
                {
                    break;
                }

                group.Add(current);

                currentIndex++;
            }

            // Минимум 2 строки.
            if (group.Count >= 2)
            {
                var columnCount =
                    group
                        .Select(row =>
                            row.Cells.Count)
                        .GroupBy(count => count)
                        .OrderByDescending(grouping =>
                            grouping.Count())
                        .ThenByDescending(grouping =>
                            grouping.Key)
                        .First()
                        .Key;

                var usableRows =
                    group
                        .Where(row =>
                            row.Cells.Count ==
                            columnCount)
                        .ToArray();

                if (usableRows.Length >= 2 &&
                    columnCount >= 2)
                {
                    var tableRows =
                        usableRows
                            .Select(row =>
                                (IReadOnlyList<string>)
                                row.Cells
                                    .Select(cell =>
                                        cell.Text)
                                    .ToArray())
                            .ToArray();

                    var tableWords =
                        usableRows
                            .SelectMany(row =>
                                row.Words)
                            .Distinct()
                            .ToArray();

                    tables.Add(
                        new OcrTable(
                            usableRows.Min(row =>
                                row.Top),
                            usableRows.Max(row =>
                                row.Bottom),
                            tableRows,
                            tableWords));
                }
            }

            if (group.Count > 1)
            {
                index += group.Count;
            }
            else
            {
                index++;
            }
        }

        return tables;
    }

    private static IReadOnlyList<CandidateRow>
        BuildCandidateRows(
            IReadOnlyList<PdfOcrService.OcrWord> words)
    {
        var result =
            new List<CandidateRow>();

        var lines =
            words
                .Where(word =>
                    word.Confidence >= 20)
                .GroupBy(word =>
                    new
                    {
                        word.BlockNumber,
                        word.ParagraphNumber,
                        word.LineNumber
                    });

        foreach (var line in lines)
        {
            var lineWords =
                line
                    .OrderBy(word =>
                        word.Left)
                    .ToArray();

            if (lineWords.Length < 2)
                continue;

            var medianHeight =
                lineWords
                    .Select(word =>
                        word.Height)
                    .OrderBy(value => value)
                    .ElementAt(
                        lineWords.Length / 2);

            // Большой горизонтальный промежуток
            // трактуем как границу между колонками.
            var gapThreshold =
                Math.Max(
                    45,
                    medianHeight * 2.2);

            var cells =
                new List<List<PdfOcrService.OcrWord>>();

            var currentCell =
                new List<PdfOcrService.OcrWord>();

            for (var i = 0;
                 i < lineWords.Length;
                 i++)
            {
                var word =
                    lineWords[i];

                if (i > 0)
                {
                    var previous =
                        lineWords[i - 1];

                    var gap =
                        word.Left -
                        previous.Right;

                    if (gap >
                        gapThreshold)
                    {
                        if (currentCell.Count > 0)
                        {
                            cells.Add(
                                currentCell);
                        }

                        currentCell =
                            new List<PdfOcrService.OcrWord>();
                    }
                }

                currentCell.Add(word);
            }

            if (currentCell.Count > 0)
            {
                cells.Add(
                    currentCell);
            }

            // Обычная строка текста,
            // а не потенциальная таблица.
            if (cells.Count < 2)
                continue;

            var cellInfos =
                cells
                    .Select(cell =>
                        new CellInfo(
                            cell.Min(word =>
                                word.Left),
                            string.Join(
                                " ",
                                cell.Select(word =>
                                    word.Text))))
                    .ToArray();

            result.Add(
                new CandidateRow(
                    lineWords.Min(word =>
                        word.Top),
                    lineWords.Max(word =>
                        word.Bottom),
                    cellInfos,
                    lineWords));
        }

        return result;
    }

    private static bool RowsAreCompatible(
        CandidateRow first,
        CandidateRow second)
    {
        if (first.Cells.Count !=
            second.Cells.Count)
        {
            return false;
        }

        var verticalGap =
            second.Top -
            first.Bottom;

        var averageHeight =
            (
                (first.Bottom - first.Top) +
                (second.Bottom - second.Top)
            ) / 2.0;

        // Слишком далеко друг от друга —
        // скорее всего уже другой блок.
        if (verticalGap >
            Math.Max(
                100,
                averageHeight * 4))
        {
            return false;
        }

        var alignedColumns = 0;

        for (var i = 0;
             i < first.Cells.Count;
             i++)
        {
            var difference =
                Math.Abs(
                    first.Cells[i].Left -
                    second.Cells[i].Left);

            if (difference <= 80)
            {
                alignedColumns++;
            }
        }

        var required =
            Math.Max(
                2,
                (int)Math.Ceiling(
                    first.Cells.Count *
                    0.66));

        return alignedColumns >= required;
    }
}