using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Converter;

public sealed class XlsxToMarkdownConverter : IFileToMarkdownConverter
{
    public void Convert(string sourcePath, string displaySourcePath, string outputPath)
    {
        MarkdownOutput.WriteAtomic(displaySourcePath, outputPath, output =>
        {
            using var fs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            using var doc = SpreadsheetDocument.Open(fs, isEditable: false);
            WorkbookPart? wbPart = doc.WorkbookPart;
            Sheets? sheets = wbPart?.Workbook?.Sheets;
            if (wbPart is null || sheets is null) return;

            SharedStringTable? sst = wbPart.SharedStringTablePart?.SharedStringTable;
            var sharedStrings = sst?.Elements<SharedStringItem>().ToArray() ?? Array.Empty<SharedStringItem>();

            bool firstSheet = true;
            foreach (Sheet sheet in sheets.Elements<Sheet>())
            {
                string? relId = sheet.Id?.Value;
                if (string.IsNullOrEmpty(relId)) continue;
                if (wbPart.GetPartById(relId) is not WorksheetPart wsPart) continue;
                SheetData? sheetData = wsPart.Worksheet?.GetFirstChild<SheetData>();

                if (!firstSheet) output.WriteLine();
                firstSheet = false;

                string name = sheet.Name?.Value ?? "Sheet";
                output.WriteLine($"## {name}");
                output.WriteLine();

                if (sheetData is null)
                {
                    output.WriteLine("_(empty sheet)_");
                    output.WriteLine();
                    continue;
                }

                WriteSheetAsTable(output, sheetData, sharedStrings);
            }
        });
    }

    private static void WriteSheetAsTable(StreamWriter output, SheetData sheetData, SharedStringItem[] sharedStrings)
    {
        var rows = sheetData.Elements<Row>().ToList();
        if (rows.Count == 0)
        {
            output.WriteLine("_(empty sheet)_");
            output.WriteLine();
            return;
        }

        int maxColIndex = 0;
        var rowDicts = new List<Dictionary<int, string>>(rows.Count);

        foreach (Row row in rows)
        {
            var cellsByCol = new Dictionary<int, string>();
            foreach (Cell cell in row.Elements<Cell>())
            {
                int col = ColumnIndexFromReference(cell.CellReference?.Value);
                cellsByCol[col] = GetCellValue(cell, sharedStrings);
                if (col > maxColIndex) maxColIndex = col;
            }
            rowDicts.Add(cellsByCol);
        }

        bool anyContent = rowDicts.Any(d => d.Values.Any(v => !string.IsNullOrEmpty(v)));
        if (!anyContent)
        {
            output.WriteLine("_(empty sheet)_");
            output.WriteLine();
            return;
        }

        int colCount = maxColIndex + 1;

        WriteRow(output, rowDicts[0], colCount);
        WriteSeparator(output, colCount);
        for (int i = 1; i < rowDicts.Count; i++)
            WriteRow(output, rowDicts[i], colCount);

        output.WriteLine();
    }

    private static void WriteRow(StreamWriter output, Dictionary<int, string> cells, int colCount)
    {
        var parts = new string[colCount];
        for (int c = 0; c < colCount; c++)
        {
            cells.TryGetValue(c, out string? v);
            parts[c] = EscapeCell(v ?? string.Empty);
        }
        output.Write("| ");
        output.Write(string.Join(" | ", parts));
        output.WriteLine(" |");
    }

    private static void WriteSeparator(StreamWriter output, int colCount)
    {
        output.Write("| ");
        output.Write(string.Join(" | ", Enumerable.Repeat("---", colCount)));
        output.WriteLine(" |");
    }

    private static string EscapeCell(string s) =>
        s.Replace("|", @"\|").Replace("\r", string.Empty).Replace("\n", " ");

    private static int ColumnIndexFromReference(string? reference)
    {
        if (string.IsNullOrEmpty(reference)) return 0;
        int result = 0;
        foreach (char c in reference)
        {
            if (c >= 'A' && c <= 'Z') result = result * 26 + (c - 'A' + 1);
            else if (c >= 'a' && c <= 'z') result = result * 26 + (c - 'a' + 1);
            else break;
        }
        return Math.Max(0, result - 1);
    }

    private static string GetCellValue(Cell cell, SharedStringItem[] sharedStrings)
    {
        CellValues? type = cell.DataType?.Value;

        if (type == CellValues.SharedString)
        {
            if (int.TryParse(cell.CellValue?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int ix)
                && ix >= 0 && ix < sharedStrings.Length)
            {
                return sharedStrings[ix].InnerText ?? string.Empty;
            }
            return string.Empty;
        }

        if (type == CellValues.InlineString)
            return cell.InlineString?.InnerText ?? string.Empty;

        if (type == CellValues.Boolean)
            return cell.CellValue?.Text == "1" ? "TRUE" : "FALSE";

        if (type == CellValues.String)
            return cell.CellValue?.Text ?? string.Empty;

        // Numeric, date, formula result, etc. We don't apply number-format strings —
        // raw values are reproducible across locales and good enough for Markdown.
        return cell.CellValue?.Text ?? cell.InnerText ?? string.Empty;
    }
}
