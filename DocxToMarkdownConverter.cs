using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Converter;

public sealed class DocxToMarkdownConverter : IFileToMarkdownConverter
{
    public void Convert(string sourcePath, string displaySourcePath, string outputPath)
    {
        MarkdownOutput.WriteAtomic(displaySourcePath, outputPath, output =>
        {
            // Use a FileStream with permissive sharing so we can still read files
            // that the user has open in Word (Word grants FileShare.Read on its lock).
            using var fs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            using var doc = WordprocessingDocument.Open(fs, isEditable: false);
            MainDocumentPart? main = doc.MainDocumentPart;
            Body? body = main?.Document?.Body;
            if (main is null || body is null) return;

            foreach (OpenXmlElement element in body.ChildElements)
            {
                switch (element)
                {
                    case Paragraph paragraph:
                        WriteParagraph(output, paragraph, main);
                        break;
                    case Table table:
                        WriteTable(output, table, main);
                        break;
                    case SectionProperties:
                        break; // section settings — not relevant to markdown
                }
            }
        });
    }

    private static void WriteParagraph(StreamWriter output, Paragraph paragraph, MainDocumentPart main)
    {
        string text = ExtractInlineText(paragraph, main).TrimEnd();
        int headingLevel = GetHeadingLevel(paragraph);

        if (headingLevel > 0 && text.Length > 0)
        {
            output.WriteLine();
            output.WriteLine(new string('#', Math.Min(headingLevel, 6)) + " " + text);
            output.WriteLine();
            return;
        }

        ListInfo? list = GetListInfo(paragraph);
        if (list is not null && text.Length > 0)
        {
            string indent = new string(' ', Math.Max(0, list.Level) * 2);
            string marker = list.IsNumbered ? "1. " : "- ";
            output.WriteLine(indent + marker + text);
            return;
        }

        if (text.Length == 0)
        {
            output.WriteLine();
            return;
        }

        output.WriteLine(text);
    }

    private static void WriteTable(StreamWriter output, Table table, MainDocumentPart main)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0) return;

        int maxCols = rows.Max(r => r.Elements<TableCell>().Count());
        if (maxCols == 0) return;

        output.WriteLine();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            TableRow row = rows[rowIndex];
            var cells = row.Elements<TableCell>().ToList();
            var cellTexts = new List<string>(maxCols);
            for (int c = 0; c < maxCols; c++)
            {
                string cellText = c < cells.Count ? ExtractCellText(cells[c], main) : string.Empty;
                cellTexts.Add(EscapeTableCell(cellText));
            }
            output.WriteLine("| " + string.Join(" | ", cellTexts) + " |");

            if (rowIndex == 0)
            {
                output.WriteLine("| " + string.Join(" | ", Enumerable.Repeat("---", maxCols)) + " |");
            }
        }
        output.WriteLine();
    }

    private static string ExtractCellText(TableCell cell, MainDocumentPart main)
    {
        var sb = new StringBuilder();
        bool firstPara = true;
        foreach (Paragraph p in cell.Elements<Paragraph>())
        {
            if (!firstPara) sb.Append("<br>");
            sb.Append(ExtractInlineText(p, main).TrimEnd());
            firstPara = false;
        }
        return sb.ToString();
    }

    private static string EscapeTableCell(string s) =>
        s.Replace("|", @"\|").Replace("\r", string.Empty).Replace("\n", " ");

    private static string ExtractInlineText(OpenXmlElement element, MainDocumentPart main)
    {
        var sb = new StringBuilder();
        WalkInline(element, main, sb);
        return sb.ToString();
    }

    private static void WalkInline(OpenXmlElement element, MainDocumentPart main, StringBuilder sb)
    {
        foreach (OpenXmlElement child in element.ChildElements)
        {
            switch (child)
            {
                case Run run:
                    AppendRun(run, sb);
                    break;
                case Hyperlink hyperlink:
                    AppendHyperlink(hyperlink, main, sb);
                    break;
                case ParagraphProperties:
                case RunProperties:
                case BookmarkStart:
                case BookmarkEnd:
                case ProofError:
                    break;
                default:
                    WalkInline(child, main, sb);
                    break;
            }
        }
    }

    private static void AppendRun(Run run, StringBuilder sb)
    {
        RunProperties? props = run.RunProperties;
        bool bold = props?.Bold is not null && props.Bold.Val?.Value != false;
        bool italic = props?.Italic is not null && props.Italic.Val?.Value != false;

        var runText = new StringBuilder();
        foreach (OpenXmlElement child in run.ChildElements)
        {
            switch (child)
            {
                case Text t: runText.Append(t.Text); break;
                case TabChar: runText.Append('\t'); break;
                case Break: runText.Append('\n'); break;
                case CarriageReturn: runText.Append('\n'); break;
            }
        }

        if (runText.Length == 0) return;
        string s = runText.ToString();
        if (bold && italic) s = $"***{s}***";
        else if (bold) s = $"**{s}**";
        else if (italic) s = $"*{s}*";
        sb.Append(s);
    }

    private static void AppendHyperlink(Hyperlink hyperlink, MainDocumentPart main, StringBuilder sb)
    {
        var inner = new StringBuilder();
        foreach (Run r in hyperlink.Elements<Run>())
            AppendRun(r, inner);

        string text = inner.ToString();
        string? url = ResolveHyperlinkUrl(hyperlink, main);

        if (!string.IsNullOrEmpty(url) && text.Length > 0)
            sb.Append('[').Append(text).Append("](").Append(url).Append(')');
        else
            sb.Append(text);
    }

    private static string? ResolveHyperlinkUrl(Hyperlink hyperlink, MainDocumentPart main)
    {
        string? id = hyperlink.Id?.Value;
        if (string.IsNullOrEmpty(id)) return null;
        try
        {
            HyperlinkRelationship? rel = main.HyperlinkRelationships.FirstOrDefault(r => r.Id == id);
            return rel?.Uri?.ToString();
        }
        catch
        {
            return null;
        }
    }

    private static int GetHeadingLevel(Paragraph paragraph)
    {
        string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrEmpty(styleId)) return 0;
        if (!styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) return 0;
        string suffix = styleId.AsSpan(7).ToString();
        return int.TryParse(suffix, out int level) ? level : 0;
    }

    private sealed record ListInfo(int Level, bool IsNumbered);

    private static ListInfo? GetListInfo(Paragraph paragraph)
    {
        NumberingProperties? numbering = paragraph.ParagraphProperties?.NumberingProperties;
        if (numbering is null) return null;
        // Without decoding the numbering definition we can't reliably tell bullet vs
        // numbered, so default to bullet — that's the safer presentation for the
        // common case where the source used a bulleted list style.
        int level = numbering.NumberingLevelReference?.Val?.Value ?? 0;
        return new ListInfo(level, IsNumbered: false);
    }
}
