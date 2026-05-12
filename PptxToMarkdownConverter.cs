using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace Converter;

public sealed class PptxToMarkdownConverter : IFileToMarkdownConverter
{
    public void Convert(string sourcePath, string displaySourcePath, string outputPath)
    {
        MarkdownOutput.WriteAtomic(displaySourcePath, outputPath, output =>
        {
            using var fs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            using var doc = PresentationDocument.Open(fs, isEditable: false);
            PresentationPart? presPart = doc.PresentationPart;
            SlideIdList? slideIds = presPart?.Presentation?.SlideIdList;
            if (presPart is null || slideIds is null) return;

            int slideNumber = 0;
            foreach (SlideId slideId in slideIds.Elements<SlideId>())
            {
                slideNumber++;
                string? relId = slideId.RelationshipId?.Value;
                if (string.IsNullOrEmpty(relId)) continue;
                if (presPart.GetPartById(relId) is not SlidePart slidePart) continue;

                output.WriteLine($"<!-- slide {slideNumber} -->");
                output.WriteLine();

                ShapeTree? tree = slidePart.Slide?.CommonSlideData?.ShapeTree;
                if (tree is not null)
                    WriteShapeTree(output, tree);

                WriteSpeakerNotes(output, slidePart.NotesSlidePart);
                output.Flush();
            }
        });
    }

    private static void WriteShapeTree(StreamWriter output, ShapeTree tree)
    {
        foreach (Shape shape in tree.Elements<Shape>())
        {
            TextBody? body = shape.TextBody;
            if (body is null) continue;
            bool isTitle = IsTitleShape(shape);
            WriteTextBody(output, body, asTitle: isTitle);
        }

        // Tables inside slides
        foreach (var graphicFrame in tree.Elements<GraphicFrame>())
        {
            var table = graphicFrame.Descendants<D.Table>().FirstOrDefault();
            if (table is null) continue;
            WriteSlideTable(output, table);
        }
    }

    private static bool IsTitleShape(Shape shape)
    {
        var ph = shape.NonVisualShapeProperties?
            .ApplicationNonVisualDrawingProperties?
            .PlaceholderShape;
        if (ph is null) return false;
        var t = ph.Type?.Value;
        return t == PlaceholderValues.Title
            || t == PlaceholderValues.CenteredTitle
            || t == PlaceholderValues.SubTitle;
    }

    private static void WriteTextBody(StreamWriter output, TextBody body, bool asTitle)
    {
        var paragraphs = body.Elements<D.Paragraph>().ToList();
        if (paragraphs.Count == 0) return;

        bool wroteAny = false;
        foreach (D.Paragraph para in paragraphs)
        {
            string text = ExtractParagraphText(para);
            if (string.IsNullOrWhiteSpace(text)) continue;

            int level = para.ParagraphProperties?.Level?.Value ?? 0;

            if (asTitle && !wroteAny)
            {
                output.WriteLine("### " + text);
                output.WriteLine();
            }
            else
            {
                string indent = new string(' ', Math.Max(0, level) * 2);
                output.WriteLine(indent + "- " + text);
            }
            wroteAny = true;
        }

        if (wroteAny) output.WriteLine();
    }

    private static string ExtractParagraphText(D.Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (OpenXmlElement child in para.ChildElements)
        {
            switch (child)
            {
                case D.Run run:
                    AppendDrawingRun(run, sb);
                    break;
                case D.Field field:
                    var ft = field.Descendants<D.Text>().FirstOrDefault();
                    if (ft is not null) sb.Append(ft.Text);
                    break;
                case D.Break:
                    sb.Append('\n');
                    break;
            }
        }
        return sb.ToString().Trim();
    }

    private static void AppendDrawingRun(D.Run run, StringBuilder sb)
    {
        string raw = run.Text?.Text ?? string.Empty;
        if (raw.Length == 0) return;

        D.RunProperties? props = run.RunProperties;
        bool bold = props?.Bold?.Value == true;
        bool italic = props?.Italic?.Value == true;

        if (bold && italic) sb.Append("***").Append(raw).Append("***");
        else if (bold) sb.Append("**").Append(raw).Append("**");
        else if (italic) sb.Append("*").Append(raw).Append("*");
        else sb.Append(raw);
    }

    private static void WriteSlideTable(StreamWriter output, D.Table table)
    {
        var rows = table.Elements<D.TableRow>().ToList();
        if (rows.Count == 0) return;
        int cols = rows.Max(r => r.Elements<D.TableCell>().Count());
        if (cols == 0) return;

        output.WriteLine();
        for (int i = 0; i < rows.Count; i++)
        {
            var cellList = rows[i].Elements<D.TableCell>().ToList();
            var parts = new string[cols];
            for (int c = 0; c < cols; c++)
            {
                string cellText = c < cellList.Count ? ExtractCellText(cellList[c]) : string.Empty;
                parts[c] = cellText.Replace("|", @"\|").Replace("\r", string.Empty).Replace("\n", " ");
            }
            output.Write("| ");
            output.Write(string.Join(" | ", parts));
            output.WriteLine(" |");
            if (i == 0)
            {
                output.Write("| ");
                output.Write(string.Join(" | ", Enumerable.Repeat("---", cols)));
                output.WriteLine(" |");
            }
        }
        output.WriteLine();
    }

    private static string ExtractCellText(D.TableCell cell)
    {
        var sb = new StringBuilder();
        bool first = true;
        foreach (D.Paragraph p in cell.Descendants<D.Paragraph>())
        {
            string t = ExtractParagraphText(p);
            if (t.Length == 0) continue;
            if (!first) sb.Append("<br>");
            sb.Append(t);
            first = false;
        }
        return sb.ToString();
    }

    private static void WriteSpeakerNotes(StreamWriter output, NotesSlidePart? notesPart)
    {
        ShapeTree? tree = notesPart?.NotesSlide?.CommonSlideData?.ShapeTree;
        if (tree is null) return;

        var sb = new StringBuilder();
        foreach (Shape shape in tree.Elements<Shape>())
        {
            TextBody? body = shape.TextBody;
            if (body is null) continue;
            foreach (D.Paragraph p in body.Elements<D.Paragraph>())
            {
                string t = ExtractParagraphText(p);
                if (t.Length == 0) continue;
                sb.AppendLine(t);
            }
        }

        string notes = sb.ToString().Trim();
        if (notes.Length == 0) return;

        output.WriteLine("> **Speaker notes:**");
        foreach (string line in notes.Split('\n'))
            output.WriteLine("> " + line.TrimEnd('\r'));
        output.WriteLine();
    }
}
