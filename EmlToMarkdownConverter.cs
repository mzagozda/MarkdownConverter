using System.Text;
using MimeKit;
using ReverseMd = ReverseMarkdown;

namespace Converter;

// Converts RFC 5322 / MIME email messages (.eml) to Markdown.
//
// Headers (From / To / Cc / Bcc / Date / Subject) are emitted as a bullet list.
// The message body is rendered as Markdown — HTML body is preferred and converted
// via ReverseMarkdown; otherwise the plain-text body is emitted as-is.
//
// Supported attachments (.eml, .pdf, .docx, .doc, .xlsx, .xls, .pptx, .ppt) are
// extracted to a temp directory, run through the matching existing converter, and
// the resulting Markdown is appended inline beneath an `## Attachment: <name>`
// heading so the whole email + attachments live in one .md file.
//
// Nested message/rfc822 parts and .eml attachments are handled recursively up to
// MaxNestingDepth. Any per-attachment failure is logged to stderr and inlined as
// a small error block — it never fails the parent EML conversion.
public sealed class EmlToMarkdownConverter : IFileToMarkdownConverter
{
    private const int MaxNestingDepth = 8;

    private readonly PdfMarkdownConverter _pdf;
    private readonly DocxToMarkdownConverter _docx;
    private readonly XlsxToMarkdownConverter _xlsx;
    private readonly PptxToMarkdownConverter _pptx;
    private readonly ReverseMd.Converter _htmlToMarkdown;

    public EmlToMarkdownConverter(
        PdfMarkdownConverter pdf,
        DocxToMarkdownConverter docx,
        XlsxToMarkdownConverter xlsx,
        PptxToMarkdownConverter pptx)
    {
        _pdf = pdf;
        _docx = docx;
        _xlsx = xlsx;
        _pptx = pptx;

        _htmlToMarkdown = new ReverseMd.Converter(new ReverseMd.Config
        {
            UnknownTags = ReverseMd.Config.UnknownTagsOption.Bypass,
            GithubFlavored = true,
            RemoveComments = true,
            SmartHrefHandling = true
        });
    }

    public void Convert(string sourcePath, string displaySourcePath, string outputPath)
    {
        MarkdownOutput.WriteAtomic(displaySourcePath, outputPath, output =>
        {
            using var fs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            MimeMessage msg = MimeMessage.Load(fs);
            WriteMessage(output, msg, displaySourcePath, level: 1, depth: 0);
        });
    }

    private void WriteMessage(StreamWriter output, MimeMessage msg, string displaySourcePath, int level, int depth)
    {
        string h = Heading(level);
        string subject = string.IsNullOrWhiteSpace(msg.Subject) ? "(no subject)" : msg.Subject.Trim();
        output.WriteLine(h + " " + EscapeForLine(subject));
        output.WriteLine();

        WriteHeader(output, "From", AddressList(msg.From));
        WriteHeader(output, "To", AddressList(msg.To));
        WriteHeader(output, "Cc", AddressList(msg.Cc));
        WriteHeader(output, "Bcc", AddressList(msg.Bcc));
        WriteHeader(output, "Date", msg.Date == default ? null : msg.Date.ToString("u"));
        WriteHeader(output, "Subject", subject);
        output.WriteLine();

        string body = ExtractBody(msg, displaySourcePath);
        if (!string.IsNullOrWhiteSpace(body))
        {
            output.WriteLine(body.TrimEnd());
            output.WriteLine();
        }
        output.Flush();

        int idx = 0;
        foreach (MimeEntity att in EnumerateAttachments(msg))
        {
            idx++;
            try
            {
                WriteAttachment(output, att, displaySourcePath, level + 1, idx, depth + 1);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(
                    $"[attachment-fail] {displaySourcePath} #{idx}: {ex.GetType().Name}: {ex.Message}");
                output.WriteLine();
                output.WriteLine(Heading(level + 1) + $" Attachment #{idx} (failed)");
                output.WriteLine();
                output.WriteLine($"- Error: `{ex.GetType().Name}`: {EscapeForLine(ex.Message)}");
                output.WriteLine();
            }
            output.Flush();
        }
    }

    // Yield every body part marked as an attachment plus any message/rfc822 parts
    // (which MimeKit does not necessarily flag as attachments when they appear in
    // multipart/mixed without a Content-Disposition).
    private static IEnumerable<MimeEntity> EnumerateAttachments(MimeMessage msg)
    {
        foreach (MimeEntity part in msg.BodyParts)
        {
            if (part is MessagePart) yield return part;
            else if (part.IsAttachment) yield return part;
        }
    }

    private string ExtractBody(MimeMessage msg, string displaySourcePath)
    {
        string? html = msg.HtmlBody;
        if (!string.IsNullOrEmpty(html))
        {
            try
            {
                return _htmlToMarkdown.Convert(html);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(
                    $"[eml-html-body] {displaySourcePath}: HTML→Markdown failed, falling back to plain text: {ex.Message}");
            }
        }
        return msg.TextBody ?? string.Empty;
    }

    private void WriteAttachment(
        StreamWriter output, MimeEntity att, string displaySourcePath, int level, int idx, int depth)
    {
        string h = Heading(level);

        // Inline nested message/rfc822
        if (att is MessagePart messagePart)
        {
            output.WriteLine();
            output.WriteLine(h + $" Attached email #{idx}");
            output.WriteLine();
            if (depth >= MaxNestingDepth)
            {
                output.WriteLine($"_(nested email skipped: depth limit {MaxNestingDepth} reached)_");
                output.WriteLine();
                return;
            }
            if (messagePart.Message is null)
            {
                output.WriteLine("_(nested email is empty)_");
                output.WriteLine();
                return;
            }
            WriteMessage(output, messagePart.Message, displaySourcePath, level + 1, depth);
            return;
        }

        if (att is not MimePart mp) return;

        string rawName = mp.FileName ?? mp.ContentDisposition?.FileName ?? $"attachment-{idx}";
        string fileName = string.IsNullOrWhiteSpace(rawName) ? $"attachment-{idx}" : rawName;
        string ext = Path.GetExtension(fileName).ToLowerInvariant();

        output.WriteLine();
        output.WriteLine(h + $" Attachment: {EscapeForLine(fileName)}");
        output.WriteLine();

        if (!IsSupportedAttachmentExt(ext))
        {
            string typeLabel = string.IsNullOrEmpty(ext) ? "<no extension>" : ext;
            output.WriteLine($"_(attachment type `{typeLabel}` is not converted)_");
            output.WriteLine();
            return;
        }

        if (depth >= MaxNestingDepth)
        {
            output.WriteLine($"_(attachment skipped: depth limit {MaxNestingDepth} reached)_");
            output.WriteLine();
            return;
        }

        string tempDir = Path.Combine(Path.GetTempPath(), "pdf2md_eml_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try
        {
            string tempSrc = Path.Combine(tempDir, SanitizeFileName(fileName));
            if (mp.Content is null)
            {
                output.WriteLine("_(attachment has no content; skipped)_");
                output.WriteLine();
                return;
            }
            using (var ofs = File.Create(tempSrc))
                mp.Content.DecodeTo(ofs);

            if (ext == ".eml")
            {
                using var efs = File.OpenRead(tempSrc);
                MimeMessage nested = MimeMessage.Load(efs);
                WriteMessage(output, nested, displaySourcePath, level + 1, depth);
                return;
            }

            ConvertBinaryAttachment(output, tempSrc, ext, displaySourcePath, fileName);
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { /* swallow */ }
        }
    }

    private void ConvertBinaryAttachment(
        StreamWriter output, string tempSrc, string ext, string displaySourcePath, string fileName)
    {
        // The inner converter writes to tempMd and, on failure, to tempDisplay + ".ex".
        // Both live inside tempDir so the cleanup in WriteAttachment removes them.
        string tempMd = tempSrc + ".md";
        string tempDisplay = tempSrc;

        try
        {
            if (LegacyOfficeUpgrader.IsLegacy(ext))
            {
                using LegacyOfficeUpgrader.UpgradedFile upgraded = LegacyOfficeUpgrader.Upgrade(tempSrc);
                string upgradedExt = Path.GetExtension(upgraded.Path).ToLowerInvariant();
                IFileToMarkdownConverter inner = ConverterFor(upgradedExt);
                inner.Convert(upgraded.Path, tempDisplay, tempMd);
            }
            else
            {
                IFileToMarkdownConverter inner = ConverterFor(ext);
                inner.Convert(tempSrc, tempDisplay, tempMd);
            }
        }
        catch (Exception ex)
        {
            // Re-thrown so the caller logs it and writes an inline error block under
            // the attachment heading. We still try to splice in any partial .md/.ex
            // output below for diagnostic value.
            AppendIfExists(output, tempMd);
            AppendIfExists(output, tempDisplay + ".ex");
            throw new InvalidOperationException(
                $"Attachment '{fileName}' conversion failed: {ex.Message}", ex);
        }

        AppendIfExists(output, tempMd);
    }

    private static void AppendIfExists(StreamWriter output, string path)
    {
        if (!File.Exists(path)) return;
        try
        {
            string content = File.ReadAllText(path);
            if (content.Length == 0) return;
            output.WriteLine(content.TrimEnd());
            output.WriteLine();
        }
        catch { /* swallow read errors — best-effort splicing */ }
    }

    private IFileToMarkdownConverter ConverterFor(string ext) => ext switch
    {
        ".pdf" => _pdf,
        ".docx" => _docx,
        ".xlsx" => _xlsx,
        ".pptx" => _pptx,
        _ => throw new NotSupportedException($"No converter registered for attachment extension '{ext}'.")
    };

    private static bool IsSupportedAttachmentExt(string ext) =>
        ext is ".eml" or ".pdf"
            or ".docx" or ".doc"
            or ".xlsx" or ".xls"
            or ".pptx" or ".ppt";

    private static string Heading(int level) => new string('#', Math.Clamp(level, 1, 6));

    private static void WriteHeader(StreamWriter output, string name, string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return;
        output.WriteLine($"- **{name}:** {EscapeForLine(value)}");
    }

    private static string AddressList(InternetAddressList? list)
    {
        if (list is null || list.Count == 0) return string.Empty;
        var sb = new StringBuilder();
        bool first = true;
        foreach (InternetAddress a in list)
        {
            if (!first) sb.Append("; ");
            sb.Append(a.ToString());
            first = false;
        }
        return sb.ToString();
    }

    private static string EscapeForLine(string s) =>
        s.Replace("\r", " ").Replace("\n", " ").Trim();

    private static string SanitizeFileName(string name)
    {
        var sb = new StringBuilder(name.Length);
        char[] invalid = Path.GetInvalidFileNameChars();
        foreach (char c in name)
            sb.Append(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
        string cleaned = sb.ToString().Trim();
        return cleaned.Length == 0 ? "attachment" : cleaned;
    }
}
