using System.Text;

namespace Converter;

public enum ExistingPolicy
{
    Skip,
    Overwrite
}

public static class Program
{
    private static readonly TimeSpan SourceNewerThreshold = TimeSpan.FromMinutes(5);

    private static readonly string[] SupportedExtensions =
    {
        ".pdf",
        ".docx", ".xlsx", ".pptx",
        ".doc",  ".xls",  ".ppt"
    };

    public static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        string? root = null;
        var policy = ExistingPolicy.Skip;
        string tessdataPath = Path.Combine(AppContext.BaseDirectory, "tessdata");
        string languages = "eng+pol";

        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            switch (a)
            {
                case "--overwrite":
                    policy = ExistingPolicy.Overwrite;
                    break;
                case "--skip":
                    policy = ExistingPolicy.Skip;
                    break;
                case "--tessdata":
                    if (++i >= args.Length) { Console.Error.WriteLine("--tessdata needs a path."); return 1; }
                    tessdataPath = args[i];
                    break;
                case "--lang":
                    if (++i >= args.Length) { Console.Error.WriteLine("--lang needs a value, e.g. eng+pol."); return 1; }
                    languages = args[i];
                    break;
                case "-h":
                case "--help":
                    PrintUsage();
                    return 0;
                default:
                    if (root is null) root = a;
                    else { Console.Error.WriteLine($"Unexpected argument: {a}"); return 1; }
                    break;
            }
        }

        if (root is null)
        {
            PrintUsage();
            return 1;
        }

        if (!Directory.Exists(root))
        {
            Console.Error.WriteLine($"Directory not found: {root}");
            return 1;
        }

        ReportLibreOfficeStatus();

        using var pdfConverter = new PdfMarkdownConverter(tessdataPath, languages);
        var docxConverter = new DocxToMarkdownConverter();
        var xlsxConverter = new XlsxToMarkdownConverter();
        var pptxConverter = new PptxToMarkdownConverter();

        int total = 0, converted = 0, skipped = 0, failed = 0;

        // Materialize so we can pre-scan for Office siblings before processing PDFs.
        var allFiles = EnumerateSupportedFiles(root).ToList();
        HashSet<string> officeBasePaths = BuildOfficeBasePathSet(allFiles);

        foreach (string srcPath in allFiles)
        {
            total++;
            string mdPath = srcPath + ".md";

            if (IsPdfShadowedByOfficeSibling(srcPath, mdPath, policy, officeBasePaths))
            {
                Console.WriteLine($"[skip] {srcPath} (Office sibling preferred)");
                skipped++;
                continue;
            }

            if (ShouldSkip(srcPath, mdPath, policy))
            {
                Console.WriteLine($"[skip] {srcPath}");
                skipped++;
                continue;
            }

            try
            {
                Console.WriteLine($"[convert] {srcPath}");
                ConvertOne(srcPath, mdPath, pdfConverter, docxConverter, xlsxConverter, pptxConverter);
                converted++;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[fail] {srcPath}: {ex.Message} (details: {srcPath}.ex)");
                failed++;
            }
        }

        Console.WriteLine($"Done. total={total} converted={converted} skipped={skipped} failed={failed}");
        return failed == 0 ? 0 : 2;
    }

    private static void ConvertOne(
        string srcPath, string mdPath,
        PdfMarkdownConverter pdf,
        DocxToMarkdownConverter docx,
        XlsxToMarkdownConverter xlsx,
        PptxToMarkdownConverter pptx)
    {
        string ext = Path.GetExtension(srcPath).ToLowerInvariant();

        if (LegacyOfficeUpgrader.IsLegacy(ext))
        {
            // Upgrade the legacy binary file to its modern XML equivalent in a temp
            // directory, then run the matching OpenXml converter. The .pdf/.docx/...
            // markdown still lands at <original>.md, and any failure produces
            // <original>.ex via MarkdownOutput's displaySourcePath.
            using LegacyOfficeUpgrader.UpgradedFile upgraded = LegacyOfficeUpgrader.Upgrade(srcPath);
            string upgradedExt = Path.GetExtension(upgraded.Path).ToLowerInvariant();
            IFileToMarkdownConverter inner = ConverterFor(upgradedExt, pdf, docx, xlsx, pptx);
            inner.Convert(upgraded.Path, srcPath, mdPath);
            return;
        }

        IFileToMarkdownConverter converter = ConverterFor(ext, pdf, docx, xlsx, pptx);
        converter.Convert(srcPath, srcPath, mdPath);
    }

    private static IFileToMarkdownConverter ConverterFor(
        string ext,
        PdfMarkdownConverter pdf,
        DocxToMarkdownConverter docx,
        XlsxToMarkdownConverter xlsx,
        PptxToMarkdownConverter pptx) => ext switch
        {
            ".pdf" => pdf,
            ".docx" => docx,
            ".xlsx" => xlsx,
            ".pptx" => pptx,
            _ => throw new NotSupportedException($"No converter registered for extension '{ext}'.")
        };

    private static IEnumerable<string> EnumerateSupportedFiles(string root)
    {
        var options = new EnumerationOptions
        {
            RecurseSubdirectories = true,
            IgnoreInaccessible = true,
            MatchCasing = MatchCasing.CaseInsensitive,
            AttributesToSkip = FileAttributes.ReparsePoint
        };
        foreach (string path in Directory.EnumerateFiles(root, "*", options))
        {
            string name = Path.GetFileName(path);

            // Microsoft Office writes a "~$<truncated-name>" owner-lock file next to a
            // document that's open in Word/Excel/PowerPoint. These are tiny metadata
            // files, not real Office packages, and trying to parse them always fails.
            if (name.StartsWith("~$", StringComparison.Ordinal)) continue;

            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (Array.IndexOf(SupportedExtensions, ext) >= 0)
                yield return path;
        }
    }

    private static bool ShouldSkip(string srcPath, string mdPath, ExistingPolicy policy)
    {
        if (!File.Exists(mdPath)) return false;

        DateTime srcTime = File.GetLastWriteTimeUtc(srcPath);
        DateTime mdTime = File.GetLastWriteTimeUtc(mdPath);
        if (srcTime - mdTime > SourceNewerThreshold) return false;

        return policy == ExistingPolicy.Skip;
    }

    private static void ReportLibreOfficeStatus()
    {
        LegacyOfficeUpgrader.ProbeResult probe = LegacyOfficeUpgrader.Probe();
        switch (probe.Status)
        {
            case LegacyOfficeUpgrader.ProbeStatus.Ok:
                Console.WriteLine(probe.Message);
                break;
            case LegacyOfficeUpgrader.ProbeStatus.NotFound:
            case LegacyOfficeUpgrader.ProbeStatus.VersionUnknown:
            case LegacyOfficeUpgrader.ProbeStatus.TooOld:
                Console.Error.WriteLine($"[warn] {probe.Message}");
                break;
        }
    }

    // Set of "C:/path/to/file" (no extension) for every Office source in the run.
    // Used to detect PDFs whose canonical source is the sibling Office file.
    private static HashSet<string> BuildOfficeBasePathSet(IEnumerable<string> allFiles)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string p in allFiles)
        {
            string ext = Path.GetExtension(p).ToLowerInvariant();
            if (ext == ".pdf") continue;
            string basePath = Path.ChangeExtension(p, null);
            if (!string.IsNullOrEmpty(basePath)) set.Add(basePath);
        }
        return set;
    }

    // A PDF is shadowed by its Office sibling when:
    //   - the file is a PDF,
    //   - the policy is --skip (the user has not asked for re-extraction),
    //   - the PDF's own .pdf.md does not yet exist,
    //   - and another supported file with the same base name and a non-.pdf extension is also being processed.
    // In that case the Office file is taken as the canonical source and the PDF is skipped.
    private static bool IsPdfShadowedByOfficeSibling(
        string srcPath, string mdPath, ExistingPolicy policy, HashSet<string> officeBasePaths)
    {
        if (policy != ExistingPolicy.Skip) return false;
        if (!string.Equals(Path.GetExtension(srcPath), ".pdf", StringComparison.OrdinalIgnoreCase)) return false;
        if (File.Exists(mdPath)) return false;
        string basePath = Path.ChangeExtension(srcPath, null);
        return !string.IsNullOrEmpty(basePath) && officeBasePaths.Contains(basePath);
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage: pdf2md <directory> [--overwrite|--skip] [--tessdata <path>] [--lang <list>]");
        Console.WriteLine();
        Console.WriteLine("  <directory>   Root directory to scan recursively for supported files.");
        Console.WriteLine("  --skip        Skip files where <file>.<ext>.md already exists (default).");
        Console.WriteLine("  --overwrite   Re-extract even when <file>.<ext>.md already exists.");
        Console.WriteLine("  --tessdata    Path to Tesseract tessdata directory (default: <appdir>/tessdata).");
        Console.WriteLine("  --lang        Tesseract language list (default: eng+pol).");
        Console.WriteLine();
        Console.WriteLine("Supported formats:");
        Console.WriteLine("  PDF    (.pdf)           text-layer extraction with OCR fallback");
        Console.WriteLine("  Word   (.docx, .doc)    direct OpenXml (modern); LibreOffice upgrade for legacy");
        Console.WriteLine("  Excel  (.xlsx, .xls)    direct OpenXml (modern); LibreOffice upgrade for legacy");
        Console.WriteLine("  PPoint (.pptx, .ppt)    direct OpenXml (modern); LibreOffice upgrade for legacy");
        Console.WriteLine();
        Console.WriteLine("Output: writes <name>.<ext>.md beside each source file, UTF-8 encoded.");
        Console.WriteLine("Source-newer rule: if source is more than 5 minutes newer than its .md, re-extract regardless of policy.");
        Console.WriteLine("Legacy formats: require LibreOffice (soffice) on PATH or installed at the default location.");
    }
}
