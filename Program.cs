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

        using var converter = new PdfMarkdownConverter(tessdataPath, languages);

        int total = 0, converted = 0, skipped = 0, failed = 0;

        foreach (string pdfPath in EnumeratePdfFiles(root))
        {
            total++;
            string mdPath = pdfPath + ".md";

            if (ShouldSkip(pdfPath, mdPath, policy))
            {
                Console.WriteLine($"[skip] {pdfPath}");
                skipped++;
                continue;
            }

            try
            {
                Console.WriteLine($"[convert] {pdfPath}");
                converter.Convert(pdfPath, mdPath);
                converted++;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[fail] {pdfPath}: {ex.Message} (details: {pdfPath}.ex)");
                failed++;
            }
        }

        Console.WriteLine($"Done. total={total} converted={converted} skipped={skipped} failed={failed}");
        return failed == 0 ? 0 : 2;
    }

    private static IEnumerable<string> EnumeratePdfFiles(string root)
    {
        var options = new EnumerationOptions
        {
            RecurseSubdirectories = true,
            IgnoreInaccessible = true,
            MatchCasing = MatchCasing.CaseInsensitive,
            AttributesToSkip = FileAttributes.ReparsePoint
        };
        return Directory.EnumerateFiles(root, "*.pdf", options);
    }

    private static bool ShouldSkip(string pdfPath, string mdPath, ExistingPolicy policy)
    {
        if (!File.Exists(mdPath)) return false;

        DateTime pdfTime = File.GetLastWriteTimeUtc(pdfPath);
        DateTime mdTime = File.GetLastWriteTimeUtc(mdPath);
        if (pdfTime - mdTime > SourceNewerThreshold) return false;

        return policy == ExistingPolicy.Skip;
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage: pdf2md <directory> [--overwrite|--skip] [--tessdata <path>] [--lang <list>]");
        Console.WriteLine();
        Console.WriteLine("  <directory>   Root directory to scan recursively for .pdf files.");
        Console.WriteLine("  --skip        Skip files where <file>.pdf.md already exists (default).");
        Console.WriteLine("  --overwrite   Re-extract even when <file>.pdf.md already exists.");
        Console.WriteLine("  --tessdata    Path to Tesseract tessdata directory (default: <appdir>/tessdata).");
        Console.WriteLine("  --lang        Tesseract language list (default: eng+pol).");
        Console.WriteLine();
        Console.WriteLine("Output: writes <name>.pdf.md beside each .pdf, UTF-8 encoded, page-by-page.");
        Console.WriteLine("Source-newer rule: if .pdf is more than 5 minutes newer than .pdf.md, re-extract regardless of policy.");
    }
}
