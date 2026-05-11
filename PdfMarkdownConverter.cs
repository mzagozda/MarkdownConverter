using System.Text;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;
using Tesseract;
using UglyToad.PdfPig;
using PdfPage = UglyToad.PdfPig.Content.Page;

namespace Converter;

public sealed class PdfMarkdownConverter : IDisposable
{
    private const int MinExtractedCharsToTrustTextLayer = 16;
    private const int OcrRenderWidth = 1700;
    private const int OcrRenderHeight = 2200;

    private readonly string _tessdataPath;
    private readonly string _languages;

    // OCR engine is process-scoped: loading language data is expensive, and warnings
    // about a missing tessdata directory should be emitted at most once per run.
    private TesseractEngine? _ocrEngine;
    private bool _ocrInitAttempted;

    public PdfMarkdownConverter(string tessdataPath, string languages)
    {
        _tessdataPath = tessdataPath;
        _languages = languages;
    }

    public void Convert(string pdfPath, string mdPath)
    {
        string tmpPath = mdPath + ".tmp";
        string exPath = pdfPath + ".ex";

        IDocReader? docReader = null;

        try
        {
            using (var output = new StreamWriter(tmpPath, append: false, new UTF8Encoding(false)))
            using (var pdf = PdfDocument.Open(pdfPath))
            {
                int pageNumber = 0;
                foreach (PdfPage page in pdf.GetPages())
                {
                    pageNumber++;

                    string text = ExtractTextLayer(page);
                    bool useOcr = CountSignificantChars(text) < MinExtractedCharsToTrustTextLayer;

                    if (useOcr)
                    {
                        TesseractEngine? engine = GetOrInitOcrEngine();
                        if (engine is not null)
                        {
                            docReader ??= TryOpenDocReader(pdfPath);
                            if (docReader is not null)
                            {
                                string ocrText = OcrPage(docReader, pageNumber - 1, engine);
                                if (!string.IsNullOrWhiteSpace(ocrText))
                                    text = ocrText;
                            }
                        }
                    }

                    WritePage(output, pageNumber, text);
                    output.Flush();
                }
            }

            if (File.Exists(mdPath)) File.Delete(mdPath);
            File.Move(tmpPath, mdPath);
            TryDelete(exPath);
        }
        catch (Exception ex)
        {
            WriteErrorFile(exPath, tmpPath, pdfPath, ex);
            throw;
        }
        finally
        {
            docReader?.Dispose();
        }
    }

    public void Dispose()
    {
        _ocrEngine?.Dispose();
        _ocrEngine = null;
    }

    private static string ExtractTextLayer(PdfPage page)
    {
        // PdfPig's page.Text returns the page's text content stream; embedded raster images are not included.
        // This satisfies "skip any images that are mixed with textual content".
        return page.Text ?? string.Empty;
    }

    private static int CountSignificantChars(string text)
    {
        int n = 0;
        foreach (char c in text)
            if (!char.IsWhiteSpace(c) && !char.IsControl(c))
                n++;
        return n;
    }

    private TesseractEngine? GetOrInitOcrEngine()
    {
        if (_ocrInitAttempted) return _ocrEngine;
        _ocrInitAttempted = true;

        if (!Directory.Exists(_tessdataPath))
        {
            Console.Error.WriteLine($"[ocr-disabled] tessdata directory not found: {_tessdataPath}");
            return null;
        }

        try
        {
            _ocrEngine = new TesseractEngine(_tessdataPath, _languages, EngineMode.Default);
            return _ocrEngine;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[ocr-disabled] failed to initialize Tesseract ({_languages}): {ex.Message}");
            return null;
        }
    }

    private static IDocReader? TryOpenDocReader(string pdfPath)
    {
        try
        {
            return DocLib.Instance.GetDocReader(pdfPath, new PageDimensions(OcrRenderWidth, OcrRenderHeight));
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[ocr-render-failed] {pdfPath}: {ex.Message}");
            return null;
        }
    }

    private static string OcrPage(IDocReader docReader, int zeroBasedPageIndex, TesseractEngine engine)
    {
        using IPageReader pageReader = docReader.GetPageReader(zeroBasedPageIndex);
        int width = pageReader.GetPageWidth();
        int height = pageReader.GetPageHeight();
        byte[] bgra = pageReader.GetImage();

        byte[] png = EncodeBgraAsPng(bgra, width, height);

        using var pix = Pix.LoadFromMemory(png);
        using var result = engine.Process(pix);
        return result.GetText() ?? string.Empty;
    }

    private static byte[] EncodeBgraAsPng(byte[] bgra, int width, int height)
    {
        using var image = Image.LoadPixelData<Bgra32>(bgra, width, height);
        using var ms = new MemoryStream(capacity: width * height);
        image.Save(ms, new PngEncoder());
        return ms.ToArray();
    }

    private static void WritePage(StreamWriter output, int pageNumber, string text)
    {
        output.WriteLine($"<!-- page {pageNumber} -->");
        output.WriteLine();
        output.WriteLine(text.Trim());
        output.WriteLine();
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); }
        catch { /* swallow cleanup errors */ }
    }

    private static void WriteErrorFile(string exPath, string tmpPath, string pdfPath, Exception ex)
    {
        try
        {
            // Append the error block to whatever partial output we captured, then rename.
            // StreamWriter(append: true) creates the file if it doesn't exist, which covers
            // the case where Convert threw before any page was written.
            using (var writer = new StreamWriter(tmpPath, append: true, new UTF8Encoding(false)))
            {
                writer.WriteLine();
                writer.WriteLine("<!-- ===== CONVERSION FAILED ===== -->");
                writer.WriteLine();
                writer.WriteLine("# Conversion failed");
                writer.WriteLine();
                writer.WriteLine($"- Source:    `{pdfPath}`");
                writer.WriteLine($"- Timestamp: {DateTime.UtcNow:O}");
                writer.WriteLine($"- Type:      `{ex.GetType().FullName}`");
                writer.WriteLine($"- Message:   {ex.Message}");
                writer.WriteLine();
                writer.WriteLine("## Details");
                writer.WriteLine();
                writer.WriteLine("```");
                writer.WriteLine(ex.ToString());
                writer.WriteLine("```");
            }

            if (File.Exists(exPath)) File.Delete(exPath);
            File.Move(tmpPath, exPath);
        }
        catch
        {
            // Last-ditch: try to leave at least a minimal error file so the failure is visible.
            try { File.WriteAllText(exPath, $"Source: {pdfPath}{Environment.NewLine}{ex}", new UTF8Encoding(false)); }
            catch { /* give up */ }
            TryDelete(tmpPath);
        }
    }
}
