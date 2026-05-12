using System.Text;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;
using Tesseract;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using PdfPage = UglyToad.PdfPig.Content.Page;

namespace Converter;

public sealed class PdfMarkdownConverter : IFileToMarkdownConverter, IDisposable
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

    public void Convert(string sourcePath, string displaySourcePath, string outputPath)
    {
        IDocReader? docReader = null;
        try
        {
            MarkdownOutput.WriteAtomic(displaySourcePath, outputPath, output =>
            {
                using var pdf = PdfDocument.Open(sourcePath);

                int totalPages = pdf.NumberOfPages;
                int successfulPages = 0;
                int failedPages = 0;
                Exception? lastPageException = null;

                // Iterate by index so we can isolate per-page failures. Common cases
                // (malformed font dictionaries, broken page resources) throw inside
                // pdf.GetPage(N); without this, one bad page would forfeit the whole
                // document. We still produce a .ex if NO page extracts successfully.
                for (int pageNumber = 1; pageNumber <= totalPages; pageNumber++)
                {
                    string text;
                    try
                    {
                        PdfPage page = pdf.GetPage(pageNumber);
                        text = ExtractTextLayer(page);
                        bool useOcr = CountSignificantChars(text) < MinExtractedCharsToTrustTextLayer;

                        if (useOcr)
                        {
                            TesseractEngine? engine = GetOrInitOcrEngine();
                            if (engine is not null)
                            {
                                docReader ??= TryOpenDocReader(sourcePath);
                                if (docReader is not null)
                                {
                                    string ocrText = OcrPage(docReader, pageNumber - 1, engine);
                                    if (!string.IsNullOrWhiteSpace(ocrText))
                                        text = ocrText;
                                }
                            }
                        }

                        successfulPages++;
                    }
                    catch (Exception ex)
                    {
                        failedPages++;
                        lastPageException = ex;
                        Console.Error.WriteLine(
                            $"[page-fail] {displaySourcePath} page {pageNumber}: {ex.GetType().Name}: {ex.Message}");
                        text = $"<!-- page extraction failed: {ex.GetType().Name}: {EscapeForComment(ex.Message)} -->";
                    }

                    WritePage(output, pageNumber, text);
                    output.Flush();
                }

                // If we couldn't extract a single page, surface the failure as a document
                // error (produces a .ex) instead of leaving a Markdown file of only error
                // markers behind. Encrypted and truly-corrupt PDFs hit this path.
                if (totalPages > 0 && successfulPages == 0 && lastPageException is not null)
                    throw new InvalidOperationException(
                        $"All {totalPages} page(s) failed to extract. Last error: {lastPageException.Message}",
                        lastPageException);
            });
        }
        finally
        {
            docReader?.Dispose();
        }
    }

    private static string EscapeForComment(string s) =>
        s.Replace("--", "- -").Replace("\r", " ").Replace("\n", " ");

    public void Dispose()
    {
        _ocrEngine?.Dispose();
        _ocrEngine = null;
    }

    // Marker inserted between words that we can deterministically say come from
    // different page regions (different blocks, columns, or table cells). Surrounding
    // spaces keep the marker as its own visual token in the resulting Markdown.
    private const string RegionBreakSeparator = " - ";

    private static string ExtractTextLayer(PdfPage page)
    {
        // Embedded raster images are not part of the word stream, so this satisfies
        // "skip any images that are mixed with textual content" by construction.
        var words = page.GetWords().ToList();
        if (words.Count == 0) return string.Empty;

        double lineHeight = ComputeMedianHeight(words);
        if (lineHeight <= 0) lineHeight = 10.0;

        var sb = new StringBuilder(EstimateCapacity(words));
        sb.Append(words[0].Text);
        for (int i = 1; i < words.Count; i++)
        {
            sb.Append(ChooseSeparator(words[i - 1], words[i], lineHeight));
            sb.Append(words[i].Text);
        }
        return sb.ToString();
    }

    private static double ComputeMedianHeight(List<Word> words)
    {
        var heights = new List<double>(words.Count);
        foreach (var w in words)
        {
            double h = w.BoundingBox.Height;
            if (h > 0) heights.Add(h);
        }
        if (heights.Count == 0) return 0;
        heights.Sort();
        return heights[heights.Count / 2];
    }

    private static int EstimateCapacity(List<Word> words)
    {
        int sum = 0;
        foreach (var w in words) sum += w.Text.Length + 1;
        return sum;
    }

    // Returns the separator to place between `prev` and `curr` based on their geometry.
    // The defaults are conservative: only emit RegionBreakSeparator when the layout
    // signal is unambiguous, so legitimate in-paragraph line wraps stay attached.
    private static string ChooseSeparator(Word prev, Word curr, double lineHeight)
    {
        var pb = prev.BoundingBox;
        var cb = curr.BoundingBox;

        double sameLineTolerance = lineHeight * 0.5;
        bool sameLine = Math.Abs(pb.Bottom - cb.Bottom) <= sameLineTolerance;

        if (sameLine)
        {
            double hGap = cb.Left - pb.Right;
            // Backward on the same line — reading order looped, distinct region.
            if (hGap < -lineHeight * 0.5) return RegionBreakSeparator;
            // Same line, very wide gap — column or table-cell boundary.
            if (hGap > lineHeight * 5.0) return RegionBreakSeparator;
            return " ";
        }

        // Next word sits above the previous one — reading order jumped to a new region.
        if (cb.Bottom > pb.Top) return RegionBreakSeparator;

        // Vertical gap to the next line; large gaps indicate a block/paragraph break.
        double vGap = pb.Bottom - cb.Top;
        if (vGap > lineHeight * 1.8) return RegionBreakSeparator;

        // Otherwise: normal next-line wrap within the same paragraph.
        return " ";
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
        return ExtractOcrTextByBlock(result);
    }

    // Tesseract already segments output into blocks; join distinct blocks with the
    // region-break separator while leaving intra-block content alone, so line-broken
    // words within a paragraph stay attached.
    private static string ExtractOcrTextByBlock(Tesseract.Page result)
    {
        var sb = new StringBuilder();
        using ResultIterator iter = result.GetIterator();
        iter.Begin();
        bool firstBlock = true;
        do
        {
            string blockText = iter.GetText(PageIteratorLevel.Block) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(blockText)) continue;
            if (!firstBlock) sb.Append(RegionBreakSeparator);
            sb.Append(blockText.Trim());
            firstBlock = false;
        } while (iter.Next(PageIteratorLevel.Block));

        return sb.Length == 0 ? (result.GetText() ?? string.Empty) : sb.ToString();
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

}
