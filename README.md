# pdf2md

A .NET 9 console application that recursively converts PDF files to Markdown.

Text-only and mixed (scanned + text overlay) PDFs use the embedded text layer. PDFs that contain no text layer fall back to OCR with `eng+pol` Tesseract language data. Pages are processed independently and flushed to disk one at a time. See [Multi-GB handling](#multi-gb-handling) for the realistic size envelope — the page-by-page benefit applies *after* PdfPig has opened the file, and the open step is the binding constraint.

## Requirements

- .NET 9 SDK (built and tested with .NET 9; .NET 10 SDK also works)
- For OCR fallback: `eng.traineddata` and `pol.traineddata` from [tessdata_fast](https://github.com/tesseract-ocr/tessdata_fast) placed in a `tessdata/` directory. Without them, OCR is silently disabled and text-only PDFs still convert normally.

## Build

```powershell
dotnet build -c Release
```

The output assembly is `pdf2md.dll` under `bin/Release/net9.0/`.

## Usage

```
pdf2md <directory> [--overwrite|--skip] [--tessdata <path>] [--lang <list>]
```

| Argument        | Description                                                                             |
| --------------- | --------------------------------------------------------------------------------------- |
| `<directory>`   | Root directory to scan recursively for `*.pdf` files.                                   |
| `--skip`        | Default. Skip files where `<name>.pdf.md` already exists.                               |
| `--overwrite`   | Re-extract even when `<name>.pdf.md` already exists.                                    |
| `--tessdata`    | Path to Tesseract `tessdata` directory. Default: `<appdir>/tessdata`.                   |
| `--lang`        | Tesseract language list. Default: `eng+pol`.                                            |

Output is written as `<name>.pdf.md` next to each source `.pdf`, UTF-8 without BOM.

### Source-newer rule

If the `.pdf` is more than **5 minutes** newer than its `.pdf.md`, the file is re-extracted regardless of the `--skip` / `--overwrite` flag. This protects against stale Markdown when the source PDF has been updated.

### Examples

Convert everything under a folder, skipping files already extracted:

```powershell
dotnet run -- "D:\Documents"
```

Force re-extraction of everything:

```powershell
dotnet run -- "D:\Documents" --overwrite
```

Point at a shared Tesseract data folder and use English only:

```powershell
dotnet run -- "D:\Documents" --tessdata "C:\tessdata" --lang eng
```

## Behavior

| PDF page contains                  | Output                                                |
| ---------------------------------- | ----------------------------------------------------- |
| Native text                        | Text layer is extracted.                              |
| Text overlay on a scanned image    | Text overlay is extracted; the image is ignored.      |
| Inline images mixed with text      | Text is extracted; images are ignored.                |
| Only a scanned image (no text)     | Page is rendered and OCR'd with `eng+pol`.            |

The threshold for "no text" is fewer than 16 significant characters on the page. Pages below this threshold trigger the OCR fallback.

## Output format

Each page is emitted as a comment-delimited block:

```markdown
<!-- page 1 -->

Page 1 text content...

<!-- page 2 -->

Page 2 text content...
```

## Multi-GB handling

**What is truly streaming:**

- **Per-page text extraction** — once the PDF is open, `PdfPig` iterates pages lazily. Object-level structures are decoded on demand.
- **OCR rendering** — `Docnet.Core` (PDFium) renders one page at a time on the OCR fallback path.
- **Output writing** — `StreamWriter` flushes after every page. The `.pdf.md` grows page-by-page on disk; nothing is buffered in memory.

**The binding constraint — PdfPig's file-open path:**

`PdfDocument.Open(string filePath)` in `UglyToad.PdfPig` `1.7.0-custom-5` reads the entire PDF into a single `byte[]` (it calls `File.ReadAllBytes` internally). Consequences:

- **Hard ceiling around ~2 GB.** A single .NET `byte[]` is capped at `~2,147,483,591` elements regardless of `gcAllowVeryLargeObjects`. A 2.5 GB PDF throws `OutOfMemoryException` at open time, before the first page is touched.
- **Below 2 GB, memory pressure spikes.** A 1.8 GB PDF puts a 1.8 GB array on the LOH at open. On a memory-constrained machine that can fail or thrash even though per-page processing afterward is cheap.
- **The OCR path is fine for size** in isolation (Docnet mmaps via PDFium), but it runs *after* PdfPig has already opened the file, so the same ceiling applies.

**Realistic envelope as currently written**: comfortable up to a few hundred MB; works up to ~1.5 GB on a machine with 8 GB+ RAM; fails above ~2 GB. Removing this ceiling requires switching to PdfPig's `IInputBytes` streaming path.

## Limitations

**Text quality and structure**

- "Markdown" is generous — output is plain text with `<!-- page N -->` separators. No headings, lists, tables, bold/italic, or links are inferred from layout.
- `page.Text` returns content-stream order. Multi-column layouts can interleave columns. Tables come out as space-separated cell text with no structure.
- RTL languages (Arabic, Hebrew) do not get logical reading order applied.
- Custom font encodings or malformed CMaps can produce garbled output. PDF/A and exotic encodings have known edge cases in PdfPig.

**OCR heuristic**

- The "no text layer" threshold is fewer than 16 significant characters per page. A scanned page with an embedded watermark string longer than 16 chars is treated as text-bearing and skips OCR. A sparse cover page with just "Chapter 5" falls to OCR unnecessarily.
- Render resolution is fixed at 1700 × 2200. Pages outside that aspect ratio are stretched, which slightly degrades OCR accuracy. Very dense pages may need higher DPI for good results.
- No image preprocessing (deskew, denoise, contrast). OCR accuracy on old or skewed scans is middling.
- OCR is silently disabled if `tessdata/` is missing — scanned PDFs in that case produce empty pages with no error.

**Operational**

- Single-threaded. Files are processed sequentially; no parallelism flag.
- Encrypted / password-protected PDFs throw at open and are logged as `[fail]`.
- No incremental progress or resume mechanism beyond `--skip`. Long runs over thousands of files cannot be paused and resumed at a specific page within a file.
- Native binaries (Tesseract, PDFium via Docnet) ship for Windows / Linux / macOS but only the Windows path has been verified.

## Failure semantics for batch runs

Each file is wrapped in its own `try` / `catch` inside the directory loop. A failure on one PDF does not stop the others.

```csharp
try { converter.Convert(pdfPath, mdPath); converted++; }
catch (Exception ex) {
    Console.Error.WriteLine($"[fail] {pdfPath}: {ex.Message}");
    failed++;
}
```

- The loop continues to the next file.
- The failing file is logged to stderr as `[fail] <path>: <message>` and counted.
- The final summary line reports `total / converted / skipped / failed`.

**Output integrity on partial failures**

Each conversion writes to `<name>.pdf.md.tmp` and only renames to the final name on success. If `Convert` throws mid-file, the `.tmp` is renamed to `<name>.pdf.ex` (see below) and the previous `<name>.pdf.md` (if any) is left intact. A half-written `.pdf.md` is not possible.

**Error files (`<name>.pdf.ex`)**

When conversion fails, a `<name>.pdf.ex` file is written next to the source PDF containing:

- A `<!-- ===== CONVERSION FAILED ===== -->` marker.
- Source path, UTC timestamp, exception type, and message.
- Full stack trace under a `## Details` section.
- Any partial Markdown that was extracted before the failure, preserved above the error block.

On a subsequent successful conversion of the same PDF, any stale `<name>.pdf.ex` is removed automatically. Writing the error file is best-effort: if it fails (e.g., disk full, read-only directory), the program falls back to a minimal `<name>.pdf.ex` containing just the source path and exception, and the failure is still logged to stderr.

**What you don't get**

- No structured report file (JSON / CSV) listing all failures across the run — only stderr lines plus per-file `.pdf.ex` files.
- A systemic failure cause (e.g., missing tessdata, out-of-disk) is not detected and short-circuited; it will fail every file individually.

## Exit codes

| Code | Meaning                                            |
| ---- | -------------------------------------------------- |
| 0    | All files processed successfully.                  |
| 1    | Invalid arguments or directory not found.          |
| 2    | One or more files failed; others may have succeeded. |

## Dependencies

| Package                    | Purpose                                              |
| -------------------------- | ---------------------------------------------------- |
| `UglyToad.PdfPig`          | Page-by-page text extraction.                        |
| `Docnet.Core`              | PDF page rendering for the OCR fallback path.        |
| `Tesseract`                | OCR engine with Polish + English language data.      |
| `SixLabors.ImageSharp`     | BGRA → PNG encoding for Tesseract input.             |
