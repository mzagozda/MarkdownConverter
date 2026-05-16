# pdf2md

A .NET 9 console application that recursively converts PDF and Office documents to Markdown.

| Format | Path |
| ------ | ---- |
| `.pdf` | Text-layer extraction with geometry-aware region separation; OCR (`eng+pol`) fallback for scan-only pages. |
| `.docx`, `.xlsx`, `.pptx` | Direct extraction from the OpenXml package — preserves headings, lists, tables, sheets, slide titles, bold/italic, hyperlinks, and speaker notes. |
| `.doc`, `.xls`, `.ppt` | Legacy binary formats are upgraded to their modern XML equivalent in a temp directory via LibreOffice headless, then run through the OpenXml path. |
| `.eml` | RFC 5322 / MIME messages parsed via MimeKit. Headers (From / To / Cc / Bcc / Date / Subject) become a bullet list; the HTML body (preferred) is converted to Markdown via ReverseMarkdown, otherwise the plain-text body is emitted as-is. Supported attachments (`.eml`, `.pdf`, `.docx`/`.doc`, `.xlsx`/`.xls`, `.pptx`/`.ppt`) are extracted to a temp dir, run through the matching converter, and inlined under an `## Attachment: <name>` heading. Nested `message/rfc822` parts and `.eml` attachments are handled recursively. |

Each file is converted independently to a `<name>.<ext>.md` next to the source.

## Requirements

- .NET 9 SDK (built and tested with .NET 9; .NET 10 SDK also works).
- **For OCR fallback (PDF only)**: `eng.traineddata` and `pol.traineddata` from [tessdata_best](https://github.com/tesseract-ocr/tessdata_best) placed in a `tessdata/` directory. The build copies it to the output folder automatically. Without it, OCR is silently disabled and text-only PDFs still convert normally.
- **For legacy `.doc` / `.xls` / `.ppt`**: [LibreOffice](https://www.libreoffice.org/) must be installed and either on `PATH` or at its default install location (`C:\Program Files\LibreOffice\program\soffice.exe` on Windows, `/usr/bin/soffice` or `/Applications/LibreOffice.app/...` elsewhere). Modern Office formats do not need LibreOffice.

## Build

```powershell
dotnet build -c Release
```

The output assembly is `pdf2md.dll` under `bin/Release/net9.0/`.

## Usage

```
pdf2md <directory> [--overwrite|--skip] [--tessdata <path>] [--lang <list>]
```

| Argument        | Description                                                                              |
| --------------- | ---------------------------------------------------------------------------------------- |
| `<directory>`   | Root directory to scan recursively for supported files.                                  |
| `--skip`        | Default. Skip files where `<name>.<ext>.md` already exists.                              |
| `--overwrite`   | Re-extract even when `<name>.<ext>.md` already exists.                                   |
| `--tessdata`    | Path to Tesseract `tessdata` directory. Default: `<appdir>/tessdata`. (PDF/OCR only.)    |
| `--lang`        | Tesseract language list. Default: `eng+pol`. (PDF/OCR only.)                             |

Output is written as `<name>.<ext>.md` next to each source file, UTF-8 without BOM.

### Source-newer rule

If a source file is more than **5 minutes** newer than its `.md`, it is re-extracted regardless of the `--skip` / `--overwrite` flag. This protects against stale Markdown when the source has been updated.

### PDF-vs-Office deduplication

When the same base name has both a PDF and an Office sibling (`report.pdf` *and* `report.docx`, for example), the PDF is treated as a derived artifact of the Office source. In **`--skip`** mode, if the PDF's own `<name>.pdf.md` is not already present, the PDF is skipped in favour of the Office file — that produces richer Markdown (real headings, lists, tables) than re-deriving structure from the PDF.

| Configuration                              | `report.docx`                | `report.pdf`                 |
| ------------------------------------------ | ---------------------------- | ---------------------------- |
| `--skip`, neither md exists                | convert → `report.docx.md`   | **skip** (sibling preferred) |
| `--skip`, only `report.docx.md` exists     | skip (md exists)             | **skip** (sibling preferred) |
| `--skip`, only `report.pdf.md` exists      | convert → `report.docx.md`   | skip (md exists)             |
| `--skip`, both mds exist                   | skip (md exists)             | skip (md exists)             |
| `--overwrite`                              | convert                      | convert                      |

The dedup applies only to PDF↔Office pairs. Two Office files with the same base name (e.g., `report.docx` and `report.xlsx`) are both converted independently.

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

### PDF (`.pdf`)

| PDF page contains                  | Output                                                |
| ---------------------------------- | ----------------------------------------------------- |
| Native text                        | Text layer is extracted.                              |
| Text overlay on a scanned image    | Text overlay is extracted; the image is ignored.      |
| Inline images mixed with text      | Text is extracted; images are ignored.                |
| Only a scanned image (no text)     | Page is rendered and OCR'd with `eng+pol`.            |

The threshold for "no text" is fewer than 16 significant characters on the page. Pages below this threshold trigger the OCR fallback.

### Word (`.docx`, `.doc`)

- Heading styles `Heading1`–`Heading6` → `#` to `######`.
- Paragraphs flow as plain text. Bold / italic runs → `**…**` / `*…*` / `***…***`.
- Hyperlinks → `[text](url)`.
- Lists → `- ` markers indented by 2 spaces per nesting level. Numbered-vs-bullet distinction is not decoded; both render as `- `.
- Tables → standard Markdown pipe tables, with `<br>` joining multi-paragraph cells.
- `.doc` (legacy binary) is upgraded to `.docx` via LibreOffice headless first, then processed the same way.

### Excel (`.xlsx`, `.xls`)

- Each worksheet becomes a `## Sheet name` section.
- Cells emit raw values (numbers as their stored representation, strings via the shared string table, booleans as `TRUE` / `FALSE`). Number-format strings are not applied — raw values are locale-independent and reproducible.
- The first row is treated as the table header; the rest are rows. Empty sheets emit `_(empty sheet)_`.
- `.xls` is upgraded to `.xlsx` via LibreOffice headless first.

### PowerPoint (`.pptx`, `.ppt`)

- Each slide is emitted as a `<!-- slide N -->` block.
- Title and subtitle placeholders → `### Heading`.
- Body text → bulleted list, indented by paragraph level.
- Tables inside slides → Markdown pipe tables.
- Speaker notes (if present) → blockquote prefixed with `> **Speaker notes:**`.
- `.ppt` is upgraded to `.pptx` via LibreOffice headless first.

## Output format

PDF pages are emitted as comment-delimited blocks, one per page. PowerPoint slides use the same convention with `<!-- slide N -->`. Word and Excel emit semantic Markdown structure directly.

```markdown
<!-- page 1 -->

Page 1 text content...

<!-- page 2 -->

Page 2 text content...
```

### Region separation

Words extracted from clearly different regions of a page (different blocks, columns, table cells) are joined by ` - ` instead of being concatenated. Words that are merely line-wrapped within a paragraph are kept attached with a normal space, so legitimate intra-paragraph breaks are not artificially split.

**Text-layer path (PdfPig)** — bounding-box geometry is compared for each pair of consecutive words. Median word height on the page is used as a line-height proxy. ` - ` is inserted only when the geometry is unambiguous:

| Signal between consecutive words                                 | Result      |
| ---------------------------------------------------------------- | ----------- |
| Next word sits *above* the previous one                          | ` - `       |
| Same baseline, next word is *left of* the previous one           | ` - `       |
| Same baseline, horizontal gap > 5 × line height                  | ` - `       |
| Vertical gap to next line > 1.8 × line height                    | ` - `       |
| Anything else (normal line wrap, normal inter-word space)        | plain space |

**OCR path (Tesseract)** — Tesseract's own block segmentation is used as the deterministic boundary. Each block's text is emitted as-is (internal line wraps preserved with normal spacing) and blocks are joined with ` - `.

This is intentionally conservative — region detection is not always possible. The geometry path can miss subtle cell boundaries that look like normal inter-word spacing, and Tesseract's block segmenter occasionally splits a sentence at a column edge. The trade-off is biased toward false negatives (missed separators) over false positives (sentences artificially split mid-flow).

## Multi-GB handling

**What is truly streaming:**

- **Per-page text extraction** — once the PDF is open, `PdfPig` iterates pages lazily. Object-level structures are decoded on demand.
- **OCR rendering** — `Docnet.Core` (PDFium) renders one page at a time on the OCR fallback path.
- **Output writing** — `StreamWriter` flushes after every page. The `.<ext>.md` grows page-by-page on disk; nothing is buffered in memory.

**The binding constraint — PdfPig's file-open path:**

`PdfDocument.Open(string filePath)` in `UglyToad.PdfPig` `1.7.0-custom-5` reads the entire PDF into a single `byte[]` (it calls `File.ReadAllBytes` internally). Consequences:

- **Hard ceiling around ~2 GB.** A single .NET `byte[]` is capped at `~2,147,483,591` elements regardless of `gcAllowVeryLargeObjects`. A 2.5 GB PDF throws `OutOfMemoryException` at open time, before the first page is touched.
- **Below 2 GB, memory pressure spikes.** A 1.8 GB PDF puts a 1.8 GB array on the LOH at open. On a memory-constrained machine that can fail or thrash even though per-page processing afterward is cheap.
- **The OCR path is fine for size** in isolation (Docnet mmaps via PDFium), but it runs *after* PdfPig has already opened the file, so the same ceiling applies.

**Realistic envelope as currently written**: comfortable up to a few hundred MB; works up to ~1.5 GB on a machine with 8 GB+ RAM; fails above ~2 GB. Removing this ceiling requires switching to PdfPig's `IInputBytes` streaming path.

## Limitations

**PDF text quality and structure**

- PDF output is plain text with `<!-- page N -->` separators and the geometry-driven region separator ` - `. Headings, lists, bold/italic, and link structure cannot be recovered from PDF.
- RTL languages (Arabic, Hebrew) do not get logical reading order applied.
- Custom font encodings or malformed CMaps can produce garbled output. PDF/A and exotic encodings have known edge cases in PdfPig.

**OCR heuristic**

- The "no text layer" threshold is fewer than 16 significant characters per page. A scanned page with an embedded watermark string longer than 16 chars is treated as text-bearing and skips OCR. A sparse cover page with just "Chapter 5" falls to OCR unnecessarily.
- Render resolution is fixed at 1700 × 2200. Pages outside that aspect ratio are stretched, which slightly degrades OCR accuracy. Very dense pages may need higher DPI for good results.
- No image preprocessing (deskew, denoise, contrast). OCR accuracy on old or skewed scans is middling.
- OCR is silently disabled if `tessdata/` is missing — scanned PDFs in that case produce empty pages with no error.

**Office formats**

- DOCX list extraction does not decode the `numbering.xml` part — bullets and numbered lists both render as `- `. Order is preserved by document order but item *numbering* (1, 2, 3, a, b, c, …) is dropped.
- DOCX nested tables, footnotes, endnotes, comments, embedded images, drawings, equations, and field codes are not extracted.
- XLSX cell values are emitted raw without applying number-format strings, currency symbols, percentage scaling, or date formatting. A cell displayed as `12,5%` may appear as `0.125` in Markdown.
- XLSX merged cells are not detected — each underlying cell is emitted in its own column. Charts, conditional formatting, and pivot tables are ignored.
- PPTX SmartArt, embedded charts, and master-slide content are not extracted. Only top-level shapes, tables in graphic frames, and speaker notes are.
- Headers / footers from Word documents and slide-master content from PowerPoint are not extracted.

**Email (`.eml`)**

- Inline images (`cid:` references) inside an HTML body are kept as their original `<img>` tag text through ReverseMarkdown; the underlying image parts are not extracted.
- S/MIME and PGP encrypted or signed parts are not decrypted — only the cleartext outer structure is read. Signature wrapper metadata is ignored.
- DKIM / ARC headers and other technical headers are not emitted; only From / To / Cc / Bcc / Date / Subject are surfaced.
- Attachments with extensions outside the supported set (`.eml`, `.pdf`, `.docx`, `.doc`, `.xlsx`, `.xls`, `.pptx`, `.ppt`) are listed by filename but their contents are not extracted. Images, archives (`.zip`, `.7z`), text/CSV, and other formats are not converted.
- Nested `message/rfc822` and `.eml` attachments are followed recursively up to 8 levels deep; deeper nesting is truncated with a notice.
- A failure converting any single attachment is logged to stderr and inlined as an error block under that attachment's heading — it never fails the parent `.eml.md` output.

**Legacy binary formats**

- `.doc` / `.xls` / `.ppt` require LibreOffice to be installed and reachable. Each upgrade spawns a `soffice` process with a 2-minute timeout per file. Each conversion uses a fresh user-profile directory, so parallel safety is supported but parallelism is not currently exposed.
- LibreOffice's `.doc` → `.docx` conversion is not 100 % lossless for documents with complex formatting, custom fonts, or unusual OLE objects.

**Operational**

- Single-threaded. Files are processed sequentially; no parallelism flag.
- Encrypted / password-protected files throw at open and are logged as `[fail]`.
- No incremental progress or resume mechanism beyond `--skip`. Long runs over thousands of files cannot be paused and resumed mid-file.
- Native binaries (Tesseract, PDFium via Docnet) ship for Windows / Linux / macOS but only the Windows path has been verified end-to-end.

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

Each conversion writes to `<name>.pdf.md.tmp` and only renames to the final name on success. If `Convert` throws mid-file, the `.tmp` is renamed to `<name>.pdf.ex` (see below) and the previous `<name>.pdf.md` (if any) is left intact. A half-written `.<ext>.md` is not possible.

**Error files (`<name>.pdf.ex`)**

When conversion fails, a `<name>.pdf.ex` file is written next to the source PDF containing:

- A `<!-- ===== CONVERSION FAILED ===== -->` marker.
- Source path, UTC timestamp, exception type, and message.
- Full stack trace under a `## Details` section.
- Any partial Markdown that was extracted before the failure, preserved above the error block.

On a subsequent successful conversion of the same PDF, any stale `<name>.pdf.ex` is removed automatically. Writing the error file is best-effort: if it fails (e.g., disk full, read-only directory), the program falls back to a minimal `<name>.pdf.ex` containing just the source path and exception, and the failure is still logged to stderr.

**What you don't get**

- No structured report file (JSON / CSV) listing all failures across the run — only stderr lines plus per-file `.<ext>.ex` files.
- A systemic failure cause (e.g., missing tessdata, out-of-disk) is not detected and short-circuited; it will fail every file individually.

## Exit codes

| Code | Meaning                                            |
| ---- | -------------------------------------------------- |
| 0    | All files processed successfully.                  |
| 1    | Invalid arguments or directory not found.          |
| 2    | One or more files failed; others may have succeeded. |

## Dependencies

| Package                    | Purpose                                                                  |
| -------------------------- | ------------------------------------------------------------------------ |
| `UglyToad.PdfPig`          | Page-by-page text extraction from PDFs.                                  |
| `Docnet.Core`              | PDF page rendering (PDFium) for the OCR fallback path.                   |
| `Tesseract`                | OCR engine with Polish + English language data.                          |
| `SixLabors.ImageSharp`     | BGRA → PNG encoding for Tesseract input.                                 |
| `DocumentFormat.OpenXml`   | Native reading of `.docx` / `.xlsx` / `.pptx` packages.                  |
| LibreOffice (external)     | Upgrades legacy `.doc` / `.xls` / `.ppt` to modern OpenXml in a temp dir. |
