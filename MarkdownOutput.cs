using System.Text;

namespace Converter;

// Shared scaffolding for "write <name>.<ext>.md atomically; on failure, leave a
// <name>.<ext>.ex error file alongside the source". Used by every converter so
// the failure semantics are identical regardless of source format.
public static class MarkdownOutput
{
    public delegate void WriteAction(StreamWriter output);

    // displaySourcePath is the original source path shown to the user (and used to
    // compute the .ex location). For legacy formats upgraded via LibreOffice, the
    // file actually being read may live in a temp directory, but the .ex must
    // appear next to the user's original .doc/.xls/.ppt.
    public static void WriteAtomic(string displaySourcePath, string outputPath, WriteAction write)
    {
        string tmpPath = outputPath + ".tmp";
        string exPath = displaySourcePath + ".ex";

        try
        {
            using (var output = new StreamWriter(tmpPath, append: false, new UTF8Encoding(false)))
            {
                write(output);
            }

            if (File.Exists(outputPath)) File.Delete(outputPath);
            File.Move(tmpPath, outputPath);
            TryDelete(exPath);
        }
        catch (Exception ex)
        {
            WriteErrorFile(exPath, tmpPath, displaySourcePath, ex);
            throw;
        }
    }

    public static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); }
        catch { /* swallow cleanup errors */ }
    }

    private static void WriteErrorFile(string exPath, string tmpPath, string sourcePath, Exception ex)
    {
        try
        {
            // Append the error block to whatever partial output we captured, then rename.
            // StreamWriter(append: true) creates the file if it doesn't exist, which covers
            // the case where Convert threw before any content was written.
            using (var writer = new StreamWriter(tmpPath, append: true, new UTF8Encoding(false)))
            {
                writer.WriteLine();
                writer.WriteLine("<!-- ===== CONVERSION FAILED ===== -->");
                writer.WriteLine();
                writer.WriteLine("# Conversion failed");
                writer.WriteLine();
                writer.WriteLine($"- Source:    `{sourcePath}`");
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
            try { File.WriteAllText(exPath, $"Source: {sourcePath}{Environment.NewLine}{ex}", new UTF8Encoding(false)); }
            catch { /* give up */ }
            TryDelete(tmpPath);
        }
    }
}
