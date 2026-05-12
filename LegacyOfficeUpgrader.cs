using System.Diagnostics;

namespace Converter;

// Converts pre-2007 binary Office formats (.doc/.xls/.ppt) to their modern XML
// equivalents (.docx/.xlsx/.pptx) by invoking LibreOffice headless. The OpenXml
// converters then handle the modern files directly.
public static class LegacyOfficeUpgrader
{
    private const int ProcessTimeoutMs = 120_000;
    private static string? _resolvedSofficePath;
    private static bool _searchAttempted;

    public static bool IsLegacy(string extension) => extension is ".doc" or ".xls" or ".ppt";

    public static string TargetExtensionFor(string legacyExtension) => legacyExtension switch
    {
        ".doc" => "docx",
        ".xls" => "xlsx",
        ".ppt" => "pptx",
        _ => throw new ArgumentException($"Not a legacy Office extension: {legacyExtension}", nameof(legacyExtension))
    };

    public sealed record UpgradedFile(string Path) : IDisposable
    {
        public void Dispose()
        {
            try
            {
                string? dir = System.IO.Path.GetDirectoryName(Path);
                if (dir is not null && Directory.Exists(dir)
                    && System.IO.Path.GetFileName(dir).StartsWith("pdf2md_office_", StringComparison.Ordinal))
                {
                    Directory.Delete(dir, recursive: true);
                }
            }
            catch { /* swallow cleanup errors */ }
        }
    }

    public static UpgradedFile Upgrade(string sourcePath)
    {
        string ext = Path.GetExtension(sourcePath).ToLowerInvariant();
        string targetExt = TargetExtensionFor(ext);

        string soffice = ResolveSofficeOrThrow();
        string tempDir = Path.Combine(Path.GetTempPath(), "pdf2md_office_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            RunSoffice(soffice, sourcePath, targetExt, tempDir);

            string baseName = Path.GetFileNameWithoutExtension(sourcePath);
            string expected = Path.Combine(tempDir, baseName + "." + targetExt);
            if (File.Exists(expected))
                return new UpgradedFile(expected);

            // LibreOffice occasionally sanitises the filename; fall back to picking
            // the single produced file in the temp directory.
            string[] produced = Directory.GetFiles(tempDir);
            if (produced.Length == 1)
                return new UpgradedFile(produced[0]);

            throw new FileNotFoundException(
                $"LibreOffice did not produce the expected upgraded file at {expected}. " +
                $"Files in temp dir: {(produced.Length == 0 ? "<none>" : string.Join(", ", produced))}");
        }
        catch
        {
            try { Directory.Delete(tempDir, true); } catch { }
            throw;
        }
    }

    private static string ResolveSofficeOrThrow()
    {
        if (!_searchAttempted)
        {
            _resolvedSofficePath = LocateSoffice();
            _searchAttempted = true;
        }
        if (_resolvedSofficePath is null)
        {
            throw new InvalidOperationException(
                "LibreOffice (soffice) was not found. Install LibreOffice or put soffice on PATH " +
                "to enable conversion of legacy .doc / .xls / .ppt files.");
        }
        return _resolvedSofficePath;
    }

    private static string? LocateSoffice()
    {
        string exe = OperatingSystem.IsWindows() ? "soffice.exe" : "soffice";

        string? fromPath = Environment.GetEnvironmentVariable("PATH");
        if (!string.IsNullOrEmpty(fromPath))
        {
            foreach (string dir in fromPath.Split(Path.PathSeparator))
            {
                if (string.IsNullOrWhiteSpace(dir)) continue;
                try
                {
                    string candidate = Path.Combine(dir, exe);
                    if (File.Exists(candidate)) return candidate;
                }
                catch { /* ignore malformed path entries */ }
            }
        }

        string[] commonLocations = OperatingSystem.IsWindows()
            ? new[]
            {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            }
            : new[]
            {
                "/usr/bin/soffice",
                "/usr/local/bin/soffice",
                "/opt/libreoffice/program/soffice",
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            };

        foreach (string p in commonLocations)
            if (File.Exists(p)) return p;

        return null;
    }

    private static void RunSoffice(string soffice, string sourcePath, string targetExt, string outDir)
    {
        var psi = new ProcessStartInfo
        {
            FileName = soffice,
            ArgumentList =
            {
                "--headless",
                "--norestore",
                "--nologo",
                "--nofirststartwizard",
                "--convert-to", targetExt,
                "--outdir", outDir,
                sourcePath
            },
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        // Some LibreOffice builds refuse to run if a user profile is locked by another
        // headless invocation. Give each run its own profile dir to keep parallel-safe.
        psi.ArgumentList.Insert(0, "-env:UserInstallation=file:///"
            + outDir.Replace('\\', '/') + "/.lo-profile");

        using Process proc = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start LibreOffice (soffice).");

        if (!proc.WaitForExit(ProcessTimeoutMs))
        {
            try { proc.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException(
                $"LibreOffice conversion timed out (>{ProcessTimeoutMs / 1000}s) for {sourcePath}.");
        }

        if (proc.ExitCode != 0)
        {
            string stderr = proc.StandardError.ReadToEnd();
            string stdout = proc.StandardOutput.ReadToEnd();
            throw new InvalidOperationException(
                $"LibreOffice failed (exit {proc.ExitCode}) for {sourcePath}. " +
                $"stderr: {stderr.Trim()}; stdout: {stdout.Trim()}");
        }
    }
}
