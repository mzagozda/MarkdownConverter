using System.Diagnostics;
using System.Text.RegularExpressions;

namespace Converter;

// Converts pre-2007 binary Office formats (.doc/.xls/.ppt) to their modern XML
// equivalents (.docx/.xlsx/.pptx) by invoking LibreOffice headless. The OpenXml
// converters then handle the modern files directly.
public static class LegacyOfficeUpgrader
{
    private const int ProcessTimeoutMs = 120_000;
    private const int VersionProbeTimeoutMs = 10_000;

    // Minimum LibreOffice release we support. 7.6 was the final 7.x line and is the
    // earliest version where `--convert-to docx/xlsx/pptx` is reliable for the binary
    // legacy formats we care about. Earlier releases work for many documents but have
    // known filter regressions; we surface a warning rather than silently using them.
    public static readonly Version MinimumLibreOfficeVersion = new(7, 6, 0);

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

    public enum ProbeStatus
    {
        Ok,
        NotFound,
        VersionUnknown,
        TooOld
    }

    public sealed record ProbeResult(ProbeStatus Status, string? Path, Version? Version, string Message);

    // Examine the local LibreOffice install once and report what we found. This is
    // safe to call before any file is processed and never throws.
    public static ProbeResult Probe()
    {
        string? path = TryResolveSoffice();
        if (path is null)
        {
            return new ProbeResult(ProbeStatus.NotFound, null, null,
                $"LibreOffice (soffice) not found on PATH or in standard install locations. " +
                $"Install version {MinimumLibreOfficeVersion} or newer to enable legacy .doc / .xls / .ppt conversion.");
        }

        Version? version = ReadSofficeVersion(path);

        if (version is null)
        {
            return new ProbeResult(ProbeStatus.VersionUnknown, path, null,
                $"LibreOffice found at {path} but `soffice --version` did not return a parseable version string. " +
                $"Minimum supported version is {MinimumLibreOfficeVersion}.");
        }

        if (version < MinimumLibreOfficeVersion)
        {
            return new ProbeResult(ProbeStatus.TooOld, path, version,
                $"LibreOffice {version} at {path} is older than the supported minimum ({MinimumLibreOfficeVersion}). " +
                $"Legacy .doc / .xls / .ppt conversion may fail or produce incorrect output. Please upgrade.");
        }

        return new ProbeResult(ProbeStatus.Ok, path, version,
            $"LibreOffice {version} detected at {path}.");
    }

    private static string? TryResolveSoffice()
    {
        if (!_searchAttempted)
        {
            _resolvedSofficePath = LocateSoffice();
            _searchAttempted = true;
        }
        return _resolvedSofficePath;
    }

    private static string ResolveSofficeOrThrow()
    {
        string? path = TryResolveSoffice();
        if (path is null)
        {
            throw new InvalidOperationException(
                $"LibreOffice (soffice) was not found. Install LibreOffice {MinimumLibreOfficeVersion} or newer, " +
                $"or put `soffice` on PATH, to enable conversion of legacy .doc / .xls / .ppt files.");
        }
        return path;
    }

    private static Version? ReadSofficeVersion(string sofficePath)
    {
        // On Windows, `soffice --version` does not reliably print to stdout — the .exe
        // detaches into the GUI subsystem and the .com console-wrapper exits with code 53
        // on `--version`. The PE file's VersionInfo, however, carries the LibreOffice
        // version reliably for any Windows install, so we try it first.
        Version? fromMetadata = TryReadVersionFromFileMetadata(sofficePath);
        if (fromMetadata is not null) return fromMetadata;

        return TryReadVersionFromCommandLine(sofficePath);
    }

    private static Version? TryReadVersionFromFileMetadata(string sofficePath)
    {
        try
        {
            FileVersionInfo info = FileVersionInfo.GetVersionInfo(sofficePath);
            string? raw = info.ProductVersion ?? info.FileVersion;
            if (string.IsNullOrWhiteSpace(raw)) return null;
            Match m = Regex.Match(raw, @"(\d+)\.(\d+)(?:\.(\d+))?(?:\.(\d+))?");
            if (!m.Success) return null;
            return BuildVersion(m);
        }
        catch
        {
            return null;
        }
    }

    private static Version? TryReadVersionFromCommandLine(string sofficePath)
    {
        var psi = new ProcessStartInfo
        {
            FileName = sofficePath,
            ArgumentList = { "--version", "--headless", "--norestore", "--nologo", "--nofirststartwizard" },
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        Process? proc;
        try { proc = Process.Start(psi); }
        catch { return null; }
        if (proc is null) return null;

        using (proc)
        {
            if (!proc.WaitForExit(VersionProbeTimeoutMs))
            {
                try { proc.Kill(entireProcessTree: true); } catch { }
                return null;
            }

            string output = proc.StandardOutput.ReadToEnd() + "\n" + proc.StandardError.ReadToEnd();
            Match m = Regex.Match(output, @"LibreOffice\s+(\d+)\.(\d+)(?:\.(\d+))?(?:\.(\d+))?");
            return m.Success ? BuildVersion(m) : null;
        }
    }

    private static Version BuildVersion(Match m)
    {
        int major = int.Parse(m.Groups[1].Value);
        int minor = int.Parse(m.Groups[2].Value);
        int build = m.Groups[3].Success ? int.Parse(m.Groups[3].Value) : 0;
        int revision = m.Groups[4].Success ? int.Parse(m.Groups[4].Value) : 0;
        return new Version(major, minor, build, revision);
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
