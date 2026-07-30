using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text.Json;

namespace JinoSupporter.Web.Services;

/// <summary>Out-of-process driver for Excel conversion and optional folder picking.
/// Conversion is always delegated to the Excel DRM Python CLI. The legacy
/// <c>JinoSupporter.ExcelHelper</c> exe is kept only for server-side folder dialogs.</summary>
public sealed class ExcelHelperRunner
{
    private static readonly string[] ExcelExtensions = [".xlsx", ".xlsm", ".xlsb", ".xls"];

    /// <summary>Path of the optional folder-picker helper exe.</summary>
    public string HelperExePath { get; }

    public string ExcelDuplicatorRoot { get; }
    public string PythonExePath { get; }
    public string PythonScriptPath { get; }

    public ExcelHelperRunner(IHostEnvironment env)
    {
        HelperExePath = Path.Combine(AppContext.BaseDirectory, "JinoSupporter.ExcelHelper.exe");

        ExcelDuplicatorRoot = ResolveExcelDuplicatorRoot(env.ContentRootPath);
        PythonScriptPath = Environment.GetEnvironmentVariable("EXCEL_DRM_CLI")
            ?? Path.Combine(ExcelDuplicatorRoot, "excel_drm_cli.py");
        PythonExePath = ResolvePythonExePath(ExcelDuplicatorRoot);
    }

    public bool HelperExists => File.Exists(HelperExePath);
    public bool FolderPickerExists => HelperExists;
    public bool ConverterExists => PythonCommandExists(PythonExePath) && File.Exists(PythonScriptPath);

    public string ConverterStatus =>
        ConverterExists
            ? $"{PythonExePath} -> {PythonScriptPath}"
            : $"Python converter not found. python={PythonExePath}, script={PythonScriptPath}";

    /// <summary>Open a folder dialog (server-side) and return the selected path.
    /// Returns null on cancel or when the helper is missing/non-zero.</summary>
    public async Task<string?> PickFolderAsync(string? initial = null,
                                               CancellationToken ct = default)
    {
        if (!HelperExists) return null;

        var args = new List<string> { "pick-folder" };
        if (!string.IsNullOrWhiteSpace(initial))
        {
            args.Add("--initial");
            args.Add(initial);
        }

        string? folder = null;
        await foreach (var evt in StreamAsync(args, ct))
        {
            if (evt.TryGetProperty("kind", out var k) && k.GetString() == "folder")
                folder = evt.TryGetProperty("path", out var p) ? p.GetString() : null;
        }
        return folder;
    }

    /// <summary>One streamed event from the helper. <see cref="Raw"/> holds the
    /// underlying JsonElement so the caller can read kind-specific fields without
    /// this wrapper having to model every payload variant.</summary>
    public sealed record HelperEvent(string Kind, JsonElement Raw);

    /// <summary>Scan <paramref name="source"/> and convert each Excel file via
    /// ExcelDuplicator's Python CLI. Outputs are written to
    /// <c>{dest}/drm_clean/{name}_clean.xlsx</c>.</summary>
    public async IAsyncEnumerable<HelperEvent> CleanAsync(
        string source, string dest,
        bool verbose = false, bool keepFormats = false,
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        if (!ConverterExists)
        {
            yield return NewEvent("fatal", new { message = ConverterStatus });
            yield break;
        }

        List<string> files = [];
        HelperEvent? setupFault = null;
        try
        {
            files = ResolveExcelFiles(source);
        }
        catch (Exception ex)
        {
            setupFault = NewEvent("fatal", new { message = ex.Message });
        }

        if (setupFault is not null)
        {
            yield return setupFault;
            yield break;
        }

        yield return NewEvent("scan", new
        {
            source,
            count = files.Count,
            converter = PythonScriptPath,
        });

        string cleanDir = "";
        try
        {
            cleanDir = Path.Combine(dest, "drm_clean");
            Directory.CreateDirectory(cleanDir);
        }
        catch (Exception ex)
        {
            setupFault = NewEvent("fatal", new { message = $"Output folder create failed: {ex.Message}" });
        }

        if (setupFault is not null)
        {
            yield return setupFault;
            yield break;
        }

        var cleaner = new ExcelDrm.ExcelDrmCleaner(PythonExePath, PythonScriptPath)
        {
            DefaultMode = ExcelDrm.ConvertMode.Clipboard,
        };

        int ok = 0;
        int fail = 0;

        for (int i = 0; i < files.Count; i++)
        {
            ct.ThrowIfCancellationRequested();

            string input = files[i];
            string output = BuildOutputPath(cleanDir, input);

            yield return NewEvent("progress", new
            {
                current = i + 1,
                total = files.Count,
                file = Path.GetFileName(input),
                source = input,
                dest = output,
            });

            ExcelDrm.ConvertResult result;
            try
            {
                // keepFormats is retained for callers, but conversion must use
                // ExcelDuplicator's Python program. Clipboard is the Python CLI's
                // default full-fidelity mode.
                result = await cleaner.ConvertAsync(input, output, ExcelDrm.ConvertMode.Clipboard, ct);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                result = new ExcelDrm.ConvertResult
                {
                    Success = false,
                    Input = input,
                    Output = output,
                    Error = ex.Message,
                    ExitCode = -1,
                };
            }

            foreach (string line in SplitLogLines(result.StdErr))
            {
                if (verbose || !result.Success)
                    yield return NewEvent("log", new { message = line });
            }

            if (result.Success) ok++;
            else fail++;

            yield return NewEvent("result", new
            {
                success = result.Success,
                source = result.Input,
                dest = result.Output ?? output,
                error = result.Error,
                elapsed = result.ElapsedSeconds,
                exitCode = result.ExitCode,
                strategy = "python:excel_drm_cli.py:clipboard",
            });
        }

        yield return NewEvent("done", new { ok, fail });
    }

    /// <summary>Convert an explicit list of Excel files and write outputs
    /// directly into <paramref name="dest"/> without creating a drm_clean child
    /// folder. This is used by the web drag/drop converter page.</summary>
    public async IAsyncEnumerable<HelperEvent> CleanFilesToDestinationRootAsync(
        IReadOnlyList<string> sourceFiles,
        string dest,
        bool verbose = false,
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        if (!ConverterExists)
        {
            yield return NewEvent("fatal", new { message = ConverterStatus });
            yield break;
        }

        List<string> files = [];
        HelperEvent? setupFault = null;
        try
        {
            files = sourceFiles
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(Path.GetFullPath)
                .Where(path => File.Exists(path) && IsExcelFile(path))
                .ToList();
        }
        catch (Exception ex)
        {
            setupFault = NewEvent("fatal", new { message = ex.Message });
        }

        if (setupFault is not null)
        {
            yield return setupFault;
            yield break;
        }

        yield return NewEvent("scan", new
        {
            source = "selected files",
            count = files.Count,
            converter = PythonScriptPath,
        });

        try
        {
            Directory.CreateDirectory(dest);
        }
        catch (Exception ex)
        {
            setupFault = NewEvent("fatal", new { message = $"Output folder create failed: {ex.Message}" });
        }

        if (setupFault is not null)
        {
            yield return setupFault;
            yield break;
        }

        var cleaner = new ExcelDrm.ExcelDrmCleaner(PythonExePath, PythonScriptPath)
        {
            DefaultMode = ExcelDrm.ConvertMode.Clipboard,
        };

        int ok = 0;
        int fail = 0;
        var usedOutputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < files.Count; i++)
        {
            ct.ThrowIfCancellationRequested();

            string input = files[i];
            string output = BuildUniqueOutputPath(dest, input, usedOutputs);

            yield return NewEvent("progress", new
            {
                current = i + 1,
                total = files.Count,
                file = Path.GetFileName(input),
                source = input,
                dest = output,
            });

            ExcelDrm.ConvertResult result;
            try
            {
                result = await cleaner.ConvertAsync(input, output, ExcelDrm.ConvertMode.Clipboard, ct);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                result = new ExcelDrm.ConvertResult
                {
                    Success = false,
                    Input = input,
                    Output = output,
                    Error = ex.Message,
                    ExitCode = -1,
                };
            }

            foreach (string line in SplitLogLines(result.StdErr))
            {
                if (verbose || !result.Success)
                    yield return NewEvent("log", new { message = line });
            }

            if (result.Success) ok++;
            else fail++;

            yield return NewEvent("result", new
            {
                success = result.Success,
                source = result.Input,
                dest = result.Output ?? output,
                error = result.Error,
                elapsed = result.ElapsedSeconds,
                exitCode = result.ExitCode,
                strategy = "python:excel_drm_cli.py:clipboard",
            });
        }

        yield return NewEvent("done", new { ok, fail });
    }

    /// <summary>Convenience: clean a single uploaded file. Writes to a temp
    /// folder, runs the Python converter, and returns the path of the cleaned output
    /// (<c>{tempDir}/drm_clean/{name}_clean.xlsx</c>) — or <c>null</c> if the
    /// strategy chain failed. The caller is responsible for deleting the temp
    /// folder when done.</summary>
    public async Task<(string? cleanedPath, string tempDir, List<string> log)> CleanSingleAsync(
        string sourceFile, bool verbose = false, bool keepFormats = true,
        CancellationToken ct = default)
    {
        // Default keepFormats=true for the single-file path: this is invoked
        // from the Input Data (Test) page where one file at a time is fine to
        // run with full per-cell format copying (fill colors, font, borders).
        // ExtractTsv reads those colors back to annotate the TSV for the AI.
        string tempDir = Path.Combine(Path.GetTempPath(),
            "jinosupp-input-test", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        var log = new List<string>();
        string? cleanedPath = null;
        await foreach (var evt in CleanAsync(sourceFile, tempDir, verbose, keepFormats, ct))
        {
            switch (evt.Kind)
            {
                case "log":
                    if (evt.Raw.TryGetProperty("message", out var m))
                        log.Add(m.GetString() ?? "");
                    break;
                case "result":
                    bool ok = evt.Raw.TryGetProperty("success", out var s) && s.GetBoolean();
                    if (ok && evt.Raw.TryGetProperty("dest", out var d))
                        cleanedPath = d.GetString();
                    else if (!ok && evt.Raw.TryGetProperty("error", out var er))
                        log.Add("✗ " + (er.GetString() ?? ""));
                    break;
                case "fatal":
                case "stderr":
                    if (evt.Raw.TryGetProperty("message", out var fm))
                        log.Add("[fatal] " + (fm.GetString() ?? ""));
                    break;
            }
        }
        return (cleanedPath, tempDir, log);
    }

    private static HelperEvent NewEvent(string kind, object payload)
    {
        using var doc = JsonDocument.Parse(JsonSerializer.Serialize(payload));
        return new HelperEvent(kind, doc.RootElement.Clone());
    }

    private static string ResolveExcelDuplicatorRoot(string contentRootPath)
    {
        string? configured = Environment.GetEnvironmentVariable("EXCEL_DUPLICATOR_ROOT");
        if (!string.IsNullOrWhiteSpace(configured) && Directory.Exists(configured))
            return configured;

        string bundled = Path.GetFullPath(Path.Combine(contentRootPath, "..", "External", "ExcelDrmCli"));
        if (File.Exists(Path.Combine(bundled, "excel_drm_cli.py")))
            return bundled;

        string sibling = Path.GetFullPath(Path.Combine(contentRootPath, "..", "..", "ExcelDuplicator"));
        if (File.Exists(Path.Combine(sibling, "excel_drm_cli.py")))
            return sibling;

        return bundled;
    }

    private static string ResolvePythonExePath(string converterRoot)
    {
        string? configured = Environment.GetEnvironmentVariable("EXCEL_DRM_PYTHON");
        if (!string.IsNullOrWhiteSpace(configured))
            return configured;

        string venvPython = Path.Combine(converterRoot, ".venv", "Scripts", "python.exe");
        if (File.Exists(venvPython))
            return venvPython;

        return "python";
    }

    private static bool PythonCommandExists(string pythonExe)
    {
        if (File.Exists(pythonExe))
            return true;

        return !string.IsNullOrWhiteSpace(pythonExe)
               && !pythonExe.Contains(Path.DirectorySeparatorChar)
               && !pythonExe.Contains(Path.AltDirectorySeparatorChar);
    }

    private static List<string> ResolveExcelFiles(string source)
    {
        if (string.IsNullOrWhiteSpace(source))
            throw new ArgumentException("Source path is empty.");

        if (File.Exists(source))
            return IsExcelFile(source) ? [Path.GetFullPath(source)] : [];

        if (!Directory.Exists(source))
            throw new DirectoryNotFoundException($"Source not found: {source}");

        return Directory.EnumerateFiles(source, "*.*", SearchOption.AllDirectories)
            .Where(IsExcelFile)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .Select(Path.GetFullPath)
            .ToList();
    }

    private static bool IsExcelFile(string path)
    {
        string name = Path.GetFileName(path);
        if (name.StartsWith("~$", StringComparison.Ordinal)) return false;

        string ext = Path.GetExtension(path);
        return ExcelExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase);
    }

    private static string BuildOutputPath(string cleanDir, string input)
    {
        string stem = Path.GetFileNameWithoutExtension(input);
        if (string.IsNullOrWhiteSpace(stem))
            stem = "workbook";

        return Path.Combine(cleanDir, $"{stem}_clean.xlsx");
    }

    private static string BuildUniqueOutputPath(string outputDir, string input, ISet<string> usedOutputs)
    {
        string output = BuildOutputPath(outputDir, input);
        if (usedOutputs.Add(output)) return output;

        string stem = Path.GetFileNameWithoutExtension(input);
        if (string.IsNullOrWhiteSpace(stem))
            stem = "workbook";

        int index = 2;
        do
        {
            output = Path.Combine(outputDir, $"{stem}_clean_{index++}.xlsx");
        }
        while (!usedOutputs.Add(output));

        return output;
    }

    private static IEnumerable<string> SplitLogLines(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            yield break;

        using var reader = new StringReader(text);
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            line = line.TrimEnd();
            if (line.Length > 0)
                yield return line;
        }
    }

    /// <summary>Spawn the helper, parse stdout line-by-line as JSON, yield each.
    /// Stderr is concatenated into a final synthetic <c>{"kind":"stderr",...}</c>
    /// event when the process exits with a non-zero code so callers can surface
    /// failures without out-of-band logging.</summary>
    private async IAsyncEnumerable<JsonElement> StreamAsync(
        IReadOnlyList<string> args,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct)
    {
        var psi = new ProcessStartInfo
        {
            FileName               = HelperExePath,
            UseShellExecute        = false,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            CreateNoWindow         = true,
            StandardOutputEncoding = System.Text.Encoding.UTF8,
            StandardErrorEncoding  = System.Text.Encoding.UTF8,
        };
        foreach (var a in args) psi.ArgumentList.Add(a);

        var proc = new Process { StartInfo = psi };
        proc.Start();
        // Assign before any awaits so even an immediate crash of the web exe
        // takes the helper with it. Job Object kills the helper *and* any
        // Excel.exe it spawned, so file locks on the helper binary release
        // the moment the parent process dies.
        ChildProcessJob.Assign(proc);

        // Cancellation must NOT wait for the next ReadLineAsync to wake up —
        // an Excel COM call can block stdout for tens of seconds. Killing the
        // process tree closes stdout, which makes the read loop exit the same
        // turn the user clicks Stop.
        using var cancelReg = ct.Register(() =>
        {
            try { if (!proc.HasExited) proc.Kill(entireProcessTree: true); }
            catch { }
        });

        try
        {
            var stderr = new System.Text.StringBuilder();
            var stderrTask = Task.Run(async () =>
            {
                string? line;
                while ((line = await proc.StandardError.ReadLineAsync()) is not null)
                    stderr.AppendLine(line);
            }, ct);

            string? stdoutLine;
            while ((stdoutLine = await proc.StandardOutput.ReadLineAsync()) is not null)
            {
                ct.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace(stdoutLine)) continue;

                JsonDocument? doc = null;
                try { doc = JsonDocument.Parse(stdoutLine); }
                catch { /* non-JSON line — skip */ }
                if (doc is not null)
                {
                    using (doc) { yield return doc.RootElement.Clone(); }
                }
            }

            await proc.WaitForExitAsync(ct);
            await stderrTask;

            if (proc.ExitCode != 0 && stderr.Length > 0)
            {
                var fault = JsonSerializer.Serialize(new
                {
                    kind = "stderr",
                    exitCode = proc.ExitCode,
                    message = stderr.ToString().TrimEnd(),
                });
                using var doc = JsonDocument.Parse(fault);
                yield return doc.RootElement.Clone();
            }
        }
        finally
        {
            // Mid-stream cancel (page navigated away, request aborted) →
            // the iterator's finally fires here. Kill the helper tree so
            // Excel.exe doesn't outlive the request.
            try { if (!proc.HasExited) proc.Kill(entireProcessTree: true); }
            catch { }
            proc.Dispose();
        }
    }
}
