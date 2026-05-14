using System.Diagnostics;
using System.Text.Json;

namespace JinoSupporter.Web.Services;

/// <summary>Out-of-process driver for the <c>JinoSupporter.ExcelHelper</c> exe.
/// Web stays at <c>net8.0</c> (the WPF launcher hardcodes that path); the helper
/// is a separate <c>net8.0-windows</c> console binary so it can use WindowsForms
/// folder dialogs and the Excel COM automation needed for DRM-locked workbooks.
///
/// All comms happen over stdout JSON-Lines. The web layer streams those events
/// to the page log; failures (helper missing, non-zero exit) surface as a single
/// fault event so callers don't have to special-case the transport.</summary>
public sealed class ExcelHelperRunner
{
    /// <summary>Path of the helper exe, resolved next to the web exe at startup.</summary>
    public string HelperExePath { get; }

    public ExcelHelperRunner()
    {
        // The helper output is copied next to the web exe via the web csproj's
        // post-build target. AppContext.BaseDirectory is the runtime folder of
        // the running web exe — same folder for both Debug and Release.
        HelperExePath = Path.Combine(AppContext.BaseDirectory, "JinoSupporter.ExcelHelper.exe");
    }

    public bool HelperExists => File.Exists(HelperExePath);

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

    /// <summary>Run <c>clean --source &lt;src&gt; --dest &lt;dst&gt;</c> and yield
    /// each emitted JSON line. <paramref name="source"/> can be a file or a folder —
    /// the helper auto-detects. Unknown kinds are passed through so the caller can
    /// decide what to render. Pass <paramref name="verbose"/> to surface step-level
    /// progress (per-chunk, per-shape, per-N-cells) — useful when diagnosing
    /// long stalls between summary lines.</summary>
    public IAsyncEnumerable<HelperEvent> CleanAsync(
        string source, string dest,
        bool verbose = false, bool keepFormats = false,
        CancellationToken ct = default)
    {
        if (!HelperExists)
            return SingleFault($"Helper not found: {HelperExePath}");

        var argList = new List<string> { "clean", "--source", source, "--dest", dest };
        if (verbose)     argList.Add("--verbose");
        if (keepFormats) argList.Add("--keep-formats");
        return MapToEvents(StreamAsync(argList, ct));
    }

    /// <summary>Convenience: clean a single uploaded file. Writes to a temp
    /// folder, runs the helper, and returns the path of the cleaned output
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

    private static async IAsyncEnumerable<HelperEvent> MapToEvents(
        IAsyncEnumerable<JsonElement> source)
    {
        await foreach (var json in source)
        {
            string kind = json.TryGetProperty("kind", out var k) && k.ValueKind == JsonValueKind.String
                ? (k.GetString() ?? "")
                : "";
            yield return new HelperEvent(kind, json.Clone());
        }
    }

    private static async IAsyncEnumerable<HelperEvent> SingleFault(string message)
    {
        var doc = JsonDocument.Parse(JsonSerializer.Serialize(new { kind = "fatal", message }));
        yield return new HelperEvent("fatal", doc.RootElement.Clone());
        await Task.CompletedTask;
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
