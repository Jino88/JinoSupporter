using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace JinoSupporter.Web.Services;

public sealed class InputDataTestBatchExtractor
{
    private static readonly string[] ExcelExtensions = [".xlsx", ".xlsm", ".xlsb", ".xls"];
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public string RepoRoot { get; }
    public string RendererPath { get; }
    public string PythonCommand { get; }
    public string SavedSessionsDir => Path.Combine(RepoRoot, "tmp", "input-data-test", "saved");

    public InputDataTestBatchExtractor(IHostEnvironment env)
    {
        RepoRoot = FindRepoRoot(env.ContentRootPath);
        RendererPath = Path.Combine(RepoRoot, "JinoSupporter.Web", "tools", "input_data_test_render_excel.py");
        PythonCommand = Environment.GetEnvironmentVariable("INPUT_DATA_TEST_PYTHON") ?? "python";
    }

    public IReadOnlyList<string> ScanExcelFiles(string folder)
    {
        if (string.IsNullOrWhiteSpace(folder))
            throw new ArgumentException("Folder path is empty.");
        if (!Directory.Exists(folder))
            throw new DirectoryNotFoundException($"Folder not found: {folder}");

        return Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
            .Where(IsExcelFile)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .Select(Path.GetFullPath)
            .ToList();
    }

    public Task<InputDataTestBatchExtractResult> ExtractFileAsync(string workbookPath, CancellationToken ct = default)
        => ExtractFileAsync(workbookPath, "", ct);

    public async Task<InputDataTestBatchExtractResult> ExtractFileAsync(
        string workbookPath,
        string sourceDataset,
        CancellationToken ct = default)
        => await ExtractFileAsync(workbookPath, sourceDataset, renderImages: true, progress: null, ct: ct);

    public async Task<InputDataTestBatchExtractResult> ExtractFileAsync(
        string workbookPath,
        string sourceDataset,
        bool renderImages,
        IProgress<string>? progress = null,
        CancellationToken ct = default,
        int? maxProgramCellsPerSheet = null)
    {
        if (string.IsNullOrWhiteSpace(workbookPath))
            throw new ArgumentException("Workbook path is empty.");
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException("Workbook not found.", workbookPath);
        if (!File.Exists(RendererPath))
            throw new FileNotFoundException("Input Data Test renderer was not found.", RendererPath);

        string safeBase = SafeFileName(Path.GetFileNameWithoutExtension(workbookPath));
        if (string.IsNullOrWhiteSpace(safeBase)) safeBase = "workbook";

        string saveStem = $"{DateTime.Now:yyyyMMdd_HHmmss}_{safeBase}_{Guid.NewGuid():N}";
        saveStem = saveStem[..Math.Min(saveStem.Length, 180)];
        string assetDir = Path.Combine(SavedSessionsDir, saveStem);
        string structureDir = Path.Combine(assetDir, "ai-structure");
        string imageDir = Path.Combine(structureDir, "images");
        string renderedRequestPath = Path.Combine(structureDir, "rendered-request.json");
        string structureJsonPath = Path.Combine(structureDir, "ai-structure.json");
        string structureTextPath = Path.Combine(structureDir, "ai-structure.txt");
        string extractedTextPath = Path.Combine(assetDir, $"{safeBase}.extracted.txt");
        string savedWorkbookPath = Path.Combine(assetDir, SafeFileName(Path.GetFileName(workbookPath)));
        string sessionPath = Path.Combine(SavedSessionsDir, $"{saveStem}.json");
        string progressPath = Path.Combine(structureDir, "extract-progress.log");

        Directory.CreateDirectory(assetDir);
        Directory.CreateDirectory(structureDir);
        File.Copy(workbookPath, savedWorkbookPath, overwrite: true);

        progress?.Report($"Staged workbook copy: {Path.GetFileName(savedWorkbookPath)}");
        RenderProcessResult renderResult = await RunRendererAsync(savedWorkbookPath, imageDir, renderedRequestPath, renderImages, progressPath, progress, ct, maxProgramCellsPerSheet);
        if (renderResult.ExitCode != 0)
        {
            string rendererError = await ReadRendererFailureAsync(renderedRequestPath, ct);
            string message = FirstNonEmpty(rendererError, renderResult.StdErr, renderResult.StdOut, $"Renderer exited with code {renderResult.ExitCode}.");
            throw new InvalidOperationException(message);
        }

        using JsonDocument renderedDoc = JsonDocument.Parse(await File.ReadAllTextAsync(renderedRequestPath, ct));
        progress?.Report("Building text extraction from rendered request JSON...");
        string structureJson = BuildProgramStructureJson(renderedDoc.RootElement);
        using JsonDocument structureDoc = JsonDocument.Parse(structureJson);

        string formattedStructureJson = JsonSerializer.Serialize(structureDoc.RootElement, JsonOptions);
        string structureText = AiStructureJsonToText(structureDoc.RootElement);
        string extractedText = BuildWorkbookExtractText(renderedDoc.RootElement, out WorkbookBatchStats stats);
        int imageCount = CountRenderedImages(renderedDoc.RootElement);
        string completedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        await File.WriteAllTextAsync(structureJsonPath, formattedStructureJson, new UTF8Encoding(false), ct);
        await File.WriteAllTextAsync(structureTextPath, structureText, new UTF8Encoding(false), ct);
        await File.WriteAllTextAsync(extractedTextPath, extractedText, new UTF8Encoding(false), ct);
        WriteSheetFiles(structureDoc.RootElement, structureDir);
        DeleteDirectorySafe(imageDir);
        progress?.Report($"Extraction text saved. Sheets={stats.SheetCount:N0}, rows={stats.RowCount:N0}, cells={stats.CellCount:N0}.");

        var session = new
        {
            savedAtUtc = DateTime.UtcNow.ToString("o"),
            workflow = "INPUT DATA (TEST)",
            generatedBy = "INPUT DATA (BATCH)",
            promptOptionsSource = "INPUT DATA (TEST) current Auto Prompt library",
            workbook = new
            {
                fileName = Path.GetFileName(workbookPath),
                sourceDataset = sourceDataset ?? "",
                workbookPath = Path.GetFullPath(savedWorkbookPath),
                extractedTextPath,
                sheetCount = stats.SheetCount,
                rowCount = stats.RowCount,
                cellCount = stats.CellCount,
                extractedTextTruncated = stats.Truncated,
                aiStructure = new
                {
                    dir = structureDir,
                    jsonPath = structureJsonPath,
                    textPath = structureTextPath,
                    renderedRequestPath,
                    imageCount,
                    completedAt
                }
            },
            reviewPurpose = "",
            reviewContext = "",
            steps = new[]
            {
                new
                {
                    step = 1,
                    prompt = "",
                    analysisText = "",
                    analysisHtml = "",
                    error = "",
                    completedAt = ""
                }
            }
        };

        Directory.CreateDirectory(SavedSessionsDir);
        await File.WriteAllTextAsync(
            sessionPath,
            JsonSerializer.Serialize(session, JsonOptions),
            new UTF8Encoding(false),
            ct);

        return new InputDataTestBatchExtractResult(
            workbookPath,
            sessionPath,
            extractedTextPath,
            structureTextPath,
            renderedRequestPath,
            stats.SheetCount,
            stats.RowCount,
            stats.CellCount,
            imageCount);
    }

    public async Task<InputDataTestBatchExtractResult> ExtractWorkbookBytesAsync(
        string fileName,
        string sourceDataset,
        byte[] data,
        CancellationToken ct = default)
    {
        if (data.Length == 0)
            throw new ArgumentException("Workbook data is empty.");

        string tempDir = Path.Combine(RepoRoot, "tmp", "input-data-test", "batch-db-work", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        string safeName = SafeFileName(string.IsNullOrWhiteSpace(fileName) ? "workbook.xlsx" : fileName);
        if (string.IsNullOrWhiteSpace(Path.GetExtension(safeName))) safeName += ".xlsx";
        string tempPath = Path.Combine(tempDir, safeName);

        await File.WriteAllBytesAsync(tempPath, data, ct);
        try
        {
            return await ExtractFileAsync(tempPath, sourceDataset, ct);
        }
        finally
        {
            DeleteDirectorySafe(tempDir);
        }
    }

    public IReadOnlyDictionary<string, InputDataTestSavedSessionSummary> LoadSavedSessionIndex()
    {
        var result = new Dictionary<string, InputDataTestSavedSessionSummary>(StringComparer.OrdinalIgnoreCase);
        if (!Directory.Exists(SavedSessionsDir)) return result;

        foreach (string sessionPath in Directory.EnumerateFiles(SavedSessionsDir, "*.json", SearchOption.TopDirectoryOnly))
        {
            InputDataTestSavedSessionSummary? summary = TryReadSessionSummary(sessionPath);
            if (summary is null || string.IsNullOrWhiteSpace(summary.FileName)) continue;

            if (!result.TryGetValue(summary.FileName, out InputDataTestSavedSessionSummary? existing)
                || string.CompareOrdinal(summary.SavedAtUtc, existing.SavedAtUtc) > 0)
            {
                result[summary.FileName] = summary;
            }
        }

        return result;
    }

    public bool ClearSessionAnalysis(string sessionPath)
    {
        if (string.IsNullOrWhiteSpace(sessionPath) || !File.Exists(sessionPath))
            return false;

        JsonNode? node = JsonNode.Parse(File.ReadAllText(sessionPath));
        if (node is not JsonObject session || session["steps"] is not JsonArray steps)
            return false;

        bool changed = false;
        foreach (JsonNode? stepNode in steps)
        {
            if (stepNode is not JsonObject step) continue;

            changed |= !string.IsNullOrWhiteSpace(NodeString(step["analysisText"]));
            changed |= !string.IsNullOrWhiteSpace(NodeString(step["analysisHtml"]));
            changed |= step.ContainsKey("parameters");
            changed |= step.ContainsKey("completedAt");

            step["analysisText"] = "";
            step["analysisHtml"] = "";
            step["error"] = "";
            step["parameters"] = new JsonObject();
            step.Remove("completedAt");
        }

        if (!changed) return false;

        session["savedAtUtc"] = DateTime.UtcNow.ToString("o");
        File.WriteAllText(
            sessionPath,
            session.ToJsonString(JsonOptions),
            new UTF8Encoding(false));
        return true;
    }

    public async Task<InputDataTestBatchAnalysisResult> AnalyzeSessionAsync(
        string sessionPath,
        string stepPrompt = "",
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(sessionPath) || !File.Exists(sessionPath))
            throw new FileNotFoundException("Saved INPUT DATA (TEST) session was not found.", sessionPath);

        JsonNode? node = JsonNode.Parse(await File.ReadAllTextAsync(sessionPath, ct));
        if (node is not JsonObject session)
            throw new InvalidOperationException("Saved session JSON is not an object.");

        string runId = Guid.NewGuid().ToString("N")[..12];
        string tmpDir = Path.Combine(RepoRoot, "tmp");
        Directory.CreateDirectory(tmpDir);
        string requestPath = Path.Combine(tmpDir, $"input_data_batch_request_{runId}.json");
        string outputPath = Path.Combine(tmpDir, $"input_data_batch_output_{runId}.txt");

        try
        {
            await WriteAnalysisRequestAsync(requestPath, runId, session, stepPrompt, ct);
            string prompt = InputDataTestAnalysisSupport.BuildCodexPrompt(requestPath);
            string raw = await RunCodexExecAsync(prompt, outputPath, ct);
            InputDataTestAnalysisResult parsed = InputDataTestAnalysisSupport.ParseStepResult(raw);

            UpdateSessionAnalysis(session, stepPrompt, parsed);
            await File.WriteAllTextAsync(
                sessionPath,
                session.ToJsonString(JsonOptions),
                new UTF8Encoding(false),
                ct);

            return new InputDataTestBatchAnalysisResult(
                sessionPath,
                parsed.AnalysisText,
                parsed.AnalysisHtml,
                parsed.Parameters,
                parsed.Error);
        }
        finally
        {
            TryDeleteFile(requestPath);
            TryDeleteFile(outputPath);
        }
    }

    private async Task WriteAnalysisRequestAsync(
        string requestPath,
        string runId,
        JsonObject session,
        string stepPrompt,
        CancellationToken ct)
    {
        JsonObject workbook = CloneObject(session["workbook"]);
        string workbookFileName = NodeString(workbook["fileName"]);
        string workbookPath = NodeString(workbook["workbookPath"]);
        string sourceDataset = NodeString(workbook["sourceDataset"]);
        string extractedTextPath = NodeString(workbook["extractedTextPath"]);
        var reviewIndex = new
        {
            indexHtmlPath = InputDataTestAnalysisSupport.DetailedReviewIndexPath,
            exists = File.Exists(InputDataTestAnalysisSupport.DetailedReviewIndexPath),
            matchKeys = InputDataTestAnalysisSupport.BuildReviewIndexMatchKeys(
                workbookFileName,
                sourceDataset,
                workbookPath,
                extractedTextPath),
            usage = "Open this HTML as the AI-generated taxonomy/reference table after workbook-text AI pre-analysis. Match by matchKeys when possible, but do not replace workbook-text AI classification with filename lookup."
        };

        var previousSteps = Array.Empty<object>();

        var request = new
        {
            runId,
            createdAt = DateTime.UtcNow.ToString("o"),
            workflow = "INPUT DATA (TEST)",
            generatedBy = "INPUT DATA (BATCH)",
            workbook,
            reviewIndex,
            reviewPurpose = NodeString(session["reviewPurpose"]),
            reviewContext = NodeString(session["reviewContext"]),
            autoPrompts = InputDataTestAnalysisSupport.LoadEnabledAutoPrompts(RepoRoot),
            currentStep = new
            {
                step = 2,
                prompt = stepPrompt ?? ""
            },
            previousSteps
        };

        await File.WriteAllTextAsync(
            requestPath,
            JsonSerializer.Serialize(request, JsonOptions),
            new UTF8Encoding(false),
            ct);
    }

    private async Task<string> RunCodexExecAsync(string prompt, string outputPath, CancellationToken ct)
    {
        string launchId = Guid.NewGuid().ToString("N")[..12];
        string promptPath = Path.Combine(RepoRoot, $"_input_data_batch_prompt_codex_{launchId}.txt");
        string scriptPath = Path.Combine(RepoRoot, $"_input_data_batch_launch_codex_{launchId}.ps1");
        string donePath = Path.Combine(Path.GetDirectoryName(outputPath) ?? RepoRoot, $"input_data_batch_done_{launchId}.json");
        string logPath = outputPath + ".log";

        try
        {
            File.WriteAllText(promptPath, prompt, new UTF8Encoding(false));
            File.WriteAllText(
                scriptPath,
                BuildVisibleCodexLauncherScript(
                    RepoRoot,
                    promptPath,
                    scriptPath,
                    "INPUT DATA (BATCH) - Codex analysis",
                    outputPath,
                    donePath),
                new UTF8Encoding(false));
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Codex launch script write failed: {ex.Message}", ex);
        }

        TryDeleteFile(outputPath);
        TryDeleteFile(logPath);
        TryDeleteFile(donePath);

        if (!TryLaunchVisibleCodexTerminal(scriptPath))
            throw new InvalidOperationException("Codex CLI terminal launch failed.");

        while (!ct.IsCancellationRequested)
        {
            if (File.Exists(donePath)) break;
            await Task.Delay(TimeSpan.FromSeconds(2), ct);
        }

        string output = File.Exists(outputPath)
            ? await File.ReadAllTextAsync(outputPath, ct)
            : "";

        if (!File.Exists(donePath))
            throw new OperationCanceledException("Waiting for Codex CLI was cancelled.", ct);

        int exitCode = ReadExitCode(await File.ReadAllTextAsync(donePath, ct));
        TryDeleteFile(donePath);

        if (exitCode != 0)
        {
            string log = File.Exists(logPath)
                ? await File.ReadAllTextAsync(logPath, ct)
                : "";
            string recoveredJson = InputDataTestAnalysisSupport.ExtractStepResultJson(output + Environment.NewLine + log);
            if (!string.IsNullOrWhiteSpace(recoveredJson))
            {
                TryDeleteFile(logPath);
                return recoveredJson;
            }

            string detail = FirstNonEmpty(output, log, $"Codex CLI exited with code {exitCode}.");
            throw new InvalidOperationException($"{detail}\nLog: {logPath}");
        }

        TryDeleteFile(logPath);
        return output;
    }

    private static bool TryLaunchVisibleCodexTerminal(string scriptPath)
    {
        string workDir = Path.GetDirectoryName(scriptPath) ?? Environment.CurrentDirectory;

        try
        {
            var cmd = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                WorkingDirectory = workDir,
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Normal,
                Arguments = $"/c powershell.exe -NoProfile -ExecutionPolicy Bypass -File {QuoteCmdArg(scriptPath)}",
            };
            Process? proc = Process.Start(cmd);
            if (proc is not null) ChildProcessJob.Assign(proc);
            return proc is not null;
        }
        catch { }

        try
        {
            var wt = new ProcessStartInfo
            {
                FileName = "wt.exe",
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Normal,
                Arguments = $"-d {QuoteCmdArg(workDir)} powershell.exe -NoProfile -ExecutionPolicy Bypass -File {QuoteCmdArg(scriptPath)}",
            };
            Process? proc = Process.Start(wt);
            if (proc is not null) ChildProcessJob.Assign(proc);
            return proc is not null;
        }
        catch
        {
            return false;
        }
    }

    private static string QuoteCmdArg(string value)
        => "\"" + (value ?? "").Replace("\"", "\\\"") + "\"";

    private static string BuildVisibleCodexLauncherScript(
        string workDir,
        string promptPath,
        string scriptPath,
        string title,
        string outputPath,
        string donePath)
    {
        static string Q(string value) => "'" + (value ?? "").Replace("'", "''") + "'";
        return string.Join(Environment.NewLine, new[]
        {
            "$ErrorActionPreference = 'Continue'",
            "$utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false",
            "[Console]::InputEncoding = $utf8NoBom",
            "[Console]::OutputEncoding = $utf8NoBom",
            "$OutputEncoding = $utf8NoBom",
            "try { chcp 65001 | Out-Null } catch {}",
            $"try {{ $Host.UI.RawUI.WindowTitle = {Q(title)} }} catch {{}}",
            $"Set-Location -LiteralPath {Q(workDir)}",
            "",
            $"$promptPath = {Q(promptPath)}",
            $"$scriptPath = {Q(scriptPath)}",
            $"$outputPath = {Q(outputPath)}",
            "$logPath = $outputPath + '.log'",
            $"$donePath = {Q(donePath)}",
            "Remove-Item -LiteralPath $outputPath -Force -ErrorAction SilentlyContinue",
            "Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue",
            "Remove-Item -LiteralPath $donePath -Force -ErrorAction SilentlyContinue",
            "$prompt = Get-Content -LiteralPath $promptPath -Raw -Encoding UTF8",
            "$exitCode = 0",
            "Write-Host 'INPUT DATA Codex analysis started. effort=high' -ForegroundColor Cyan",
            "Write-Host ('Prompt: ' + $promptPath) -ForegroundColor DarkGray",
            "Write-Host ('Output: ' + $outputPath) -ForegroundColor DarkGray",
            "try {",
            $"    $prompt | & codex exec -c 'model_reasoning_effort=\"high\"' --cd {Q(workDir)} --dangerously-bypass-approvals-and-sandbox --skip-git-repo-check --color never --output-last-message $outputPath - 2>&1 | Tee-Object -FilePath $logPath",
            "    $exitCode = if ($null -eq $LASTEXITCODE) { 0 } else { $LASTEXITCODE }",
            "} catch {",
            "    $exitCode = 1",
            "    $_ | Out-String | Tee-Object -FilePath $outputPath -Append",
            "}",
            "if ($exitCode -ne 0) {",
            "    Write-Host ('Codex exited with code ' + $exitCode + '.') -ForegroundColor Yellow",
            "    if ((-not (Test-Path -LiteralPath $outputPath)) -or ((Get-Item -LiteralPath $outputPath).Length -eq 0)) {",
            "        if (Test-Path -LiteralPath $logPath) { Copy-Item -LiteralPath $logPath -Destination $outputPath -Force }",
            "    }",
            "}",
            "if ($exitCode -eq 0) { Write-Host 'INPUT DATA Codex analysis finished.' -ForegroundColor Green }",
            "$done = @{ exitCode = $exitCode; completedAt = [DateTime]::UtcNow.ToString('o') } | ConvertTo-Json -Compress",
            "Set-Content -LiteralPath $donePath -Value $done -Encoding UTF8",
            "Remove-Item -LiteralPath $promptPath -Force -ErrorAction SilentlyContinue",
            "Remove-Item -LiteralPath $scriptPath -Force -ErrorAction SilentlyContinue",
        });
    }

    private static int ReadExitCode(string doneJson)
    {
        try
        {
            using JsonDocument doc = JsonDocument.Parse(doneJson);
            if (doc.RootElement.TryGetProperty("exitCode", out JsonElement exit)
                && exit.ValueKind == JsonValueKind.Number
                && exit.TryGetInt32(out int value))
            {
                return value;
            }
        }
        catch { }
        return 0;
    }

    private static void UpdateSessionAnalysis(
        JsonObject session,
        string stepPrompt,
        InputDataTestAnalysisResult parsed)
    {
        if (session["steps"] is not JsonArray steps)
        {
            steps = [];
            session["steps"] = steps;
        }

        JsonObject? target = null;
        foreach (JsonNode? stepNode in steps)
        {
            if (stepNode is JsonObject step && NodeInt(step["step"]) == 1)
            {
                target = step;
                break;
            }
        }

        if (target is null)
        {
            target = new JsonObject { ["step"] = 1 };
            steps.Add(target);
        }

        target["prompt"] = stepPrompt ?? "";
        target["parameters"] = InputDataTestAnalysisSupport.ParametersToJsonObject(parsed.Parameters);
        target["analysisText"] = parsed.AnalysisText;
        target["analysisHtml"] = parsed.AnalysisHtml;
        target["error"] = parsed.Error;
        target["completedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        session["savedAtUtc"] = DateTime.UtcNow.ToString("o");
    }

    private async Task<RenderProcessResult> RunRendererAsync(
        string workbookPath,
        string imageDir,
        string renderedRequestPath,
        bool renderImages,
        string progressPath,
        IProgress<string>? progress,
        CancellationToken ct,
        int? maxProgramCellsPerSheet)
    {
        var psi = new ProcessStartInfo
        {
            FileName = PythonCommand,
            WorkingDirectory = RepoRoot,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };
        if (maxProgramCellsPerSheet is > 0)
            psi.Environment["JINO_MAX_PROGRAM_CELLS_PER_SHEET"] = maxProgramCellsPerSheet.Value.ToString(CultureInfo.InvariantCulture);

        psi.ArgumentList.Add(RendererPath);
        psi.ArgumentList.Add(workbookPath);
        psi.ArgumentList.Add(imageDir);
        psi.ArgumentList.Add(renderedRequestPath);
        psi.ArgumentList.Add("36");
        psi.ArgumentList.Add("3");
        if (!renderImages)
            psi.ArgumentList.Add("--text-only");
        psi.ArgumentList.Add("--progress");
        psi.ArgumentList.Add(progressPath);

        using var proc = new Process { StartInfo = psi };
        try
        {
            progress?.Report(renderImages
                ? "Launching Excel COM renderer with visual image capture..."
                : "Launching Excel COM text-only extractor...");
            proc.Start();
            ChildProcessJob.Assign(proc);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Renderer launch failed: {ex.Message}", ex);
        }

        using var cancelReg = ct.Register(() =>
        {
            try { if (!proc.HasExited) proc.Kill(entireProcessTree: true); }
            catch { }
        });

        Task<string> stdoutTask = ReadRendererStreamAsync(proc.StandardOutput, progress, ct);
        Task<string> stderrTask = proc.StandardError.ReadToEndAsync(ct);
        await proc.WaitForExitAsync(ct);

        string stdout = await stdoutTask;
        string stderr = await stderrTask;
        return new RenderProcessResult(proc.ExitCode, stdout.Trim(), stderr.Trim());
    }

    private static async Task<string> ReadRendererStreamAsync(
        StreamReader reader,
        IProgress<string>? progress,
        CancellationToken ct)
    {
        var sb = new StringBuilder();
        while (true)
        {
            string? line = await reader.ReadLineAsync(ct);
            if (line is null) break;

            if (line.StartsWith("PROGRESS\t", StringComparison.Ordinal))
            {
                string message = line["PROGRESS\t".Length..].Trim();
                if (!string.IsNullOrWhiteSpace(message))
                    progress?.Report(message);
                continue;
            }

            sb.AppendLine(line);
        }

        return sb.ToString();
    }

    private static async Task<string> ReadRendererFailureAsync(string renderedRequestPath, CancellationToken ct)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(renderedRequestPath) || !File.Exists(renderedRequestPath))
                return "";

            using JsonDocument doc = JsonDocument.Parse(await File.ReadAllTextAsync(renderedRequestPath, ct));
            string error = JsonString(doc.RootElement, "error");
            if (!string.IsNullOrWhiteSpace(error)) return error;
            return JsonString(doc.RootElement, "traceback");
        }
        catch
        {
            return "";
        }
    }

    private static string BuildProgramStructureJson(JsonElement renderedRoot)
    {
        var sheetPayloads = new List<object>();
        if (renderedRoot.TryGetProperty("sheets", out JsonElement sheets) && sheets.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement sheet in sheets.EnumerateArray())
            {
                int sheetIndex = JsonInt(sheet, "sheetIndex");
                string sheetName = JsonString(sheet, "sheetName");
                string sourceRange = BuildSourceRange(sheet);
                var visibleText = new List<object>();
                var cellRows = new List<string[]>();
                bool truncated = false;

                if (sheet.TryGetProperty("programExtract", out JsonElement programExtract))
                {
                    truncated = JsonBool(programExtract, "truncated");
                    if (programExtract.TryGetProperty("cells", out JsonElement cells) && cells.ValueKind == JsonValueKind.Array)
                    {
                        int order = 1;
                        foreach (JsonElement cell in cells.EnumerateArray())
                        {
                            string text = JsonString(cell, "text");
                            if (string.IsNullOrWhiteSpace(text) && string.IsNullOrWhiteSpace(JsonString(cell, "formula")))
                                continue;

                            string address = JsonString(cell, "address");
                            int row = JsonInt(cell, "row");
                            int column = JsonInt(cell, "column");
                            cellRows.Add([address, row.ToString(), column.ToString(), text]);

                            if (visibleText.Count < 250)
                            {
                                visibleText.Add(new
                                {
                                    order,
                                    role = InferTextRole(text),
                                    text
                                });
                            }
                            order++;
                        }
                    }
                }

                var mergedCells = new List<object>();
                if (sheet.TryGetProperty("mergedCells", out JsonElement merged) && merged.ValueKind == JsonValueKind.Array)
                {
                    foreach (JsonElement item in merged.EnumerateArray())
                    {
                        int rowStart = JsonInt(item, "rowStart");
                        int rowEnd = JsonInt(item, "rowEnd");
                        int colStart = JsonInt(item, "columnStart");
                        int colEnd = JsonInt(item, "columnEnd");
                        mergedCells.Add(new
                        {
                            text = JsonString(item, "text"),
                            range = JsonString(item, "address", "range"),
                            source = JsonString(item, "source"),
                            meaning = "Excel merged range from program extraction",
                            appliesTo = new[] { $"rows {rowStart}-{rowEnd}, columns {colStart}-{colEnd}" }
                        });
                    }
                }

                var notes = new List<string>
                {
                    "Generated without AI from Excel COM programExtract cells and merged ranges.",
                    "Merged cell values are expanded to every row and column in their merged range."
                };
                if (truncated)
                    notes.Add("ProgramExtract reached the per-sheet cell limit.");

                sheetPayloads.Add(new
                {
                    sheetIndex,
                    sheetName,
                    chunksRead = Array.Empty<int>(),
                    visibleText,
                    sections = Array.Empty<object>(),
                    tables = new[]
                    {
                        new
                        {
                            title = "ProgramExtract visible cells",
                            sourceRange,
                            layout = new
                            {
                                headerRows = Array.Empty<string[]>(),
                                bodyRows = Array.Empty<string[]>()
                            },
                            mergedCells,
                            normalizedRows = Array.Empty<object>(),
                            headers = new[] { "Address", "Row", "Column", "Text" },
                            rows = cellRows,
                            notes
                        }
                    },
                    unreadableAreas = truncated
                        ? new[] { "ProgramExtract truncated before reading every visible cell." }
                        : Array.Empty<string>()
                });
            }
        }

        var payload = new
        {
            fileName = JsonString(renderedRoot, "fileName"),
            sheets = sheetPayloads,
            extractionNotes = new[]
            {
                "INPUT DATA (BATCH) Step 1 preprocessing only.",
                "No AI or Codex CLI was used for this extraction.",
                "Merged cell values are treated as the same value for every cell in the merged range."
            }
        };

        return JsonSerializer.Serialize(payload, JsonOptions);
    }

    private static string BuildWorkbookExtractText(JsonElement renderedRoot, out WorkbookBatchStats stats)
    {
        var sb = new StringBuilder();
        int sheetCount = 0;
        int cellCount = 0;
        bool truncated = false;
        var rows = new HashSet<string>(StringComparer.Ordinal);

        string fileName = JsonString(renderedRoot, "fileName");
        if (!string.IsNullOrWhiteSpace(fileName))
            sb.Append("File: ").AppendLine(fileName);
        sb.AppendLine("MergedCellMode: merged cell values are expanded to every row and column in the merged range.");

        if (renderedRoot.TryGetProperty("sheets", out JsonElement sheets) && sheets.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement sheet in sheets.EnumerateArray())
            {
                sheetCount++;
                int sheetIndex = JsonInt(sheet, "sheetIndex", sheetCount);
                string sheetName = JsonString(sheet, "sheetName");
                if (sb.Length > 0) sb.AppendLine();
                sb.Append("## Sheet: ").AppendLine(sheetName);
                sb.AppendLine("Address\tRow\tColumn\tText");

                if (!sheet.TryGetProperty("programExtract", out JsonElement programExtract))
                    continue;

                truncated = truncated || JsonBool(programExtract, "truncated");
                if (!programExtract.TryGetProperty("cells", out JsonElement cells) || cells.ValueKind != JsonValueKind.Array)
                    continue;

                foreach (JsonElement cell in cells.EnumerateArray())
                {
                    string text = CleanCellText(JsonString(cell, "text"));
                    if (string.IsNullOrWhiteSpace(text) && string.IsNullOrWhiteSpace(JsonString(cell, "formula")))
                        continue;

                    int row = JsonInt(cell, "row");
                    int column = JsonInt(cell, "column");
                    rows.Add($"{sheetIndex}:{row}");
                    cellCount++;
                    sb.Append(JsonString(cell, "address"))
                        .Append('\t')
                        .Append(row)
                        .Append('\t')
                        .Append(column)
                        .Append('\t')
                        .AppendLine(text);
                }
            }
        }

        stats = new WorkbookBatchStats(sheetCount, rows.Count, cellCount, truncated);
        return sb.ToString().TrimEnd() + Environment.NewLine;
    }

    private static string AiStructureJsonToText(JsonElement root)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# AI Excel Structure Extraction");

        string fileName = JsonString(root, "fileName");
        if (!string.IsNullOrWhiteSpace(fileName))
            sb.AppendLine($"File: {AiEvidenceText(fileName)}");

        if (root.TryGetProperty("sheets", out JsonElement sheets) && sheets.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement sheet in sheets.EnumerateArray())
            {
                sb.AppendLine();
                int sheetIndex = JsonInt(sheet, "sheetIndex");
                string sheetName = JsonString(sheet, "sheetName");
                sb.AppendLine(sheetIndex > 0
                    ? $"## Sheet {sheetIndex}: {AiEvidenceText(sheetName)}"
                    : $"## Sheet: {AiEvidenceText(sheetName)}");

                if (sheet.TryGetProperty("tables", out JsonElement tables) && tables.ValueKind == JsonValueKind.Array)
                {
                    foreach (JsonElement table in tables.EnumerateArray())
                    {
                        string title = JsonString(table, "title");
                        if (!string.IsNullOrWhiteSpace(title))
                        {
                            sb.AppendLine();
                            sb.AppendLine($"### {AiEvidenceText(title)}");
                        }

                        if (table.TryGetProperty("mergedCells", out JsonElement merged) && merged.ValueKind == JsonValueKind.Array && merged.GetArrayLength() > 0)
                        {
                            sb.AppendLine("Merged cells:");
                            foreach (JsonElement item in merged.EnumerateArray())
                                sb.Append("- ").Append(JsonString(item, "range")).Append(": ").AppendLine(AiEvidenceText(JsonString(item, "text")));
                        }

                        if (table.TryGetProperty("rows", out JsonElement rows) && rows.ValueKind == JsonValueKind.Array)
                        {
                            sb.AppendLine("Address\tRow\tColumn\tText");
                            foreach (JsonElement row in rows.EnumerateArray())
                            {
                                if (row.ValueKind != JsonValueKind.Array) continue;
                                var values = row.EnumerateArray().Select(JsonValueAiEvidenceText).ToArray();
                                sb.AppendLine(string.Join('\t', values));
                            }
                        }
                    }
                }
            }
        }

        if (root.TryGetProperty("extractionNotes", out JsonElement notes) && notes.ValueKind == JsonValueKind.Array)
        {
            sb.AppendLine();
            sb.AppendLine("## Extraction Notes");
            foreach (JsonElement note in notes.EnumerateArray())
                sb.Append("- ").AppendLine(AiEvidenceText(JsonValueText(note)));
        }

        return sb.ToString().TrimEnd() + Environment.NewLine;
    }

    private static void WriteSheetFiles(JsonElement root, string structureDir)
    {
        try
        {
            if (!root.TryGetProperty("sheets", out JsonElement sheets) || sheets.ValueKind != JsonValueKind.Array)
                return;

            string sheetDir = Path.Combine(structureDir, "sheets");
            string tableDir = Path.Combine(structureDir, "tables");
            Directory.CreateDirectory(sheetDir);
            Directory.CreateDirectory(tableDir);

            foreach (JsonElement sheet in sheets.EnumerateArray())
            {
                int sheetIndex = JsonInt(sheet, "sheetIndex");
                string sheetName = JsonString(sheet, "sheetName");
                string sheetStem = SafeFileName($"sheet_{sheetIndex:00}_{sheetName}");
                if (string.IsNullOrWhiteSpace(sheetStem)) sheetStem = $"sheet_{sheetIndex:00}";

                File.WriteAllText(
                    Path.Combine(sheetDir, $"{sheetStem}.json"),
                    JsonSerializer.Serialize(sheet, JsonOptions),
                    new UTF8Encoding(false));

                if (!sheet.TryGetProperty("tables", out JsonElement tables) || tables.ValueKind != JsonValueKind.Array)
                    continue;

                int tableIndex = 1;
                foreach (JsonElement table in tables.EnumerateArray())
                {
                    string tableTitle = JsonString(table, "title");
                    string tableStem = SafeFileName($"sheet_{sheetIndex:00}_table_{tableIndex:00}_{tableTitle}");
                    if (string.IsNullOrWhiteSpace(tableStem))
                        tableStem = $"sheet_{sheetIndex:00}_table_{tableIndex:00}";

                    File.WriteAllText(
                        Path.Combine(tableDir, $"{tableStem}.json"),
                        JsonSerializer.Serialize(table, JsonOptions),
                        new UTF8Encoding(false));
                    tableIndex++;
                }
            }
        }
        catch
        {
            // Auxiliary sheet files are for inspection only. The main JSON/TXT remains the source of truth.
        }
    }

    private static int CountRenderedImages(JsonElement renderedRoot)
    {
        int count = 0;
        if (!renderedRoot.TryGetProperty("sheets", out JsonElement sheets) || sheets.ValueKind != JsonValueKind.Array)
            return 0;

        foreach (JsonElement sheet in sheets.EnumerateArray())
        {
            if (!sheet.TryGetProperty("chunks", out JsonElement chunks) || chunks.ValueKind != JsonValueKind.Array)
                continue;
            foreach (JsonElement chunk in chunks.EnumerateArray())
            {
                if (!string.IsNullOrWhiteSpace(JsonString(chunk, "imagePath"))) count++;
            }
        }
        return count;
    }

    private static string BuildSourceRange(JsonElement sheet)
    {
        if (!sheet.TryGetProperty("usedRange", out JsonElement usedRange) || usedRange.ValueKind != JsonValueKind.Object)
            return "programExtract visible cells";

        int rowStart = JsonInt(usedRange, "rowStart");
        int rowEnd = JsonInt(usedRange, "rowEnd");
        int columnStart = JsonInt(usedRange, "columnStart");
        int columnEnd = JsonInt(usedRange, "columnEnd");
        return $"usedRange R{rowStart}C{columnStart}:R{rowEnd}C{columnEnd}";
    }

    private static string InferTextRole(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return "other";
        string trimmed = text.Trim();
        if (trimmed.StartsWith("I.", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("II.", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("III.", StringComparison.OrdinalIgnoreCase)
            || trimmed.Contains("Purpose", StringComparison.OrdinalIgnoreCase)
            || trimmed.Contains("Result", StringComparison.OrdinalIgnoreCase))
        {
            return "section_header";
        }
        if (trimmed.Length >= 20) return "paragraph";
        return "table_text";
    }

    private static bool IsExcelFile(string path)
    {
        string name = Path.GetFileName(path);
        if (name.StartsWith("~$", StringComparison.Ordinal)) return false;

        string ext = Path.GetExtension(path);
        return ExcelExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase);
    }

    private static string JsonString(JsonElement element, params string[] names)
    {
        foreach (string name in names)
        {
            if (!element.TryGetProperty(name, out JsonElement value)) continue;
            return JsonValueText(value);
        }
        return "";
    }

    private static int JsonInt(JsonElement element, string name, int fallback = 0)
    {
        if (!element.TryGetProperty(name, out JsonElement value)) return fallback;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out int i)) return i;
        if (value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), out i)) return i;
        return fallback;
    }

    private static bool JsonBool(JsonElement element, string name)
    {
        if (!element.TryGetProperty(name, out JsonElement value)) return false;
        if (value.ValueKind == JsonValueKind.True) return true;
        if (value.ValueKind == JsonValueKind.False) return false;
        return value.ValueKind == JsonValueKind.String && bool.TryParse(value.GetString(), out bool b) && b;
    }

    private static string JsonValueText(JsonElement value)
    {
        string text = value.ValueKind switch
        {
            JsonValueKind.String => value.GetString() ?? "",
            JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False => value.ToString(),
            JsonValueKind.Null or JsonValueKind.Undefined => "",
            _ => value.ToString()
        };
        return Flatten(text);
    }

    private static string JsonValueAiEvidenceText(JsonElement value)
        => AiEvidenceText(JsonValueText(value));

    private static string CleanCellText(string value)
        => NormalizeAiEvidenceText((value ?? "")
            .Replace('\t', ' ')
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Trim());

    private static string AiEvidenceText(string value)
        => NormalizeAiEvidenceText(Flatten(value));

    private static string NormalizeAiEvidenceText(string value)
    {
        string text = value ?? "";
        if (text.Length == 0) return "";

        text = text
            .Replace("℃", " degC")
            .Replace("℉", " degF")
            .Replace("°C", " degC", StringComparison.OrdinalIgnoreCase)
            .Replace("°F", " degF", StringComparison.OrdinalIgnoreCase)
            .Replace("㎏", " kg")
            .Replace("㎎", " mg")
            .Replace("㎜", " mm")
            .Replace("㎛", " um")
            .Replace("㎝", " cm")
            .Replace("㎟", " mm2")
            .Replace("㎠", " cm2")
            .Replace("㎡", " m2")
            .Replace("㎥", " m3")
            .Replace("㎪", " kPa");

        return Regex.Replace(text, @"[ \u00A0]{2,}", " ").Trim();
    }

    private static string Flatten(string value)
        => (value ?? "")
            .Replace('\t', ' ')
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Trim();

    private static string SafeFileName(string fileName)
    {
        string safe = string.Join("_", (fileName ?? "workbook").Split(Path.GetInvalidFileNameChars()));
        safe = safe.Trim();
        return string.IsNullOrWhiteSpace(safe) ? "workbook" : safe;
    }

    private static InputDataTestSavedSessionSummary? TryReadSessionSummary(string sessionPath)
    {
        try
        {
            JsonNode? node = JsonNode.Parse(File.ReadAllText(sessionPath));
            if (node is not JsonObject session) return null;
            JsonObject workbook = CloneObject(session["workbook"]);
            string fileName = NodeString(workbook["fileName"]);
            if (string.IsNullOrWhiteSpace(fileName)) return null;

            bool hasStructure = false;
            if (workbook["aiStructure"] is JsonObject aiStructure)
            {
                string textPath = NodeString(aiStructure["textPath"]);
                hasStructure = !string.IsNullOrWhiteSpace(textPath) && File.Exists(textPath);
            }

            bool hasAnalysis = false;
            if (session["steps"] is JsonArray steps)
            {
                foreach (JsonNode? stepNode in steps)
                {
                    if (stepNode is not JsonObject step) continue;
                    if (!string.IsNullOrWhiteSpace(NodeString(step["analysisHtml"])))
                    {
                        hasAnalysis = true;
                        break;
                    }
                }
            }

            return new InputDataTestSavedSessionSummary(
                fileName,
                sessionPath,
                NodeString(session["savedAtUtc"]),
                hasStructure,
                hasAnalysis);
        }
        catch
        {
            return null;
        }
    }

    private static JsonObject CloneObject(JsonNode? node)
    {
        if (node is not JsonObject obj) return [];
        JsonNode? clone = JsonNode.Parse(obj.ToJsonString());
        return clone as JsonObject ?? [];
    }

    private static string NodeString(JsonNode? node)
        => node is null ? "" : node.GetValueKind() == JsonValueKind.String ? node.GetValue<string>() ?? "" : node.ToString();

    private static int NodeInt(JsonNode? node)
    {
        if (node is null) return 0;
        try
        {
            if (node.GetValueKind() == JsonValueKind.Number)
                return node.GetValue<int>();
        }
        catch { }
        return int.TryParse(NodeString(node), out int value) ? value : 0;
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                File.Delete(path);
        }
        catch { }
    }

    private static string FirstNonEmpty(params string[] values)
        => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";

    private static string FindRepoRoot(string startDir)
    {
        string dir = Path.GetFullPath(startDir);
        for (int i = 0; i < 8; i++)
        {
            if (File.Exists(Path.Combine(dir, "JinoSupporter.sln"))) return dir;
            string? parent = Directory.GetParent(dir)?.FullName;
            if (string.IsNullOrWhiteSpace(parent) || parent == dir) break;
            dir = parent;
        }
        return Path.GetFullPath(startDir);
    }

    private static void DeleteDirectorySafe(string path)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                Directory.Delete(path, recursive: true);
        }
        catch { }
    }

    private readonly record struct WorkbookBatchStats(int SheetCount, int RowCount, int CellCount, bool Truncated);
    private readonly record struct RenderProcessResult(int ExitCode, string StdOut, string StdErr);
}

public sealed record InputDataTestBatchExtractResult(
    string WorkbookPath,
    string SessionPath,
    string ExtractedTextPath,
    string StructureTextPath,
    string RenderedRequestPath,
    int SheetCount,
    int RowCount,
    int CellCount,
    int ImageCount);

public sealed record InputDataTestBatchAnalysisResult(
    string SessionPath,
    string AnalysisText,
    string AnalysisHtml,
    InputDataTestAnalysisParameters Parameters,
    string Error);

public sealed record InputDataTestSavedSessionSummary(
    string FileName,
    string SessionPath,
    string SavedAtUtc,
    bool HasStructure,
    bool HasAnalysis);
