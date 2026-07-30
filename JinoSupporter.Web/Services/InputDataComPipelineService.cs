using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace JinoSupporter.Web.Services;

public sealed record InputDataComPipelineRequest(
    string InputPath,
    string DatasetName,
    int Limit,
    bool RunAiReviewCase);

public sealed class InputDataComPipelineService(
    IWebHostEnvironment env,
    WebRepository repo,
    ExcelHelperRunner excelHelper,
    IConfiguration config)
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
    };

    private static readonly string[] ExcelExtensions = [".xlsx", ".xlsm", ".xlsb", ".xls"];
    private static readonly UTF8Encoding Utf8WithBom = new(encoderShouldEmitUTF8Identifier: true);
    private const int ExtractorMaxAttempts = 5;

    public async Task<InputDataComPipelineResult> RunAsync(
        InputDataComPipelineRequest request,
        Action<string>? onOutput,
        CancellationToken ct)
    {
        string startedAt = DateTime.UtcNow.ToString("O");
        string dataset = string.IsNullOrWhiteSpace(request.DatasetName)
            ? "InputDataCom"
            : request.DatasetName.Trim();

        IReadOnlyList<string> files = ResolveInputFiles(request.InputPath, request.Limit);
        if (files.Count == 0)
            throw new InvalidOperationException("No Excel files were found for COM extraction.");

        string runId = Guid.NewGuid().ToString("N")[..12];
        string tmpDir = Path.Combine(env.ContentRootPath, "tmp", "input-data-com", runId);
        Directory.CreateDirectory(tmpDir);

        var results = new List<InputDataComPipelineFileResult>();
        int reviewCaseCount = 0;

        for (int i = 0; i < files.Count; i++)
        {
            ct.ThrowIfCancellationRequested();

            string sourcePath = files[i];
            string fileName = Path.GetFileName(sourcePath);
            string outputJsonPath = Path.Combine(tmpDir, $"{i + 1:000}_{SafeFileStem(fileName)}.com-grid.json");
            onOutput?.Invoke($"[{i + 1}/{files.Count}] Excel COM extract: {fileName}");

            try
            {
                await RunExtractorAsync(sourcePath, outputJsonPath, onOutput, ct);

                using JsonDocument doc = JsonDocument.Parse(await File.ReadAllTextAsync(outputJsonPath, ct));
                InputDataComExtractionSave save = BuildSave(dataset, sourcePath, outputJsonPath, doc.RootElement);
                InputDataComExtractionStoreResult stored = repo.SaveInputDataComExtraction(save);
                onOutput?.Invoke(
                    $"Stored raw grid: workbookId={stored.WorkbookId}, sheets={stored.SheetCount:N0}, cells={stored.TotalCells:N0}, merges={stored.MergeCount:N0}, candidates={stored.CandidateCount:N0}");

                bool aiSaved = false;
                string reviewCaseStatus = "";
                if (request.RunAiReviewCase)
                {
                    onOutput?.Invoke($"Starting AI recognition and Ask AI extraction-readiness check: workbookId={stored.WorkbookId}");
                    (aiSaved, reviewCaseStatus) = await RunAiReviewCaseAsync(stored, save, outputJsonPath, tmpDir, onOutput, ct);
                    if (aiSaved) reviewCaseCount++;
                }

                results.Add(new InputDataComPipelineFileResult(
                    sourcePath,
                    fileName,
                    true,
                    "",
                    stored.WorkbookId,
                    stored.SheetCount,
                    stored.TotalRows,
                    stored.TotalCells,
                    stored.NonEmptyCells,
                    stored.MergeCount,
                    stored.CandidateCount,
                    request.RunAiReviewCase,
                    aiSaved,
                    reviewCaseStatus));
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                onOutput?.Invoke($"FAILED: {fileName} / {ex.Message}");
                results.Add(new InputDataComPipelineFileResult(
                    sourcePath,
                    fileName,
                    false,
                    ex.Message,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    request.RunAiReviewCase,
                    false,
                    ""));
            }
        }

        string finishedAt = DateTime.UtcNow.ToString("O");
        return new InputDataComPipelineResult(
            dataset,
            files.Count,
            results.Count(r => r.Success),
            results.Count(r => !r.Success),
            results.Sum(r => r.SheetCount),
            results.Sum(r => r.TotalRows),
            results.Sum(r => r.TotalCells),
            results.Sum(r => r.NonEmptyCells),
            results.Sum(r => r.MergeCount),
            results.Sum(r => r.CandidateCount),
            reviewCaseCount,
            startedAt,
            finishedAt,
            results);
    }

    private async Task RunExtractorAsync(string sourcePath, string outputJsonPath, Action<string>? onOutput, CancellationToken ct)
    {
        Exception? lastError = null;
        for (int attempt = 1; attempt <= ExtractorMaxAttempts; attempt++)
        {
            ct.ThrowIfCancellationRequested();
            TryDelete(outputJsonPath);
            onOutput?.Invoke($"Excel COM extract attempt {attempt}/{ExtractorMaxAttempts}");

            try
            {
                await RunExtractorAttemptAsync(sourcePath, outputJsonPath, onOutput, ct);
                return;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex) when (attempt < ExtractorMaxAttempts)
            {
                lastError = ex;
                onOutput?.Invoke($"Excel COM extract retry {attempt}/{ExtractorMaxAttempts} failed: {ex.Message}");
                await Task.Delay(TimeSpan.FromSeconds(Math.Min(10, attempt * 2)), ct);
            }
            catch (Exception ex)
            {
                lastError = ex;
            }
        }

        throw new InvalidOperationException($"Excel COM extraction failed after {ExtractorMaxAttempts} attempts: {lastError?.Message}", lastError);
    }

    private async Task RunExtractorAttemptAsync(string sourcePath, string outputJsonPath, Action<string>? onOutput, CancellationToken ct)
    {
        string scriptPath = Path.Combine(env.ContentRootPath, "tools", "input_data_excel_com_extract.py");
        if (!File.Exists(scriptPath))
            throw new FileNotFoundException("Input Data COM extractor script was not found.", scriptPath);

        string python = string.IsNullOrWhiteSpace(excelHelper.PythonExePath)
            ? "python"
            : excelHelper.PythonExePath;

        var psi = new ProcessStartInfo
        {
            FileName = python,
            WorkingDirectory = env.ContentRootPath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };
        psi.ArgumentList.Add(scriptPath);
        psi.ArgumentList.Add("--input");
        psi.ArgumentList.Add(sourcePath);
        psi.ArgumentList.Add("--output");
        psi.ArgumentList.Add(outputJsonPath);
        psi.ArgumentList.Add("--covered-cell-mode");
        psi.ArgumentList.Add("blank");

        string stdout = "";
        string stderr = "";
        using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };
        process.OutputDataReceived += (_, e) =>
        {
            if (e.Data is null) return;
            stdout += e.Data + Environment.NewLine;
            onOutput?.Invoke(e.Data);
        };
        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data is null) return;
            stderr += e.Data + Environment.NewLine;
            onOutput?.Invoke(e.Data);
        };

        try
        {
            process.Start();
            ChildProcessJob.Assign(process);
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Excel COM extractor launch failed: {ex.Message}", ex);
        }

        await process.WaitForExitAsync(ct);
        if (process.ExitCode != 0 || !File.Exists(outputJsonPath))
        {
            string detail = FirstNonBlank(stdout, stderr, $"Extractor exited with code {process.ExitCode}.");
            throw new InvalidOperationException(detail.Trim());
        }
    }

    private async Task<(bool Saved, string Status)> RunAiReviewCaseAsync(
        InputDataComExtractionStoreResult stored,
        InputDataComExtractionSave save,
        string rawJsonPath,
        string tmpDir,
        Action<string>? onOutput,
        CancellationToken ct)
    {
        string generationRequestPath = Path.Combine(tmpDir, $"reviewcase-generate-{stored.WorkbookId}.request.json");
        string generationOutputPath = Path.Combine(tmpDir, $"reviewcase-generate-{stored.WorkbookId}.json");
        string preanalysisDraftPath = Path.Combine(tmpDir, $"reviewcase-draft-{stored.WorkbookId}.json");
        string verificationRequestPath = Path.Combine(tmpDir, $"reviewcase-verify-{stored.WorkbookId}.request.json");
        string verificationOutputPath = Path.Combine(tmpDir, $"reviewcase-verify-{stored.WorkbookId}.json");
        string repoRoot = AiPromptRegistry.FindRepositoryRoot();
        string reviewCaseRulesPath = ExistingFilePath(AiPromptRegistry.ResolvePath("input-data/reviewcase-com-rules.md"));
        string legacyReviewCasePromptPath = ExistingFilePath(AiPromptRegistry.ResolvePath("data-inference/reviewcase-ai-analysis.md"));
        string calibrationReferencePath = ExistingFilePath(Path.Combine(repoRoot, "ASK_AI_REVIEWCASE_NEXT_STEPS.md"));
        string auditDecisionPath = ExistingFilePath(Path.Combine(repoRoot, "REVIEWCASE_AI_AUDIT_DECISIONS.md"));
        string verifiedReviewCaseManifestPath = ExistingFilePath(Path.Combine(repoRoot, "REVIEWCASE_AI_DRAFTS", "verified", "reviewcase_ai_verification_manifest.json"));
        object? knownUserDecision = ReadKnownUserDecisionByFileName(save.SourceFileName, auditDecisionPath);
        object preanalysisDraft = BuildCliStyleReviewCaseDraft(stored, save, knownUserDecision);
        string preanalysisDraftJson = JsonSerializer.Serialize(preanalysisDraft, JsonOptions);
        await File.WriteAllTextAsync(preanalysisDraftPath, preanalysisDraftJson, Utf8WithBom, ct);

        var request = new
        {
            workflow = "JinoSupporter Input Data COM recognition for Ask AI extraction",
            workbookId = stored.WorkbookId,
            datasetName = stored.DatasetName,
            sourceFile = stored.SourceFileName,
            sourcePath = save.SourcePath,
            rawGridJsonPath = rawJsonPath,
            preanalysisDraftJsonPath = preanalysisDraftPath,
            reviewCaseRulesPath,
            legacyReviewCasePromptPath,
            calibrationReferencePath,
            auditDecisionPath,
            verifiedReviewCaseManifestPath,
            knownUserDecision,
            rule = "Read preanalysisDraftJsonPath first, then reviewCaseRulesPath, legacyReviewCasePromptPath, calibrationReferencePath, auditDecisionPath, and rawGridJsonPath. The preanalysis draft is the CLI-style program draft. Improve it only with cited raw grid evidence. Existing user audit/calibration decisions must be honored. Judge only data capture, grouping clarity, and Ask AI extraction readiness; do not judge product/process quality.",
            candidates = save.Candidates.Take(250).Select(c => new
            {
                c.CandidateKind,
                c.SheetName,
                c.RowNumber,
                c.Label,
                c.Confidence,
                c.EvidenceCells,
                c.RawText
            }).ToArray()
        };
        await File.WriteAllTextAsync(generationRequestPath, JsonSerializer.Serialize(request, JsonOptions), Utf8WithBom, ct);

        string generationPrompt = BuildGenerationPrompt(generationRequestPath);
        string generationRaw = await RunCodexExecAsync(generationPrompt, generationOutputPath, onOutput, ct);
        string generationJson;
        try
        {
            generationJson = ExtractJsonObject(generationRaw);
            if (!HasReviewCases(generationJson) && HasReviewCases(preanalysisDraftJson))
            {
                onOutput?.Invoke("AI generation returned no ReviewCase; using CLI-style program draft for verification.");
                generationJson = preanalysisDraftJson;
            }
        }
        catch (JsonException ex)
        {
            onOutput?.Invoke($"AI generation did not return valid JSON; using CLI-style program draft for verification. {ex.Message}");
            generationJson = preanalysisDraftJson;
        }
        await File.WriteAllTextAsync(generationOutputPath, generationJson, Utf8WithBom, ct);

        var verifyRequest = new
        {
            workflow = "JinoSupporter Input Data COM recognition verification for Ask AI extraction",
            workbookId = stored.WorkbookId,
            datasetName = stored.DatasetName,
            sourceFile = stored.SourceFileName,
            rawGridJsonPath = rawJsonPath,
            preanalysisDraftJsonPath = preanalysisDraftPath,
            generatedReviewCaseJsonPath = generationOutputPath,
            reviewCaseRulesPath,
            legacyReviewCasePromptPath,
            calibrationReferencePath,
            auditDecisionPath,
            verifiedReviewCaseManifestPath,
            knownUserDecision,
            rule = "Read rules/calibration/audit paths, then verify every cited row/cell against raw grid. Merge metadata is helpful evidence, not a mandatory approval condition. Approve only when extracted data, grouping, units, and metric ownership are clear enough for later Ask AI extraction."
        };
        await File.WriteAllTextAsync(verificationRequestPath, JsonSerializer.Serialize(verifyRequest, JsonOptions), Utf8WithBom, ct);

        string verificationPrompt = BuildVerificationPrompt(verificationRequestPath);
        string verificationRaw = await RunCodexExecAsync(verificationPrompt, verificationOutputPath, onOutput, ct);
        string verificationJson = ExtractJsonObject(verificationRaw);
        await File.WriteAllTextAsync(verificationOutputPath, verificationJson, Utf8WithBom, ct);

        string status = FirstNonBlank(
            JsonPropertyString(verificationJson, "aiReviewCaseStatus"),
            JsonPropertyString(verificationJson, "reviewCaseStatus"),
            JsonPropertyString(generationJson, "reviewCaseStatus"),
            "needs_review");
        bool approved = JsonPropertyBool(verificationJson, "approvedForAskAi")
                        || status.Equals("verified", StringComparison.OrdinalIgnoreCase);
        string reviewCaseId = FirstNonBlank(
            FirstReviewCaseId(generationJson),
            $"input-com-{stored.WorkbookId}-rc-1");
        string now = DateTime.UtcNow.ToString("O");

        repo.SaveInputDataReviewCase(new InputDataReviewCaseSave(
            stored.WorkbookId,
            reviewCaseId,
            status,
            approved,
            generationJson,
            verificationJson,
            now,
            now));

        onOutput?.Invoke($"AI recognition saved: status={status}, approvedForAskAi={approved}");
        return (true, status);
    }

    private async Task<string> RunCodexExecAsync(string prompt, string outputPath, Action<string>? onOutput, CancellationToken ct)
    {
        string repoRoot = AiPromptRegistry.FindRepositoryRoot();
        string logPath = outputPath + ".log";
        TryDelete(outputPath);
        TryDelete(logPath);

        string codexPath = ResolveCodexCliPath();
        var psi = BuildCodexStartInfo(codexPath, repoRoot, outputPath);

        var output = new StringBuilder();
        var error = new StringBuilder();
        using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };
        process.OutputDataReceived += (_, e) =>
        {
            if (e.Data is null) return;
            output.AppendLine(e.Data);
            onOutput?.Invoke(e.Data);
        };
        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data is null) return;
            error.AppendLine(e.Data);
            onOutput?.Invoke(e.Data);
        };

        try
        {
            process.Start();
            ChildProcessJob.Assign(process);
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            await process.StandardInput.WriteAsync(prompt);
            await process.StandardInput.FlushAsync(ct);
            process.StandardInput.Close();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Codex CLI launch failed: {ex.Message}", ex);
        }

        await process.WaitForExitAsync(ct);

        string combinedLog = output + Environment.NewLine + error;
        await File.WriteAllTextAsync(logPath, combinedLog, Utf8WithBom, ct);

        string final = File.Exists(outputPath)
            ? await File.ReadAllTextAsync(outputPath, ct)
            : combinedLog;
        if (process.ExitCode != 0)
            throw new InvalidOperationException(FirstNonBlank(final, combinedLog, $"Codex CLI exited with code {process.ExitCode}."));

        return final;
    }

    private ProcessStartInfo BuildCodexStartInfo(string codexPath, string repoRoot, string outputPath)
    {
        string[] args =
        [
            "exec",
            "-c",
            "model_reasoning_effort=\"high\"",
            "--cd",
            repoRoot,
            "--dangerously-bypass-approvals-and-sandbox",
            "--skip-git-repo-check",
            "--color",
            "never",
            "--output-last-message",
            outputPath,
            "-"
        ];

        string ext = Path.GetExtension(codexPath);
        var psi = new ProcessStartInfo
        {
            WorkingDirectory = repoRoot,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };

        if (ext.Equals(".cmd", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".bat", StringComparison.OrdinalIgnoreCase))
        {
            psi.FileName = "powershell.exe";
            psi.ArgumentList.Add("-NoProfile");
            psi.ArgumentList.Add("-ExecutionPolicy");
            psi.ArgumentList.Add("Bypass");
            psi.ArgumentList.Add("-Command");
            psi.ArgumentList.Add(BuildPowerShellCodexCommand(codexPath, args));
        }
        else if (ext.Equals(".ps1", StringComparison.OrdinalIgnoreCase))
        {
            psi.FileName = "powershell.exe";
            psi.ArgumentList.Add("-NoProfile");
            psi.ArgumentList.Add("-ExecutionPolicy");
            psi.ArgumentList.Add("Bypass");
            psi.ArgumentList.Add("-Command");
            psi.ArgumentList.Add(BuildPowerShellCodexCommand(codexPath, args));
        }
        else
        {
            psi.FileName = codexPath;
            foreach (string arg in args)
                psi.ArgumentList.Add(arg);
        }

        return psi;
    }

    private static string BuildPowerShellCodexCommand(string codexPath, IReadOnlyList<string> args)
    {
        string joinedArgs = string.Join(" ", args.Select(QuotePowerShell));
        return string.Join("; ", new[]
        {
            "$ErrorActionPreference = 'Stop'",
            "$inputText = [Console]::In.ReadToEnd()",
            $"$inputText | & {QuotePowerShell(codexPath)} {joinedArgs}",
            "exit $LASTEXITCODE"
        });
    }

    private string ResolveCodexCliPath()
    {
        string configured = FirstNonBlank(
            config["Codex:CliPath"],
            config["OpenAI:CodexCliPath"],
            Environment.GetEnvironmentVariable("CODEX_CLI_PATH"),
            Environment.GetEnvironmentVariable("OPENAI_CODEX_CLI_PATH"));
        if (!string.IsNullOrWhiteSpace(configured) && File.Exists(configured))
            return Path.GetFullPath(configured);

        string? fromPath = FindExecutableOnPath("codex.cmd")
                           ?? FindExecutableOnPath("codex.exe")
                           ?? FindExecutableOnPath("codex")
                           ?? FindExecutableOnPath("codex.ps1");
        if (!string.IsNullOrWhiteSpace(fromPath))
            return fromPath;

        string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        string[] candidates =
        [
            Path.Combine(appData, "npm", "codex.cmd"),
            Path.Combine(appData, "npm", "codex.exe"),
            Path.Combine(appData, "npm", "codex"),
            Path.Combine(userProfile, ".npm-global", "codex.cmd"),
            Path.Combine(userProfile, ".bun", "bin", "codex.exe"),
            Path.Combine(userProfile, ".bun", "bin", "codex")
        ];

        string? found = candidates.FirstOrDefault(File.Exists);
        if (!string.IsNullOrWhiteSpace(found))
            return found;

        throw new FileNotFoundException(
            "Codex CLI was not found. Set Codex:CliPath or CODEX_CLI_PATH to the full path of codex.cmd.",
            "codex");
    }

    private static string? FindExecutableOnPath(string fileName)
    {
        string pathValue = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (string rawDir in pathValue.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            string dir = rawDir.Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(dir)) continue;
            try
            {
                string candidate = Path.Combine(dir, fileName);
                if (File.Exists(candidate))
                    return candidate;
            }
            catch { }
        }

        return null;
    }

    private static string QuotePowerShell(string value)
        => "'" + (value ?? "").Replace("'", "''", StringComparison.Ordinal) + "'";

    private static InputDataComExtractionSave BuildSave(string dataset, string sourcePath, string rawJsonPath, JsonElement root)
    {
        string fileName = JsonString(root, "fileName", Path.GetFileName(sourcePath));
        long fileSize = JsonLong(root, "fileSize", new FileInfo(sourcePath).Length);
        long mtimeNs = JsonLong(root, "mtimeNs", File.GetLastWriteTimeUtc(sourcePath).Ticks);
        string extractedAt = JsonString(root, "extractedAt", DateTime.UtcNow.ToString("O"));
        JsonElement totals = root.TryGetProperty("totals", out JsonElement t) ? t : default;

        var sheets = new List<InputDataComSheetSave>();
        var cells = new List<InputDataComCellSave>();
        var merges = new List<InputDataComMergeSave>();
        var candidateRows = new List<RowCandidateSource>();

        if (root.TryGetProperty("sheets", out JsonElement sheetArray) && sheetArray.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement sheet in sheetArray.EnumerateArray())
            {
                string sheetName = JsonString(sheet, "sheetName");
                JsonElement used = sheet.TryGetProperty("usedRange", out JsonElement u) ? u : default;
                sheets.Add(new InputDataComSheetSave(
                    JsonInt(sheet, "sheetIndex"),
                    sheetName,
                    JsonInt(used, "top"),
                    JsonInt(used, "left"),
                    JsonInt(used, "bottom"),
                    JsonInt(used, "right"),
                    JsonInt(used, "rowCount"),
                    JsonInt(used, "columnCount"),
                    JsonInt(sheet, "nonEmptyCells"),
                    JsonInt(sheet, "mergeCount")));

                if (sheet.TryGetProperty("merges", out JsonElement mergeArray) && mergeArray.ValueKind == JsonValueKind.Array)
                {
                    foreach (JsonElement merge in mergeArray.EnumerateArray())
                    {
                        merges.Add(new InputDataComMergeSave(
                            sheetName,
                            JsonString(merge, "address"),
                            JsonInt(merge, "top"),
                            JsonInt(merge, "left"),
                            JsonInt(merge, "bottom"),
                            JsonInt(merge, "right"),
                            JsonInt(merge, "rowSpan"),
                            JsonInt(merge, "columnSpan"),
                            JsonAnyToString(merge, "value")));
                    }
                }

                if (sheet.TryGetProperty("rows", out JsonElement rows) && rows.ValueKind == JsonValueKind.Array)
                {
                    foreach (JsonElement row in rows.EnumerateArray())
                    {
                        int rowNumber = JsonInt(row, "rowNumber");
                        var rowValues = new List<string>();
                        var rowAddresses = new List<string>();

                        if (row.TryGetProperty("cells", out JsonElement cellArray) && cellArray.ValueKind == JsonValueKind.Array)
                        {
                            foreach (JsonElement cell in cellArray.EnumerateArray())
                            {
                                string address = JsonString(cell, "address");
                                string value = JsonAnyToString(cell, "value");
                                string rawValue = JsonAnyToString(cell, "rawValue");
                                JsonElement merge = cell.TryGetProperty("merge", out JsonElement m) ? m : default;
                                JsonElement anchor = merge.ValueKind == JsonValueKind.Object && merge.TryGetProperty("anchor", out JsonElement a) ? a : default;

                                string mergeRole = JsonString(merge, "role", "none");
                                string mergeAddress = JsonString(merge, "address");
                                int? anchorRow = anchor.ValueKind == JsonValueKind.Object ? JsonNullableInt(anchor, "row") : null;
                                int? anchorCol = anchor.ValueKind == JsonValueKind.Object ? JsonNullableInt(anchor, "column") : null;

                                cells.Add(new InputDataComCellSave(
                                    sheetName,
                                    JsonInt(cell, "row"),
                                    JsonInt(cell, "column"),
                                    JsonString(cell, "colLabel"),
                                    address,
                                    value,
                                    rawValue,
                                    mergeRole,
                                    mergeAddress,
                                    anchorRow,
                                    anchorCol));

                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    rowValues.Add(value);
                                    if (!string.IsNullOrWhiteSpace(address)) rowAddresses.Add(address);
                                }
                            }
                        }

                        if (rowValues.Count > 0)
                        {
                            candidateRows.Add(new RowCandidateSource(
                                sheetName,
                                rowNumber,
                                string.Join(" | ", rowValues),
                                rowAddresses.Take(40).ToArray()));
                        }
                    }
                }
            }
        }

        IReadOnlyList<InputDataReviewCandidateSave> candidates = BuildCandidates(candidateRows);

        return new InputDataComExtractionSave(
            dataset,
            Path.GetFullPath(sourcePath),
            fileName,
            fileSize,
            mtimeNs,
            $"{fileSize}:{mtimeNs}",
            "OK",
            "",
            JsonInt(totals, "sheetCount", sheets.Count),
            JsonInt(totals, "rowCount", sheets.Sum(s => s.RowCount)),
            JsonInt(totals, "cellCount", cells.Count),
            JsonInt(totals, "nonEmptyCells", candidateRows.Sum(r => r.EvidenceCells.Count)),
            JsonInt(totals, "mergeCount", merges.Count),
            Path.GetFullPath(rawJsonPath),
            extractedAt,
            sheets,
            cells,
            merges,
            candidates);
    }

    private static IReadOnlyList<InputDataReviewCandidateSave> BuildCandidates(IReadOnlyList<RowCandidateSource> rows)
    {
        var candidates = new List<InputDataReviewCandidateSave>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (RowCandidateSource row in rows)
        {
            string lower = row.RowText.ToLowerInvariant();
            AddIfMatch(candidates, seen, row, "review_context", "high",
                lower.Contains("title") || lower.Contains("purpose") || lower.Contains("objective")
                || lower.Contains("content") || lower.Contains("result") || lower.Contains("decision")
                || lower.Contains("conclusion"));
            AddIfMatch(candidates, seen, row, "comparison_condition", "high",
                ContainsAny(lower, "normal", "test", "before", "after", "control", "sample", "standard"));
            AddIfMatch(candidates, seen, row, "metric_table", "high",
                ContainsAny(lower, "input", "qty", "quantity")
                && ContainsAny(lower, "ng", "ok", "rate", "fail", "defect"));
            AddIfMatch(candidates, seen, row, "measurement_table", "medium",
                Regex.IsMatch(lower, @"\b(?:min|max|avg|average|spec|dimension|gauss|tension|strength|spl|thd|imp|f0|fo|hz|kgf|mg|mm)\b"));
            AddIfMatch(candidates, seen, row, "changed_factor_hint", "medium",
                Regex.IsMatch(lower, @"\b(?:material|supplier|vendor|lot|jig|fixture|machine|m/c|process|condition|method|bond|dry|uv|film|magnet|yoke)\b"));

            if (candidates.Count >= 1500) break;
        }

        return candidates;
    }

    private static void AddIfMatch(
        List<InputDataReviewCandidateSave> candidates,
        HashSet<string> seen,
        RowCandidateSource row,
        string kind,
        string confidence,
        bool matched)
    {
        if (!matched) return;
        string key = $"{kind}|{row.SheetName}|{row.RowNumber}";
        if (!seen.Add(key)) return;
        candidates.Add(new InputDataReviewCandidateSave(
            kind,
            row.SheetName,
            row.RowNumber,
            Truncate(row.RowText, 260),
            confidence,
            row.EvidenceCells,
            row.RowText));
    }

    private static IReadOnlyList<string> ResolveInputFiles(string inputPath, int limit)
    {
        string trimmed = (inputPath ?? "").Trim().Trim('"');
        if (string.IsNullOrWhiteSpace(trimmed))
            throw new ArgumentException("Input path is empty.");

        var files = new List<string>();
        if (File.Exists(trimmed))
        {
            if (!IsExcelFile(trimmed))
                throw new ArgumentException($"Input file is not an Excel workbook: {trimmed}");
            files.Add(Path.GetFullPath(trimmed));
        }
        else if (Directory.Exists(trimmed))
        {
            files.AddRange(Directory.EnumerateFiles(trimmed, "*.*", SearchOption.AllDirectories)
                .Where(IsExcelFile)
                .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .Select(Path.GetFullPath));
        }
        else
        {
            throw new FileNotFoundException("Input path was not found.", trimmed);
        }

        if (limit > 0)
            files = files.Take(limit).ToList();
        return files;
    }

    private static bool IsExcelFile(string path)
        => ExcelExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);

    private static object BuildCliStyleReviewCaseDraft(
        InputDataComExtractionStoreResult stored,
        InputDataComExtractionSave save,
        object? knownUserDecision)
    {
        List<ComRowView> rows = BuildComRows(save);
        HashSet<string> availableRefs = rows.Select(r => r.RowId).ToHashSet(StringComparer.OrdinalIgnoreCase);
        List<object> contextRows = BuildContextRows(rows);
        object modelReview = BuildModelReview(save, rows);
        List<object> comparisonCandidates = BuildComparisonCandidates(rows);
        List<object> metricCandidates = BuildMetricCandidates(rows);
        List<MeasurementCandidateView> measurementCandidates = BuildMeasurementCandidates(rows);

        var sourceStats = new
        {
            sheetCount = save.SheetCount,
            sheetNames = save.Sheets.Select(s => s.SheetName).ToArray(),
            sheetRowCount = rows.Count,
            sheetCellCount = save.TotalCells,
            nonEmptyCells = save.NonEmptyCells,
            mergeCount = save.MergeCount,
            comparisonCandidateCount = comparisonCandidates.Count,
            metricCandidateCount = metricCandidates.Count,
            measurementCandidateCount = measurementCandidates.Count
        };

        if (IsExcludedDecision(knownUserDecision))
        {
            return new
            {
                reviewCaseStatus = "excluded",
                sourceWorkbookId = stored.WorkbookId,
                sourceFile = save.SourceFileName,
                sourcePath = save.SourcePath,
                generatedAt = DateTime.UtcNow.ToString("O"),
                generationMode = "cli-style com-grid preanalysis",
                auditDecision = knownUserDecision,
                sourceStats,
                contextRows,
                reviewCases = Array.Empty<object>(),
                verification = new
                {
                    status = "passed",
                    checkedEvidenceRows = Array.Empty<string>(),
                    issues = new[] { "User audit decision excludes this exact workbook." }
                }
            };
        }

        List<object> outcomes = BuildDraftOutcomes(comparisonCandidates, metricCandidates, measurementCandidates);
        object? reviewCase = null;
        if (outcomes.Count > 0)
        {
            object changedFactor = BuildDraftChangedFactor(save, rows, comparisonCandidates, metricCandidates, measurementCandidates);
            string[] outcomeRows = ExtractRowsFromOutcomes(outcomes).Where(availableRefs.Contains).Distinct(StringComparer.OrdinalIgnoreCase).Take(300).ToArray();
            string[] factorRows = ExtractStringArray(changedFactor, "evidenceRows").Where(availableRefs.Contains).ToArray();
            string[] allRows = factorRows.Concat(outcomeRows).Distinct(StringComparer.OrdinalIgnoreCase).Take(300).ToArray();

            reviewCase = new
            {
                reviewCaseId = $"input-com-{stored.WorkbookId}-rc-1",
                reviewTitle = TitleFromRows(save.SourceFileName, rows),
                reviewPurpose = PurposeFromRows(rows),
                changedFactors = new[] { changedFactor },
                outcomes,
                evidenceRows = allRows,
                limitations = new[]
                {
                    "Program draft generated from Excel COM raw grid; AI must verify grouping before Ask AI approval.",
                    "Program candidates are source-backed hints and may still need user confirmation for ambiguous baseline/control mapping."
                }
            };
        }

        string[] checkedRows = reviewCase is null
            ? Array.Empty<string>()
            : ExtractStringArray(reviewCase, "evidenceRows").Where(availableRefs.Contains).ToArray();
        var issues = new List<string>();
        if (outcomes.Count == 0)
            issues.Add("No comparison, metric, or measurement candidates were generated from the COM raw grid.");
        if (checkedRows.Length == 0 && outcomes.Count > 0)
            issues.Add("Generated candidates have no row evidence references.");

        return new
        {
            reviewCaseStatus = outcomes.Count > 0 ? "needs_ai_verification" : "needs_review",
            sourceWorkbookId = stored.WorkbookId,
            sourceFile = save.SourceFileName,
            sourcePath = save.SourcePath,
            generatedAt = DateTime.UtcNow.ToString("O"),
            generationMode = "cli-style com-grid preanalysis",
            auditDecision = knownUserDecision,
            sourceStats,
            modelReview,
            contextRows,
            comparisonCandidates,
            metricCandidates,
            measurementCandidates = measurementCandidates.Select(ToSerializableMeasurement).ToArray(),
            reviewCases = reviewCase is null ? Array.Empty<object>() : new[] { reviewCase },
            verification = new
            {
                status = issues.Count == 0 ? "passed_for_source_row_existence" : "needs_review",
                checkedEvidenceRows = checkedRows,
                missingEvidenceRows = Array.Empty<string>(),
                issues = issues.ToArray()
            }
        };
    }

    private static List<ComRowView> BuildComRows(InputDataComExtractionSave save)
    {
        return save.Cells
            .GroupBy(c => new { c.SheetName, c.RowNumber })
            .OrderBy(g => save.Sheets.FirstOrDefault(s => string.Equals(s.SheetName, g.Key.SheetName, StringComparison.OrdinalIgnoreCase))?.SheetIndex ?? int.MaxValue)
            .ThenBy(g => g.Key.RowNumber)
            .Select(g =>
            {
                List<ComCellView> cells = g
                    .OrderBy(c => c.ColNumber)
                    .Select(c => new ComCellView(c.ColNumber, c.ColLabel, c.CellAddress, c.CellValue))
                    .ToList();
                List<ComCellView> nonEmpty = cells.Where(c => !string.IsNullOrWhiteSpace(c.Value)).ToList();
                string rowText = string.Join(" | ", nonEmpty.Select(c => $"{c.Address}={c.Value}"));
                return new ComRowView(g.Key.SheetName, g.Key.RowNumber, RowRef(g.Key.SheetName, g.Key.RowNumber), rowText, cells);
            })
            .Where(r => r.Cells.Any(c => !string.IsNullOrWhiteSpace(c.Value)))
            .ToList();
    }

    private static List<object> BuildContextRows(IReadOnlyList<ComRowView> rows)
    {
        string[] terms =
        [
            "title", "purpose", "objective", "content", "condition", "standard", "spec",
            "result", "decision", "conclusion", "remark", "note", "problem", "reason",
            "before", "after", "normal", "test", "sample", "lot", "supplier", "material",
            "machine", "m/c", "jig", "base", "laser", "coating", "gauss", "tension",
            "function", "vision", "repair", "method", "dry", "uv", "press"
        ];

        var result = new List<object>();
        foreach (ComRowView row in rows)
        {
            string lower = row.RowText.ToLowerInvariant();
            if (result.Count < 12 || terms.Any(t => lower.Contains(t, StringComparison.OrdinalIgnoreCase)))
            {
                result.Add(new
                {
                    rowId = row.RowId,
                    sheetName = row.SheetName,
                    rowNumber = row.RowNumber,
                    rowText = Truncate(row.RowText, 520),
                    cells = row.Cells
                        .Where(c => !string.IsNullOrWhiteSpace(c.Value))
                        .Select(c => new { address = c.Address, value = c.Value })
                        .ToArray()
                });
            }

            if (result.Count >= 100) break;
        }

        return result;
    }

    private static object BuildModelReview(InputDataComExtractionSave save, IReadOnlyList<ComRowView> rows)
    {
        var candidates = new Dictionary<string, ModelCandidateView>(StringComparer.OrdinalIgnoreCase);
        AddModelCandidates(candidates, save.SourceFileName, "file_name", "medium", "");
        AddModelCandidates(candidates, save.SourcePath, "source_path", "low", "");
        foreach (ComRowView row in rows.Take(100))
            AddModelCandidates(candidates, row.RowText, "sheet_rows", row.RowNumber <= 12 ? "medium" : "low", row.RowId);

        List<ModelCandidateView> ordered = candidates.Values
            .OrderByDescending(c => ConfidenceRank(c.Confidence))
            .ThenBy(c => c.Model, StringComparer.OrdinalIgnoreCase)
            .Take(12)
            .ToList();
        string[] candidateModels = ordered.Select(c => c.Model).ToArray();
        string[] exact = candidateModels.Where(m => !IsAmbiguousModel(m)).ToArray();
        string[] selected = exact.Length == 1 && candidateModels.All(m => IsAmbiguousModel(m) || string.Equals(m, exact[0], StringComparison.OrdinalIgnoreCase))
            ? [exact[0]]
            : [];

        return new
        {
            sourceModels = "",
            selectedModels = selected,
            mappingStatus = selected.Length > 0 ? "confirmed" : candidateModels.Length > 0 ? "needs_user_mapping" : "missing",
            confidence = selected.Length > 0 ? "medium" : "low",
            candidates = ordered.Select(c => new
            {
                model = c.Model,
                confidence = c.Confidence,
                sources = c.Sources.ToArray(),
                evidence = c.Evidence.ToArray(),
                evidenceRows = c.EvidenceRows.ToArray(),
                ambiguous = IsAmbiguousModel(c.Model)
            }).ToArray(),
            evidenceRows = ordered.SelectMany(c => c.EvidenceRows).Distinct(StringComparer.OrdinalIgnoreCase).Take(8).ToArray()
        };
    }

    private static List<object> BuildComparisonCandidates(IReadOnlyList<ComRowView> rows)
    {
        var result = new List<object>();
        foreach (ComRowView row in rows)
        {
            string lower = row.RowText.ToLowerInvariant();
            if (!ContainsAny(lower, "normal", "test", "before", "after", "control", "standard", "old", "new"))
                continue;
            result.Add(new
            {
                candidateId = $"comparison-{result.Count + 1}",
                sheetName = row.SheetName,
                rowNumber = row.RowNumber,
                rowId = row.RowId,
                conditionText = Truncate(row.RowText, 420),
                evidenceRows = new[] { row.RowId },
                evidenceCells = row.Cells.Where(c => !string.IsNullOrWhiteSpace(c.Value)).Select(c => c.Address).ToArray()
            });
            if (result.Count >= 300) break;
        }

        return result;
    }

    private static List<object> BuildMetricCandidates(IReadOnlyList<ComRowView> rows)
    {
        var result = new List<object>();
        foreach (ComRowView row in rows)
        {
            string lower = row.RowText.ToLowerInvariant();
            bool hasQty = ContainsAny(lower, "input", "qty", "quantity", "ok", "ng", "rate", "fail", "defect");
            bool hasNumber = row.Cells.Any(c => TryParseNumber(c.Value, out _));
            if (!hasQty || !hasNumber) continue;

            result.Add(new
            {
                metricId = $"metric-{result.Count + 1}",
                sheetName = row.SheetName,
                rowNumber = row.RowNumber,
                rowId = row.RowId,
                tableTitle = NearestTitle(rows, row),
                conditionLabel = ConditionLabelFromRow(row),
                rawRow = Truncate(row.RowText, 520),
                evidenceRows = new[] { row.RowId },
                evidenceCells = row.Cells.Where(c => !string.IsNullOrWhiteSpace(c.Value)).Select(c => c.Address).ToArray()
            });
            if (result.Count >= 300) break;
        }

        return result;
    }

    private static List<MeasurementCandidateView> BuildMeasurementCandidates(IReadOnlyList<ComRowView> rows)
    {
        var result = new List<MeasurementCandidateView>();
        Dictionary<string, List<ComRowView>> bySheet = rows.GroupBy(r => r.SheetName).ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        foreach (List<ComRowView> sheetRows in bySheet.Values)
        {
            Dictionary<int, ComRowView> rowMap = sheetRows.ToDictionary(r => r.RowNumber);
            foreach (ComRowView row in sheetRows)
            {
                List<ComCellView> numericCells = row.Cells
                    .Where(c => TryParseNumber(c.Value, out _))
                    .ToList();
                if (numericCells.Count < 2) continue;

                ComRowView? headerRow = null;
                List<(ComCellView Cell, string Header)> measured = [];
                for (int offset = 1; offset <= 4 && headerRow is null; offset++)
                {
                    if (!rowMap.TryGetValue(row.RowNumber - offset, out ComRowView? possibleHeader)) continue;
                    List<(ComCellView Cell, string Header)> matches = [];
                    foreach (ComCellView cell in numericCells)
                    {
                        string header = possibleHeader.Cells.FirstOrDefault(c => c.Column == cell.Column)?.Value ?? "";
                        if (!string.IsNullOrWhiteSpace(header) && !TryParseNumber(header, out _))
                            matches.Add((cell, header));
                    }

                    if (matches.Count >= 2)
                    {
                        headerRow = possibleHeader;
                        measured = matches;
                    }
                }

                if (headerRow is null || measured.Count < 2) continue;
                string itemLabel = BuildMeasurementItemLabel(row, measured.Select(m => m.Cell.Column).ToHashSet());
                string headerKey = string.Join(" | ", measured.Select(m => m.Header).Distinct(StringComparer.OrdinalIgnoreCase));
                object[] values = measured.Select(m => new
                {
                    header = m.Header,
                    value = m.Cell.Value,
                    cell = m.Cell.Address
                }).Cast<object>().ToArray();

                result.Add(new MeasurementCandidateView(
                    result.Count + 1,
                    row.SheetName,
                    row.RowNumber,
                    row.RowId,
                    itemLabel,
                    headerKey,
                    values,
                    measured.Select(m => m.Cell.Address).Concat(row.Cells.Where(c => !string.IsNullOrWhiteSpace(c.Value) && !measured.Any(m => m.Cell.Column == c.Column)).Select(c => c.Address)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray(),
                    Truncate(row.RowText, 620)));

                if (result.Count >= 300) return result;
            }
        }

        return result;
    }

    private static List<object> BuildDraftOutcomes(
        IReadOnlyList<object> comparisonCandidates,
        IReadOnlyList<object> metricCandidates,
        IReadOnlyList<MeasurementCandidateView> measurementCandidates)
    {
        var outcomes = new List<object>();
        if (comparisonCandidates.Count > 0)
        {
            outcomes.Add(new
            {
                outcomeId = "comparison-1",
                changedFactorId = "cf-1",
                outcomeDomain = "condition comparison",
                outcomeMetric = "condition rows detected from workbook text",
                comparisonRows = ExtractRowsFromCandidates(comparisonCandidates),
                subResults = comparisonCandidates.Take(80).ToArray(),
                limitations = new[] { "Comparison candidates need AI grouping verification before Ask AI extraction." }
            });
        }

        if (metricCandidates.Count > 0)
        {
            outcomes.Add(new
            {
                outcomeId = "metric-1",
                changedFactorId = "cf-1",
                outcomeDomain = "defect/result metric",
                outcomeMetric = "metric rows detected from workbook text",
                comparisonRows = ExtractRowsFromCandidates(metricCandidates),
                subResults = metricCandidates.Take(80).ToArray(),
                limitations = new[] { "Metric candidates need AI grouping verification before Ask AI extraction." }
            });
        }

        if (measurementCandidates.Count > 0)
        {
            var groups = measurementCandidates.GroupBy(m => NormalizeKey(m.HeaderKey)).ToList();
            int idx = 1;
            foreach (IGrouping<string, MeasurementCandidateView> group in groups.Take(20))
            {
                MeasurementCandidateView first = group.First();
                outcomes.Add(new
                {
                    outcomeId = $"measurement-{idx++}",
                    changedFactorId = "cf-1",
                    outcomeDomain = "measurement",
                    outcomeMetric = FirstNonBlank(first.HeaderKey, "measurement table"),
                    comparisonRows = group.Select(m => m.RowId).Distinct(StringComparer.OrdinalIgnoreCase).Take(100).ToArray(),
                    subResults = group.Take(80).Select(ToSerializableMeasurement).ToArray(),
                    limitations = new[] { "Measurement candidates need AI grouping and unit verification before Ask AI extraction." }
                });
            }
        }

        return outcomes;
    }

    private static object BuildDraftChangedFactor(
        InputDataComExtractionSave save,
        IReadOnlyList<ComRowView> rows,
        IReadOnlyList<object> comparisonCandidates,
        IReadOnlyList<object> metricCandidates,
        IReadOnlyList<MeasurementCandidateView> measurementCandidates)
    {
        string textBlob = string.Join(" | ", new[]
        {
            save.SourceFileName,
            string.Join(" | ", rows.Take(20).Select(r => r.RowText)),
            string.Join(" | ", measurementCandidates.Take(20).Select(m => m.HeaderKey))
        });
        string[] domains = InferDomains(textBlob);
        string changed = measurementCandidates.Count > 0
            ? string.Join(", ", measurementCandidates.Select(m => m.HeaderKey).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Take(8))
            : "test/changed condition not explicit";
        string baseline = comparisonCandidates.Count > 0
            ? "baseline/control candidate exists in source rows"
            : "baseline/normal condition not explicit";
        string[] evidenceRows = rows
            .Where(r => IsContextLike(r.RowText))
            .Select(r => r.RowId)
            .Concat(measurementCandidates.Take(20).Select(m => m.RowId))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(100)
            .ToArray();

        return new
        {
            changedFactorId = "cf-1",
            changeDomain = domains,
            changedFactor = Truncate(TitleFromRows(save.SourceFileName, rows), 220),
            baselineCondition = baseline,
            changedCondition = FirstNonBlank(changed, "condition values not explicit"),
            subgroupKeys = measurementCandidates
                .Select(m => m.ItemLabel)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Take(30)
                .ToArray(),
            evidenceRows
        };
    }

    private static bool IsExcludedDecision(object? knownUserDecision)
    {
        if (knownUserDecision is not IDictionary<string, object?> dict) return false;
        return dict.TryGetValue("decision", out object? value)
               && string.Equals(Convert.ToString(value, CultureInfo.InvariantCulture), "exclude", StringComparison.OrdinalIgnoreCase);
    }

    private static object ToSerializableMeasurement(MeasurementCandidateView item)
        => new
        {
            statId = item.Id,
            sheetName = item.SheetName,
            rowNumber = item.RowNumber,
            rowId = item.RowId,
            itemLabel = item.ItemLabel,
            conditionLabel = item.HeaderKey,
            values = item.Values,
            rawRow = item.RawRow,
            evidenceRows = new[] { item.RowId },
            evidenceCells = item.EvidenceCells
        };

    private static string[] ExtractRowsFromCandidates(IReadOnlyList<object> candidates)
        => candidates.SelectMany(c => ExtractStringArray(c, "evidenceRows")).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();

    private static string[] ExtractRowsFromOutcomes(IReadOnlyList<object> outcomes)
        => outcomes.SelectMany(o => ExtractStringArray(o, "comparisonRows")).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();

    private static string[] ExtractStringArray(object source, string propertyName)
    {
        try
        {
            using JsonDocument doc = JsonDocument.Parse(JsonSerializer.Serialize(source));
            if (doc.RootElement.TryGetProperty(propertyName, out JsonElement value) && value.ValueKind == JsonValueKind.Array)
            {
                return value.EnumerateArray()
                    .Select(e => e.ValueKind == JsonValueKind.String ? e.GetString() ?? "" : e.GetRawText())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToArray();
            }
        }
        catch { }
        return [];
    }

    private static void AddModelCandidates(
        Dictionary<string, ModelCandidateView> candidates,
        string text,
        string source,
        string confidence,
        string evidenceRow)
    {
        if (string.IsNullOrWhiteSpace(text)) return;
        Regex[] patterns =
        [
            new(@"\bBRS[-\s]?\d{4,6}(?:[-\s]?(?:DT|GMI|TF))?\b", RegexOptions.IgnoreCase),
            new(@"\bTIU[-\s]?C11[-\s]?20(?:[-\s]?[LR])?\b", RegexOptions.IgnoreCase),
            new(@"\bC11[-\s]?20(?:[-\s]?[LR])?\b", RegexOptions.IgnoreCase),
            new(@"\bTIU[-\s]?L5S3[-\s]?01(?:[-\s]?[LR])?\b", RegexOptions.IgnoreCase),
            new(@"\bL5S3[-\s]?01(?:[-\s]?[LR])?\b", RegexOptions.IgnoreCase),
            new(@"\b(?:MSU[-\s]?)?L?20S15[-\s]?07(?:[-\s]?(?:DT|GMI))?\b", RegexOptions.IgnoreCase),
            new(@"\b(?:MSU[-\s]?)?20S15[-\s]?07\b", RegexOptions.IgnoreCase),
            new(@"\bMSU[-\s]?201507(?:[-\s]?DT)?\b", RegexOptions.IgnoreCase),
            new(@"\bTIM[-\s]?\d{3}(?:[-\s]?[A-Z])?\b", RegexOptions.IgnoreCase)
        ];

        foreach (Regex pattern in patterns)
        {
            foreach (Match match in pattern.Matches(text))
            {
                string model = NormalizeModelName(match.Value);
                if (string.IsNullOrWhiteSpace(model)) continue;
                if (!candidates.TryGetValue(model, out ModelCandidateView? row))
                {
                    row = new ModelCandidateView(model, confidence);
                    candidates[model] = row;
                }

                if (ConfidenceRank(confidence) > ConfidenceRank(row.Confidence))
                    row.Confidence = confidence;
                if (!row.Sources.Contains(source, StringComparer.OrdinalIgnoreCase))
                    row.Sources.Add(source);
                string evidence = Truncate(text, 180);
                if (!string.IsNullOrWhiteSpace(evidence) && !row.Evidence.Contains(evidence, StringComparer.OrdinalIgnoreCase) && row.Evidence.Count < 8)
                    row.Evidence.Add(evidence);
                if (!string.IsNullOrWhiteSpace(evidenceRow) && !row.EvidenceRows.Contains(evidenceRow, StringComparer.OrdinalIgnoreCase) && row.EvidenceRows.Count < 8)
                    row.EvidenceRows.Add(evidenceRow);
            }
        }
    }

    private static string NormalizeModelName(string value)
    {
        string raw = Regex.Replace(value ?? "", @"\s+", " ").Trim().Trim(' ', '.', ',', '_', ';', ':', '/', '\\', '|', '(', ')', '[', ']', '{', '}').ToUpperInvariant();
        raw = raw.Replace("_", " ", StringComparison.Ordinal).Replace("BRS ", "BRS-", StringComparison.Ordinal).Replace("TIU ", "TIU-", StringComparison.Ordinal).Replace("MSU ", "MSU-", StringComparison.Ordinal);
        raw = Regex.Replace(raw, @"\s*-\s*", "-");

        Match brs = Regex.Match(raw, @"\bBRS-?(\d{4,6})(?:-?(?:DT|GMI|TF))?\b", RegexOptions.IgnoreCase);
        if (brs.Success) return $"BRS-{brs.Groups[1].Value}";
        Match tiuC = Regex.Match(raw, @"\b(?:TIU-?)?C11-?20(?:-?([LR]))?\b", RegexOptions.IgnoreCase);
        if (tiuC.Success) return "TIU-C11-20" + (tiuC.Groups[1].Success ? "-" + tiuC.Groups[1].Value : "");
        Match tiuL = Regex.Match(raw, @"\b(?:TIU-?)?L5S3-?01(?:-?([LR]))?\b", RegexOptions.IgnoreCase);
        if (tiuL.Success) return "TIU-L5S3-01" + (tiuL.Groups[1].Success ? "-" + tiuL.Groups[1].Value : "");
        if (Regex.IsMatch(raw, @"\b(?:MSU-?)?L?20S15-?07\b", RegexOptions.IgnoreCase)) return "L20S15-07";
        if (Regex.IsMatch(raw, @"\bMSU-?201507(?:-?DT)?\b", RegexOptions.IgnoreCase)) return "MSU-201507";
        Match tim = Regex.Match(raw, @"\bTIM-?(\d{3})(?:-?([A-Z]))?\b", RegexOptions.IgnoreCase);
        if (tim.Success) return $"TIM-{tim.Groups[1].Value}" + (tim.Groups[2].Success ? "-" + tim.Groups[2].Value : "");
        return "";
    }

    private static int ConfidenceRank(string confidence)
        => confidence.Trim().ToLowerInvariant() switch
        {
            "high" => 3,
            "medium" => 2,
            "low" => 1,
            _ => 0
        };

    private static bool IsAmbiguousModel(string model)
        => Regex.IsMatch(model ?? "", @"^BRS-\d{4}$", RegexOptions.IgnoreCase);

    private static string NearestTitle(IReadOnlyList<ComRowView> rows, ComRowView row)
    {
        return rows
            .Where(r => string.Equals(r.SheetName, row.SheetName, StringComparison.OrdinalIgnoreCase) && r.RowNumber <= row.RowNumber)
            .OrderByDescending(r => r.RowNumber)
            .Select(r => r.RowText)
            .FirstOrDefault(text => IsContextLike(text)) ?? "";
    }

    private static string ConditionLabelFromRow(ComRowView row)
    {
        return string.Join(" | ", row.Cells
            .Where(c => !string.IsNullOrWhiteSpace(c.Value) && !TryParseNumber(c.Value, out _))
            .Select(c => c.Value)
            .Take(8));
    }

    private static string BuildMeasurementItemLabel(ComRowView row, HashSet<int> measurementColumns)
    {
        string label = string.Join(" | ", row.Cells
            .Where(c => !measurementColumns.Contains(c.Column) && !string.IsNullOrWhiteSpace(c.Value))
            .Select(c => c.Value)
            .Take(8));
        return Truncate(label, 220);
    }

    private static bool TryParseNumber(string value, out double number)
    {
        number = 0;
        string text = (value ?? "").Trim();
        if (string.IsNullOrWhiteSpace(text)) return false;
        Match match = Regex.Match(text, @"[-+]?\d+(?:\.\d+)?");
        return match.Success && double.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out number);
    }

    private static string NormalizeKey(string text)
    {
        string value = Regex.Replace(text ?? "", @"[^a-z0-9]+", " ", RegexOptions.IgnoreCase).ToLowerInvariant();
        return Regex.Replace(value, @"\s+", " ").Trim();
    }

    private static string[] InferDomains(string text)
    {
        string lower = (text ?? "").ToLowerInvariant();
        var domains = new List<string>();
        if (ContainsAny(lower, "supplier", "vendor")) domains.Add("supplier");
        if (ContainsAny(lower, "material", "raw material", "mtr", "yoke", "film", "magnet")) domains.Add("material");
        if (ContainsAny(lower, "coating", "plating", "tin", "polish")) domains.Add("coating");
        if (ContainsAny(lower, "machine", "equipment", "m/c", " mc ", "repair")) domains.Add("equipment");
        if (ContainsAny(lower, "jig", "fixture")) domains.Add("jig");
        if (ContainsAny(lower, "mold", "mould")) domains.Add("mold");
        if (ContainsAny(lower, "condition", "method", "press", "dry", "uv", "plasma", "bonding", "assembly", "assy", "ass'y")) domains.Add("process condition");
        if (ContainsAny(lower, "dimension", "spec", "thickness", "cutting", "height", "width")) domains.Add("dimension/spec");
        if (ContainsAny(lower, " lot", "lot ")) domains.Add("lot");
        return domains.Count == 0 ? ["unknown"] : domains.Distinct(StringComparer.OrdinalIgnoreCase).Take(5).ToArray();
    }

    private static bool IsContextLike(string text)
    {
        string lower = (text ?? "").ToLowerInvariant();
        return ContainsAny(lower,
            "title", "purpose", "objective", "content", "condition", "standard", "spec",
            "result", "decision", "conclusion", "remark", "note", "problem", "reason",
            "before", "after", "normal", "test", "sample", "lot", "supplier", "material",
            "machine", "m/c", "jig", "method", "dry", "uv", "press");
    }

    private static string TitleFromRows(string fileName, IReadOnlyList<ComRowView> rows)
    {
        foreach (ComRowView row in rows.Take(20))
        {
            string text = row.RowText;
            if (text.Contains("title", StringComparison.OrdinalIgnoreCase) || text.Length > 20)
            {
                string cleaned = Regex.Replace(text, @"^.*?\bTITLE\s*\|\s*", "", RegexOptions.IgnoreCase);
                return Truncate(FirstNonBlank(cleaned, text), 220);
            }
        }

        return Path.GetFileNameWithoutExtension(fileName).Replace("_clean", "", StringComparison.OrdinalIgnoreCase);
    }

    private static string PurposeFromRows(IReadOnlyList<ComRowView> rows)
    {
        for (int i = 0; i < rows.Count; i++)
        {
            if (!rows[i].RowText.Contains("purpose", StringComparison.OrdinalIgnoreCase)) continue;
            return string.Join(" / ", rows.Skip(i).Take(5).Select(r => Truncate(r.RowText, 260)).Where(x => !string.IsNullOrWhiteSpace(x)));
        }

        return "";
    }

    private static string RowRef(string sheet, int rowNumber)
        => string.IsNullOrWhiteSpace(sheet) ? rowNumber.ToString(CultureInfo.InvariantCulture) : $"{sheet}!{rowNumber}";

    private static bool HasReviewCases(string json)
    {
        try
        {
            using JsonDocument doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("reviewCases", out JsonElement cases) && cases.ValueKind == JsonValueKind.Array)
                return cases.GetArrayLength() > 0;
        }
        catch { }
        return false;
    }

    private static string BuildGenerationPrompt(string requestPath)
        => """
        Run JinoSupporter Input Data recognition and Ask AI extraction-readiness generation.

        RequestJson=__REQUEST_PATH__

        Read RequestJson first. Then read these paths when present:
        - RequestJson.preanalysisDraftJsonPath
        - RequestJson.reviewCaseRulesPath
        - RequestJson.legacyReviewCasePromptPath
        - RequestJson.calibrationReferencePath
        - RequestJson.auditDecisionPath
        - RequestJson.rawGridJsonPath

        Encoding rule:
        - All JSON/Markdown inputs are UTF-8. If you inspect files with Windows PowerShell, always use Get-Content -Encoding UTF8 or [System.IO.File]::ReadAllText(path, [System.Text.Encoding]::UTF8). Never use PowerShell's default Get-Content encoding for these files.

        The raw grid was extracted by Excel COM from the source workbook in read-only mode. It contains UsedRange rows, cells, and explicit merged-cell metadata. Treat that raw grid as source authority.

        The preanalysis draft is a CLI-style program draft equivalent to the old generate_reviewcase_batch.py layer. It contains source-backed contextRows, comparisonCandidates, metricCandidates, measurementCandidates, modelReview, and initial reviewCases. Start from this draft. Do not ignore it and do not restart from filename-only reasoning.

        Program candidates in RequestJson and in the preanalysis draft are hints, but the draft is the required starting structure. You may correct it only when raw grid evidence proves the correction.

        Scope rule:
        - Your job is not to judge whether the product/process result is good, bad, improved, worsened, pass, or fail.
        - Your job is to judge whether the Excel COM extraction captured the workbook data correctly enough, whether data groups are separated clearly enough, and whether Ask AI can later extract the needed data without losing table structure.
        - Treat source-provided words such as OK, NG, pass, fail, improved, or worsened as workbook data values only. Do not create a new quality judgement unless the workbook explicitly contains that exact decision.

        Existing CLI calibration and user mapping rules are mandatory context:
        - If RequestJson.knownUserDecision exists, honor it as a user-confirmed decision for this exact source file.
        - Use calibration examples for reasoning style and grouping policy, not as hard-coded filename or keyword rules.
        - Do not exclude only because mergeCount is 0, merged headers are absent, existing comparison_pairs are absent, or the workbook lacks explicit "Normal/Test" wording.
        - Measurement-only, DOE, multi-arm, repeated-run, equipment validation, material/supplier/coating, treatment-condition, and dimension/spec workbooks can be Ask AI-usable data packages when rows/cells support condition/result evidence.

        Language rule:
        - Return JSON property names exactly as requested in English.
        - Write all human-readable values in Korean: reviewTitle, reviewPurpose, changeDomain, changedFactor, baselineCondition, changedCondition, outcomes, limitations, and verification issue text.
        - Keep sheet names, cell addresses, file names, units, and source values exactly as they appear in the workbook.

        Your job:
        - Refine the preanalysis draft into final JSON that represents the AI-recognized data structure for Ask AI.
        - Decide whether the extracted workbook contains one usable data package, multiple usable data packages, or no Ask AI-usable data package.
        - Determine what the workbook data is about, what condition/comparison groups exist, which rows are baseline/control-like, which rows are test/comparison-like, and which result/measurement tables belong to each group.
        - For every outcome, write resultSummary as a data recognition statement: what columns/metrics/rows were recognized and whether the separation is clear enough for later Ask AI extraction.
        - For every outcome, include recognizedTables with visible workbook values. These tables are for the user screen, so they must show the recognized headers and data values directly, not only evidence row/cell IDs. Keep source row/cell IDs separately in evidenceRows/evidenceCells.
        - For tension/measurement outcomes, recognizedTables must include the actual measurement matrix, not only Q'ty check rows. Include Min, Max, AVG, and Std Dev rows when the workbook shows or allows calculating them from cited measurement rows.
        - Keep row evidence IDs exactly in the preanalysis format: SheetName!RowNumber, for example Sheet1!6. Do not invent Sheet1!R6 or other row-id variants.
        - If there is no single baseline/control but the workbook still has citeable condition combinations and result/measurement rows, preserve those condition rows and mark needs_review only for unresolved extraction/grouping.
        - Preserve separate outcome domains. Do not merge process defect, function defect, and measurement evidence unless the workbook explicitly links them.
        - Cite row/cell evidence using sheet name, row number, and cell addresses from the raw grid.
        - Run a self-check before returning.

        Return only compact JSON with this shape:
        {
          "reviewCaseStatus": "verified | needs_review | excluded",
          "sourceWorkbookId": 0,
          "sourceFile": "",
          "reviewCases": [
            {
              "reviewCaseId": "",
              "reviewTitle": "",
              "reviewPurpose": "",
              "changedFactors": [
                {
                  "changedFactorId": "",
                  "changeDomain": "",
                  "changedFactor": "",
                  "baselineCondition": "",
                  "changedCondition": "",
                  "subgroupKeys": [],
                  "evidenceRows": [],
                  "evidenceCells": []
                }
              ],
              "outcomes": [
                {
                  "outcomeId": "",
                  "changedFactorId": "",
                  "outcomeDomain": "",
                  "outcomeMetric": "",
                  "resultSummary": "",
                  "recognizedTables": [
                    {
                      "title": "",
                      "columns": [],
                      "rows": [
                        []
                      ]
                    }
                  ],
                  "comparisonRows": [],
                  "evidenceRows": [],
                  "evidenceCells": [],
                  "limitations": []
                }
              ],
              "evidenceRows": [],
              "evidenceCells": [],
              "limitations": []
            }
          ],
          "verification": {
            "status": "passed | needs_review | failed",
            "checkedEvidenceRows": [],
            "checkedEvidenceCells": [],
            "issues": []
          }
        }
        """.Replace("__REQUEST_PATH__", requestPath, StringComparison.Ordinal);

    private static string BuildVerificationPrompt(string requestPath)
        => """
        Run JinoSupporter Input Data recognition and Ask AI extraction-readiness verification.

        RequestJson=__REQUEST_PATH__

        Read RequestJson. Then read these paths when present:
        - RequestJson.preanalysisDraftJsonPath
        - RequestJson.reviewCaseRulesPath
        - RequestJson.legacyReviewCasePromptPath
        - RequestJson.calibrationReferencePath
        - RequestJson.auditDecisionPath
        - RequestJson.rawGridJsonPath
        - RequestJson.generatedReviewCaseJsonPath

        Encoding rule:
        - All JSON/Markdown inputs are UTF-8. If you inspect files with Windows PowerShell, always use Get-Content -Encoding UTF8 or [System.IO.File]::ReadAllText(path, [System.Text.Encoding]::UTF8). Never use PowerShell's default Get-Content encoding for these files.

        Verify the generated data recognition JSON against both the CLI-style preanalysis draft and the Excel COM raw grid:
        - Every cited row/cell must exist.
        - Row evidence IDs must match the draft/raw-row format exactly: SheetName!RowNumber, for example Sheet1!6.
        - If the generated JSON drops valid source-backed data areas from the preanalysis draft without evidence, mark needs_review and describe the drop.
        - Every numeric value must appear in cited evidence or be calculable from cited numerator/denominator.
        - recognizedTables must contain the actual recognized workbook headers and data values, not only row/cell references.
        - For tension/measurement outcomes, recognizedTables must include the measurement matrix and Min/Max/AVG/Std Dev rows when those statistics are present or calculable from cited rows.
        - Baseline/control-like and test/comparison-like groups must be supported by source evidence such as row/cell layout, table headers, labels, repeated sections, source text, or merge metadata.
        - Merge metadata is helpful structural evidence, but merged headers are not required for Ask AI-usable data recognition.
        - Do not reject or exclude only because mergeCount is 0, comparison_pairs are absent, or Normal/Test wording is implicit.
        - Measurement-only, DOE, multi-arm, repeated-run, and descriptive condition-combination workbooks can be Ask AI-usable when the cited rows/cells support clear extraction.
        - If RequestJson.knownUserDecision exists, verify that the output honors that user-confirmed decision.
        - Do not judge whether product/process results are good, bad, improved, worsened, pass, or fail. Only check whether those values were captured and separated correctly enough for later Ask AI extraction.
        - If grouping, units, condition boundaries, or metric ownership are ambiguous, do not approve for Ask AI.
        - If the workbook has no extractable condition/result/measurement data for Ask AI, mark excluded.

        Language rule:
        - Return JSON property names exactly as requested in English.
        - Write all human-readable values in Korean: summary, issues, requiredUserQuestions, correctionPlan, and evidencePolicy.
        - summary must directly answer: 데이터가 제대로 입력됐는지, 구분이 잘 됐는지, 나중에 Ask AI가 데이터를 잘 추출할 수 있는지.
        - Keep sheet names, cell addresses, file names, units, and source values exactly as they appear in the workbook.

        Return only compact JSON:
        {
          "sourceWorkbookId": 0,
          "aiReviewCaseStatus": "verified | needs_review | excluded",
          "verificationStatus": "passed | needs_review | failed",
          "approvedForAskAi": false,
          "confidence": "high | medium | low",
          "summary": "",
          "issues": [],
          "requiredUserQuestions": [],
          "correctionPlan": [],
          "evidencePolicy": ""
        }
        """.Replace("__REQUEST_PATH__", requestPath, StringComparison.Ordinal);

    private static string ExistingFilePath(string path)
        => File.Exists(path) ? Path.GetFullPath(path) : "";

    private static object? ReadKnownUserDecisionByFileName(string sourceFileName, string auditDecisionPath)
    {
        if (string.IsNullOrWhiteSpace(sourceFileName) || string.IsNullOrWhiteSpace(auditDecisionPath) || !File.Exists(auditDecisionPath))
            return null;

        string source = NormalizeAuditFileName(sourceFileName);
        foreach (string line in File.ReadLines(auditDecisionPath))
        {
            Match m = Regex.Match(line, @"^\|\s*(?<id>\d+)\s*\|\s*(?<decision>[^|]+)\|\s*(?<reason>[^|]+)\|\s*`?(?<file>[^|`]+)`?\s*\|");
            if (!m.Success) continue;

            string decision = m.Groups["decision"].Value.Trim();
            if (!decision.Equals("keep", StringComparison.OrdinalIgnoreCase)
                && !decision.Equals("exclude", StringComparison.OrdinalIgnoreCase))
                continue;

            string fileName = NormalizeAuditFileName(m.Groups["file"].Value);
            if (!string.Equals(source, fileName, StringComparison.OrdinalIgnoreCase))
                continue;

            long.TryParse(m.Groups["id"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long fileId);
            return new Dictionary<string, object?>
            {
                ["fileId"] = fileId,
                ["decision"] = decision,
                ["reason"] = m.Groups["reason"].Value.Trim(),
                ["fileName"] = fileName,
                ["source"] = "REVIEWCASE_AI_AUDIT_DECISIONS.md"
            };
        }

        return null;
    }

    private static string NormalizeAuditFileName(string value)
    {
        string fileName = Path.GetFileName((value ?? "").Trim().Trim('`'));
        return Regex.Replace(fileName, @"\s+", " ").Trim();
    }

    private static string ExtractJsonObject(string raw)
    {
        string text = (raw ?? "").Trim();
        if (text.StartsWith("```", StringComparison.Ordinal))
        {
            int firstNewline = text.IndexOf('\n');
            if (firstNewline >= 0) text = text[(firstNewline + 1)..];
            int lastFence = text.LastIndexOf("```", StringComparison.Ordinal);
            if (lastFence >= 0) text = text[..lastFence];
            text = text.Trim();
        }

        string? balanced = TryExtractFirstJsonObject(text);
        if (!string.IsNullOrWhiteSpace(balanced))
            text = balanced;

        using JsonDocument _ = JsonDocument.Parse(text);
        return text;
    }

    private static string? TryExtractFirstJsonObject(string text)
    {
        for (int start = 0; start < text.Length; start++)
        {
            if (text[start] != '{') continue;

            bool inString = false;
            bool escaped = false;
            int depth = 0;
            for (int i = start; i < text.Length; i++)
            {
                char ch = text[i];
                if (inString)
                {
                    if (escaped)
                    {
                        escaped = false;
                    }
                    else if (ch == '\\')
                    {
                        escaped = true;
                    }
                    else if (ch == '"')
                    {
                        inString = false;
                    }
                    continue;
                }

                if (ch == '"')
                {
                    inString = true;
                    continue;
                }

                if (ch == '{')
                {
                    depth++;
                }
                else if (ch == '}')
                {
                    depth--;
                    if (depth == 0)
                    {
                        string candidate = text[start..(i + 1)].Trim();
                        try
                        {
                            using JsonDocument _ = JsonDocument.Parse(candidate);
                            return candidate;
                        }
                        catch (JsonException)
                        {
                            break;
                        }
                    }
                }
            }
        }

        return null;
    }

    private static string FirstReviewCaseId(string json)
    {
        try
        {
            using JsonDocument doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("reviewCases", out JsonElement cases)
                && cases.ValueKind == JsonValueKind.Array
                && cases.GetArrayLength() > 0)
            {
                JsonElement first = cases[0];
                return JsonString(first, "reviewCaseId");
            }
        }
        catch { }
        return "";
    }

    private static string JsonPropertyString(string json, string property)
    {
        try
        {
            using JsonDocument doc = JsonDocument.Parse(json);
            return JsonString(doc.RootElement, property);
        }
        catch { return ""; }
    }

    private static bool JsonPropertyBool(string json, string property)
    {
        try
        {
            using JsonDocument doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty(property, out JsonElement value) && value.ValueKind is JsonValueKind.True or JsonValueKind.False)
                return value.GetBoolean();
        }
        catch { }
        return false;
    }

    private static string JsonString(JsonElement element, string propertyName, string fallback = "")
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(propertyName, out JsonElement value))
            return fallback;
        return value.ValueKind == JsonValueKind.String ? value.GetString() ?? fallback : JsonValueToString(value);
    }

    private static string JsonAnyToString(JsonElement element, string propertyName)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(propertyName, out JsonElement value))
            return "";
        return JsonValueToString(value);
    }

    private static string JsonValueToString(JsonElement value)
        => value.ValueKind switch
        {
            JsonValueKind.Null or JsonValueKind.Undefined => "",
            JsonValueKind.String => value.GetString() ?? "",
            JsonValueKind.Number => value.TryGetInt64(out long l)
                ? l.ToString(CultureInfo.InvariantCulture)
                : value.GetDouble().ToString("G17", CultureInfo.InvariantCulture),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            _ => value.GetRawText()
        };

    private static int JsonInt(JsonElement element, string propertyName, int fallback = 0)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(propertyName, out JsonElement value))
            return fallback;
        return value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out int result) ? result : fallback;
    }

    private static int? JsonNullableInt(JsonElement element, string propertyName)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(propertyName, out JsonElement value))
            return null;
        return value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out int result) ? result : null;
    }

    private static long JsonLong(JsonElement element, string propertyName, long fallback = 0)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(propertyName, out JsonElement value))
            return fallback;
        return value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out long result) ? result : fallback;
    }

    private static string SafeFileStem(string fileName)
        => Regex.Replace(Path.GetFileNameWithoutExtension(fileName), @"[^A-Za-z0-9_.-]+", "_").Trim('_', '.', '-');

    private static string Truncate(string value, int maxLength)
    {
        value = Regex.Replace(value ?? "", @"\s+", " ").Trim();
        return value.Length <= maxLength ? value : value[..maxLength].TrimEnd();
    }

    private static bool ContainsAny(string text, params string[] terms)
        => terms.Any(term => text.Contains(term, StringComparison.OrdinalIgnoreCase));

    private static string FirstNonBlank(params string[] values)
        => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch { }
    }

    private sealed record RowCandidateSource(
        string SheetName,
        int RowNumber,
        string RowText,
        IReadOnlyList<string> EvidenceCells);

    private sealed record ComCellView(
        int Column,
        string ColLabel,
        string Address,
        string Value);

    private sealed record ComRowView(
        string SheetName,
        int RowNumber,
        string RowId,
        string RowText,
        IReadOnlyList<ComCellView> Cells);

    private sealed class ModelCandidateView(string model, string confidence)
    {
        public string Model { get; } = model;
        public string Confidence { get; set; } = confidence;
        public List<string> Sources { get; } = [];
        public List<string> Evidence { get; } = [];
        public List<string> EvidenceRows { get; } = [];
    }

    private sealed record MeasurementCandidateView(
        int Id,
        string SheetName,
        int RowNumber,
        string RowId,
        string ItemLabel,
        string HeaderKey,
        IReadOnlyList<object> Values,
        IReadOnlyList<string> EvidenceCells,
        string RawRow);
}
