using PortalGNLAB.Web.Services;
using PortalGNLAB.Web.Services.GraphMaker;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>JSON endpoints for the React tools. Each group uses the existing menu permission.</summary>
public static class ToolPageEndpoints
{
    public sealed record PastedImage(string MediaType, string Base64);
    private sealed record TranslationResult(string SourceText, string Translation, CodexTokenUsage? Usage, string Provider, string Model);
    public sealed record TranslateInput(string Text, string TargetLanguage, List<PastedImage> Images);
    public sealed record GraphInput(string Text, string Delimiter, int HeaderRowNumber, bool HeaderNotUse, GraphMakerRequest Request);
    public sealed record ReportInput(ReviewReport Report, string? Request);

    public static IEndpointRouteBuilder MapToolPages(this IEndpointRouteBuilder app)
    {
        var translate = ToolGroup(app, "translate", AppMenus.Translate);
        translate.MapGet("", (CodexApiService api, AiProviderSettingsService settings) => new
        {
            api.Model,
            api.TranslateEffort,
            isCli = IsCli(api),
            available = IsCli(api) ? settings.CodexCliEnabled : api.IsConfigured
        });
        translate.MapPost("", TranslateAsync);

        var reports = ToolGroup(app, "review-report", AppMenus.ReviewReport);
        reports.MapGet("", (ReviewReportService service) => Results.Ok(new
        {
            configured = service.IsConfigured,
            languages = ReviewReportService.Languages,
            reports = service.List()
        }));
        reports.MapGet("/{id}", (string id, ReviewReportService service) =>
            service.Get(id) is { } report ? Results.Ok(report) : Results.NotFound());
        reports.MapDelete("/{id}", (string id, ReviewReportService service) =>
        {
            service.Delete(id);
            return Results.NoContent();
        });
        reports.MapPost("/{action}", ReportActionAsync);

        var graph = ToolGroup(app, "graph-maker", AppMenus.GraphMaker);
        graph.MapGet("/example/{kind:int}/{index:int}", GetExampleAsync);
        graph.MapPost("/parse", (GraphInput body) => Results.Ok(
            GraphMakerParser.Parse(body.Text, new(body.Delimiter, body.HeaderRowNumber, body.HeaderNotUse))));
        graph.MapPost("/render", (GraphInput body) =>
        {
            var table = GraphMakerParser.Parse(body.Text, new(body.Delimiter, body.HeaderRowNumber, body.HeaderNotUse));
            return Results.Ok(GraphMakerPayloadBuilder.Build(table, body.Request));
        });
        return app;
    }

    private static RouteGroupBuilder ToolGroup(IEndpointRouteBuilder app, string path, string menu)
    {
        var group = app.MapGroup("/api/portal/tools/" + path).RequireAuthorization();
        group.AddEndpointFilter(async (ctx, next) =>
        {
            var context = ctx.HttpContext;
            var permissions = context.RequestServices.GetRequiredService<MenuPermissionService>();
            if (!permissions.IsAllowed(context.User, menu))
                return Results.StatusCode(StatusCodes.Status403Forbidden);
            if (!HttpMethods.IsGet(context.Request.Method) &&
                context.Request.Headers["X-Requested-With"] != "PortalGNLAB")
                return Results.BadRequest(new { error = "Required request header missing." });
            try
            {
                return await next(ctx);
            }
            catch (OperationCanceledException)
            {
                return Results.BadRequest(new { error = "Request cancelled or timed out." });
            }
            catch (Exception ex)
            {
                context.RequestServices.GetRequiredService<ILoggerFactory>()
                    .CreateLogger("ToolPages").LogError(ex, "Tool request failed");
                return Results.BadRequest(new { error = "Tool request failed. Check input and provider settings, then retry." });
            }
        });
        return group;
    }

    private static bool IsCli(CodexApiService api) =>
        !api.IsConfigured || api.Model.Contains("gemini", StringComparison.OrdinalIgnoreCase);

    private static async Task<IResult> TranslateAsync(
        TranslateInput body, CodexApiService api, AiProviderSettingsService settings,
        WebRepository repo, HttpContext context)
    {
        if (string.IsNullOrWhiteSpace(body.Text) && body.Images.Count == 0)
            return Results.BadRequest(new { error = "Enter text or add images." });
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(context.RequestAborted);
        timeout.CancelAfter(TimeSpan.FromMinutes(10));
        TranslationResult result;
        if (IsCli(api))
        {
            if (!settings.CodexCliEnabled)
                return Results.BadRequest(new { error = "CLI is disabled in settings." });
            string actor = $"user={context.User.Identity?.Name ?? "unknown"}; ip={context.Connection.RemoteIpAddress}";
            result = await TranslateWithCodexCliAsync(
                body.Text, body.TargetLanguage, body.Images, api.Model, actor, timeout.Token);
        }
        else
        {
            var value = await api.TranslateAsync(body.Text, body.TargetLanguage,
                body.Images.Select(i => (i.MediaType, i.Base64)).ToList(), timeout.Token);
            SaveCodexApiUsage(repo, value.Usage);
            result = new(value.SourceText, value.Translation, value.Usage, "Codex API", api.Model);
        }
        return Results.Ok(new { result.SourceText, result.Translation, usage = FormatUsage(result) });
    }

    private static async Task<IResult> ReportActionAsync(
        string action, ReportInput body, ReviewReportService service, HttpContext context)
    {
        var report = body.Report;
        var existing = string.IsNullOrEmpty(report.Id) ? null : service.Get(report.Id);
        if (!string.IsNullOrEmpty(report.Id) && existing is null)
            return Results.NotFound();
        report.CreatedBy = existing?.CreatedBy ?? context.User.Identity?.Name ?? "unknown";
        report.CreatedAt = existing?.CreatedAt ?? DateTime.Now;
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(context.RequestAborted);
        timeout.CancelAfter(TimeSpan.FromMinutes(10));
        switch (action)
        {
            case "generate":
                report.Doc = await service.GenerateAsync(report.InputText, report.Images,
                    report.Language, report.ExtraInstruction, timeout.Token);
                report.Revisions = [];
                break;
            case "revise":
                report.Doc = await service.ReviseAsync(report, body.Request ?? "", timeout.Token);
                report.Revisions.Add(new()
                {
                    At = DateTime.Now,
                    By = context.User.Identity?.Name ?? "unknown",
                    Request = body.Request ?? ""
                });
                break;
            case "preview":
                return Preview(report);
            case "markdown":
                return Results.Ok(new { text = ReviewReportService.ToMarkdown(report) });
            case "export":
                return Results.File(ReviewReportService.ExportXlsx(report),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "review-report.xlsx");
            case "save":
                break;
            default:
                return Results.NotFound();
        }
        if (action != "save")
            report.Model = $"{ReviewReportService.Model} · {ReviewReportService.Effort}";
        service.Save(report);
        return Results.Ok(report);
    }

    private static IResult Preview(ReviewReport report)
    {
        if (report.Doc is null) return Results.Ok(Array.Empty<object>());
        return Results.Ok(ReviewReportService.ResultTables(report.Doc).Select(table => new
        {
            table.Title,
            showDate = table.Arms.Any(a => !string.IsNullOrWhiteSpace(a.Date)),
            showModel = table.Arms.Any(a => !string.IsNullOrWhiteSpace(a.Model)),
            dateSpans = ReviewReportService.Spans(table.Arms.Select(a => a.Date.Trim()).ToList()),
            modelSpans = ReviewReportService.Spans(table.Arms.Select(a => a.Model.Trim()).ToList()),
            columns = table.Columns.Select(c => new { c.Header, c.HasBelow, c.BelowHeader }),
            rows = table.Arms.Select(a => new
            {
                a.Date,
                a.Model,
                a.Label,
                test = a.Role == "test",
                cells = table.Columns.Select(c => new
                {
                    main = c.Main.ValueOf(a.Id.Trim()),
                    below = c.Below(a.Id.Trim()),
                    c.HasBelow,
                    rate = c.Main.Kind == "rate"
                })
            })
        }));
    }

    private static async Task<IResult> GetExampleAsync(int kind, int index)
    {
        var candidates = GraphExamples.Where(e => (int)e.Kind == kind).ToArray();
        if (candidates.Length == 0)
            return Results.NotFound(new { error = "No example for this graph type." });
        var example = candidates[Math.Abs(index % candidates.Length)];
        string[] roots =
        [
            Path.GetDirectoryName(example.Path)!,
            Path.Combine(AppContext.BaseDirectory, "GraphMakerExamples"),
            Path.Combine(AppContext.BaseDirectory, "wwwroot", "GraphMakerExamples"),
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "GraphMakerExamples")
        ];
        var path = roots.Select(root => Path.Combine(root, Path.GetFileName(example.Path))).FirstOrDefault(File.Exists);
        if (path is null)
            return Results.NotFound(new { error = $"Example file not found: {Path.GetFileName(example.Path)}" });
        return Results.Ok(new
        {
            text = await File.ReadAllTextAsync(path),
            name = Path.GetFileName(path),
            example.HeaderNotUse
        });
    }

    private static readonly (GraphMakerGraphType Kind, string Path, bool HeaderNotUse)[] GraphExamples =
    [
        (GraphMakerGraphType.Line, @"D:\000. MyWorks\005. Program\GRAPH_spl.txt", false),
        (GraphMakerGraphType.Scatter, @"D:\000. MyWorks\005. Program\SCATTER_NOHEADER_BRKINNER TENSION.txt", true),
        (GraphMakerGraphType.Scatter, @"D:\000. MyWorks\005. Program\SCATTER_hookdimensiondata.txt", false),
        (GraphMakerGraphType.ProcessTrend, @"D:\000. MyWorks\005. Program\PROCESSTREND_gausstrend5.txt", false),
        (GraphMakerGraphType.HeatMap, @"C:\Users\jhbyun\Desktop\새 폴더 (5)\Merged_VP_Test_Data_records.json", false)
    ];

    private static async Task<TranslationResult> TranslateWithCodexCliAsync(
        string text,
        string targetLanguage,
        IReadOnlyList<PastedImage> images,
        string model,
        string actor,
        CancellationToken ct)
    {
        string workDir = FindRepoRoot();
        string tmpDir = Path.Combine(workDir, "tmp", "translate-codex", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tmpDir);

        string promptPath = Path.Combine(tmpDir, "prompt.txt");
        string scriptPath = Path.Combine(tmpDir, "run_codex_translate.ps1");
        var imageFiles = new List<string>();

        try
        {
            LogTranslate($"prepare tmp={tmpDir}", actor);
            for (int i = 0; i < images.Count; i++)
            {
                string ext = ExtensionFromMediaType(images[i].MediaType);
                string imagePath = Path.Combine(tmpDir, $"image_{i + 1:00}{ext}");
                byte[] bytes = Convert.FromBase64String(images[i].Base64);
                await File.WriteAllBytesAsync(imagePath, bytes, ct);
                imageFiles.Add(imagePath);
                LogTranslate($"wrote image {i + 1}/{images.Count}: {Path.GetFileName(imagePath)}, bytes={bytes.Length:N0}", actor);
            }

            string prompt = BuildCodexTranslatePrompt(text, targetLanguage, imageFiles);
            await File.WriteAllTextAsync(promptPath, prompt, new System.Text.UTF8Encoding(false), ct);
            await File.WriteAllTextAsync(scriptPath, BuildCodexTranslateScript(workDir, promptPath, imageFiles, model), new System.Text.UTF8Encoding(false), ct);
            LogTranslate($"prompt ready chars={prompt.Length:N0}, images={imageFiles.Count}, script={scriptPath}", actor);

            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powershell.exe",
                WorkingDirectory = workDir,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8,
            };
            psi.ArgumentList.Add("-NoProfile");
            psi.ArgumentList.Add("-ExecutionPolicy");
            psi.ArgumentList.Add("Bypass");
            psi.ArgumentList.Add("-File");
            psi.ArgumentList.Add(scriptPath);

            using var proc = new System.Diagnostics.Process { StartInfo = psi };
            LogTranslate("starting CLI process", actor);
            proc.Start();
            ChildProcessJob.Assign(proc);
            LogTranslate($"started pid={proc.Id}", actor);

            Task<string> stdoutTask = proc.StandardOutput.ReadToEndAsync(ct);
            Task<string> stderrTask = proc.StandardError.ReadToEndAsync(ct);

            using var killReg = ct.Register(() =>
            {
                try { if (!proc.HasExited) proc.Kill(entireProcessTree: true); }
                catch { }
            });

            await proc.WaitForExitAsync(ct);
            string stdout = await stdoutTask;
            string stderr = await stderrTask;
            LogTranslate($"process exited code={proc.ExitCode}, stdoutChars={stdout.Length:N0}, stderrChars={stderr.Length:N0}", actor);
            if (!string.IsNullOrWhiteSpace(stderr))
                LogTranslate($"stderr: {TrimForLog(stderr)}", actor);

            if (proc.ExitCode != 0)
                throw new InvalidOperationException(string.IsNullOrWhiteSpace(stderr)
                    ? $"CLI exited with code {proc.ExitCode}."
                    : $"CLI exited with code {proc.ExitCode}: {stderr.Trim()}");

            TranslationResult result = ParseCodexTranslationResult(stdout, model);
            if (string.IsNullOrWhiteSpace(result.Translation))
                throw new InvalidOperationException("CLI returned an empty translation.");

            return result;
        }
        finally
        {
            try { if (Directory.Exists(tmpDir)) Directory.Delete(tmpDir, recursive: true); } catch { }
            LogTranslate($"cleanup tmp={tmpDir}", actor);
        }
    }

    private static void LogTranslate(string message, string actor)
        => Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [Translate] [{actor}] {message}");

    private static string FormatUsage(TranslationResult result)
    {
        if (result.Usage is null)
            return $"{result.Provider} {result.Model}: token usage unavailable";

        return $"{result.Provider} {result.Model}: input {FormatToken(result.Usage.InputTokens)}, output {FormatToken(result.Usage.OutputTokens)}, total {FormatToken(result.Usage.TotalTokens)} tokens";
    }

    private static string FormatToken(int? tokens)
        => tokens.HasValue ? tokens.Value.ToString("N0") : "-";

    private static void SaveCodexApiUsage(WebRepository Repo, CodexTokenUsage? usage)
    {
        if (usage is null) return;

        long input = ReadLongSetting(Repo, "AiUsages:CodexApiInputTokens") + (usage.InputTokens ?? 0);
        long output = ReadLongSetting(Repo, "AiUsages:CodexApiOutputTokens") + (usage.OutputTokens ?? 0);
        long total = ReadLongSetting(Repo, "AiUsages:CodexApiTotalTokens") + (usage.TotalTokens ?? ((usage.InputTokens ?? 0) + (usage.OutputTokens ?? 0)));
        string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        Repo.SetSetting("AiUsages:CodexApiInputTokens", input.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Repo.SetSetting("AiUsages:CodexApiOutputTokens", output.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Repo.SetSetting("AiUsages:CodexApiTotalTokens", total.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Repo.SetSetting("AiUsages:CodexApiLastRequestAt", ts);
    }

    private static long ReadLongSetting(WebRepository Repo, string key)
    {
        string value = Repo.GetSetting(key) ?? "";
        value = value.Replace(",", "").Trim();
        return long.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out long result)
            ? result
            : 0L;
    }

    private static string TrimForLog(string text)
    {
        text = (text ?? "").Trim();
        const int Max = 4000;
        return text.Length <= Max ? text : text[..Max] + "...";
    }

    private static string BuildCodexTranslatePrompt(string text, string targetLanguage, IReadOnlyList<string> imageFiles)
    {
        string images = imageFiles.Count == 0
            ? "(none)"
            : string.Join(Environment.NewLine, imageFiles.Select((path, i) => $"{i + 1}. {path}"));

        return AiPromptRegistry.Render(
            "translation/codex-ocr-translation.md",
            ("targetLanguage", targetLanguage),
            ("imageFiles", images),
            ("sourceText", string.IsNullOrWhiteSpace(text) ? "(none)" : text.Trim()));
    }

    private static string BuildCodexTranslateScript(string workDir, string promptPath, IReadOnlyList<string> imageFiles, string model)
    {
        static string Q(string value) => "'" + (value ?? "").Replace("'", "''") + "'";
        var imgCmds = new List<string>();
        foreach (var img in imageFiles)
        {
            imgCmds.Add("-i");
            imgCmds.Add(Q(img));
        }
        string imageArgsStr = imgCmds.Count > 0 ? string.Join(" ", imgCmds) : "";

        return string.Join(Environment.NewLine, new[]
        {
            "$ErrorActionPreference = 'Stop'",
            "$utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false",
            "[Console]::InputEncoding = $utf8NoBom",
            "[Console]::OutputEncoding = $utf8NoBom",
            "$OutputEncoding = $utf8NoBom",
            "try { chcp 65001 | Out-Null } catch {}",
            $"Set-Location -LiteralPath {Q(workDir)}",
            $"$prompt = Get-Content -LiteralPath {Q(promptPath)} -Raw -Encoding UTF8",
            "if (Get-Command agy -ErrorAction SilentlyContinue) {",
            $"    & agy --model {Q(model)} --effort high --dangerously-skip-permissions --prompt $prompt",
            "} else {",
            $"    $prompt | & codex exec {imageArgsStr} -c 'model_reasoning_effort=\"high\"' --cd {Q(workDir)} --dangerously-bypass-approvals-and-sandbox --skip-git-repo-check -",
            "}"
        });
    }

    private static string CleanCodexOutput(string output)
    {
        string text = (output ?? "").Trim();
        int firstFence = text.IndexOf("```", StringComparison.Ordinal);
        if (firstFence >= 0)
        {
            int firstNewline = text.IndexOf('\n', firstFence);
            int lastFence = text.LastIndexOf("```", StringComparison.Ordinal);
            if (firstNewline >= 0 && lastFence > firstNewline)
            {
                text = text[(firstNewline + 1)..lastFence];
            }
        }
        return text.Trim();
    }

    private static TranslationResult ParseCodexTranslationResult(string output, string model)
    {
        string text = CleanCodexOutput(output);
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(text);
            var root = doc.RootElement;
            string source = root.TryGetProperty("sourceText", out var s) ? s.GetString() ?? "" : "";
            string translation = "";
            if (root.TryGetProperty("translation", out var t) && !string.IsNullOrWhiteSpace(t.GetString()))
                translation = t.GetString()!;
            else if (root.TryGetProperty("translatedText", out var tt) && !string.IsNullOrWhiteSpace(tt.GetString()))
                translation = tt.GetString()!;
            else if (root.TryGetProperty("result", out var tr) && !string.IsNullOrWhiteSpace(tr.GetString()))
                translation = tr.GetString()!;

            return new TranslationResult(source.Trim(), translation.Trim(), null, "Antigravity CLI", model);
        }
        catch
        {
            int firstBrace = text.IndexOf('{');
            int lastBrace = text.LastIndexOf('}');
            if (firstBrace >= 0 && lastBrace > firstBrace)
            {
                try
                {
                    string subJson = text[firstBrace..(lastBrace + 1)];
                    using var doc = System.Text.Json.JsonDocument.Parse(subJson);
                    var root = doc.RootElement;
                    string source = root.TryGetProperty("sourceText", out var s) ? s.GetString() ?? "" : "";
                    string translation = "";
                    if (root.TryGetProperty("translation", out var t) && !string.IsNullOrWhiteSpace(t.GetString()))
                        translation = t.GetString()!;
                    else if (root.TryGetProperty("translatedText", out var tt) && !string.IsNullOrWhiteSpace(tt.GetString()))
                        translation = tt.GetString()!;
                    else if (root.TryGetProperty("result", out var tr) && !string.IsNullOrWhiteSpace(tr.GetString()))
                        translation = tr.GetString()!;

                    if (!string.IsNullOrWhiteSpace(translation))
                        return new TranslationResult(source.Trim(), translation.Trim(), null, "Antigravity CLI", model);
                }
                catch { }
            }

            return new TranslationResult("", text, null, "Antigravity CLI", model);
        }
    }

    private static string ExtensionFromMediaType(string mediaType)
        => mediaType.ToLowerInvariant() switch
        {
            "image/jpeg" => ".jpg",
            "image/jpg" => ".jpg",
            "image/webp" => ".webp",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            _ => ".png",
        };

    private static string FindRepoRoot()
    {
        string dir = AppContext.BaseDirectory;
        for (int i = 0; i < 8; i++)
        {
            if (File.Exists(Path.Combine(dir, "PortalGNLAB.sln"))) return dir;
            string? parent = Path.GetDirectoryName(dir.TrimEnd('\\', '/'));
            if (string.IsNullOrEmpty(parent) || parent == dir) break;
            dir = parent;
        }

        string current = Directory.GetCurrentDirectory();
        dir = current;
        for (int i = 0; i < 8; i++)
        {
            if (File.Exists(Path.Combine(dir, "PortalGNLAB.sln"))) return dir;
            string? parent = Path.GetDirectoryName(dir.TrimEnd('\\', '/'));
            if (string.IsNullOrEmpty(parent) || parent == dir) break;
            dir = parent;
        }

        return current;
    }


}
