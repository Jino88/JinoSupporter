using System.Data;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;
using PortalGNLAB.Web.Services;
namespace PortalGNLAB.Web.Endpoints;

public static class AdminToolEndpoints
{
    // State logic below is ported verbatim from the Blazor pages; services remain unchanged.
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);
    private static object View(object state) => state.GetType().GetFields().Where(f => !f.IsStatic)
        .ToDictionary(f => f.Name.TrimStart('_'), f => f.GetValue(state));
    private static void Apply(object state, JsonElement body, string[] allowed)
    {
        foreach (var key in allowed)
            if (body.TryGetProperty(key, out var value))
            {
                var field = state.GetType().GetField("_" + key)!;
                field.SetValue(state, value.Deserialize(field.FieldType, Json));
            }
    }
    public static void MapAdminTools(this WebApplication app)
    {
        var api = app.MapGroup("/api/portal/admin").RequireAuthorization(p => p.RequireRole(AppRoles.Admin));
        api.AddEndpointFilter(async (ctx, next) => {
            if (!HttpMethods.IsGet(ctx.HttpContext.Request.Method) && ctx.HttpContext.Request.Headers["X-Requested-With"] != "PortalGNLAB")
                return Results.BadRequest(new { error = "Required request header is missing." });
            return await next(ctx);
        });
        api.MapGet("/ai-usages", (WebRepository repo, ClaudeUsageScraper claude, CodexUsageScraper codex) => {
            var s = new AiState(repo, claude, codex); s.Initialize(); return AiView(s);
        });
        api.MapPost("/ai-usages/{action}", async (string action, JsonElement body, WebRepository repo, ClaudeUsageScraper claude, CodexUsageScraper codex) => {
            var s = new AiState(repo, claude, codex); s.Initialize();
            Apply(s, body, ["apiCreditBalance","sessionUsage","weeklyAllModels","weeklySonnet","weeklyClaudeDesign","extraUsage","extraBalance","codexPlanUsage","codexPlanLimitUsage","codexPlanNote","codexApiInputTokens","codexApiOutputTokens","codexApiTotalTokens"]);
            switch(action) {
                case "save-api": await s.SaveApiAsync(); break;
                case "save-plan": await s.SavePlanAsync(); break;
                case "save-codex-plan": await s.SaveCodexPlanAsync(); break;
                case "save-codex-api": await s.SaveCodexApiAsync(); break;
                case "claude-login": await s.LoginAsync(); break;
                case "claude-refresh": await s.RefreshAsync(); break;
                case "claude-delete": s.DeleteSession(); break;
                case "codex-login": await s.LoginCodexAsync(); break;
                case "codex-refresh": await s.RefreshCodexAsync(); break;
                case "codex-delete": s.DeleteCodexSession(); break;
                default: return Results.NotFound();
            }
            return Results.Ok(AiView(s));
        });
        api.MapGet("/db-query", (AppPathsService paths) => { var s = new DbState(paths); s.Initialize(); return View(s); });
        api.MapPost("/db-query/{action}", async (string action, JsonElement body, AppPathsService paths) => {
            var s = new DbState(paths); s.Initialize();
            Apply(s, body, ["server","port","database","userId","password","timeoutSeconds","encrypt","trustServerCertificate","query","saveToLocalDb","localTableName"]);
            if (action == "test") await s.TestConnectionAsync();
            else if(action == "query") await s.RunQueryAsync();
            else return Results.NotFound();
            return Results.Ok(View(s));
        });
        api.MapGet("/test-excel-converter", (AppPathsService paths, ExcelHelperRunner helper) => new {
            sourceDir = paths.Current.TestExcelConverterSourceDir, destDir = paths.Current.TestExcelConverterDestDir,
            helper.ConverterExists, helper.ConverterStatus, helper.FolderPickerExists });
        api.MapPost("/test-excel-converter/settings", (ExcelSettings body, AppPathsService paths) => {
            var config = JsonSerializer.Deserialize<AppPathsConfig>(JsonSerializer.Serialize(paths.Current))!;
            config.TestExcelConverterSourceDir = body.SourceDir.Trim(); config.TestExcelConverterDestDir = body.DestDir.Trim(); paths.Save(config);
            return Results.Ok(body);
        });
        api.MapPost("/test-excel-converter/pick", async (PickFolder body, ExcelHelperRunner helper, CancellationToken ct) =>
            Results.Ok(new { path = await helper.PickFolderAsync(body.Initial, ct) }));
        api.MapPost("/test-excel-converter/convert", async (HttpContext ctx, ExcelHelperRunner helper) => {
            if (!helper.ConverterExists) { ctx.Response.StatusCode = 400; await ctx.Response.WriteAsJsonAsync(new { error = helper.ConverterStatus }); return; }
            var form = await ctx.Request.ReadFormAsync(ctx.RequestAborted);
            if (form.Files.Count is < 1 or > 200 || string.IsNullOrWhiteSpace(form["destDir"])) { ctx.Response.StatusCode = 400; return; }
            string folder = Path.Combine(Path.GetTempPath(), "portal-admin-excel", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            try {
                var files = new List<string>();
                foreach(var file in form.Files) {
                    if (file.Length > 1024L*1024*1024 || !new[]{".xlsx",".xlsm",".xlsb",".xls"}.Contains(Path.GetExtension(file.FileName).ToLowerInvariant())) continue;
                    string path = Path.Combine(folder, Guid.NewGuid().ToString("N") + Path.GetExtension(file.FileName));
                    await using(var output = File.Create(path)) await file.CopyToAsync(output, ctx.RequestAborted);
                    files.Add(path);
                }
                ctx.Response.ContentType = "application/x-ndjson";
                await foreach(var evt in helper.CleanFilesToDestinationRootAsync(files, form["destDir"].ToString(), form["verbose"] == "true", ctx.RequestAborted)) {
                    await ctx.Response.WriteAsync(JsonSerializer.Serialize(evt.Raw) + "\n", ctx.RequestAborted); await ctx.Response.Body.FlushAsync(ctx.RequestAborted);
                }
            } finally { Directory.Delete(folder, true); }
        });
    }
    public sealed record ExcelSettings(string SourceDir, string DestDir);
    public sealed record PickFolder(string? Initial);
    private static object AiView(AiState s) => new { values = View(s), s.ClaudeSaved, s.CodexSaved, s.ClaudeProfile, s.CodexProfile, s.DebugDirectory };
public sealed class AiState(WebRepository Repo, ClaudeUsageScraper Scraper, CodexUsageScraper CodexScraper) {

    public string  _apiCreditBalance = string.Empty;
    public string? _apiUpdatedAt;
    public bool    _apiSaving;
    public bool    _apiSaved;
    public string  _apiError = string.Empty;

    public string  _sessionUsage       = string.Empty;
    public string  _weeklyAllModels    = string.Empty;
    public string  _weeklySonnet       = string.Empty;
    public string  _weeklyClaudeDesign = string.Empty;
    public string  _extraUsage         = string.Empty;
    public string  _extraBalance       = string.Empty;
    public string? _planUpdatedAt;
    public bool    _planSaving;
    public bool    _planSaved;
    public string  _planError          = string.Empty;

    public string  _codexPlanUsage      = string.Empty;
    public string  _codexPlanLimitUsage = string.Empty;
    public string  _codexPlanNote       = string.Empty;
    public string? _codexPlanUpdatedAt;
    public bool    _codexPlanSaving;
    public bool    _codexPlanSaved;
    public string  _codexPlanError      = string.Empty;

    public string  _codexApiInputTokens  = string.Empty;
    public string  _codexApiOutputTokens = string.Empty;
    public string  _codexApiTotalTokens  = string.Empty;
    public string? _codexApiLastRequestAt;
    public bool    _codexApiSaving;
    public bool    _codexApiSaved;
    public string  _codexApiError        = string.Empty;
    public string  _codexPlanScrapeError = string.Empty;
    public string  _codexApiScrapeError  = string.Empty;
    public string  _codexPlanSnippet     = string.Empty;
    public string  _codexApiSnippet      = string.Empty;

    public bool   _loginBusy;
    public bool   _scrapeBusy;
    public string _sessionStatus      = string.Empty;
    public bool   _sessionError;
    public bool   _imported;
    public string _billingScrapeError = string.Empty;
    public string _usageScrapeError   = string.Empty;
    public string _billingSnippet     = string.Empty;
    public string _usageSnippet       = string.Empty;

    public bool _codexLoginBusy;
    public bool _codexScrapeBusy;
    public string _codexSessionStatus = string.Empty;
    public bool _codexSessionError;

    public void Initialize()
    {
        _apiCreditBalance = Repo.GetSetting("AiUsages:ApiCreditBalance") ?? string.Empty;
        _apiUpdatedAt     = Repo.GetSetting("AiUsages:ApiUpdatedAt");

        _sessionUsage       = Repo.GetSetting("AiUsages:SessionUsage")       ?? string.Empty;
        _weeklyAllModels    = Repo.GetSetting("AiUsages:WeeklyAllModels")    ?? string.Empty;
        _weeklySonnet       = Repo.GetSetting("AiUsages:WeeklySonnet")       ?? string.Empty;
        _weeklyClaudeDesign = Repo.GetSetting("AiUsages:WeeklyClaudeDesign") ?? string.Empty;
        _extraUsage         = Repo.GetSetting("AiUsages:ExtraUsage")         ?? string.Empty;
        _extraBalance       = Repo.GetSetting("AiUsages:ExtraBalance")       ?? string.Empty;
        _planUpdatedAt      = Repo.GetSetting("AiUsages:PlanUpdatedAt");

        _codexPlanUsage      = Repo.GetSetting("AiUsages:CodexPlanUsage")      ?? string.Empty;
        _codexPlanLimitUsage = Repo.GetSetting("AiUsages:CodexPlanLimitUsage") ?? string.Empty;
        _codexPlanNote       = Repo.GetSetting("AiUsages:CodexPlanNote")       ?? string.Empty;
        _codexPlanUpdatedAt  = Repo.GetSetting("AiUsages:CodexPlanUpdatedAt");

        _codexApiInputTokens   = CleanTokenSetting(Repo.GetSetting("AiUsages:CodexApiInputTokens"));
        _codexApiOutputTokens  = CleanTokenSetting(Repo.GetSetting("AiUsages:CodexApiOutputTokens"));
        _codexApiTotalTokens   = CleanTokenSetting(Repo.GetSetting("AiUsages:CodexApiTotalTokens"));
        _codexApiLastRequestAt = Repo.GetSetting("AiUsages:CodexApiLastRequestAt");
    }

    public async Task LoginAsync()
    {
        _loginBusy = true;
        _sessionError = false;
        _sessionStatus = "Opening browser. Sign in on both tabs, then close it.";

        try
        {
            await Scraper.OpenLoginBrowserAsync();
            _sessionStatus = "Session saved. Fetching once without opening another browser window.";
            await RefreshCoreAsync(autoTriggered: true);
        }
        catch (Exception ex)
        {
            _sessionError = true;
            _sessionStatus = $"Login failed: {ex.Message}";
        }
        finally { _loginBusy = false; }
    }

    public Task RefreshAsync() => RefreshCoreAsync(autoTriggered: false);

    public async Task RefreshCoreAsync(bool autoTriggered)
    {
        if (_scrapeBusy || !Scraper.SessionProfileExists) return;

        _scrapeBusy = true;
        _sessionError = false;
        _sessionStatus = autoTriggered ? "Auto-fetching Claude usage..." : "Fetching Claude usage...";
        _billingScrapeError = _usageScrapeError = string.Empty;


        try
        {
            var result = await Scraper.ScrapeAsync();
            if (!autoTriggered && NeedsVisibleLogin(result.BillingError, result.UsageError))
            {
                _sessionStatus = "Hidden refresh needs login or security check. Opening Claude browser...";

                await Scraper.OpenLoginBrowserAsync();
                result = await Scraper.ScrapeAsync();
            }

            if (!string.IsNullOrEmpty(result.ApiCreditBalance))   _apiCreditBalance = result.ApiCreditBalance;
            if (!string.IsNullOrEmpty(result.SessionUsage))       _sessionUsage = result.SessionUsage;
            if (!string.IsNullOrEmpty(result.WeeklyAllModels))    _weeklyAllModels = result.WeeklyAllModels;
            if (!string.IsNullOrEmpty(result.WeeklySonnet))       _weeklySonnet = result.WeeklySonnet;
            if (!string.IsNullOrEmpty(result.WeeklyClaudeDesign)) _weeklyClaudeDesign = result.WeeklyClaudeDesign;
            if (!string.IsNullOrEmpty(result.ExtraUsage))         _extraUsage = result.ExtraUsage;
            if (!string.IsNullOrEmpty(result.ExtraBalance))       _extraBalance = result.ExtraBalance;

            _billingScrapeError = result.BillingError ?? string.Empty;
            _usageScrapeError   = result.UsageError   ?? string.Empty;
            _billingSnippet     = result.BillingRawSnippet;
            _usageSnippet       = result.UsageRawSnippet;

            SaveFetchedClaudeValues();
            _imported = true;
            _sessionStatus = result.BillingError is null && result.UsageError is null
                ? "Claude auto-fetch complete and saved."
                : "Some Claude pages could not be fetched. Saved available values.";
            _sessionError = result.BillingError is not null || result.UsageError is not null;
            if (NeedsVisibleLogin(result.BillingError, result.UsageError))
            {
                _sessionStatus = autoTriggered
                    ? "Claude hidden refresh needs login or security check. Click Claude Login once."
                    : "Claude refresh still needs login or security check. Click Claude Login once.";
            }
        }
        catch (Exception ex)
        {
            _sessionError = true;
            _sessionStatus = $"Refresh failed: {ex.Message}";
        }
        finally
        {
            _scrapeBusy = false;

        }
    }

    public async Task LoginCodexAsync()
    {
        _codexLoginBusy = true;
        _codexSessionError = false;
        _codexSessionStatus = "Opening browser. Sign in to ChatGPT and OpenAI Platform, then close it.";

        try
        {
            await CodexScraper.OpenLoginBrowserAsync();
            _codexSessionStatus = "Codex/OpenAI session saved. Fetching once without opening another browser window.";
            await RefreshCodexCoreAsync(autoTriggered: true);
        }
        catch (Exception ex)
        {
            _codexSessionError = true;
            _codexSessionStatus = $"Codex/OpenAI login failed: {ex.Message}";
        }
        finally { _codexLoginBusy = false; }
    }

    public Task RefreshCodexAsync() => RefreshCodexCoreAsync(autoTriggered: false);

    public async Task RefreshCodexCoreAsync(bool autoTriggered)
    {
        if (_codexScrapeBusy || !CodexScraper.SessionProfileExists) return;

        _codexScrapeBusy = true;
        _codexSessionError = false;
        _codexSessionStatus = autoTriggered ? "Auto-fetching Codex/OpenAI usage..." : "Fetching Codex/OpenAI usage...";
        _codexPlanScrapeError = _codexApiScrapeError = string.Empty;


        try
        {
            var result = await CodexScraper.ScrapeAsync();
            if (!autoTriggered && NeedsVisibleLogin(result.PlanError, result.ApiError))
            {
                _codexSessionStatus = "Hidden refresh needs login or security check. Opening Codex/OpenAI browser...";

                await CodexScraper.OpenLoginBrowserAsync();
                result = await CodexScraper.ScrapeAsync();
            }

            if (!string.IsNullOrEmpty(result.PlanUsage))      _codexPlanUsage = result.PlanUsage;
            if (!string.IsNullOrEmpty(result.PlanLimitUsage)) _codexPlanLimitUsage = result.PlanLimitUsage;
            if (!string.IsNullOrEmpty(result.PlanNote))       _codexPlanNote = result.PlanNote;
            if (IsPlainTokenCount(result.ApiInputTokens))      _codexApiInputTokens = result.ApiInputTokens;
            if (IsPlainTokenCount(result.ApiOutputTokens))     _codexApiOutputTokens = result.ApiOutputTokens;
            if (IsPlainTokenCount(result.ApiTotalTokens))      _codexApiTotalTokens = result.ApiTotalTokens;

            _codexPlanScrapeError = result.PlanError ?? string.Empty;
            _codexApiScrapeError  = result.ApiError  ?? string.Empty;
            _codexPlanSnippet     = result.PlanRawSnippet;
            _codexApiSnippet      = result.ApiRawSnippet;

            SaveFetchedCodexValues();
            _codexSessionStatus = result.PlanError is null && result.ApiError is null
                ? "Codex/OpenAI auto-fetch complete and saved."
                : "Some Codex/OpenAI pages could not be fetched. Saved available values.";
            _codexSessionError = result.PlanError is not null || result.ApiError is not null;

            if (NeedsVisibleLogin(result.PlanError, result.ApiError))
            {
                _codexSessionStatus = autoTriggered
                    ? "Codex/OpenAI hidden refresh needs login or security check. Click Codex/OpenAI Login once."
                    : "Codex/OpenAI refresh still needs login or security check. Click Codex/OpenAI Login once.";
            }
        }
        catch (Exception ex)
        {
            _codexSessionError = true;
            _codexSessionStatus = $"Codex/OpenAI refresh failed: {ex.Message}";
        }
        finally
        {
            _codexScrapeBusy = false;

        }
    }

    public void DeleteSession()
    {
        Scraper.DeleteSession();
        _sessionStatus = "Session deleted. Please Login again.";
        _sessionError = false;
        _imported = false;
    }

    public void DeleteCodexSession()
    {
        CodexScraper.DeleteSession();
        _codexSessionStatus = "Codex/OpenAI session deleted. Please Login again.";
        _codexSessionError = false;
    }

    public void SaveFetchedClaudeValues()
    {
        string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        Repo.SetSetting("AiUsages:ApiCreditBalance", (_apiCreditBalance ?? "").Trim());
        Repo.SetSetting("AiUsages:ApiUpdatedAt", ts);
        Repo.SetSetting("AiUsages:SessionUsage",       (_sessionUsage       ?? "").Trim());
        Repo.SetSetting("AiUsages:WeeklyAllModels",    (_weeklyAllModels    ?? "").Trim());
        Repo.SetSetting("AiUsages:WeeklySonnet",       (_weeklySonnet       ?? "").Trim());
        Repo.SetSetting("AiUsages:WeeklyClaudeDesign", (_weeklyClaudeDesign ?? "").Trim());
        Repo.SetSetting("AiUsages:ExtraUsage",         (_extraUsage         ?? "").Trim());
        Repo.SetSetting("AiUsages:ExtraBalance",       (_extraBalance       ?? "").Trim());
        Repo.SetSetting("AiUsages:PlanUpdatedAt", ts);
        _apiUpdatedAt = ts;
        _planUpdatedAt = ts;
    }

    public void SaveFetchedCodexValues()
    {
        string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        Repo.SetSetting("AiUsages:CodexPlanUsage",      (_codexPlanUsage      ?? "").Trim());
        Repo.SetSetting("AiUsages:CodexPlanLimitUsage", (_codexPlanLimitUsage ?? "").Trim());
        Repo.SetSetting("AiUsages:CodexPlanNote",       (_codexPlanNote       ?? "").Trim());
        Repo.SetSetting("AiUsages:CodexPlanUpdatedAt", ts);
        Repo.SetSetting("AiUsages:CodexApiInputTokens",  (_codexApiInputTokens  ?? "").Trim());
        Repo.SetSetting("AiUsages:CodexApiOutputTokens", (_codexApiOutputTokens ?? "").Trim());
        Repo.SetSetting("AiUsages:CodexApiTotalTokens",  (_codexApiTotalTokens  ?? "").Trim());
        Repo.SetSetting("AiUsages:CodexApiLastRequestAt", ts);
        _codexPlanUpdatedAt = ts;
        _codexApiLastRequestAt = ts;
    }

    public async Task SaveApiAsync()
    {
        _apiSaving = true; _apiSaved = false; _apiError = string.Empty;
        try
        {
            string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Repo.SetSetting("AiUsages:ApiCreditBalance", (_apiCreditBalance ?? "").Trim());
            Repo.SetSetting("AiUsages:ApiUpdatedAt", ts);
            _apiUpdatedAt = ts;
            _apiSaved = true;
        }
        catch (Exception ex) { _apiError = ex.Message; }
        finally { _apiSaving = false; }
        await Task.CompletedTask;
    }

    public async Task SavePlanAsync()
    {
        _planSaving = true; _planSaved = false; _planError = string.Empty;
        try
        {
            string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Repo.SetSetting("AiUsages:SessionUsage",       (_sessionUsage       ?? "").Trim());
            Repo.SetSetting("AiUsages:WeeklyAllModels",    (_weeklyAllModels    ?? "").Trim());
            Repo.SetSetting("AiUsages:WeeklySonnet",       (_weeklySonnet       ?? "").Trim());
            Repo.SetSetting("AiUsages:WeeklyClaudeDesign", (_weeklyClaudeDesign ?? "").Trim());
            Repo.SetSetting("AiUsages:ExtraUsage",         (_extraUsage         ?? "").Trim());
            Repo.SetSetting("AiUsages:ExtraBalance",       (_extraBalance       ?? "").Trim());
            Repo.SetSetting("AiUsages:PlanUpdatedAt", ts);
            _planUpdatedAt = ts;
            _planSaved = true;
        }
        catch (Exception ex) { _planError = ex.Message; }
        finally { _planSaving = false; }
        await Task.CompletedTask;
    }

    public async Task SaveCodexPlanAsync()
    {
        _codexPlanSaving = true; _codexPlanSaved = false; _codexPlanError = string.Empty;
        try
        {
            string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Repo.SetSetting("AiUsages:CodexPlanUsage",      (_codexPlanUsage      ?? "").Trim());
            Repo.SetSetting("AiUsages:CodexPlanLimitUsage", (_codexPlanLimitUsage ?? "").Trim());
            Repo.SetSetting("AiUsages:CodexPlanNote",       (_codexPlanNote       ?? "").Trim());
            Repo.SetSetting("AiUsages:CodexPlanUpdatedAt", ts);
            _codexPlanUpdatedAt = ts;
            _codexPlanSaved = true;
        }
        catch (Exception ex) { _codexPlanError = ex.Message; }
        finally { _codexPlanSaving = false; }
        await Task.CompletedTask;
    }

    public async Task SaveCodexApiAsync()
    {
        _codexApiSaving = true; _codexApiSaved = false; _codexApiError = string.Empty;
        try
        {
            Repo.SetSetting("AiUsages:CodexApiInputTokens",  (_codexApiInputTokens  ?? "").Trim());
            Repo.SetSetting("AiUsages:CodexApiOutputTokens", (_codexApiOutputTokens ?? "").Trim());
            Repo.SetSetting("AiUsages:CodexApiTotalTokens",  (_codexApiTotalTokens  ?? "").Trim());
            if (string.IsNullOrWhiteSpace(_codexApiLastRequestAt))
                _codexApiLastRequestAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Repo.SetSetting("AiUsages:CodexApiLastRequestAt", _codexApiLastRequestAt);
            _codexApiSaved = true;
        }
        catch (Exception ex) { _codexApiError = ex.Message; }
        finally { _codexApiSaving = false; }
        await Task.CompletedTask;
    }

    public static string FormatTs(string? ts)
        => string.IsNullOrWhiteSpace(ts) ? "(no record)" : ts;

    public static bool IsSessionExpiredError(string? message)
        => !string.IsNullOrWhiteSpace(message) &&
           (message.Contains("session expired", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("signed out", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("Login again", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("/login", StringComparison.OrdinalIgnoreCase));

    public static bool IsBrowserChallengeError(string? message)
        => !string.IsNullOrWhiteSpace(message) &&
           (message.Contains("Cloudflare", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("bot detection", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("CAPTCHA", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("security check", StringComparison.OrdinalIgnoreCase));

    public static bool NeedsVisibleLogin(params string?[] messages)
        => messages.Any(message => IsSessionExpiredError(message) || IsBrowserChallengeError(message));

    public static bool IsPlainTokenCount(string? value)
        => !string.IsNullOrWhiteSpace(value) &&
           value.All(ch => char.IsDigit(ch) || ch == ',');

    public static string CleanTokenSetting(string? value)
        => IsPlainTokenCount(value) ? value!.Trim() : string.Empty;

    public static double? TryParsePct(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        string t = s.Trim().TrimEnd('%').Trim();
        return double.TryParse(t, System.Globalization.NumberStyles.Float,
                               System.Globalization.CultureInfo.InvariantCulture, out double v)
            ? v : null;
    }

    public ValueTask DisposeAsync() => ValueTask.CompletedTask;

 public bool ClaudeSaved => Scraper.SessionProfileExists; public bool CodexSaved => CodexScraper.SessionProfileExists; public string ClaudeProfile => Scraper.ProfileDirectory; public string CodexProfile => CodexScraper.ProfileDirectory; public string DebugDirectory => Scraper.DebugDirectory;

}public sealed class DbState(AppPathsService AppPaths) {

    public const int _maxRows = 500;

    public string _server = string.Empty;
    public int _port = 1433;
    public string _database = "master";
    public string _userId = string.Empty;
    public string _password = string.Empty;
    public int _timeoutSeconds = 15;
    public bool _encrypt = true;
    public bool _trustServerCertificate = true;
    public bool _showPassword;
    public bool _saveToLocalDb;
    public string _localTableName = "ImportedQueryResult";

    public string _query = "SELECT TOP (100) name, database_id, create_date FROM sys.databases ORDER BY name;";
    public bool _busy;
    public string _busyAction = string.Empty;
    public bool _connectionOk;
    public string _connectionMessage = string.Empty;
    public string _resultMessage = string.Empty;
    public string _error = string.Empty;
    public List<string> _columns = [];
    public List<List<string>> _rows = [];

    public void Initialize()
    {
        LoadSavedConnection();
    }

    public async Task TestConnectionAsync()
    {
        ClearMessages();
        _busy = true;
        _busyAction = "test";
        try
        {
            await using var conn = new SqlConnection(BuildConnectionString());
            await conn.OpenAsync();
            SaveConnectionSettings();
            _connectionOk = true;
            _connectionMessage = "Connected. Connection saved.";
        }
        catch (Exception ex)
        {
            _connectionOk = false;
            _connectionMessage = "Connection failed.";
            _error = ex.Message;
        }
        finally
        {
            _busy = false;
            _busyAction = string.Empty;
        }
    }

    public async Task RunQueryAsync()
    {
        ClearMessages();
        _columns = [];
        _rows = [];

        if (string.IsNullOrWhiteSpace(_query))
        {
            _error = "Query is required.";
            return;
        }
        if (!IsReadOnlyQuery(_query))
        {
            _error = "Only read-only SELECT/WITH queries are allowed. Server write commands are blocked.";
            return;
        }

        _busy = true;
        _busyAction = "query";
        try
        {
            await using var conn = new SqlConnection(BuildConnectionString());
            await conn.OpenAsync();
            SaveConnectionSettings();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = _query;
            cmd.CommandTimeout = Math.Clamp(_timeoutSeconds, 1, 120);

            await using var reader = await cmd.ExecuteReaderAsync(CommandBehavior.SequentialAccess);
            _columns = BuildUniqueColumnNames(reader);

            LocalTableWriter? localWriter = null;
            if (_saveToLocalDb)
            {
                if (_columns.Count == 0)
                    throw new InvalidOperationException("Local DB Save requires a query result set.");
                string tableName = SanitizeIdentifier(_localTableName);
                localWriter = CreateLocalTableWriter(AppPaths.Current.MainDbPath, tableName, _columns);
                _localTableName = tableName;
            }

            long totalRows = 0;
            try
            {
                while (await reader.ReadAsync())
                {
                    var row = new List<string>(reader.FieldCount);
                    for (int i = 0; i < reader.FieldCount; i++)
                        row.Add(await FormatCellAsync(reader, i));

                    if (_rows.Count < _maxRows)
                        _rows.Add(row);

                    if (localWriter is not null)
                        localWriter.Insert(row);

                    totalRows++;
                }

                if (localWriter is not null)
                    await localWriter.CommitAsync();
            }
            finally
            {
                if (localWriter is not null)
                    await localWriter.DisposeAsync();
            }

            _connectionOk = true;
            _connectionMessage = "Connected. Connection saved.";
            _resultMessage = _columns.Count == 0
                ? $"Query completed. Rows affected: {reader.RecordsAffected:N0}."
                : _saveToLocalDb
                    ? $"Loaded preview {_rows.Count:N0} row(s). Saved {totalRows:N0} row(s) to local table [{_localTableName}] in {AppPaths.Current.MainDbPath}."
                    : $"Loaded {_rows.Count:N0} of {totalRows:N0} row(s).";
        }
        catch (Exception ex)
        {
            _error = ex.Message;
        }
        finally
        {
            _busy = false;
            _busyAction = string.Empty;
        }
    }

    public string BuildConnectionString()
    {
        if (string.IsNullOrWhiteSpace(_server))
            throw new InvalidOperationException("Server is required.");
        if (string.IsNullOrWhiteSpace(_userId))
            throw new InvalidOperationException("User ID is required.");

        var builder = new SqlConnectionStringBuilder
        {
            DataSource = $"{_server.Trim()},{Math.Clamp(_port, 1, 65535)}",
            InitialCatalog = string.IsNullOrWhiteSpace(_database) ? "master" : _database.Trim(),
            UserID = _userId.Trim(),
            Password = _password,
            ConnectTimeout = Math.Clamp(_timeoutSeconds, 1, 120),
            Encrypt = _encrypt,
            TrustServerCertificate = _trustServerCertificate,
            MultipleActiveResultSets = false,
            ApplicationName = "PortalGNLAB.Web DB Query"
        };

        return builder.ConnectionString;
    }

    public void LoadSavedConnection()
    {
        var c = AppPaths.Current;
        _server = c.AdminDbQueryServer;
        _port = c.AdminDbQueryPort > 0 ? c.AdminDbQueryPort : 1433;
        _database = string.IsNullOrWhiteSpace(c.AdminDbQueryDatabase) ? "master" : c.AdminDbQueryDatabase;
        _userId = c.AdminDbQueryUserId;
        _password = c.AdminDbQueryPassword;
        _timeoutSeconds = c.AdminDbQueryTimeoutSeconds > 0 ? c.AdminDbQueryTimeoutSeconds : 15;
        _encrypt = c.AdminDbQueryEncrypt;
        _trustServerCertificate = c.AdminDbQueryTrustServerCertificate;
    }

    public void SaveConnectionSettings()
    {
        var c = AppPaths.Current;
        AppPaths.Save(new AppPathsConfig
        {
            DbRootDirectory = c.DbRootDirectory,
            MainDbPath = c.MainDbPath,
            NgRateDbSaveDirectory = c.NgRateDbSaveDirectory,
            NgRateRoutingFilePath = c.NgRateRoutingFilePath,
            NgRateReasonFilePath = c.NgRateReasonFilePath,
            NgRateSettingsDbDirectory = c.NgRateSettingsDbDirectory,
            ModelBmesJsonFolder = c.ModelBmesJsonFolder,
            TestExcelConverterSourceDir = c.TestExcelConverterSourceDir,
            TestExcelConverterDestDir = c.TestExcelConverterDestDir,
            AdminDbQueryServer = _server.Trim(),
            AdminDbQueryPort = Math.Clamp(_port, 1, 65535),
            AdminDbQueryDatabase = string.IsNullOrWhiteSpace(_database) ? "master" : _database.Trim(),
            AdminDbQueryUserId = _userId.Trim(),
            AdminDbQueryPassword = _password,
            AdminDbQueryTimeoutSeconds = Math.Clamp(_timeoutSeconds, 1, 120),
            AdminDbQueryEncrypt = _encrypt,
            AdminDbQueryTrustServerCertificate = _trustServerCertificate,
        });
    }

    public static List<string> BuildUniqueColumnNames(SqlDataReader reader)
    {
        var columns = new List<string>(reader.FieldCount);
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < reader.FieldCount; i++)
        {
            string baseName = string.IsNullOrWhiteSpace(reader.GetName(i))
                ? $"Column{i + 1}"
                : reader.GetName(i).Trim();
            string name = baseName;
            if (seen.TryGetValue(baseName, out int count))
            {
                count++;
                seen[baseName] = count;
                name = $"{baseName}_{count}";
            }
            else
            {
                seen[baseName] = 1;
            }
            columns.Add(name);
        }
        return columns;
    }

    public static async Task<string> FormatCellAsync(SqlDataReader reader, int ordinal)
    {
        if (await reader.IsDBNullAsync(ordinal))
            return string.Empty;

        object value = reader.GetValue(ordinal);
        return value switch
        {
            DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss"),
            DateTimeOffset dto => dto.ToString("yyyy-MM-dd HH:mm:ss zzz"),
            byte[] bytes => $"0x{Convert.ToHexString(bytes)}",
            _ => Convert.ToString(value, System.Globalization.CultureInfo.CurrentCulture) ?? string.Empty
        };
    }

    public static bool IsReadOnlyQuery(string sql)
    {
        string normalized = Regex.Replace(sql.Trim(), @"\s+", " ");
        if (normalized.Length == 0) return false;
        if (!Regex.IsMatch(normalized, @"^(select|with)\b", RegexOptions.IgnoreCase))
            return false;

        return !Regex.IsMatch(
            normalized,
            @"\b(insert|update|delete|merge|exec|execute|alter|drop|truncate|create|grant|revoke|backup|restore|into)\b",
            RegexOptions.IgnoreCase);
    }

    public static string SanitizeIdentifier(string raw)
    {
        string s = Regex.Replace((raw ?? string.Empty).Trim(), @"[^\p{L}\p{Nd}_]+", "_");
        s = Regex.Replace(s, "_{2,}", "_").Trim('_');
        if (string.IsNullOrWhiteSpace(s)) s = "ImportedQueryResult";
        if (char.IsDigit(s[0])) s = "_" + s;
        return s.Length > 80 ? s[..80] : s;
    }

    public static string QuoteSqliteIdentifier(string name) =>
        "\"" + name.Replace("\"", "\"\"") + "\"";

    public static LocalTableWriter CreateLocalTableWriter(
        string dbPath,
        string tableName,
        IReadOnlyList<string> columns)
    {
        if (string.IsNullOrWhiteSpace(dbPath))
            throw new InvalidOperationException("Local DB path is not configured.");

        string? dir = Path.GetDirectoryName(dbPath);
        if (!string.IsNullOrWhiteSpace(dir))
            Directory.CreateDirectory(dir);

        var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        var tx = conn.BeginTransaction();
        try
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = $"DROP TABLE IF EXISTS {QuoteSqliteIdentifier(tableName)};";
                cmd.ExecuteNonQuery();
            }

            var safeColumns = BuildUniqueSqliteColumnNames(columns);
            string colsSql = string.Join(", ", safeColumns.Select(c => $"{QuoteSqliteIdentifier(c)} TEXT"));
            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = $"CREATE TABLE {QuoteSqliteIdentifier(tableName)} ({colsSql});";
                cmd.ExecuteNonQuery();
            }

            var insert = conn.CreateCommand();
            insert.Transaction = tx;
            string names = string.Join(", ", safeColumns.Select(QuoteSqliteIdentifier));
            string values = string.Join(", ", columns.Select((_, i) => $"@p{i}"));
            insert.CommandText = $"INSERT INTO {QuoteSqliteIdentifier(tableName)} ({names}) VALUES ({values});";
            for (int i = 0; i < columns.Count; i++)
                insert.Parameters.Add(new SqliteParameter($"@p{i}", DbType.String));

            return new LocalTableWriter(conn, tx, insert);
        }
        catch
        {
            tx.Dispose();
            conn.Dispose();
            throw;
        }
    }

    public static List<string> BuildUniqueSqliteColumnNames(IReadOnlyList<string> columns)
    {
        var result = new List<string>(columns.Count);
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var col in columns)
        {
            string baseName = SanitizeIdentifier(col);
            string name = baseName;
            if (seen.TryGetValue(baseName, out int count))
            {
                count++;
                seen[baseName] = count;
                name = $"{baseName}_{count}";
            }
            else
            {
                seen[baseName] = 1;
            }
            result.Add(name);
        }
        return result;
    }

    public sealed class LocalTableWriter : IAsyncDisposable
    {
        public readonly SqliteConnection _conn;
        public readonly SqliteTransaction _tx;
        public readonly SqliteCommand _insert;
        public bool _completed;

        public LocalTableWriter(SqliteConnection conn, SqliteTransaction tx, SqliteCommand insert)
        {
            _conn = conn;
            _tx = tx;
            _insert = insert;
        }

        public void Insert(IReadOnlyList<string> values)
        {
            for (int i = 0; i < values.Count; i++)
                _insert.Parameters[i].Value = values[i];
            _insert.ExecuteNonQuery();
        }

        public async Task CommitAsync()
        {
            await _tx.CommitAsync();
            _completed = true;
        }

        public async ValueTask DisposeAsync()
        {
            if (!_completed)
                await _tx.RollbackAsync();
            await _insert.DisposeAsync();
            await _tx.DisposeAsync();
            await _conn.DisposeAsync();
        }
    }

    public void ClearMessages()
    {
        _connectionMessage = string.Empty;
        _resultMessage = string.Empty;
        _error = string.Empty;
    }

    public void Dispose()
    {
    }

}
}