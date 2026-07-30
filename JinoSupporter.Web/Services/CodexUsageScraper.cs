using System.Text.RegularExpressions;
using Microsoft.Playwright;

namespace JinoSupporter.Web.Services;

public sealed class CodexUsageScraper
{
    public sealed record ScrapeResult(
        string PlanUsage,
        string PlanLimitUsage,
        string PlanNote,
        string ApiInputTokens,
        string ApiOutputTokens,
        string ApiTotalTokens,
        string? PlanError,
        string? ApiError,
        string PlanRawSnippet,
        string ApiRawSnippet);

    private const string PlanUrl = "https://chatgpt.com/";
    private const string ApiUsageUrl = "https://platform.openai.com/usage";

    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly string _profileDir;
    private bool _playwrightReady;

    public CodexUsageScraper()
    {
        _profileDir = AppStoragePaths.Combine("JinoSupporter", "codex-usages-profile");
        Directory.CreateDirectory(_profileDir);
    }

    public bool SessionProfileExists =>
        Directory.Exists(_profileDir) && Directory.EnumerateFileSystemEntries(_profileDir).Any();

    public string ProfileDirectory => _profileDir;
    public string DebugDirectory => Path.GetFullPath(Path.Combine(_profileDir, "..", "codex-usages-debug"));

    private void EnsurePlaywright()
    {
        if (_playwrightReady) return;
        int exit = Microsoft.Playwright.Program.Main(new[] { "install-deps" });
        _playwrightReady = true;
        _ = exit;
    }

    private async Task<IBrowserContext> LaunchContextAsync(IPlaywright pw, bool headless)
    {
        string[] args = { "--disable-blink-features=AutomationControlled" };
        string[] ignore = { "--enable-automation" };

        string? saved = ReadChannel();
        List<string> channels = new();
        if (saved is not null) channels.Add(saved);
        foreach (string c in new[] { "msedge", "chrome" })
            if (!channels.Contains(c)) channels.Add(c);

        Exception? lastErr = null;
        foreach (string ch in channels)
        {
            try
            {
                var ctx = await pw.Chromium.LaunchPersistentContextAsync(_profileDir,
                    new BrowserTypeLaunchPersistentContextOptions
                    {
                        Headless = headless,
                        Channel = ch,
                        Args = args,
                        IgnoreDefaultArgs = ignore,
                        ViewportSize = new() { Width = 1200, Height = 850 },
                    });
                WriteChannel(ch);
                return ctx;
            }
            catch (Exception ex) { lastErr = ex; }
        }

        try
        {
            int exit = Microsoft.Playwright.Program.Main(new[] { "install", "chromium" });
            if (exit != 0) throw new InvalidOperationException($"Fallback Chromium install failed (exit={exit})");
            return await pw.Chromium.LaunchPersistentContextAsync(_profileDir,
                new BrowserTypeLaunchPersistentContextOptions
                {
                    Headless = headless,
                    Args = args,
                    IgnoreDefaultArgs = ignore,
                    ViewportSize = new() { Width = 1200, Height = 850 },
                });
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "Failed to launch Edge/Chrome/Chromium for Codex/OpenAI usage login. " +
                $"Cause: {lastErr?.Message ?? ex.Message}", ex);
        }
    }

    private string ChannelFile => Path.Combine(_profileDir, ".channel");

    private string? ReadChannel()
    {
        try { return File.Exists(ChannelFile) ? File.ReadAllText(ChannelFile).Trim() : null; }
        catch { return null; }
    }

    private void WriteChannel(string ch)
    {
        try { File.WriteAllText(ChannelFile, ch); } catch { }
    }

    public async Task OpenLoginBrowserAsync(CancellationToken ct = default)
    {
        await _gate.WaitAsync(ct);
        try
        {
            EnsurePlaywright();
            using IPlaywright pw = await Playwright.CreateAsync();
            IBrowserContext ctx = await LaunchContextAsync(pw, headless: false);

            IPage planTab = ctx.Pages.Count > 0 ? ctx.Pages[0] : await ctx.NewPageAsync();
            try { await planTab.GotoAsync(PlanUrl, new PageGotoOptions { Timeout = 30_000 }); } catch { }

            IPage apiTab = await ctx.NewPageAsync();
            try { await apiTab.GotoAsync(ApiUsageUrl, new PageGotoOptions { Timeout = 30_000 }); } catch { }

            var closeTcs = new TaskCompletionSource();
            ctx.Close += (_, _) => closeTcs.TrySetResult();

            using (ct.Register(() => { try { _ = ctx.CloseAsync(); } catch { } }))
                await closeTcs.Task;
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task<ScrapeResult> ScrapeAsync(CancellationToken ct = default)
    {
        await _gate.WaitAsync(ct);
        try
        {
            EnsurePlaywright();
            using IPlaywright pw = await Playwright.CreateAsync();
            IBrowserContext ctx = await LaunchContextAsync(pw, headless: true);

            string planUsage = "", planLimitUsage = "", planNote = "";
            string apiInput = "", apiOutput = "", apiTotal = "";
            string? planErr = null, apiErr = null;
            string planSnippet = "", apiSnippet = "";

            try
            {
                var (text, finalUrl) = await LoadBodyTextAsync(ctx, PlanUrl, "codex-plan");
                planSnippet = Truncate(text, 1500);

                if (LooksLikeAuthRedirect(finalUrl, PlanUrl) || LooksLikeSignedOut(text))
                    planErr = $"Codex plan session expired or signed out (url: {finalUrl}). Login again.";
                else
                {
                    planUsage = ExtractPercentNear(text, new[] { "Codex", "usage", "limit", "remaining" });
                    planLimitUsage = ExtractPercentNear(text, new[] { "weekly", "monthly", "limit" });
                    planNote = "Signed in to chatgpt.com";
                    if (string.IsNullOrEmpty(planUsage) && string.IsNullOrEmpty(planLimitUsage))
                        planErr = "Signed in to chatgpt.com, but Codex plan usage was not found on the page.";
                }
            }
            catch (Exception ex) { planErr = ex.Message; }

            try
            {
                var (text, finalUrl) = await LoadBodyTextAsync(ctx, ApiUsageUrl, "codex-api");
                apiSnippet = Truncate(text, 1500);

                if (LooksLikeAuthRedirect(finalUrl, ApiUsageUrl) || LooksLikeSignedOut(text))
                    apiErr = $"OpenAI platform session expired or signed out (url: {finalUrl}). Login again.";
                else
                {
                    apiInput = ExtractTokenNear(text, new[] { "input tokens", "prompt tokens" });
                    apiOutput = ExtractTokenNear(text, new[] { "output tokens", "completion tokens" });
                    apiTotal = ExtractTokenNear(text, new[] { "total tokens" });
                }
            }
            catch (Exception ex) { apiErr = ex.Message; }

            await ctx.CloseAsync();

            return new ScrapeResult(
                planUsage, planLimitUsage, planNote,
                apiInput, apiOutput, apiTotal,
                planErr, apiErr,
                planSnippet, apiSnippet);
        }
        finally
        {
            _gate.Release();
        }
    }

    private async Task<(string text, string finalUrl)> LoadBodyTextAsync(IBrowserContext ctx, string url, string debugSlug)
    {
        IPage page = await ctx.NewPageAsync();
        try
        {
            await page.GotoAsync(url, new PageGotoOptions
            {
                Timeout = 30_000,
                WaitUntil = WaitUntilState.DOMContentLoaded,
            });
            await page.WaitForTimeoutAsync(7000);

            string text = await page.InnerTextAsync("body");
            await DumpDebugAsync(debugSlug, text, page);
            return (text, page.Url);
        }
        finally
        {
            try { await page.CloseAsync(); } catch { }
        }
    }

    private static bool LooksLikeAuthRedirect(string finalUrl, string target)
    {
        try
        {
            var t = new Uri(target);
            var f = new Uri(finalUrl);
            if (!string.Equals(t.Host, f.Host, StringComparison.OrdinalIgnoreCase)) return true;
            return f.AbsolutePath.Contains("login", StringComparison.OrdinalIgnoreCase) ||
                   f.AbsolutePath.Contains("auth", StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private static bool LooksLikeSignedOut(string text) =>
        text.Contains("Log in", StringComparison.OrdinalIgnoreCase) ||
        text.Contains("Sign in", StringComparison.OrdinalIgnoreCase) ||
        text.Contains("login", StringComparison.OrdinalIgnoreCase);

    private static string ExtractPercentNear(string text, string[] labels)
    {
        const string PCT = @"(\d+(?:\.\d+)?)\s*%";
        foreach (string label in labels)
        {
            var m1 = Regex.Match(text, Regex.Escape(label) + @"[\s\S]{0,160}?" + PCT,
                RegexOptions.IgnoreCase);
            if (m1.Success) return m1.Groups[1].Value + "%";

            var m2 = Regex.Match(text, PCT + @"[\s\S]{0,80}?" + Regex.Escape(label),
                RegexOptions.IgnoreCase);
            if (m2.Success) return m2.Groups[1].Value + "%";
        }
        return "";
    }

    private static string ExtractTokenNear(string text, string[] labels)
    {
        const string NUMBER = @"([0-9][0-9,]*(?:\.\d+)?\s*[KMB]?)";
        foreach (string label in labels)
        {
            var m1 = Regex.Match(text, Regex.Escape(label) + @"[\s\S]{0,160}?" + NUMBER,
                RegexOptions.IgnoreCase);
            if (m1.Success) return NormalizeNumber(m1.Groups[1].Value);

            var m2 = Regex.Match(text, NUMBER + @"[\s\S]{0,80}?" + Regex.Escape(label),
                RegexOptions.IgnoreCase);
            if (m2.Success) return NormalizeNumber(m2.Groups[1].Value);
        }
        return "";
    }

    private static string NormalizeNumber(string raw)
        => raw.Replace(" ", "").Trim();

    private static string Truncate(string s, int max)
        => string.IsNullOrEmpty(s) ? "" : (s.Length <= max ? s : s[..max] + "...");

    private async Task DumpDebugAsync(string slug, string text, IPage page)
    {
        try
        {
            string dir = DebugDirectory;
            Directory.CreateDirectory(dir);
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            await File.WriteAllTextAsync(Path.Combine(dir, $"{slug}_{stamp}.txt"), text);
            await page.ScreenshotAsync(new PageScreenshotOptions
            {
                Path = Path.Combine(dir, $"{slug}_{stamp}.png"),
                FullPage = true,
            });
        }
        catch { }
    }

    public void DeleteSession()
    {
        try
        {
            if (Directory.Exists(_profileDir))
                Directory.Delete(_profileDir, recursive: true);
            Directory.CreateDirectory(_profileDir);
        }
        catch { }
    }
}
