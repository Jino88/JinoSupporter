using System.Text.RegularExpressions;
using Microsoft.Playwright;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Scrapes Anthropic billing page and Claude.ai usage page via a persistent
/// Playwright browser context. Admin logs in once through <see cref="OpenLoginBrowserAsync"/>;
/// subsequent <see cref="ScrapeAsync"/> calls reuse saved cookies.
/// </summary>
public sealed class ClaudeUsageScraper
{
    public sealed record ScrapeResult(
        string  ApiCreditBalance,
        string  SessionUsage,        // 현재 세션 %
        string  WeeklyAllModels,     // 주간 한도: 모든 모델 %
        string  WeeklySonnet,        // 주간 한도: Sonnet만 %
        string  WeeklyClaudeDesign,  // 주간 한도: Claude Design %
        string  ExtraUsage,          // 추가 사용량 (USD)
        string  ExtraBalance,        // 현재 잔액 (USD)
        string? BillingError,
        string? UsageError,
        string  BillingRawSnippet,
        string  UsageRawSnippet);

    private const string BillingUrl = "https://platform.claude.com/settings/billing";
    private const string UsageUrl   = "https://claude.ai/settings/usage";

    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly string        _profileDir;
    private bool                   _playwrightReady;

    public ClaudeUsageScraper()
    {
        _profileDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "JinoSupporter", "ai-usages-profile");
        Directory.CreateDirectory(_profileDir);
    }

    public bool SessionProfileExists =>
        Directory.Exists(_profileDir) && Directory.EnumerateFileSystemEntries(_profileDir).Any();

    public string ProfileDirectory => _profileDir;

    /// <summary>Ensures Playwright drivers are installed (no Chromium — we use installed Edge/Chrome).</summary>
    private void EnsurePlaywright()
    {
        if (_playwrightReady) return;
        // Install only the PW driver; we rely on locally installed Edge/Chrome via Channel=
        // (Chromium bundled by Playwright triggers Google's "insecure browser" block on OAuth.)
        int exit = Microsoft.Playwright.Program.Main(new[] { "install-deps" });
        // install-deps may no-op on Windows — that's fine. Treat non-zero as warning.
        _playwrightReady = true;
        _ = exit;
    }

    /// <summary>
    /// Launch persistent context with real Edge/Chrome so Google OAuth works.
    /// Uses a locked-in channel (saved during first successful login) to ensure
    /// session cookies persist between login → scrape.
    /// </summary>
    private async Task<IBrowserContext> LaunchContextAsync(
        IPlaywright pw, bool headless)
    {
        string[] args   = { "--disable-blink-features=AutomationControlled" };
        string[] ignore = { "--enable-automation" };

        // Try previously-chosen channel first — session cookies are per-browser-channel.
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
                        Headless          = headless,
                        Channel           = ch,
                        Args              = args,
                        IgnoreDefaultArgs = ignore,
                        ViewportSize      = new() { Width = 1200, Height = 850 },
                    });
                WriteChannel(ch);
                return ctx;
            }
            catch (Exception ex) { lastErr = ex; }
        }

        // Last-ditch fallback: bundled Chromium.
        try
        {
            int exit = Microsoft.Playwright.Program.Main(new[] { "install", "chromium" });
            if (exit != 0) throw new InvalidOperationException($"Fallback Chromium install failed (exit={exit})");
            return await pw.Chromium.LaunchPersistentContextAsync(_profileDir,
                new BrowserTypeLaunchPersistentContextOptions
                {
                    Headless          = headless,
                    Args              = args,
                    IgnoreDefaultArgs = ignore,
                    ViewportSize      = new() { Width = 1200, Height = 850 },
                });
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "Failed to launch Edge/Chrome/Chromium. Windows Edge is typically installed by default. " +
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

    /// <summary>
    /// Opens a HEADED browser window on the server machine pointed at the login
    /// URLs. Returns when the user closes the browser.
    /// </summary>
    public async Task OpenLoginBrowserAsync(CancellationToken ct = default)
    {
        await _gate.WaitAsync(ct);
        try
        {
            EnsurePlaywright();
            using IPlaywright pw = await Playwright.CreateAsync();

            IBrowserContext ctx = await LaunchContextAsync(pw, headless: false);

            // Open BOTH target URLs so the user logs into each domain independently.
            // platform.claude.com and claude.ai are separate auth domains — session cookies
            // do NOT transfer between them.
            IPage billingTab = ctx.Pages.Count > 0 ? ctx.Pages[0] : await ctx.NewPageAsync();
            try { await billingTab.GotoAsync(BillingUrl, new PageGotoOptions { Timeout = 30_000 }); }
            catch { }

            IPage usageTab = await ctx.NewPageAsync();
            try { await usageTab.GotoAsync(UsageUrl, new PageGotoOptions { Timeout = 30_000 }); }
            catch { }

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

    /// <summary>
    /// Opens a HEADLESS browser using the saved profile and scrapes both pages.
    /// Returns per-page error messages in the result if scraping failed.
    /// </summary>
    public async Task<ScrapeResult> ScrapeAsync(CancellationToken ct = default)
    {
        await _gate.WaitAsync(ct);
        try
        {
            EnsurePlaywright();
            using IPlaywright pw = await Playwright.CreateAsync();

            // NOTE: Cloudflare blocks headless browsers on claude.ai ("보안 확인 수행 중").
            // Running non-headless (with the real profile) reliably passes the challenge,
            // usually without any user interaction. A small browser window flashes during
            // refresh — acceptable for a local admin tool.
            IBrowserContext ctx = await LaunchContextAsync(pw, headless: false);

            string? billingErr = null, usageErr = null;
            string  apiBalance = "";
            string  sessionPct = "", weeklyAll = "", weeklySonnet = "", weeklyDesign = "";
            string  extraUsage = "", extraBalance = "";
            string  billingSnippet = "", usageSnippet = "";

            try
            {
                var (billingText, finalUrl) = await LoadBodyTextAsync(ctx, BillingUrl, "billing");
                billingSnippet = Truncate(billingText, 1500);

                if (IsCloudflareChallenge(billingText))
                {
                    billingErr = "Blocked by Cloudflare bot detection. Solve the CAPTCHA in the browser window, then click Refresh again.";
                }
                else if (LooksLikeAuthRedirect(finalUrl, BillingUrl))
                {
                    billingErr = $"Session expired (redirected to: {finalUrl}). Please Login again.";
                }
                else
                {
                    apiBalance = ExtractBalance(billingText);
                    if (string.IsNullOrEmpty(apiBalance))
                        billingErr = "Could not find the Balance value. Session may be expired or the DOM structure may have changed. (See Raw snippet below)";
                }
            }
            catch (Exception ex) { billingErr = ex.Message; }

            try
            {
                var (usageText, finalUrl) = await LoadBodyTextAsync(ctx, UsageUrl, "usage");
                usageSnippet = Truncate(usageText, 1500);

                if (IsCloudflareChallenge(usageText))
                {
                    usageErr = "Blocked by Cloudflare bot detection. While the browser window is open, click the CAPTCHA checkbox directly, or reconnect via the Login button and complete manual verification, then click Refresh.";
                }
                else if (LooksLikeAuthRedirect(finalUrl, UsageUrl))
                {
                    usageErr = $"claude.ai session expired (redirected to: {finalUrl}). Click Login and sign in on the **second tab (claude.ai)** as well, then click Refresh again.";
                }
                else
                {
                    // Label-based extraction, mapped to the real claude.ai/settings/usage layout.
                    sessionPct   = ExtractPercentNear(usageText, new[] { "현재 세션", "current session" });
                    weeklyAll    = ExtractPercentNear(usageText, new[] { "모든 모델", "all models" });
                    weeklySonnet = ExtractPercentNear(usageText, new[] { "Sonnet만", "Sonnet 만", "Sonnet only" });
                    weeklyDesign = ExtractPercentNear(usageText, new[] { "Claude Design" });

                    extraUsage   = ExtractDollarNear(usageText, new[] { "추가 사용량", "additional usage" });
                    extraBalance = ExtractDollarNear(usageText, new[] { "현재 잔액", "current balance", "잔액" });

                    bool nothingFound = string.IsNullOrEmpty(sessionPct) &&
                                        string.IsNullOrEmpty(weeklyAll)  &&
                                        string.IsNullOrEmpty(weeklySonnet) &&
                                        string.IsNullOrEmpty(weeklyDesign) &&
                                        string.IsNullOrEmpty(extraUsage)   &&
                                        string.IsNullOrEmpty(extraBalance);
                    if (nothingFound)
                        usageErr = "Could not find values near the labels (current session / all models / Sonnet only / balance, etc.) on the usage page. Session may be expired or the page structure may have changed. (See Raw snippet)";
                }
            }
            catch (Exception ex) { usageErr = ex.Message; }

            await ctx.CloseAsync();

            return new ScrapeResult(
                apiBalance,
                sessionPct, weeklyAll, weeklySonnet, weeklyDesign,
                extraUsage, extraBalance,
                billingErr, usageErr,
                billingSnippet, usageSnippet);
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
            // NOTE: claude.ai / platform.claude.com are streaming SPAs; NetworkIdle never fires.
            await page.GotoAsync(url, new PageGotoOptions
            {
                Timeout   = 30_000,
                WaitUntil = WaitUntilState.DOMContentLoaded,
            });
            await page.WaitForTimeoutAsync(7000);

            string text = await page.InnerTextAsync("body");

            // Cloudflare "Just a moment..." / "보안 확인 수행 중" challenge — wait longer and retry.
            if (IsCloudflareChallenge(text))
            {
                for (int i = 0; i < 4 && IsCloudflareChallenge(text); i++)
                {
                    await page.WaitForTimeoutAsync(5000);
                    text = await page.InnerTextAsync("body");
                }
            }

            await DumpDebugAsync(debugSlug, text, page);
            return (text, page.Url);
        }
        finally
        {
            try { await page.CloseAsync(); } catch { }
        }
    }

    /// <summary>
    /// Heuristic: final URL lands outside the settings path (e.g. / or /plans or
    /// /login) → we were redirected because auth failed.
    /// </summary>
    private static bool LooksLikeAuthRedirect(string finalUrl, string target)
    {
        try
        {
            var t = new Uri(target);
            var f = new Uri(finalUrl);
            if (!string.Equals(t.Host, f.Host, StringComparison.OrdinalIgnoreCase)) return true;
            // If we expected /settings/... but ended elsewhere
            if (target.Contains("/settings/", StringComparison.OrdinalIgnoreCase)
                && !f.AbsolutePath.StartsWith("/settings/", StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }
        catch { return false; }
    }

    private static bool IsCloudflareChallenge(string text) =>
        text.Contains("보안 확인 수행 중", StringComparison.Ordinal) ||
        text.Contains("Just a moment",  StringComparison.OrdinalIgnoreCase) ||
        text.Contains("Checking your browser", StringComparison.OrdinalIgnoreCase) ||
        (text.Contains("Cloudflare", StringComparison.OrdinalIgnoreCase) &&
         text.Contains("Ray ID",     StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Extracts API credit balance from billing-page text. Matches the $ amount
    /// nearest a "balance" / "잔액" / "credits" style label (label→$ or $→label,
    /// within 0–60 chars). Supports both "$2.93" and "US$2.93". Falls back to the
    /// largest $ if no label context is found.
    /// </summary>
    private static string ExtractBalance(string text)
    {
        // Regex snippet that matches "$X.XX" or "US$X.XX" (or with a space between).
        const string DOLLAR = @"(?:US)?\s*\$\s*([0-9,]+(?:\.[0-9]{1,2})?)";

        // Labels in order of specificity. The first match wins.
        string[] labels =
        {
            "크레딧 잔액",   // "credit balance" (KR) — most specific
            "크레딧 잔고",
            "credit balance",
            "available credit",
            "available balance",
            "잔액",          // "balance" (KR)
            "balance",
            "available",
            "credits",
        };

        foreach (string lbl in labels)
        {
            // label THEN $
            var m1 = Regex.Match(text, Regex.Escape(lbl) + @"[\s\S]{0,60}?" + DOLLAR,
                RegexOptions.IgnoreCase);
            if (m1.Success) return Normalize(m1.Groups[1].Value);

            // $ THEN label (e.g. "US$2.93 … 잔액")
            var m2 = Regex.Match(text, DOLLAR + @"[\s\S]{0,30}?" + Regex.Escape(lbl),
                RegexOptions.IgnoreCase);
            if (m2.Success) return Normalize(m2.Groups[1].Value);
        }

        // Fallback: largest $ amount (often wrong — use only as last resort).
        double best = -1;
        foreach (Match m in Regex.Matches(text, DOLLAR))
        {
            string s = m.Groups[1].Value.Replace(",", "");
            if (double.TryParse(s, System.Globalization.NumberStyles.Float,
                                System.Globalization.CultureInfo.InvariantCulture, out double v)
                && v > best) best = v;
        }
        return best < 0 ? "" : best.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);

        static string Normalize(string raw) =>
            double.TryParse(raw.Replace(",", ""), System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out double v)
                ? v.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                : raw.Replace(",", "").Trim();
    }

    private static string Truncate(string s, int max) =>
        string.IsNullOrEmpty(s) ? "" : (s.Length <= max ? s : s[..max] + "…");

    /// <summary>
    /// Find `N%` within ~150 chars AFTER any of the given labels, or ~60 chars
    /// BEFORE the label. First match wins. Returns "42%" or "".
    /// </summary>
    private static string ExtractPercentNear(string text, string[] labels)
    {
        const string PCT = @"(\d+(?:\.\d+)?)\s*%";
        foreach (string lbl in labels)
        {
            var m1 = Regex.Match(text, Regex.Escape(lbl) + @"[\s\S]{0,150}?" + PCT,
                RegexOptions.IgnoreCase);
            if (m1.Success) return m1.Groups[1].Value + "%";

            var m2 = Regex.Match(text, PCT + @"[\s\S]{0,60}?" + Regex.Escape(lbl),
                RegexOptions.IgnoreCase);
            if (m2.Success) return m2.Groups[1].Value + "%";
        }
        return "";
    }

    /// <summary>
    /// Find `$X.XX` (optionally `US$`) near any of the given labels. Label can be
    /// before ($ within ~80 chars after) or after ($ within ~40 chars before).
    /// </summary>
    private static string ExtractDollarNear(string text, string[] labels)
    {
        const string DOLLAR = @"(?:US)?\s*\$\s*([0-9,]+(?:\.[0-9]{1,2})?)";
        foreach (string lbl in labels)
        {
            var m1 = Regex.Match(text, Regex.Escape(lbl) + @"[\s\S]{0,80}?" + DOLLAR,
                RegexOptions.IgnoreCase);
            if (m1.Success) return NormalizeDollar(m1.Groups[1].Value);

            var m2 = Regex.Match(text, DOLLAR + @"[\s\S]{0,40}?" + Regex.Escape(lbl),
                RegexOptions.IgnoreCase);
            if (m2.Success) return NormalizeDollar(m2.Groups[1].Value);
        }
        return "";

        static string NormalizeDollar(string raw) =>
            double.TryParse(raw.Replace(",", ""), System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out double v)
                ? v.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
                : raw.Replace(",", "").Trim();
    }

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
                Path     = Path.Combine(dir, $"{slug}_{stamp}.png"),
                FullPage = true,
            });
        }
        catch { /* best-effort */ }
    }

    public string DebugDirectory => Path.GetFullPath(Path.Combine(_profileDir, "..", "ai-usages-debug"));

    public void DeleteSession()
    {
        try
        {
            if (Directory.Exists(_profileDir))
                Directory.Delete(_profileDir, recursive: true);
            Directory.CreateDirectory(_profileDir);
        }
        catch { /* best-effort */ }
    }
}
