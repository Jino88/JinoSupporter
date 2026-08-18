using System.Net;
using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Extensions.Options;

namespace JinoSupporter.Web.Services.BmesReports;

/// <summary>
/// Host-cutover controls for the staged BMES React migration. Setting
/// <see cref="ReactViewerEnabled"/> to false is the immediate rollback path and does not
/// invalidate already generated report tokens.
/// </summary>
public sealed class BmesReportViewerOptions
{
    public const string SectionName = "BmesReport";

    public bool ReactViewerEnabled { get; set; } = true;

    /// <summary>Deployment version appended only to the stable ESM/CSS URLs.</summary>
    public string ViewerAssetVersion { get; set; } = "20260818-1";
}

public enum BmesReportViewerMode
{
    Legacy,
    React,
}

public sealed record BmesReportViewerSelection(BmesReportViewerMode Mode, string Reason);

public sealed record BmesReportBootstrapDocument(string Html, string ContentSecurityPolicy);

/// <summary>
/// Selects React only when the flag, current report JSON and both built assets are ready.
/// Any missing prerequisite remains a same-token legacy fallback; a runtime ESM/mount
/// failure redirects to the explicit <c>?legacy=true</c> branch.
/// </summary>
public sealed class BmesReportViewerBootstrap
{
    public const string JavaScriptRequestPath = "/bmes-report/bmes-report.js";
    public const string CssRequestPath = "/bmes-report/bmes-report.css";

    private readonly IWebHostEnvironment _environment;
    private readonly IOptionsMonitor<BmesReportViewerOptions> _options;
    private readonly ILogger<BmesReportViewerBootstrap> _logger;

    public BmesReportViewerBootstrap(
        IWebHostEnvironment environment,
        IOptionsMonitor<BmesReportViewerOptions> options,
        ILogger<BmesReportViewerBootstrap> logger)
    {
        _environment = environment;
        _options = options;
        _logger = logger;
    }

    public BmesReportViewerSelection Select(BmesReportArtifacts artifacts, bool forceLegacy)
    {
        if (forceLegacy)
            return new(BmesReportViewerMode.Legacy, "explicit-legacy-query");

        BmesReportViewerOptions options = _options.CurrentValue;
        if (!options.ReactViewerEnabled)
            return new(BmesReportViewerMode.Legacy, "feature-flag-disabled");

        if (!artifacts.IsCurrentContract || artifacts.ReportJsonPath is null)
            return Fallback("current-report-json-unavailable");

        string webRoot = string.IsNullOrWhiteSpace(_environment.WebRootPath)
            ? Path.Combine(_environment.ContentRootPath, "wwwroot")
            : _environment.WebRootPath;
        string javascriptPath = Path.Combine(webRoot, "bmes-report", "bmes-report.js");
        string cssPath = Path.Combine(webRoot, "bmes-report", "bmes-report.css");
        if (!IsNonEmptyFile(javascriptPath) || !IsNonEmptyFile(cssPath))
            return Fallback("viewer-assets-unavailable");

        return new(BmesReportViewerMode.React, "react-ready");
    }

    public BmesReportBootstrapDocument BuildReactDocument(string token)
    {
        if (!BmesReportHtmlExportService.IsValidToken(token))
            throw new ArgumentException("A valid BMES report token is required.", nameof(token));

        string assetVersion = Uri.EscapeDataString(_options.CurrentValue.ViewerAssetVersion.Trim());
        string javascriptUrl = $"{JavaScriptRequestPath}?v={assetVersion}";
        string cssUrl = $"{CssRequestPath}?v={assetVersion}";
        string reportUrl = $"/report/bmes/data/{token}";
        string fallbackUrl = $"/report/bmes/view/{token}?legacy=true";
        string nonce = Convert.ToHexString(RandomNumberGenerator.GetBytes(16));
        string html = $$"""
            <!doctype html>
            <html lang="ko">
            <head>
              <meta charset="utf-8">
              <meta name="viewport" content="width=device-width, initial-scale=1">
              <title>BMES Report</title>
              <link rel="stylesheet" href="{{WebUtility.HtmlEncode(cssUrl)}}">
            </head>
            <body>
              <div id="bmes-report-root" data-report-url="{{WebUtility.HtmlEncode(reportUrl)}}"></div>
              <script type="module" nonce="{{nonce}}">
                const root = document.getElementById("bmes-report-root");
                const fallbackUrl = {{JsonSerializer.Serialize(fallbackUrl)}};
                try {
                  const viewer = await import({{JsonSerializer.Serialize(javascriptUrl)}});
                  if (!root || typeof viewer.mount !== "function") {
                    throw new Error("BMES React viewer mount export is unavailable.");
                  }
                  viewer.mount(root, { reportUrl: root.dataset.reportUrl });
                } catch (error) {
                  console.error("BMES React viewer failed; loading legacy report.", error);
                  window.location.replace(fallbackUrl);
                }
              </script>
            </body>
            </html>
            """;
        string contentSecurityPolicy =
            $"default-src 'none'; style-src 'self' 'unsafe-inline'; script-src 'self' 'nonce-{nonce}'; " +
            "connect-src 'self'; img-src 'self' data:; font-src 'self'; base-uri 'none'; frame-ancestors 'self'";
        return new(html, contentSecurityPolicy);
    }

    private BmesReportViewerSelection Fallback(string reason)
    {
        _logger.LogWarning("BMES React viewer fallback selected: {Reason}", reason);
        return new(BmesReportViewerMode.Legacy, reason);
    }

    private static bool IsNonEmptyFile(string path)
    {
        try
        {
            return new FileInfo(path) is { Exists: true, Length: > 0 };
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }
}
