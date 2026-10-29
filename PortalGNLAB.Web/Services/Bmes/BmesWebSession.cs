using System.Net;
using System.Text.RegularExpressions;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// The one place the app logs in to the BMES web server. Every BMES screen call goes
/// through a session opened here and then through <see cref="BmesGateway"/>, so when BMES
/// hands out a real API only this folder (Services/Bmes) has to change.
///
/// Data screens that have a PLATFORM_API Tool are served through <see cref="BmesApi"/> using the
/// credentials <see cref="BmesApiCredentialProvider"/> resolves, so opening a session does NOT log
/// in to the web. The web login happens lazily, the first time a screen that still needs web
/// parsing calls <see cref="EnsureWebLoginAsync"/> (IPG, material master, routing master, PDM,
/// worker attendance, LPA photos).
///
/// Web flow (same as the BMES login page): read the anti-forgery token from the home page,
/// post it with the configured credentials to MES000000/LoginCheck, keep the cookie.
/// Each session gets its own cookie container; browser cookies or copied SSO tokens are
/// never used.
/// </summary>
public sealed class BmesWebSession : IDisposable
{
    public const string BaseUrl = "https://bmes.bujeon.com";
    public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(300);

    private readonly HttpClientHandler _handler;
    private readonly HttpClient _client;
    private readonly BmesWebSessionOptions _options;
    private readonly SemaphoreSlim _loginLock = new(1, 1);
    private bool _webLoggedIn;
    private (bool Ok, BmesLoginFailure Failure, string Diagnostic)? _webLoginFailure;

    /// <summary>The login the caller gave to <see cref="OpenAsync"/> (BMES web login).</summary>
    public string LoginId { get; }
    public string Password { get; }

    /// <summary>The web client. Only valid after <see cref="EnsureWebLoginAsync"/> succeeded.</summary>
    public HttpClient Client => _webLoggedIn
        ? _client
        : throw new InvalidOperationException("BMES web session is not logged in; call EnsureWebLoginAsync first.");

    /// <summary>PLATFORM_API login for this session (settings API login, else this session's web login); null when none.</summary>
    public BmesApiCredentials? ApiCredentials => BmesApiCredentialProvider.Resolve(LoginId, Password);

    private BmesWebSession(string loginId, string password, BmesWebSessionOptions options)
    {
        LoginId = loginId;
        Password = password;
        _options = options;
        _handler = new HttpClientHandler
        {
            UseCookies = true,
            CookieContainer = new CookieContainer(),
            AutomaticDecompression = options.AutomaticDecompression,
        };
        _client = new HttpClient(_handler) { Timeout = options.Timeout ?? DefaultTimeout };
        if (options.AjaxHeader)
            _client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");
        if (options.Referrer is not null)
            _client.DefaultRequestHeaders.Referrer = options.Referrer;
    }

    /// <summary>Opens a session. No server call happens here: API calls log in on their first use and
    /// the web login waits for <see cref="EnsureWebLoginAsync"/>. The result carries no session only when
    /// neither the BMES web login nor an API login is configured.</summary>
    public static Task<BmesLoginOutcome> OpenAsync(
        string loginId,
        string password,
        BmesWebSessionOptions? options = null,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        bool webLogin = !string.IsNullOrWhiteSpace(loginId) && !string.IsNullOrWhiteSpace(password);
        if (!webLogin && !BmesApiCredentialProvider.HasApiLogin)
            return Task.FromResult(BmesLoginOutcome.Failed(BmesLoginFailure.MissingCredentials, "BMES credentials are not configured."));

        var session = new BmesWebSession(loginId ?? string.Empty, password ?? string.Empty, options ?? new BmesWebSessionOptions());
        return Task.FromResult(new BmesLoginOutcome(session, BmesLoginFailure.None, string.Empty));
    }

    /// <summary>Logs in to the BMES web (once per session) for screens that are still parsed from web pages.
    /// A failure is remembered so a later call does not log in again.</summary>
    public async Task<BmesLoginOutcome> EnsureWebLoginAsync(
        IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        await _loginLock.WaitAsync(cancellationToken);
        try
        {
            if (_webLoggedIn) return new BmesLoginOutcome(this, BmesLoginFailure.None, string.Empty);
            if (_webLoginFailure is { } failed) return BmesLoginOutcome.Failed(failed.Failure, failed.Diagnostic);

            if (string.IsNullOrWhiteSpace(LoginId) || string.IsNullOrWhiteSpace(Password))
                return Remember(BmesLoginFailure.MissingCredentials, "BMES web credentials are not configured (this screen has no API Tool).");

            if (_options.TokenMessage is { Length: > 0 } tokenMessage) progress?.Report(tokenMessage);
            string token = await GetTokenAsync(cancellationToken);
            if (string.IsNullOrEmpty(token))
                return Remember(BmesLoginFailure.Token, "Failed to read BMES login token.");

            if (_options.LoginMessage is { Length: > 0 } loginMessage) progress?.Report(loginMessage);
            (bool ok, string diagnostic) = await LoginAsync(token, LoginId, Password, cancellationToken);
            if (!ok) return Remember(BmesLoginFailure.Login, diagnostic);

            _webLoggedIn = true;
            return new BmesLoginOutcome(this, BmesLoginFailure.None, string.Empty);
        }
        finally
        {
            _loginLock.Release();
        }
    }

    private BmesLoginOutcome Remember(BmesLoginFailure failure, string diagnostic)
    {
        _webLoginFailure = (false, failure, diagnostic);
        return BmesLoginOutcome.Failed(failure, diagnostic);
    }

    public void Dispose()
    {
        _client.Dispose();
        _handler.Dispose();
        _loginLock.Dispose();
    }

    private async Task<string> GetTokenAsync(CancellationToken cancellationToken)
    {
        try
        {
            string html = await _client.GetStringAsync(BaseUrl, cancellationToken);
            // The token input may list name before value or value before name.
            Match m = Regex.Match(html,
                @"<input[^>]+name=""__RequestVerificationToken""[^>]+value=""([^""]+)""",
                RegexOptions.IgnoreCase);
            if (!m.Success)
                m = Regex.Match(html,
                    @"<input[^>]+value=""([^""]+)""[^>]+name=""__RequestVerificationToken""",
                    RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value : string.Empty;
        }
        catch { return string.Empty; }
    }

    private async Task<(bool Ok, string Diagnostic)> LoginAsync(
        string token, string loginId, string password, CancellationToken cancellationToken)
    {
        using var content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["UserInfo[USRID]"] = loginId,
            ["UserInfo[PWNO]"]  = password,
            ["UserInfo[LANG]"]  = "EN",
            ["UserInfo[FACCO]"] = "GN",
            ["UserInfo[STYPE]"] = "P",
            ["UserInfo[VTYPE]"] = "P",
            ["__RequestVerificationToken"] = token,
        });
        try
        {
            using var response = await _client.PostAsync(BaseUrl + "/MES000000/LoginCheck", content, cancellationToken);
            string body = await response.Content.ReadAsStringAsync(cancellationToken);
            bool ok = body.Contains("\"Result\":\"M\"", StringComparison.OrdinalIgnoreCase);
            return (ok, BuildLoginDiagnostic(response, body));
        }
        catch (Exception ex) { return (false, $"request error: {ex.GetType().Name}"); }
    }

    private static string BuildLoginDiagnostic(HttpResponseMessage response, string body)
    {
        string status = $"HTTP {(int)response.StatusCode} {response.ReasonPhrase}".Trim();
        Match result = Regex.Match(body, "\"Result\"\\s*:\\s*\"([^\"]*)\"", RegexOptions.IgnoreCase);
        Match message = Regex.Match(body, "\"(?:Message|ErrorMessage)\"\\s*:\\s*\"([^\"]*)\"", RegexOptions.IgnoreCase);
        if (result.Success) status += $", Result={result.Groups[1].Value}";
        if (message.Success) status += $", Message={SanitizeLoginMessage(message.Groups[1].Value)}";
        return status;
    }

    private static string SanitizeLoginMessage(string value)
    {
        string normalized = Regex.Replace(value, @"\s+", " ").Trim();
        return normalized.Length <= 160 ? normalized : normalized[..160] + "...";
    }
}

public sealed record BmesWebSessionOptions
{
    /// <summary>Null = <see cref="BmesWebSession.DefaultTimeout"/>.</summary>
    public TimeSpan? Timeout { get; init; }

    /// <summary>BMES screens answer SearchList calls only as AJAX requests.</summary>
    public bool AjaxHeader { get; init; } = true;

    public Uri? Referrer { get; init; }

    /// <summary>PDM file downloads arrive compressed; the SearchList screens do not need it.</summary>
    public DecompressionMethods AutomaticDecompression { get; init; } = DecompressionMethods.None;

    /// <summary>Optional progress lines reported before each login step.</summary>
    public string? TokenMessage { get; init; }
    public string? LoginMessage { get; init; }
}

public enum BmesLoginFailure { None, MissingCredentials, Token, Login }

/// <param name="Session">Logged-in session, or null when a step failed. The caller owns
/// and disposes it.</param>
/// <param name="Diagnostic">For <see cref="BmesLoginFailure.Login"/>: HTTP status plus the
/// server's Result/Message, safe to show (no credentials).</param>
public sealed record BmesLoginOutcome(BmesWebSession? Session, BmesLoginFailure Failure, string Diagnostic)
{
    public bool Succeeded => Session is not null;

    internal static BmesLoginOutcome Failed(BmesLoginFailure failure, string diagnostic) =>
        new(null, failure, diagnostic);
}
