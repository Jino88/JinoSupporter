using System.Net;
using System.Text.RegularExpressions;

namespace JinoSupporter.Web.Services;

/// <summary>
/// The one place the app logs in to the BMES web server. Every BMES screen call goes
/// through a session opened here and then through <see cref="BmesGateway"/>, so when BMES
/// hands out a real API only this folder (Services/Bmes) has to change.
///
/// Flow (same as the BMES login page): read the anti-forgery token from the home page,
/// post it with the configured credentials to MES000000/LoginCheck, keep the cookie.
/// Each session gets its own cookie container; browser cookies or copied SSO tokens are
/// never used.
/// </summary>
public sealed class BmesWebSession : IDisposable
{
    public const string BaseUrl = "https://bmes.bujeon.com";
    public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(300);

    private readonly HttpClientHandler _handler;

    public HttpClient Client { get; }

    private BmesWebSession(BmesWebSessionOptions options)
    {
        _handler = new HttpClientHandler
        {
            UseCookies = true,
            CookieContainer = new CookieContainer(),
            AutomaticDecompression = options.AutomaticDecompression,
        };
        Client = new HttpClient(_handler) { Timeout = options.Timeout ?? DefaultTimeout };
        if (options.AjaxHeader)
            Client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");
        if (options.Referrer is not null)
            Client.DefaultRequestHeaders.Referrer = options.Referrer;
    }

    /// <summary>Opens a logged-in session. On failure the result carries no session and
    /// says which step failed, so callers can keep their own error wording.
    /// Like every BMES login before the gateway existed, a timeout or cancellation during
    /// login is a failed step, not an exception.</summary>
    public static async Task<BmesLoginOutcome> OpenAsync(
        string loginId,
        string password,
        BmesWebSessionOptions? options = null,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
            return BmesLoginOutcome.Failed(BmesLoginFailure.MissingCredentials, "BMES credentials are not configured.");

        options ??= new BmesWebSessionOptions();
        var session = new BmesWebSession(options);
        try
        {
            if (options.TokenMessage is { Length: > 0 } tokenMessage) progress?.Report(tokenMessage);
            string token = await session.GetTokenAsync(cancellationToken);
            if (string.IsNullOrEmpty(token))
            {
                session.Dispose();
                return BmesLoginOutcome.Failed(BmesLoginFailure.Token, "Failed to read BMES login token.");
            }

            if (options.LoginMessage is { Length: > 0 } loginMessage) progress?.Report(loginMessage);
            (bool ok, string diagnostic) = await session.LoginAsync(token, loginId, password, cancellationToken);
            if (!ok)
            {
                session.Dispose();
                return BmesLoginOutcome.Failed(BmesLoginFailure.Login, diagnostic);
            }

            return new BmesLoginOutcome(session, BmesLoginFailure.None, string.Empty);
        }
        catch
        {
            session.Dispose();
            throw;
        }
    }

    public void Dispose()
    {
        Client.Dispose();
        _handler.Dispose();
    }

    private async Task<string> GetTokenAsync(CancellationToken cancellationToken)
    {
        try
        {
            string html = await Client.GetStringAsync(BaseUrl, cancellationToken);
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
            using var response = await Client.PostAsync(BaseUrl + "/MES000000/LoginCheck", content, cancellationToken);
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
