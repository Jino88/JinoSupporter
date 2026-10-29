using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;

namespace PortalGNLAB.Web.Services;

/// <summary>PLATFORM_API login: facility (BN / GN), BMES user id and password.</summary>
public sealed record BmesApiCredentials(string Facility, string UserId, string Password);

/// <summary>PLATFORM_API failure. Status 0 = connection failure / timeout; TraceId is the number to quote to the API team.</summary>
public sealed class BmesApiException(int status, string code, string message, string? traceId = null) : Exception(message)
{
    public int Status { get; } = status;
    public string Code { get; } = code;
    public string? TraceId { get; } = traceId;
}

/// <summary>One result set of a Tool answer: column names plus rows of raw JSON values in column order.</summary>
public sealed record BmesApiResultSet(IReadOnlyList<string> Columns, IReadOnlyList<JsonElement[]> Rows)
{
    public int Column(string name)
    {
        for (int i = 0; i < Columns.Count; i++)
            if (string.Equals(Columns[i], name, StringComparison.OrdinalIgnoreCase)) return i;
        return -1;
    }
}

/// <param name="Result">S success, B parameter problem, E no result / error.</param>
public sealed record BmesApiResult(string ToolName, string Result, string Message, IReadOnlyList<BmesApiResultSet> ResultSets)
{
    /// <summary>"No result" is a normal empty answer, not a failure (message code IVDT_M_018).</summary>
    public bool IsEmpty => Result == "E" && Message.Contains("IVDT_M_018", StringComparison.OrdinalIgnoreCase);
}

/// <summary>
/// The one PLATFORM_API (https://api.bujeon.com) client of the app. Rules it follows:
/// login once per credential set (access token 15 min, refresh token), on 401 refresh exactly once
/// (never two refreshes at the same time), one call at a time with at least one second between calls,
/// 60 seconds per call, X-Client header on every request.
/// Gateway methods call <see cref="ExecuteAsync"/>; nothing outside Services/Bmes talks to the API.
/// </summary>
public static class BmesApi
{
    public const string BaseUrl = "https://api.bujeon.com";
    public const string ClientName = "python-web";
    public static readonly TimeSpan CallSpacing = TimeSpan.FromSeconds(1);
    public static readonly TimeSpan CallTimeout = TimeSpan.FromSeconds(60);

    /// <summary>Longest period one call may cover (Tool rule: 31 days or less).</summary>
    public const int MaxPeriodDays = 31;

    private sealed class Tokens(string access, string refresh)
    {
        public string Access = access;
        public string Refresh = refresh;
    }

    private static readonly HttpClient Http = new() { Timeout = CallTimeout };
    private static readonly SemaphoreSlim Gate = new(1, 1);          // serialises calls, login and refresh
    private static readonly Dictionary<string, Tokens> Sessions = new();
    private static DateTime _lastCallUtc = DateTime.MinValue;

    /// <summary>Runs one Tool and returns its result sets. An "E" answer other than "no result" and every
    /// "B" answer throw <see cref="BmesApiException"/>; "no result" returns an empty result.</summary>
    public static async Task<BmesApiResult> ExecuteAsync(
        BmesApiCredentials credentials, string toolName, IReadOnlyDictionary<string, object> parameters,
        string language = "EN", CancellationToken cancellationToken = default)
    {
        await Gate.WaitAsync(cancellationToken);
        try
        {
            string key = credentials.Facility + "|" + credentials.UserId;
            if (!Sessions.TryGetValue(key, out Tokens? tokens))
            {
                tokens = await LoginAsync(credentials, cancellationToken);
                Sessions[key] = tokens;
            }

            string text;
            try
            {
                text = await PostToolAsync(tokens.Access, toolName, language, parameters, cancellationToken);
            }
            catch (BmesApiException e) when (e.Status == 401)
            {
                try { await RefreshAsync(tokens, cancellationToken); }
                catch (BmesApiException r) when (r.Status == 401)
                {
                    // the API session ended: log in again once
                    tokens = await LoginAsync(credentials, cancellationToken);
                    Sessions[key] = tokens;
                }
                text = await PostToolAsync(tokens.Access, toolName, language, parameters, cancellationToken);
            }

            BmesApiResult result = BmesToolRows.ParseExecuteResponse(text);
            if (result.Result == "S" || result.IsEmpty) return result;
            throw new BmesApiException(0, result.Result == "B" ? "BAD_PARAMETER" : "TOOL_ERROR",
                $"{toolName}: {result.Message}");
        }
        finally
        {
            Gate.Release();
        }
    }

    private static async Task<Tokens> LoginAsync(BmesApiCredentials c, CancellationToken ct)
    {
        string text = await SendAsync("/api/v1/auth/login", null, new
        {
            facilityCode = c.Facility,
            userId = c.UserId,
            password = c.Password,
            language = "EN",
            client = ClientName,
            deviceId = "PORTAL-" + Environment.MachineName,
        }, ct);
        return ReadTokens(text);
    }

    private static async Task RefreshAsync(Tokens tokens, CancellationToken ct)
    {
        string text = await SendAsync("/api/v1/auth/refresh", null, new { refreshToken = tokens.Refresh }, ct);
        Tokens fresh = ReadTokens(text);
        tokens.Access = fresh.Access;
        tokens.Refresh = fresh.Refresh;
    }

    private static Task<string> PostToolAsync(string access, string toolName, string language,
        IReadOnlyDictionary<string, object> parameters, CancellationToken ct) =>
        SendAsync("/api/v1/tools/execute", access, new { toolName, language, parameters }, ct);

    private static async Task<string> SendAsync(string path, string? access, object body, CancellationToken ct)
    {
        // at least one second between two API calls
        TimeSpan wait = _lastCallUtc + CallSpacing - DateTime.UtcNow;
        if (wait > TimeSpan.Zero) await Task.Delay(wait, ct);

        using var request = new HttpRequestMessage(HttpMethod.Post, BaseUrl + path);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        request.Headers.Add("X-Client", ClientName);
        if (access is not null) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", access);
        request.Content = JsonContent.Create(body);

        try
        {
            using HttpResponseMessage response = await Http.SendAsync(request, ct);
            string text = await response.Content.ReadAsStringAsync(ct);
            if (!response.IsSuccessStatusCode) throw ToError(response.StatusCode, text);
            return text;
        }
        catch (TaskCanceledException) when (!ct.IsCancellationRequested)
        {
            throw new BmesApiException(0, "TIMEOUT", "BMES API did not answer within 60 seconds.");
        }
        catch (HttpRequestException)
        {
            throw new BmesApiException(0, "CONNECTION_FAILED", "Could not connect to BMES API (api.bujeon.com).");
        }
        finally
        {
            _lastCallUtc = DateTime.UtcNow;
        }
    }

    private static Tokens ReadTokens(string text)
    {
        using JsonDocument doc = JsonDocument.Parse(text);
        JsonElement root = doc.RootElement;
        string? access = root.TryGetProperty("accessToken", out var a) ? a.GetString() : null;
        string? refresh = root.TryGetProperty("refreshToken", out var r) ? r.GetString() : null;
        if (string.IsNullOrEmpty(access) || string.IsNullOrEmpty(refresh))
            throw new BmesApiException(0, "BAD_RESPONSE", "BMES API login answer has no tokens.");
        return new Tokens(access, refresh);
    }

    /// <summary>Common error shape { error: { code, message }, traceId }.</summary>
    private static BmesApiException ToError(HttpStatusCode status, string text)
    {
        try
        {
            using JsonDocument doc = JsonDocument.Parse(text);
            JsonElement root = doc.RootElement;
            string? trace = root.TryGetProperty("traceId", out var t) && t.ValueKind == JsonValueKind.String ? t.GetString() : null;
            if (root.TryGetProperty("error", out var error) && error.ValueKind == JsonValueKind.Object)
            {
                string code = error.TryGetProperty("code", out var c) && c.ValueKind == JsonValueKind.String ? c.GetString()! : "ERROR";
                string message = error.TryGetProperty("message", out var m) && m.ValueKind == JsonValueKind.String ? m.GetString()! : "";
                return new BmesApiException((int)status, code, message, trace);
            }
        }
        catch (JsonException) { }
        return new BmesApiException((int)status, "HTTP_ERROR", $"HTTP {(int)status}");
    }
}

/// <summary>
/// Where the API login comes from. Program.cs hands it the settings once; feature services keep
/// calling <c>BmesWebSession.OpenAsync(loginId, password, ...)</c> unchanged.
/// API id / password from the settings DB (BmesApi:*) win; when they are empty the BMES web login
/// of the same call is used. Facility default GN = the facility saved in the old portal's BMES settings.
/// </summary>
public static class BmesApiCredentialProvider
{
    public const string DefaultFacility = "GN";

    private static NgRateSettingsService? _settings;

    public static void Configure(NgRateSettingsService settings) => _settings = settings;

    /// <summary>Null when neither the API login nor the web login is available.</summary>
    public static BmesApiCredentials? Resolve(string? webLoginId, string? webPassword)
    {
        string facility = _settings?.ApiFacility.Trim() ?? "";
        if (facility.Length == 0) facility = DefaultFacility;

        if (_settings is { } s && !string.IsNullOrWhiteSpace(s.ApiLoginId) && !string.IsNullOrWhiteSpace(s.ApiPassword))
            return new BmesApiCredentials(facility, s.ApiLoginId.Trim(), s.ApiPassword);
        if (!string.IsNullOrWhiteSpace(webLoginId) && !string.IsNullOrWhiteSpace(webPassword))
            return new BmesApiCredentials(facility, webLoginId.Trim(), webPassword);
        return null;
    }

    /// <summary>True when the API login alone is enough (no BMES web login needed).</summary>
    public static bool HasApiLogin => _settings is { } s && s.IsApiCredentialsConfigured;
}
