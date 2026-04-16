using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace JinoSupporter.Web.Services;

/// <summary>
/// BMES MES072900/SearchPopupDetail — 근태(출퇴근) 조회 서비스
/// </summary>
public sealed class WorkerStatusService(NgRateSettingsService settings)
{
    private readonly NgRateSettingsService _settings = settings;
    private const string BaseUrl = "http://bmes.bujeon.com";

    // ── 응답 모델 ─────────────────────────────────────────────────────────────────

    public sealed class WorkerRecord
    {
        /// <summary>사번 (EMPNO)</summary>
        public string EmpNo       { get; init; } = string.Empty;
        /// <summary>이름 (FNAME)</summary>
        public string Name        { get; init; } = string.Empty;
        /// <summary>근무상태 (WKSTA_TX, e.g. "In office")</summary>
        public string WorkStatus  { get; init; } = string.Empty;
        /// <summary>근무유형 (DTYPE_TX, e.g. "Normal")</summary>
        public string DayType     { get; init; } = string.Empty;
        /// <summary>출근 예정 시간 (WSTIM)</summary>
        public string SchedStart  { get; init; } = string.Empty;
        /// <summary>퇴근 예정 시간 (WETIM)</summary>
        public string SchedEnd    { get; init; } = string.Empty;
        /// <summary>실제 출근 시간 (SDATM)</summary>
        public string CheckIn     { get; init; } = string.Empty;
        /// <summary>실제 퇴근 시간 (EDATM)</summary>
        public string CheckOut    { get; init; } = string.Empty;
        /// <summary>공장코드 (FACCO)</summary>
        public string Factory     { get; init; } = string.Empty;
        /// <summary>원본 JSON 전체</summary>
        public Dictionary<string, string> Raw { get; init; } = new();
    }

    public sealed class FetchResult
    {
        public bool               IsSuccess    { get; set; }
        public string             ErrorMessage { get; set; } = string.Empty;
        public List<WorkerRecord> Records      { get; set; } = new();
        public int                TotalCount   { get; set; }
        public DateTime           FetchedAt    { get; set; } = DateTime.Now;
    }

    // ── Public API ────────────────────────────────────────────────────────────────

    public async Task<FetchResult> FetchAsync(
        DateTime date,
        IProgress<string>? progress = null)
    {
        var result = new FetchResult();

        if (!_settings.IsCredentialsConfigured)
        {
            result.ErrorMessage = "BMES credentials not configured. Go to BMES → Setting.";
            return result;
        }

        using var handler = new HttpClientHandler
        {
            UseCookies      = true,
            CookieContainer = new CookieContainer(),
        };
        using var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(60) };
        client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");

        // 1. Token
        progress?.Report("Fetching verification token…");
        string token = await GetTokenAsync(client);
        if (string.IsNullOrEmpty(token))
        {
            result.ErrorMessage = "Failed to obtain CSRF token.";
            return result;
        }

        // 2. Login
        progress?.Report("Logging in to BMES…");
        if (!await LoginAsync(client, token))
        {
            result.ErrorMessage = "Login failed — check credentials in BMES Setting.";
            return result;
        }
        progress?.Report("Login successful.");

        // 3. Fetch
        progress?.Report($"Fetching Worker Status for {date:yyyy-MM-dd}…");
        try
        {
            var records = await FetchWorkerStatusAsync(client, date);
            result.Records    = records;
            result.TotalCount = records.Count;
            result.IsSuccess  = true;
            result.FetchedAt  = DateTime.Now;
            progress?.Report($"Done — {records.Count:N0} records.");
        }
        catch (Exception ex)
        {
            result.ErrorMessage = $"Fetch error: {ex.Message}";
            progress?.Report($"[ERROR] {ex.Message}");
        }

        return result;
    }

    // ── Private: HTTP ──────────────────────────────────────────────────────────────

    private static async Task<string> GetTokenAsync(HttpClient client)
    {
        try
        {
            string html = await client.GetStringAsync(BaseUrl);
            var m = Regex.Match(html,
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

    private async Task<bool> LoginAsync(HttpClient client, string token)
    {
        var content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["UserInfo.USRID"] = _settings.LoginId,
            ["UserInfo.PWNO"]  = _settings.Password,
            ["UserInfo.LANG"]  = "EN",
            ["UserInfo.FACCO"] = "GN",
            ["UserInfo.STYPE"] = "P",
            ["UserInfo.VTYPE"] = "P",
            ["__RequestVerificationToken"] = token,
        });
        try
        {
            var response = await client.PostAsync(BaseUrl + "/MES000000/LoginCheck", content);
            string body  = await response.Content.ReadAsStringAsync();
            return body.Contains("\"Result\":\"M\"");
        }
        catch { return false; }
    }

    private static async Task<List<WorkerRecord>> FetchWorkerStatusAsync(
        HttpClient client,
        DateTime date)
    {
        string dateStr = date.ToString("yyyy-MM-dd");
        var bodyObj = new
        {
            Condition = new
            {
                ZTYPE = "A",
                STDAT = dateStr,
                DEPNO = "G0011",
                FACCO = "GN",
                KWORD = "",
            }
        };

        string bodyJson = JsonSerializer.Serialize(bodyObj);
        using var requestContent = new StringContent(bodyJson, Encoding.UTF8, "application/json");

        var response = await client.PostAsync(BaseUrl + "/MES072900/SearchPopupDetail", requestContent);
        response.EnsureSuccessStatusCode();

        string json = await response.Content.ReadAsStringAsync();
        return ParseResponse(json);
    }

    // ── Parse ──────────────────────────────────────────────────────────────────────

    private static List<WorkerRecord> ParseResponse(string json)
    {
        var records = new List<WorkerRecord>();
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            // { "result": true, "data": { "contents": [...] } }
            JsonElement? arr = null;
            if (root.ValueKind == JsonValueKind.Array)
                arr = root;
            else if (root.TryGetProperty("data", out var data))
            {
                if (data.TryGetProperty("contents", out var contents))
                    arr = contents;
                else if (data.ValueKind == JsonValueKind.Array)
                    arr = data;
            }
            else if (root.TryGetProperty("rows", out var rows))
                arr = rows;
            else if (root.TryGetProperty("contents", out var contents2))
                arr = contents2;

            if (arr is null) return records;

            foreach (var item in arr.Value.EnumerateArray())
            {
                var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var prop in item.EnumerateObject())
                    raw[prop.Name] = prop.Value.ValueKind == JsonValueKind.Null
                        ? string.Empty : prop.Value.ToString();

                records.Add(new WorkerRecord
                {
                    EmpNo      = raw.GetValueOrDefault("EMPNO",    string.Empty),
                    Name       = raw.GetValueOrDefault("FNAME",    string.Empty),
                    WorkStatus = raw.GetValueOrDefault("WKSTA_TX", string.Empty),
                    DayType    = raw.GetValueOrDefault("DTYPE_TX", string.Empty),
                    SchedStart = raw.GetValueOrDefault("WSTIM",    string.Empty),
                    SchedEnd   = raw.GetValueOrDefault("WETIM",    string.Empty),
                    CheckIn    = raw.GetValueOrDefault("SDATM",    string.Empty),
                    CheckOut   = raw.GetValueOrDefault("EDATM",    string.Empty),
                    Factory    = raw.GetValueOrDefault("FACCO",    string.Empty),
                    Raw        = raw,
                });
            }
        }
        catch { }
        return records;
    }
}
