using System.Net;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Fetches material master records (MES020010/SearchList) from BMES and persists
/// them to the BmesMaterials table. Reuses the same login/token flow as NgRateService.
/// </summary>
public sealed class BmesMaterialService(
    NgRateSettingsService settings,
    WebRepository repo)
{
    private const string BaseUrl = "http://bmes.bujeon.com";

    /// <summary>
    /// Fetches all materials from MES020010/SearchList and upserts them.
    /// Returns the number of rows saved, or -1 on failure.
    /// </summary>
    public async Task<int> FetchAllAsync(IProgress<string>? progress = null)
    {
        string loginId  = settings.LoginId;
        string password = settings.Password;
        if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
        {
            progress?.Report("[ERROR] BMES credentials not configured.");
            return -1;
        }

        using var handler = new HttpClientHandler
        {
            UseCookies      = true,
            CookieContainer = new CookieContainer(),
        };
        using var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(300) };
        client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");

        progress?.Report("Fetching verification token…");
        string token = await GetTokenAsync(client);
        if (string.IsNullOrEmpty(token))
        {
            progress?.Report("[ERROR] Failed to obtain token.");
            return -1;
        }

        progress?.Report("Logging in…");
        if (!await LoginAsync(client, token, loginId, password))
        {
            progress?.Report("[ERROR] Login failed — check credentials in BMES Setting.");
            return -1;
        }

        progress?.Report("Calling MES020010/SearchList…");
        string url =
            $"{BaseUrl}/MES020010/SearchList?perPage=" +
            "&Condition.MATNR=&Condition.MAKTX=&Condition.SAVYN=Y&Condition.KEYWO=" +
            "&Condition.GRCOD=&Condition.GRCOD_TX=&Condition.MTYPE=&Condition.BTYPE=&Condition.INJTP=" +
            "&Condition.SDATE=1900-01-01&Condition.EDATE=2050-12-31&page=1";

        HttpResponseMessage resp;
        try   { resp = await client.GetAsync(url); }
        catch (Exception ex)
        {
            progress?.Report($"[ERROR] Request failed: {ex.Message}");
            return -1;
        }
        if (!resp.IsSuccessStatusCode)
        {
            progress?.Report($"[ERROR] HTTP {(int)resp.StatusCode} {resp.StatusCode}");
            return -1;
        }

        string json = await resp.Content.ReadAsStringAsync();
        List<BmesMaterial> list;
        try
        {
            using var doc = JsonDocument.Parse(json);
            JsonElement contents = doc.RootElement.GetProperty("data").GetProperty("contents");
            if (contents.ValueKind != JsonValueKind.Array || contents.GetArrayLength() == 0)
            {
                progress?.Report("[ERROR] No contents in response.");
                return -1;
            }

            string fetchedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            list = new List<BmesMaterial>(contents.GetArrayLength());
            foreach (var item in contents.EnumerateArray())
            {
                list.Add(new BmesMaterial
                {
                    Matnr     = ReadStr(item, "MATNR"),
                    Maktx     = ReadStr(item, "MAKTX"),
                    Meins     = ReadStr(item, "MEINS"),
                    Injtp     = ReadStr(item, "INJTP"),
                    Mtype     = ReadStr(item, "MTYPE"),
                    Btype     = ReadStr(item, "BTYPE"),
                    MngCode   = ReadStr(item, "MNGCODE").Trim(),
                    ModNameB  = ReadStr(item, "MODNAME_B"),
                    LotQt     = ReadStr(item, "LOTQT"),
                    Bunch     = ReadStr(item, "BUNCH"),
                    NgTar     = ReadStr(item, "NGTAR"),
                    McLv1Tx   = ReadStr(item, "MCLV1_TX"),
                    McLv2Tx   = ReadStr(item, "MCLV2_TX"),
                    McLv3Tx   = ReadStr(item, "MCLV3_TX"),
                    McLv4Tx   = ReadStr(item, "MCLV4_TX"),
                    McLv5Tx   = ReadStr(item, "MCLV5_TX"),
                    McLv6Tx   = ReadStr(item, "MCLV6_TX"),
                    Ernam     = ReadStr(item, "ERNAM"),
                    Erdat     = ReadStr(item, "ERDAT"),
                    Grcod     = ReadStr(item, "GRCOD"),
                    Grnam     = ReadStr(item, "GRNAM"),
                    MfPhi     = ReadStr(item, "MFPHI"),
                    FetchedAt = fetchedAt,
                });
            }
        }
        catch (Exception ex)
        {
            progress?.Report($"[ERROR] Parse failed: {ex.Message}");
            return -1;
        }

        progress?.Report($"Parsed {list.Count:N0} rows. Saving to DB…");
        int saved = await Task.Run(() => repo.UpsertBmesMaterials(list));
        progress?.Report($"✓ Saved {saved:N0} material(s).");
        return saved;
    }

    private static string ReadStr(JsonElement obj, string key)
    {
        if (!obj.TryGetProperty(key, out var prop)) return string.Empty;
        return prop.ValueKind switch
        {
            JsonValueKind.Null      => string.Empty,
            JsonValueKind.String    => prop.GetString() ?? string.Empty,
            _                       => prop.ToString(),
        };
    }

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

    private static async Task<bool> LoginAsync(
        HttpClient client, string token, string id, string password)
    {
        var content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["UserInfo.USRID"] = id,
            ["UserInfo.PWNO"]  = password,
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
}
