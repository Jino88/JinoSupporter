using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Fetches LPA (Layered Process Audit) records from BMES MES073260/SearchList.
///
/// Unlike the other BMES scrapers this one posts a JSON body rather than a query string,
/// and it does not persist anything — the page just renders what comes back.
///
/// The response schema is not pinned down anywhere, so rows are read generically: every
/// property of every returned object becomes a column, ordered by first appearance. That
/// way a server-side schema change shows up as an extra column instead of silently
/// dropping data.
/// </summary>
public sealed class BmesLpaScrapeService(NgRateSettingsService settings)
{
    private const string BaseUrl        = "https://bmes.bujeon.com";
    private const string ListEndpoint   = "/MES073260/SearchList";
    private const string DetailEndpoint = "/MES073261/SearchList";

    /// <summary>Search conditions. Defaults mirror the "all records" query.</summary>
    public sealed record LpaQuery
    {
        public string Facco   { get; init; } = "GN";
        public string DateFrom { get; init; } = "19000101";
        public string DateTo   { get; init; } = "20501231";
        public string Lqrno   { get; init; } = "";
        public string Lqbno   { get; init; } = "";
        public string Auloc   { get; init; } = "";
        public string Implv   { get; init; } = "";
        public string Dicno   { get; init; } = "";
        public string Laseq   { get; init; } = "";
        public string Chker   { get; init; } = "";
        public string Zstat   { get; init; } = "";
        public string UseYn   { get; init; } = "Y";
    }

    public sealed record LpaResult(
        IReadOnlyList<string> Columns,
        IReadOnlyList<IReadOnlyDictionary<string, string>> Rows,
        string? Error)
    {
        public bool IsSuccess => Error is null;
        public static LpaResult Fail(string message) => new([], [], message);
    }

    /// <summary>Conditions for the per-audit checklist behind one LQRNO.</summary>
    public sealed record LpaDetailQuery
    {
        public required string Lqrno { get; init; }
        public string Facco { get; init; } = "GN";
        public string Dicno { get; init; } = "";
        public string Auloc { get; init; } = "";
        public string Implv { get; init; } = "";
        public string Lcitm { get; init; } = "";
    }

    private static object ListCondition(LpaQuery query) => new
    {
        Condition = new
        {
            FACCO    = query.Facco,
            AUDAT_FR = query.DateFrom,
            AUDAT_TO = query.DateTo,
            LQRNO    = query.Lqrno,
            LQBNO    = query.Lqbno,
            AULOC    = query.Auloc,
            IMPLV    = query.Implv,
            DICNO    = query.Dicno,
            LASEQ    = query.Laseq,
            CHKER    = query.Chker,
            ZSTAT    = query.Zstat,
            USEYN    = query.UseYn,
        },
    };

    private static object DetailCondition(LpaDetailQuery query) => new
    {
        Condition = new
        {
            LQRNO = query.Lqrno,
            FACCO = query.Facco,
            DICNO = query.Dicno,
            AULOC = query.Auloc,
            IMPLV = query.Implv,
            LCITM = query.Lcitm,
        },
    };

    public async Task<LpaResult> FetchAsync(LpaQuery query, IProgress<string>? progress = null)
    {
        var (session, error) = await OpenSessionAsync(progress);
        if (session is null) return LpaResult.Fail(error!);
        using (session)
            return await PostSearchAsync(session, ListEndpoint, ListCondition(query), progress);
    }

    public async Task<LpaResult> FetchDetailAsync(LpaDetailQuery query, IProgress<string>? progress = null)
    {
        var (session, error) = await OpenSessionAsync(progress);
        if (session is null) return LpaResult.Fail(error!);
        using (session)
            return await PostSearchAsync(session, DetailEndpoint, DetailCondition(query), progress);
    }

    /// <summary>
    /// How many detail lookups run at once inside one batch. The token+login handshake is
    /// already paid once for the whole batch, so this only bounds how hard BMES gets hit —
    /// kept low deliberately: MES073261 is a legacy endpoint and the point of the batch is
    /// to avoid hammering it, not to race it.
    /// </summary>
    private const int DetailBatchConcurrency = 4;

    /// <summary>
    /// Fetch the checklist behind many LQRNOs over a SINGLE authenticated session.
    ///
    /// <see cref="FetchDetailAsync"/> opens a fresh session per call, and that handshake
    /// (GET token → POST login) costs more than the search itself. Prefetching a whole
    /// result set one row at a time would therefore mean one login per row; this logs in
    /// once and reuses the cookie for every lookup.
    ///
    /// Returns one entry per distinct LQRNO. A row that failed carries its own
    /// <see cref="LpaResult.Error"/> instead of failing the batch — a partial prefetch is
    /// still worth showing.
    /// </summary>
    public async Task<IReadOnlyDictionary<string, LpaResult>> FetchDetailsAsync(
        IReadOnlyList<LpaDetailQuery> queries, IProgress<string>? progress = null,
        IProgress<(int Done, int Total)>? counter = null)
    {
        var results = new Dictionary<string, LpaResult>(StringComparer.Ordinal);
        if (queries.Count == 0) return results;

        var (session, error) = await OpenSessionAsync(progress);
        if (session is null)
        {
            foreach (var q in queries) results[q.Lqrno] = LpaResult.Fail(error!);
            return results;
        }

        using (session)
        {
            int done = 0;
            using var gate = new SemaphoreSlim(DetailBatchConcurrency);
            var tasks = queries.Select(async q =>
            {
                await gate.WaitAsync();
                try
                {
                    // No per-call progress: the batch reports its own count instead, which
                    // is the only number that means anything across hundreds of lookups.
                    LpaResult r = await PostSearchAsync(session, DetailEndpoint, DetailCondition(q), null);
                    int n = Interlocked.Increment(ref done);
                    if (n % 10 == 0 || n == queries.Count)
                    {
                        progress?.Report($"Fetching detail {n:N0} / {queries.Count:N0}…");
                        counter?.Report((n, queries.Count));
                    }
                    return (q.Lqrno, r);
                }
                finally { gate.Release(); }
            }).ToList();

            foreach ((string lqrno, LpaResult r) in await Task.WhenAll(tasks))
                results[lqrno] = r;   // duplicate LQRNOs collapse; they carry the same detail
        }

        return results;
    }

    /// <summary>An authenticated BMES connection: cookie container + logged-in session.</summary>
    private sealed class Session(HttpClientHandler handler, HttpClient client) : IDisposable
    {
        public HttpClient Client { get; } = client;

        public void Dispose()
        {
            Client.Dispose();
            handler.Dispose();
        }
    }

    /// <summary>GET verification token → POST login. Returns the error text on failure.</summary>
    private async Task<(Session? Session, string? Error)> OpenSessionAsync(IProgress<string>? progress)
    {
        string loginId  = settings.LoginId;
        string password = settings.Password;
        if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
            return (null, "BMES credentials are not configured. Set them in Setting.");

        var handler = new HttpClientHandler
        {
            UseCookies      = true,
            CookieContainer = new System.Net.CookieContainer(),
        };
        var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(300) };
        client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");
        var session = new Session(handler, client);

        progress?.Report("Fetching verification token…");
        string token = await GetTokenAsync(client);
        if (string.IsNullOrEmpty(token))
        {
            session.Dispose();
            return (null, "Failed to obtain the verification token from BMES.");
        }

        progress?.Report("Logging in…");
        if (!await LoginAsync(client, token, loginId, password))
        {
            session.Dispose();
            return (null, "BMES login failed — check the credentials in Setting.");
        }

        return (session, null);
    }

    /// <summary>JSON POST → generic row parse, on an already-authenticated session. Shared by
    /// both endpoints, which differ only in path and condition object.</summary>
    private static async Task<LpaResult> PostSearchAsync(
        Session session, string endpoint, object condition, IProgress<string>? progress)
    {
        string payload = JsonSerializer.Serialize(condition);

        progress?.Report($"Calling {endpoint}…");
        HttpResponseMessage response;
        try
        {
            using var content = new StringContent(payload, Encoding.UTF8, "application/json");
            response = await session.Client.PostAsync(BaseUrl + endpoint, content);
        }
        catch (Exception ex)
        {
            return LpaResult.Fail($"Request failed: {ex.Message}");
        }

        if (!response.IsSuccessStatusCode)
            return LpaResult.Fail($"HTTP {(int)response.StatusCode} {response.StatusCode}");

        string json = await response.Content.ReadAsStringAsync();
        try
        {
            var (columns, rows) = ParseRows(json);
            progress?.Report($"✓ {rows.Count:N0} row(s), {columns.Count} column(s).");
            return new LpaResult(columns, rows, null);
        }
        catch (Exception ex)
        {
            return LpaResult.Fail($"Failed to parse the response: {ex.Message}");
        }
    }

    /// <summary>
    /// Pulls the row array out of the response and flattens it.
    ///
    /// The wrapper shape is not documented and differs between BMES endpoints
    /// (<c>data.contents</c>, <c>data</c>, a bare array…), so rather than guessing a path
    /// this walks the whole document and takes the largest array of objects it finds. On
    /// a result payload that is the row set by a wide margin — metadata arrays alongside
    /// it are short.
    /// </summary>
    private static (List<string> Columns, List<IReadOnlyDictionary<string, string>> Rows) ParseRows(string json)
    {
        using var doc = JsonDocument.Parse(json);

        JsonElement? best = null;
        int bestCount = 0;
        FindLargestObjectArray(doc.RootElement, 0, ref best, ref bestCount);

        var columns = new List<string>();
        var seen    = new HashSet<string>(StringComparer.Ordinal);
        var rows    = new List<IReadOnlyDictionary<string, string>>();
        if (best is not { } array) return (columns, rows);

        foreach (JsonElement item in array.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object) continue;
            // Case-insensitive: callers look columns up by name and BMES casing varies.
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (JsonProperty prop in item.EnumerateObject())
            {
                if (seen.Add(prop.Name)) columns.Add(prop.Name);
                row[prop.Name] = Stringify(prop.Value);
            }
            rows.Add(row);
        }

        return (columns, rows);
    }

    private const int MaxScanDepth = 8;

    private static void FindLargestObjectArray(
        JsonElement element, int depth, ref JsonElement? best, ref int bestCount)
    {
        if (depth > MaxScanDepth) return;

        switch (element.ValueKind)
        {
            case JsonValueKind.Array:
                int objects = 0;
                foreach (JsonElement item in element.EnumerateArray())
                    if (item.ValueKind == JsonValueKind.Object) objects++;

                if (objects > bestCount)
                {
                    best = element;
                    bestCount = objects;
                }
                // Rows can still be nested one level deeper (e.g. grouped results).
                foreach (JsonElement item in element.EnumerateArray())
                    FindLargestObjectArray(item, depth + 1, ref best, ref bestCount);
                break;

            case JsonValueKind.Object:
                foreach (JsonProperty prop in element.EnumerateObject())
                    FindLargestObjectArray(prop.Value, depth + 1, ref best, ref bestCount);
                break;
        }
    }

    private static string Stringify(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null or JsonValueKind.Undefined => string.Empty,
        JsonValueKind.String => value.GetString() ?? string.Empty,
        JsonValueKind.True   => "Y",
        JsonValueKind.False  => "N",
        _ => value.ToString(),
    };

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
