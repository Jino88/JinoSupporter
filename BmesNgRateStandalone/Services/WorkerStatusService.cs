using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace BmesNgRateStandalone.Services;

/// <summary>
/// BMES MES072900/SearchPopupDetail — attendance (check-in/out) fetch service.
/// </summary>
public sealed class WorkerStatusService(NgRateSettingsService settings)
{
    private readonly NgRateSettingsService _settings = settings;
    private const string BaseUrl = "http://bmes.bujeon.com";

    // ── Response model ────────────────────────────────────────────────────────────

    public sealed class WorkerRecord
    {
        /// <summary>The date this record belongs to (set by the fetcher, not from JSON).</summary>
        public DateTime Date      { get; set; }
        /// <summary>Employee number (EMPNO)</summary>
        public string EmpNo       { get; init; } = string.Empty;
        /// <summary>Name (FNAME)</summary>
        public string Name        { get; init; } = string.Empty;
        /// <summary>Department number used for SearchPopupDetail (DEPNO)</summary>
        public string DepartmentNo   { get; set; } = string.Empty;
        /// <summary>Group/part code from SearchList (DPCOD)</summary>
        public string DepartmentCode { get; set; } = string.Empty;
        /// <summary>Department name from SearchList (P_DEPTX/G_DEPTX/DEPTX)</summary>
        public string DepartmentName { get; set; } = string.Empty;
        /// <summary>Department type label from SearchList (DTYNM)</summary>
        public string DepartmentType { get; set; } = string.Empty;
        /// <summary>Work status (WKSTA_TX, e.g. "In office")</summary>
        public string WorkStatus  { get; init; } = string.Empty;
        /// <summary>Day type (DTYPE_TX, e.g. "Normal")</summary>
        public string DayType     { get; init; } = string.Empty;
        /// <summary>Scheduled check-in time (WSTIM)</summary>
        public string SchedStart  { get; init; } = string.Empty;
        /// <summary>Scheduled check-out time (WETIM)</summary>
        public string SchedEnd    { get; init; } = string.Empty;
        /// <summary>Actual check-in time (SDATM)</summary>
        public string CheckIn     { get; init; } = string.Empty;
        /// <summary>Actual check-out time (EDATM)</summary>
        public string CheckOut    { get; init; } = string.Empty;
        /// <summary>Factory code (FACCO)</summary>
        public string Factory     { get; init; } = string.Empty;
        /// <summary>Full original JSON row</summary>
        public Dictionary<string, string> Raw { get; init; } = new();
    }

    private sealed class DepartmentRecord
    {
        public string DepartmentNo   { get; init; } = string.Empty;
        public string DepartmentCode { get; init; } = string.Empty;
        public string DepartmentName { get; init; } = string.Empty;
        public string DepartmentType { get; init; } = string.Empty;
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

    public Task<FetchResult> FetchAsync(
        DateTime date,
        IProgress<string>? progress = null)
        => FetchRangeAsync(date, date, progress);

    public async Task<FetchResult> FetchRangeAsync(
        DateTime startDate,
        DateTime endDate,
        IProgress<string>? progress = null)
    {
        var result = new FetchResult();

        if (endDate < startDate)
            (startDate, endDate) = (endDate, startDate);

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

        // 3. Fetch - read all people first, then resolve department membership only for new people.
        var allRecords = new List<WorkerRecord>();
        int totalDays = (endDate - startDate).Days + 1;
        int dayIdx    = 0;

        List<DepartmentRecord> departments;
        try
        {
            departments = new List<DepartmentRecord>();
            for (var date = startDate.Date; date <= endDate.Date; date = date.AddDays(1))
            {
                progress?.Report($"Fetching departments for {date:yyyy-MM-dd}...");
                departments = await FetchDepartmentsAsync(client, date);
                if (departments.Count > 0)
                    break;
            }

            progress?.Report($"Found {departments.Count:N0} G09 department(s).");
        }
        catch (Exception ex)
        {
            result.ErrorMessage = $"Department fetch error: {ex.Message}";
            progress?.Report($"[ERROR] {ex.Message}");
            return result;
        }

        if (departments.Count == 0)
        {
            result.Records   = new List<WorkerRecord>();
            result.IsSuccess = true;
            result.FetchedAt = DateTime.Now;
            progress?.Report("No G09 departments found.");
            return result;
        }

        var aggregateDepartment = SelectAggregateDepartment(departments);
        progress?.Report(
            $"Using {aggregateDepartment.DepartmentCode} / {aggregateDepartment.DepartmentNo} for daily worker status.");

        for (var date = startDate.Date; date <= endDate.Date; date = date.AddDays(1))
        {
            dayIdx++;
            progress?.Report($"Fetching worker status for {date:yyyy-MM-dd} ({dayIdx}/{totalDays})...");
            try
            {
                var records = await FetchWorkerStatusAsync(client, date, aggregateDepartment);
                foreach (var r in records) r.Date = date;
                allRecords.AddRange(records);
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Fetch error on {date:yyyy-MM-dd}: {ex.Message}";
                progress?.Report($"[ERROR] {date:yyyy-MM-dd}: {ex.Message}");
                // keep what we have so far; continue
            }
        }

        var employeeFirstSeenDates = GetEmployeeFirstSeenDates(allRecords);
        progress?.Report(
            $"Mapping department membership for {employeeFirstSeenDates.Count:N0} people across {totalDays} day(s)...");
        var employeeDepartments = await BuildEmployeeDepartmentMapAsync(
            client,
            employeeFirstSeenDates,
            departments,
            aggregateDepartment,
            progress);
        ApplyDepartmentMap(allRecords, employeeDepartments, aggregateDepartment);

        result.Records    = SortRecords(DeduplicateRecords(allRecords)).ToList();
        result.TotalCount = result.Records.Count;
        result.IsSuccess  = result.Records.Count > 0 || string.IsNullOrEmpty(result.ErrorMessage);
        result.FetchedAt  = DateTime.Now;
        progress?.Report($"Done - {result.Records.Count:N0} records over {totalDays} day(s).");

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

    private static async Task<List<DepartmentRecord>> FetchDepartmentsAsync(
        HttpClient client,
        DateTime date)
    {
        string dateStr = date.ToString("yyyy-MM-dd");
        var bodyObj = new
        {
            Condition = new
            {
                FACCO = "GN",
                STDAT = dateStr,
                KWORD = "",
            }
        };

        string bodyJson = JsonSerializer.Serialize(bodyObj);
        using var requestContent = new StringContent(bodyJson, Encoding.UTF8, "application/json");

        var response = await client.PostAsync(BaseUrl + "/MES072900/SearchList", requestContent);
        response.EnsureSuccessStatusCode();

        string json = await response.Content.ReadAsStringAsync();
        return ParseDepartmentList(json)
            .Where(d => d.DepartmentCode.Contains("G09", StringComparison.OrdinalIgnoreCase))
            .Where(d => !string.IsNullOrWhiteSpace(d.DepartmentNo))
            .GroupBy(d => d.DepartmentNo, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.OrderByDescending(d => DepartmentSpecificity(d.DepartmentCode)).First())
            .OrderBy(d => d.DepartmentCode, StringComparer.OrdinalIgnoreCase)
            .ThenBy(d => d.DepartmentNo, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static async Task<List<WorkerRecord>> FetchWorkerStatusAsync(
        HttpClient client,
        DateTime date,
        DepartmentRecord department)
    {
        string dateStr = date.ToString("yyyy-MM-dd");
        var bodyObj = new
        {
            Condition = new
            {
                ZTYPE = "A",
                STDAT = dateStr,
                DEPNO = department.DepartmentNo,
                FACCO = "GN",
                KWORD = "",
            }
        };

        string bodyJson = JsonSerializer.Serialize(bodyObj);
        using var requestContent = new StringContent(bodyJson, Encoding.UTF8, "application/json");

        var response = await client.PostAsync(BaseUrl + "/MES072900/SearchPopupDetail", requestContent);
        response.EnsureSuccessStatusCode();

        string json = await response.Content.ReadAsStringAsync();
        return ParseResponse(json, department);
    }

    private static DepartmentRecord SelectAggregateDepartment(IReadOnlyList<DepartmentRecord> departments)
    {
        return departments.FirstOrDefault(d =>
                   string.Equals(d.DepartmentCode, "G09", StringComparison.OrdinalIgnoreCase))
               ?? departments
                   .OrderBy(d => DepartmentSpecificity(d.DepartmentCode))
                   .ThenBy(d => d.DepartmentCode, StringComparer.OrdinalIgnoreCase)
                   .ThenBy(d => d.DepartmentNo, StringComparer.OrdinalIgnoreCase)
                   .First();
    }

    private static Dictionary<string, DateTime> GetEmployeeFirstSeenDates(IEnumerable<WorkerRecord> records)
    {
        var result = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
        foreach (var record in records)
        {
            string key = EmployeeKey(record);
            if (string.IsNullOrEmpty(key)) continue;

            var date = record.Date.Date;
            if (!result.TryGetValue(key, out var existing) || date < existing)
                result[key] = date;
        }
        return result;
    }

    private static async Task<Dictionary<string, DepartmentRecord>> BuildEmployeeDepartmentMapAsync(
        HttpClient client,
        IReadOnlyDictionary<string, DateTime> employeeFirstSeenDates,
        IReadOnlyList<DepartmentRecord> departments,
        DepartmentRecord aggregateDepartment,
        IProgress<string>? progress)
    {
        var result = new Dictionary<string, DepartmentRecord>(StringComparer.OrdinalIgnoreCase);
        if (employeeFirstSeenDates.Count == 0) return result;

        var detailDepartments = departments
            .Where(d => !string.Equals(d.DepartmentNo, aggregateDepartment.DepartmentNo, StringComparison.OrdinalIgnoreCase))
            .OrderBy(d => DepartmentSpecificity(d.DepartmentCode))
            .ThenBy(d => d.DepartmentCode, StringComparer.OrdinalIgnoreCase)
            .ThenBy(d => d.DepartmentNo, StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var dateGroup in employeeFirstSeenDates
                     .GroupBy(kv => kv.Value.Date)
                     .OrderBy(g => g.Key))
        {
            var targetKeys = dateGroup
                .Select(kv => kv.Key)
                .Where(k => !result.ContainsKey(k))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (targetKeys.Count == 0) continue;

            progress?.Report(
                $"Mapping {targetKeys.Count:N0} new people from {dateGroup.Key:yyyy-MM-dd}...");

            foreach (var department in detailDepartments)
            {
                var records = await FetchWorkerStatusAsync(client, dateGroup.Key, department);
                foreach (var record in records)
                {
                    string key = EmployeeKey(record);
                    if (!targetKeys.Contains(key)) continue;

                    if (!result.TryGetValue(key, out var existing) ||
                        DepartmentSpecificity(department.DepartmentCode) >=
                        DepartmentSpecificity(existing.DepartmentCode))
                    {
                        result[key] = department;
                    }
                }
            }
        }

        return result;
    }

    private static void ApplyDepartmentMap(
        IEnumerable<WorkerRecord> records,
        IReadOnlyDictionary<string, DepartmentRecord> employeeDepartments,
        DepartmentRecord fallbackDepartment)
    {
        foreach (var record in records)
        {
            string key = EmployeeKey(record);
            var department = !string.IsNullOrEmpty(key) &&
                             employeeDepartments.TryGetValue(key, out var mapped)
                ? mapped
                : fallbackDepartment;

            record.DepartmentNo   = department.DepartmentNo;
            record.DepartmentCode = department.DepartmentCode;
            record.DepartmentName = department.DepartmentName;
            record.DepartmentType = department.DepartmentType;
        }
    }

    // ── Parse ──────────────────────────────────────────────────────────────────────

    private static List<DepartmentRecord> ParseDepartmentList(string json)
    {
        var departments = new List<DepartmentRecord>();
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (!TryGetContentsArray(doc.RootElement, out var arr)) return departments;

            foreach (var item in arr.EnumerateArray())
            {
                var raw = ReadStringMap(item);
                string departmentName =
                    FirstNonEmpty(
                        raw.GetValueOrDefault("P_DEPTX", string.Empty),
                        raw.GetValueOrDefault("G_DEPTX", string.Empty),
                        raw.GetValueOrDefault("DEPTX", string.Empty));

                departments.Add(new DepartmentRecord
                {
                    DepartmentNo   = raw.GetValueOrDefault("DEPNO", string.Empty).Trim(),
                    DepartmentCode = raw.GetValueOrDefault("DPCOD", string.Empty).Trim(),
                    DepartmentName = departmentName.Trim(),
                    DepartmentType = raw.GetValueOrDefault("DTYNM", string.Empty).Trim(),
                    Raw            = raw,
                });
            }
        }
        catch { }
        return departments;
    }

    private static List<WorkerRecord> ParseResponse(string json, DepartmentRecord department)
    {
        var records = new List<WorkerRecord>();
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (!TryGetContentsArray(doc.RootElement, out var arr)) return records;

            foreach (var item in arr.EnumerateArray())
            {
                var raw = ReadStringMap(item);

                records.Add(new WorkerRecord
                {
                    EmpNo          = raw.GetValueOrDefault("EMPNO",    string.Empty),
                    Name           = raw.GetValueOrDefault("FNAME",    string.Empty),
                    DepartmentNo   = department.DepartmentNo,
                    DepartmentCode = department.DepartmentCode,
                    DepartmentName = department.DepartmentName,
                    DepartmentType = department.DepartmentType,
                    WorkStatus     = raw.GetValueOrDefault("WKSTA_TX", string.Empty),
                    DayType        = raw.GetValueOrDefault("DTYPE_TX", string.Empty),
                    SchedStart     = raw.GetValueOrDefault("WSTIM",    string.Empty),
                    SchedEnd       = raw.GetValueOrDefault("WETIM",    string.Empty),
                    CheckIn        = raw.GetValueOrDefault("SDATM",    string.Empty),
                    CheckOut       = raw.GetValueOrDefault("EDATM",    string.Empty),
                    Factory        = raw.GetValueOrDefault("FACCO",    string.Empty),
                    Raw            = raw,
                });
            }
        }
        catch { }
        return records;
    }

    private static bool TryGetContentsArray(JsonElement root, out JsonElement arr)
    {
        if (root.ValueKind == JsonValueKind.Array)
        {
            arr = root;
            return true;
        }

        if (root.TryGetProperty("data", out var data))
        {
            if (data.TryGetProperty("contents", out var contents) &&
                contents.ValueKind == JsonValueKind.Array)
            {
                arr = contents;
                return true;
            }

            if (data.ValueKind == JsonValueKind.Array)
            {
                arr = data;
                return true;
            }
        }

        if (root.TryGetProperty("rows", out var rows) && rows.ValueKind == JsonValueKind.Array)
        {
            arr = rows;
            return true;
        }

        if (root.TryGetProperty("contents", out var contents2) &&
            contents2.ValueKind == JsonValueKind.Array)
        {
            arr = contents2;
            return true;
        }

        arr = default;
        return false;
    }

    private static Dictionary<string, string> ReadStringMap(JsonElement item)
    {
        var raw = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in item.EnumerateObject())
            raw[prop.Name] = prop.Value.ValueKind == JsonValueKind.Null
                ? string.Empty : prop.Value.ToString();
        return raw;
    }

    private static string EmployeeKey(WorkerRecord record)
        => !string.IsNullOrWhiteSpace(record.EmpNo)
            ? record.EmpNo.Trim()
            : record.Name.Trim();

    private static IEnumerable<WorkerRecord> DeduplicateRecords(IEnumerable<WorkerRecord> records)
    {
        return records
            .GroupBy(r => (
                Date: r.Date.Date,
                EmpNo: EmployeeKey(r)),
                StringTupleComparer.Instance)
            .Select(g => g
                .OrderByDescending(r => DepartmentSpecificity(r.DepartmentCode))
                .ThenBy(r => r.DepartmentCode, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.DepartmentNo, StringComparer.OrdinalIgnoreCase)
                .First());
    }

    private static IEnumerable<WorkerRecord> SortRecords(IEnumerable<WorkerRecord> records)
    {
        return records
            .OrderBy(r => r.DepartmentCode, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.DepartmentNo, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.Name, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.EmpNo, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.Date);
    }

    private static int DepartmentSpecificity(string departmentCode)
        => string.IsNullOrWhiteSpace(departmentCode)
            ? 0
            : departmentCode.Count(c => c == '-') + 1;

    private static string FirstNonEmpty(params string[] values)
        => values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? string.Empty;

    private sealed class StringTupleComparer : IEqualityComparer<(DateTime Date, string EmpNo)>
    {
        public static readonly StringTupleComparer Instance = new();

        public bool Equals((DateTime Date, string EmpNo) x, (DateTime Date, string EmpNo) y)
            => x.Date == y.Date && string.Equals(x.EmpNo, y.EmpNo, StringComparison.OrdinalIgnoreCase);

        public int GetHashCode((DateTime Date, string EmpNo) obj)
            => HashCode.Combine(obj.Date, StringComparer.OrdinalIgnoreCase.GetHashCode(obj.EmpNo));
    }
}
