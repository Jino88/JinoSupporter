using System.Text;
using System.Text.Json;

namespace JinoSupporter.Web.Services;

public static partial class BmesGateway
{
    /// <summary>
    /// MES072900/SearchList: department list for worker attendance.
    /// Input: <paramref name="date"/> (STDAT), plant GN, empty keyword.
    /// Output: every department row (number, code, name, type, raw columns); the caller filters.
    /// Throws on HTTP failure; an unparseable body yields an empty list.
    /// </summary>
    public static async Task<List<WorkerStatusService.DepartmentRecord>> FetchWorkerDepartments(
        BmesWebSession session,
        DateTime date,
        CancellationToken cancellationToken = default)
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

        var response = await session.Client.PostAsync(
            BmesWebSession.BaseUrl + "/MES072900/SearchList", requestContent, cancellationToken);
        response.EnsureSuccessStatusCode();

        string json = await response.Content.ReadAsStringAsync(cancellationToken);
        return ParseDepartmentList(json);
    }

    /// <summary>
    /// MES072900/SearchPopupDetail: attendance (check-in / check-out) rows of one department.
    /// Input: <paramref name="date"/> (STDAT) and the department number (DEPNO), plant GN, ZTYPE=A.
    /// Output: one record per worker with raw columns; Date is left for the caller to set.
    /// Throws on HTTP failure; an unparseable body yields an empty list.
    /// </summary>
    public static async Task<List<WorkerStatusService.WorkerRecord>> FetchWorkerAttendance(
        BmesWebSession session,
        DateTime date,
        WorkerStatusService.DepartmentRecord department,
        CancellationToken cancellationToken = default)
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

        var response = await session.Client.PostAsync(
            BmesWebSession.BaseUrl + "/MES072900/SearchPopupDetail", requestContent, cancellationToken);
        response.EnsureSuccessStatusCode();

        string json = await response.Content.ReadAsStringAsync(cancellationToken);
        return ParseWorkerAttendance(json, department);
    }

    private static List<WorkerStatusService.DepartmentRecord> ParseDepartmentList(string json)
    {
        var departments = new List<WorkerStatusService.DepartmentRecord>();
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

                departments.Add(new WorkerStatusService.DepartmentRecord
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

    private static List<WorkerStatusService.WorkerRecord> ParseWorkerAttendance(string json, WorkerStatusService.DepartmentRecord department)
    {
        var records = new List<WorkerStatusService.WorkerRecord>();
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (!TryGetContentsArray(doc.RootElement, out var arr)) return records;

            foreach (var item in arr.EnumerateArray())
            {
                var raw = ReadStringMap(item);

                records.Add(new WorkerStatusService.WorkerRecord
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

    private static string FirstNonEmpty(params string[] values)
        => values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? string.Empty;
}
