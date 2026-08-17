using System.Globalization;
using System.Text.Json;

namespace JinoSupporter.Web.Services;

public sealed class ProcessMaterialMappingService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
    };

    private readonly object _lock = new();
    private readonly string _path = AppStoragePaths.Combine("process-material-mappings.json");
    private readonly string _processPath = AppStoragePaths.Combine("process-material-processes.json");
    private List<ProcessMaterialMappingRow>? _cache;
    private List<ProcessMaterialProcessRow>? _processCache;

    public List<ProcessMaterialMappingRow> GetAll()
    {
        lock (_lock)
            return LoadLocked().Select(Clone).ToList();
    }

    public ProcessMaterialMappingRow Upsert(ProcessMaterialMappingEdit edit)
    {
        lock (_lock)
        {
            var rows = LoadLocked();
            DateTime now = DateTime.UtcNow;
            string nowText = now.ToString("o", CultureInfo.InvariantCulture);
            ProcessMaterialMappingRow row;

            if (edit.Id > 0)
            {
                row = rows.FirstOrDefault(r => r.Id == edit.Id)
                    ?? throw new InvalidOperationException("Mapping row not found.");
                Apply(row, edit);
                row.UpdatedAt = nowText;
            }
            else
            {
                row = new ProcessMaterialMappingRow
                {
                    Id = rows.Count == 0 ? 1 : rows.Max(r => r.Id) + 1,
                    CreatedAt = nowText,
                    UpdatedAt = nowText,
                };
                Apply(row, edit);
                rows.Add(row);
            }

            SaveLocked(rows);
            return Clone(row);
        }
    }

    public bool Delete(long id)
    {
        lock (_lock)
        {
            var rows = LoadLocked();
            int removed = rows.RemoveAll(r => r.Id == id);
            if (removed == 0)
                return false;
            SaveLocked(rows);
            return true;
        }
    }

    public List<ProcessMaterialProcessRow> GetProcessRows()
    {
        lock (_lock)
            return LoadProcessesLocked().Select(Clone).ToList();
    }

    public ProcessMaterialProcessRow UpsertProcess(ProcessMaterialProcessEdit edit)
    {
        lock (_lock)
        {
            var rows = LoadProcessesLocked();
            DateTime now = DateTime.UtcNow;
            string nowText = now.ToString("o", CultureInfo.InvariantCulture);
            string modelName = edit.ModelName.Trim();
            string processCode = edit.ProcessCode.Trim();
            string processName = edit.ProcessName.Trim();

            var row = rows.FirstOrDefault(r =>
                string.Equals(r.ModelName, modelName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(r.ProcessCode, processCode, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(r.ProcessName, processName, StringComparison.OrdinalIgnoreCase));

            if (row is null)
            {
                row = new ProcessMaterialProcessRow
                {
                    Id = rows.Count == 0 ? 1 : rows.Max(r => r.Id) + 1,
                    CreatedAt = nowText,
                    UpdatedAt = nowText,
                };
                rows.Add(row);
            }

            Apply(row, edit);
            row.UpdatedAt = nowText;
            SaveProcessesLocked(rows);
            return Clone(row);
        }
    }

    public bool DeleteProcess(long id)
    {
        lock (_lock)
        {
            var rows = LoadProcessesLocked();
            int removed = rows.RemoveAll(r => r.Id == id);
            if (removed == 0)
                return false;
            SaveProcessesLocked(rows);
            return true;
        }
    }

    private List<ProcessMaterialMappingRow> LoadLocked()
    {
        if (_cache is not null)
            return _cache;

        try
        {
            if (!File.Exists(_path))
            {
                _cache = [];
                return _cache;
            }

            string json = File.ReadAllText(_path);
            _cache = JsonSerializer.Deserialize<List<ProcessMaterialMappingRow>>(json, JsonOptions) ?? [];
        }
        catch
        {
            _cache = [];
        }

        return _cache;
    }

    private void SaveLocked(List<ProcessMaterialMappingRow> rows)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        string json = JsonSerializer.Serialize(rows.OrderBy(r => r.ModelName).ThenBy(r => r.ProcessName).ThenBy(r => r.RawMaterialName), JsonOptions);
        File.WriteAllText(_path, json);
        _cache = rows;
    }

    private List<ProcessMaterialProcessRow> LoadProcessesLocked()
    {
        if (_processCache is not null)
            return _processCache;

        try
        {
            if (!File.Exists(_processPath))
            {
                _processCache = [];
                return _processCache;
            }

            string json = File.ReadAllText(_processPath);
            _processCache = JsonSerializer.Deserialize<List<ProcessMaterialProcessRow>>(json, JsonOptions) ?? [];
        }
        catch
        {
            _processCache = [];
        }

        return _processCache;
    }

    private void SaveProcessesLocked(List<ProcessMaterialProcessRow> rows)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_processPath)!);
        string json = JsonSerializer.Serialize(
            rows
                .OrderBy(r => r.ModelName)
                .ThenBy(r => r.ProcessNo)
                .ThenBy(r => r.ProcessName),
            JsonOptions);
        File.WriteAllText(_processPath, json);
        _processCache = rows;
    }

    private static void Apply(ProcessMaterialMappingRow row, ProcessMaterialMappingEdit edit)
    {
        row.ModelName = edit.ModelName.Trim();
        row.ProcessCode = edit.ProcessCode.Trim();
        row.ProcessName = edit.ProcessName.Trim();
        row.RawMaterialCode = edit.RawMaterialCode.Trim();
        row.RawMaterialName = edit.RawMaterialName.Trim();
        row.UsageQty = edit.UsageQty;
        row.UsageUnit = edit.UsageUnit.Trim();
        row.Note = edit.Note.Trim();
    }

    private static void Apply(ProcessMaterialProcessRow row, ProcessMaterialProcessEdit edit)
    {
        row.ModelName = edit.ModelName.Trim();
        row.ProcessCode = edit.ProcessCode.Trim();
        row.ProcessName = edit.ProcessName.Trim();
        row.ProcessNo = edit.ProcessNo.Trim();
        row.ReferenceProcessNo = edit.ReferenceProcessNo.Trim();
        row.LaneCode = edit.LaneCode.Trim();
        row.MergeProcessNo = edit.MergeProcessNo.Trim();
    }

    private static ProcessMaterialMappingRow Clone(ProcessMaterialMappingRow row)
        => new()
        {
            Id = row.Id,
            ModelName = row.ModelName,
            ProcessCode = row.ProcessCode,
            ProcessName = row.ProcessName,
            RawMaterialCode = row.RawMaterialCode,
            RawMaterialName = row.RawMaterialName,
            UsageQty = row.UsageQty,
            UsageUnit = row.UsageUnit,
            Note = row.Note,
            CreatedAt = row.CreatedAt,
            UpdatedAt = row.UpdatedAt,
        };

    private static ProcessMaterialProcessRow Clone(ProcessMaterialProcessRow row)
        => new()
        {
            Id = row.Id,
            ModelName = row.ModelName,
            ProcessCode = row.ProcessCode,
            ProcessName = row.ProcessName,
            ProcessNo = row.ProcessNo,
            ReferenceProcessNo = row.ReferenceProcessNo,
            LaneCode = row.LaneCode,
            MergeProcessNo = row.MergeProcessNo,
            CreatedAt = row.CreatedAt,
            UpdatedAt = row.UpdatedAt,
        };
}

public sealed class ProcessMaterialMappingRow
{
    public long Id { get; set; }
    public string ModelName { get; set; } = string.Empty;
    public string ProcessCode { get; set; } = string.Empty;
    public string ProcessName { get; set; } = string.Empty;
    public string RawMaterialCode { get; set; } = string.Empty;
    public string RawMaterialName { get; set; } = string.Empty;
    public decimal UsageQty { get; set; }
    public string UsageUnit { get; set; } = "PC";
    public string Note { get; set; } = string.Empty;
    public string CreatedAt { get; set; } = string.Empty;
    public string UpdatedAt { get; set; } = string.Empty;
}

public sealed class ProcessMaterialMappingEdit
{
    public long Id { get; set; }
    public string ModelName { get; set; } = string.Empty;
    public string ProcessCode { get; set; } = string.Empty;
    public string ProcessName { get; set; } = string.Empty;
    public string RawMaterialCode { get; set; } = string.Empty;
    public string RawMaterialName { get; set; } = string.Empty;
    public decimal UsageQty { get; set; }
    public string UsageUnit { get; set; } = "PC";
    public string Note { get; set; } = string.Empty;
}

public sealed class ProcessMaterialProcessRow
{
    public long Id { get; set; }
    public string ModelName { get; set; } = string.Empty;
    public string ProcessCode { get; set; } = string.Empty;
    public string ProcessName { get; set; } = string.Empty;
    public string ProcessNo { get; set; } = string.Empty;
    public string ReferenceProcessNo { get; set; } = string.Empty;
    public string LaneCode { get; set; } = string.Empty;
    public string MergeProcessNo { get; set; } = string.Empty;
    public string CreatedAt { get; set; } = string.Empty;
    public string UpdatedAt { get; set; } = string.Empty;
}

public sealed class ProcessMaterialProcessEdit
{
    public string ModelName { get; set; } = string.Empty;
    public string ProcessCode { get; set; } = string.Empty;
    public string ProcessName { get; set; } = string.Empty;
    public string ProcessNo { get; set; } = string.Empty;
    public string ReferenceProcessNo { get; set; } = string.Empty;
    public string LaneCode { get; set; } = string.Empty;
    public string MergeProcessNo { get; set; } = string.Empty;
}
