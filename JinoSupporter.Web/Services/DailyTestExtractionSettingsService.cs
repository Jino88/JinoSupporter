using System.Text.Json;
using System.Text.Json.Serialization;

namespace JinoSupporter.Web.Services;

public static class DailyTestExtractionModes
{
    public const string ExcelComMergedCells = "excel-com-merged-cells";
    public const string SavedExtractText = "saved-extract-text";

    public static string Normalize(string? value)
        => string.Equals(value, SavedExtractText, StringComparison.OrdinalIgnoreCase)
            ? SavedExtractText
            : ExcelComMergedCells;

    public static bool IsSavedExtractText(string? value)
        => string.Equals(Normalize(value), SavedExtractText, StringComparison.Ordinal);
}

public static class DailyTestSavedExtractRefreshModes
{
    public const string RefreshLimited = "refresh-limited";
    public const string RefreshAll = "refresh-all";
    public const string Never = "never";

    public static string Normalize(string? value)
    {
        if (string.Equals(value, RefreshAll, StringComparison.OrdinalIgnoreCase)) return RefreshAll;
        if (string.Equals(value, Never, StringComparison.OrdinalIgnoreCase)) return Never;
        return RefreshLimited;
    }

    public static bool IsNever(string? value)
        => string.Equals(Normalize(value), Never, StringComparison.Ordinal);

    public static bool IsRefreshAll(string? value)
        => string.Equals(Normalize(value), RefreshAll, StringComparison.Ordinal);
}

public sealed class DailyTestExtractionSettingsSnapshot
{
    public string ExtractionMode { get; set; } = DailyTestExtractionModes.ExcelComMergedCells;
    public int MaxProgramCellsPerSheet { get; set; } = DailyTestExtractionSettingsService.DefaultMaxProgramCellsPerSheet;
    public string SavedExtractRefreshMode { get; set; } = DailyTestSavedExtractRefreshModes.RefreshLimited;
}

public sealed class DailyTestExtractionSettingsService
{
    public const int LegacyLimitedCellThreshold = 5000;
    public const int DefaultMaxProgramCellsPerSheet = 100000;
    public const int MinMaxProgramCellsPerSheet = 1000;
    public const int MaxMaxProgramCellsPerSheet = 1000000;

    private readonly object _gate = new();
    private readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
    };

    private DailyTestExtractionSettingsSnapshot? _cached;
    private DateTime _lastWriteUtc = DateTime.MinValue;

    public DailyTestExtractionSettingsService()
    {
        SettingsPath = Path.Combine(FindRepoRoot(), "daily-test-extraction-settings.json");
        EnsureFile();
    }

    public string SettingsPath { get; }

    public DailyTestExtractionSettingsSnapshot GetSnapshot()
    {
        lock (_gate)
        {
            DateTime currentWriteUtc = GetLastWriteUtcNoLock();
            if (_cached is null || currentWriteUtc != _lastWriteUtc)
            {
                _cached = LoadNoLock();
                _lastWriteUtc = currentWriteUtc;
            }

            return Clone(_cached);
        }
    }

    public void Save(DailyTestExtractionSettingsSnapshot snapshot)
    {
        lock (_gate)
        {
            DailyTestExtractionSettingsSnapshot normalized = Normalize(snapshot);
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath) ?? ".");
            string json = JsonSerializer.Serialize(normalized, _jsonOptions);
            File.WriteAllText(SettingsPath, json, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            _cached = Clone(normalized);
            _lastWriteUtc = GetLastWriteUtcNoLock();
        }
    }

    public static DailyTestExtractionSettingsSnapshot Normalize(DailyTestExtractionSettingsSnapshot? source)
    {
        source ??= new DailyTestExtractionSettingsSnapshot();
        return new DailyTestExtractionSettingsSnapshot
        {
            ExtractionMode = DailyTestExtractionModes.Normalize(source.ExtractionMode),
            MaxProgramCellsPerSheet = Math.Clamp(
                source.MaxProgramCellsPerSheet,
                MinMaxProgramCellsPerSheet,
                MaxMaxProgramCellsPerSheet),
            SavedExtractRefreshMode = DailyTestSavedExtractRefreshModes.Normalize(source.SavedExtractRefreshMode),
        };
    }

    private void EnsureFile()
    {
        lock (_gate)
        {
            if (File.Exists(SettingsPath))
            {
                _cached = LoadNoLock();
                _lastWriteUtc = GetLastWriteUtcNoLock();
                return;
            }

            _cached = new DailyTestExtractionSettingsSnapshot();
            Save(_cached);
        }
    }

    private DailyTestExtractionSettingsSnapshot LoadNoLock()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return new DailyTestExtractionSettingsSnapshot();

            string json = File.ReadAllText(SettingsPath);
            var snapshot = JsonSerializer.Deserialize<DailyTestExtractionSettingsSnapshot>(
                json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            return Normalize(snapshot);
        }
        catch
        {
            return new DailyTestExtractionSettingsSnapshot();
        }
    }

    private static DailyTestExtractionSettingsSnapshot Clone(DailyTestExtractionSettingsSnapshot source)
        => new()
        {
            ExtractionMode = source.ExtractionMode,
            MaxProgramCellsPerSheet = source.MaxProgramCellsPerSheet,
            SavedExtractRefreshMode = source.SavedExtractRefreshMode,
        };

    private static string FindRepoRoot()
    {
        string dir = AppContext.BaseDirectory;
        for (int i = 0; i < 10; i++)
        {
            if (File.Exists(Path.Combine(dir, "JinoSupporter.sln"))) return dir;
            string? parent = Path.GetDirectoryName(dir.TrimEnd('\\', '/'));
            if (string.IsNullOrWhiteSpace(parent) || parent == dir) break;
            dir = parent;
        }

        string current = Directory.GetCurrentDirectory();
        dir = current;
        for (int i = 0; i < 10; i++)
        {
            if (File.Exists(Path.Combine(dir, "JinoSupporter.sln"))) return dir;
            string? parent = Path.GetDirectoryName(dir.TrimEnd('\\', '/'));
            if (string.IsNullOrWhiteSpace(parent) || parent == dir) break;
            dir = parent;
        }

        return current;
    }

    private DateTime GetLastWriteUtcNoLock()
        => File.Exists(SettingsPath) ? File.GetLastWriteTimeUtc(SettingsPath) : DateTime.MinValue;
}
