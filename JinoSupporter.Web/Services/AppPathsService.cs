using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Central, persistent registry of every file/folder path the JinoSupporter.Web app
/// reads or writes. Backed by <c>AppStoragePaths.RootDirectory\webapp-paths.json</c>.
/// Edited via the Admin Paths page; consumed by services + Program.cs at startup.
/// </summary>
public sealed class AppPathsConfig
{
    public string DbRootDirectory           { get; set; } = string.Empty;
    public string MainDbPath                { get; set; } = string.Empty;
    public string NgRateDbSaveDirectory     { get; set; } = string.Empty;
    public string NgRateRoutingFilePath     { get; set; } = string.Empty;
    public string NgRateReasonFilePath      { get; set; } = string.Empty;
    public string NgRateSettingsDbDirectory { get; set; } = string.Empty;
    public string ModelBmesJsonFolder       { get; set; } = string.Empty;

    // Test Excel Converter — admin-only DRM-clean tool. Source folder is scanned
    // recursively for *.xlsx / *.xlsm and the cleaned output goes under
    // <DestDir>/drm_clean/<name>_clean.xlsx.
    public string TestExcelConverterSourceDir { get; set; } = string.Empty;
    public string TestExcelConverterDestDir   { get; set; } = string.Empty;

    public string AdminDbQueryServer { get; set; } = string.Empty;
    public int AdminDbQueryPort { get; set; } = 1433;
    public string AdminDbQueryDatabase { get; set; } = string.Empty;
    public string AdminDbQueryUserId { get; set; } = string.Empty;
    public string AdminDbQueryPassword { get; set; } = string.Empty;
    public int AdminDbQueryTimeoutSeconds { get; set; } = 15;
    public bool AdminDbQueryEncrypt { get; set; } = true;
    public bool AdminDbQueryTrustServerCertificate { get; set; } = true;

    public string QrBakoDataServer { get; set; } = string.Empty;
    public int QrBakoDataPort { get; set; } = 1430;
    public string QrBakoDataDatabase { get; set; } = string.Empty;
    public string QrBakoDataUserId { get; set; } = string.Empty;
    public string QrBakoDataPassword { get; set; } = string.Empty;
    public int QrBakoDataTimeoutSeconds { get; set; } = 60;
    public int QrBakoDataDefaultMaxRows { get; set; } = 1000;
    public bool QrBakoDataEncrypt { get; set; } = false;
    public bool QrBakoDataTrustServerCertificate { get; set; } = true;
}

public sealed class AppPathsService
{
    public const string NgRateFolderName = "01. NG RATE";
    public const string FCostFolderName = "02. FCOST";
    public const string WorkerStatusFolderName = "03. WK STATUS";
    private const string LegacyNgRateFolderName = "01. NGRATE";
    private const string LegacyWorkerStatusFolderName = "3. Wk Status";

    private static readonly string[] RequiredNgRatePathProperties =
    [
        nameof(AppPathsConfig.NgRateDbSaveDirectory),
        nameof(AppPathsConfig.NgRateSettingsDbDirectory),
    ];

    private static readonly Dictionary<string, string> RequiredNgRatePathLabels = new()
    {
        [nameof(AppPathsConfig.NgRateDbSaveDirectory)] = "NG Rate DB Save Directory",
        [nameof(AppPathsConfig.NgRateSettingsDbDirectory)] = "NG Rate Settings DB Directory",
    };

    private static readonly string ConfigDir = AppStoragePaths.RootDirectory;

    private static readonly string ConfigFile = Path.Combine(ConfigDir, "webapp-paths.json");
    private static readonly string BootstrapNgRateSettingsDbDirectory = Path.Combine(
        ConfigDir,
        "standalone-bootstrap",
        "ModelBmes");

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
    };

    private AppPathsConfig _current;

    private sealed record LegacyNgRatePaths(
        string DbSaveDirectory,
        string FCostDbSaveDirectory,
        string RoutingFilePath,
        string ReasonFilePath);

    public AppPathsService()
    {
        _current = LoadOrCreate();
        EnsureConfiguredDirectories(_current);
        MigrateKnownLegacyLayouts(_current);
    }

    /// <summary>The currently effective path configuration. Always non-null;
    /// first-run storage paths stay blank until the user sets them.</summary>
    public AppPathsConfig Current => _current;

    public string ConfigFilePath => ConfigFile;

    public bool IsNgRateStorageConfigured => HasNgRateStorageConfigured(Current);

    public IReadOnlyList<string> GetMissingRequiredNgRatePathLabels() =>
        GetMissingRequiredNgRatePathLabels(Current);

    public static string DefaultDbRootDirectory => Path.Combine(ConfigDir, "db");

    public static string GetNgRateDirectory(string root) =>
        Path.Combine((root ?? string.Empty).Trim(), NgRateFolderName);

    public static string GetFCostDirectory(string root) =>
        Path.Combine((root ?? string.Empty).Trim(), FCostFolderName);

    public static string GetWorkerStatusDirectory(string root) =>
        Path.Combine((root ?? string.Empty).Trim(), WorkerStatusFolderName);

    /// <summary>Baseline used for initial population and for any field the user leaves blank.
    /// Storage paths intentionally start empty so a new standalone install does not inherit
    /// developer-machine directories.</summary>
    public static AppPathsConfig Defaults()
    {
        string mainDb = WpfSettingsReader.TryGetDatabasePath() ?? string.Empty;

        return new AppPathsConfig
        {
            DbRootDirectory           = string.Empty,
            MainDbPath                = mainDb,
            NgRateDbSaveDirectory     = string.Empty,
            NgRateRoutingFilePath     = string.Empty,
            NgRateReasonFilePath      = string.Empty,
            NgRateSettingsDbDirectory = string.Empty,
            ModelBmesJsonFolder       = string.Empty,

            // Default source/dest = same folder; the cleaned files go to a
            // `drm_clean` sub-folder so they don't shadow the originals.
            TestExcelConverterSourceDir = string.Empty,
            TestExcelConverterDestDir   = string.Empty,

            AdminDbQueryServer = string.Empty,
            AdminDbQueryPort = 1433,
            AdminDbQueryDatabase = "master",
            AdminDbQueryUserId = string.Empty,
            AdminDbQueryPassword = string.Empty,
            AdminDbQueryTimeoutSeconds = 15,
            AdminDbQueryEncrypt = true,
            AdminDbQueryTrustServerCertificate = true,

            QrBakoDataServer = "10.6.0.22",
            QrBakoDataPort = 1430,
            QrBakoDataDatabase = "TCDB",
            QrBakoDataUserId = "TCDB",
            QrBakoDataPassword = string.Empty,
            QrBakoDataTimeoutSeconds = 60,
            QrBakoDataDefaultMaxRows = 1000,
            QrBakoDataEncrypt = false,
            QrBakoDataTrustServerCertificate = true,
        };
    }

    private static AppPathsConfig LoadOrCreate()
    {
        AppPathsConfig def = Defaults();
        if (!File.Exists(ConfigFile)) return def;
        try
        {
            string json = File.ReadAllText(ConfigFile);
            var loaded = JsonSerializer.Deserialize<AppPathsConfig>(json, JsonOpts);
            if (loaded is null) return def;
            return Merge(def, loaded);
        }
        catch
        {
            // Corrupt or unreadable file → fall back to defaults silently.
            return def;
        }
    }

    public void Save(AppPathsConfig cfg)
    {
        AppPathsConfig previous = _current;
        bool previousUsedUnifiedRoot = !string.IsNullOrWhiteSpace(previous.DbRootDirectory);
        string oldSettingsDbDirectory = string.IsNullOrWhiteSpace(previous.NgRateSettingsDbDirectory)
            ? BootstrapNgRateSettingsDbDirectory
            : previous.NgRateSettingsDbDirectory;
        LegacyNgRatePaths legacyPaths = ReadLegacyNgRatePaths(oldSettingsDbDirectory);

        var merged = Merge(Defaults(), cfg ?? new AppPathsConfig());
        Directory.CreateDirectory(ConfigDir);
        EnsureConfiguredDirectories(merged);
        MigrateDbFile(previous.MainDbPath, merged.MainDbPath);
        MigrateSettingsDbFile(oldSettingsDbDirectory, merged.NgRateSettingsDbDirectory, "ngrate_settings.db");
        MigrateSettingsDbFile(oldSettingsDbDirectory, merged.NgRateSettingsDbDirectory, "bmes_routing_raw.db");

        string oldNgRateDbDirectory = previousUsedUnifiedRoot
            ? previous.NgRateDbSaveDirectory
            : FirstNonBlank(legacyPaths.DbSaveDirectory, previous.NgRateDbSaveDirectory);
        string oldFCostDbDirectory = previousUsedUnifiedRoot
            ? previous.NgRateDbSaveDirectory
            : FirstNonBlank(legacyPaths.FCostDbSaveDirectory, legacyPaths.DbSaveDirectory, previous.NgRateDbSaveDirectory);
        string oldRoutingFilePath = previousUsedUnifiedRoot
            ? previous.NgRateRoutingFilePath
            : FirstNonBlank(legacyPaths.RoutingFilePath, previous.NgRateRoutingFilePath);
        string oldReasonFilePath = previousUsedUnifiedRoot
            ? previous.NgRateReasonFilePath
            : FirstNonBlank(legacyPaths.ReasonFilePath, previous.NgRateReasonFilePath);

        string newFCostDbDirectory = GetEffectiveFCostDirectory(merged);
        string newWorkerStatusDirectory = GetEffectiveWorkerStatusDirectory(merged);

        MigrateNgRateStorage(oldNgRateDbDirectory, merged.NgRateDbSaveDirectory);
        MigrateFCostStorage(oldFCostDbDirectory, newFCostDbDirectory);
        string oldWorkerStatusDirectory = string.IsNullOrWhiteSpace(oldNgRateDbDirectory)
            ? string.Empty
            : Path.Combine(oldNgRateDbDirectory, LegacyWorkerStatusFolderName);
        MigrateWorkerStatusStorage(oldWorkerStatusDirectory, newWorkerStatusDirectory);
        MigrateFileIfMissing(oldRoutingFilePath, merged.NgRateRoutingFilePath);
        MigrateFileIfMissing(oldReasonFilePath, merged.NgRateReasonFilePath);
        MigrateKnownLegacyLayouts(merged);

        string json = JsonSerializer.Serialize(merged, JsonOpts);
        File.WriteAllText(ConfigFile, json);
        _current = merged;
    }

    /// <summary>Returns a copy where each blank field in <paramref name="user"/> is replaced
    /// with the corresponding value from <paramref name="defaults"/>.</summary>
    private static AppPathsConfig Merge(AppPathsConfig defaults, AppPathsConfig user)
    {
        bool hasSavedAdminDbQuery = !string.IsNullOrWhiteSpace(user.AdminDbQueryServer)
            || !string.IsNullOrWhiteSpace(user.AdminDbQueryUserId)
            || !string.IsNullOrWhiteSpace(user.AdminDbQueryPassword);
        bool hasSavedQrBakoData = !string.IsNullOrWhiteSpace(user.QrBakoDataServer)
            || !string.IsNullOrWhiteSpace(user.QrBakoDataUserId)
            || !string.IsNullOrWhiteSpace(user.QrBakoDataPassword);

        var merged = new AppPathsConfig
        {
            DbRootDirectory           = Pick(user.DbRootDirectory,           defaults.DbRootDirectory),
            MainDbPath                = Pick(user.MainDbPath,                defaults.MainDbPath),
            NgRateDbSaveDirectory     = Pick(user.NgRateDbSaveDirectory,     defaults.NgRateDbSaveDirectory),
            NgRateRoutingFilePath     = Pick(user.NgRateRoutingFilePath,     defaults.NgRateRoutingFilePath),
            NgRateReasonFilePath      = Pick(user.NgRateReasonFilePath,      defaults.NgRateReasonFilePath),
            NgRateSettingsDbDirectory = Pick(user.NgRateSettingsDbDirectory, defaults.NgRateSettingsDbDirectory),
            ModelBmesJsonFolder       = Pick(user.ModelBmesJsonFolder,       defaults.ModelBmesJsonFolder),
            TestExcelConverterSourceDir = Pick(user.TestExcelConverterSourceDir, defaults.TestExcelConverterSourceDir),
            TestExcelConverterDestDir   = Pick(user.TestExcelConverterDestDir,   defaults.TestExcelConverterDestDir),
            AdminDbQueryServer = Pick(user.AdminDbQueryServer, defaults.AdminDbQueryServer),
            AdminDbQueryPort = user.AdminDbQueryPort > 0 ? user.AdminDbQueryPort : defaults.AdminDbQueryPort,
            AdminDbQueryDatabase = Pick(user.AdminDbQueryDatabase, defaults.AdminDbQueryDatabase),
            AdminDbQueryUserId = Pick(user.AdminDbQueryUserId, defaults.AdminDbQueryUserId),
            AdminDbQueryPassword = user.AdminDbQueryPassword ?? defaults.AdminDbQueryPassword,
            AdminDbQueryTimeoutSeconds = user.AdminDbQueryTimeoutSeconds > 0
                ? user.AdminDbQueryTimeoutSeconds
                : defaults.AdminDbQueryTimeoutSeconds,
            AdminDbQueryEncrypt = hasSavedAdminDbQuery ? user.AdminDbQueryEncrypt : defaults.AdminDbQueryEncrypt,
            AdminDbQueryTrustServerCertificate = hasSavedAdminDbQuery
                ? user.AdminDbQueryTrustServerCertificate
                : defaults.AdminDbQueryTrustServerCertificate,
            QrBakoDataServer = Pick(user.QrBakoDataServer, defaults.QrBakoDataServer),
            QrBakoDataPort = user.QrBakoDataPort > 0 ? user.QrBakoDataPort : defaults.QrBakoDataPort,
            QrBakoDataDatabase = Pick(user.QrBakoDataDatabase, defaults.QrBakoDataDatabase),
            QrBakoDataUserId = Pick(user.QrBakoDataUserId, defaults.QrBakoDataUserId),
            QrBakoDataPassword = user.QrBakoDataPassword ?? defaults.QrBakoDataPassword,
            QrBakoDataTimeoutSeconds = user.QrBakoDataTimeoutSeconds > 0
                ? user.QrBakoDataTimeoutSeconds
                : defaults.QrBakoDataTimeoutSeconds,
            QrBakoDataDefaultMaxRows = user.QrBakoDataDefaultMaxRows > 0
                ? user.QrBakoDataDefaultMaxRows
                : defaults.QrBakoDataDefaultMaxRows,
            QrBakoDataEncrypt = hasSavedQrBakoData ? user.QrBakoDataEncrypt : defaults.QrBakoDataEncrypt,
            QrBakoDataTrustServerCertificate = hasSavedQrBakoData
                ? user.QrBakoDataTrustServerCertificate
                : defaults.QrBakoDataTrustServerCertificate,
        }.WithUnifiedDbDirectory();

        return ResolveDirectoryFileConflicts(merged);
    }

    private static string Pick(string? user, string fallback) =>
        string.IsNullOrWhiteSpace(user) ? fallback : user.Trim();

    private static string FirstNonBlank(params string?[] values)
    {
        foreach (string? value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
                return value.Trim();
        }
        return string.Empty;
    }

    public static AppPathsConfig ApplyUnifiedDbDirectory(AppPathsConfig cfg) =>
        ResolveDirectoryFileConflicts((cfg ?? new AppPathsConfig()).WithUnifiedDbDirectory());

    public static bool DirectoryPathConflictsWithFile(string? directory)
    {
        if (string.IsNullOrWhiteSpace(directory)) return false;
        try
        {
            return File.Exists(directory.Trim());
        }
        catch
        {
            return false;
        }
    }

    private static AppPathsConfig ResolveDirectoryFileConflicts(AppPathsConfig cfg)
    {
        if (DirectoryPathConflictsWithFile(cfg.DbRootDirectory))
        {
            cfg.DbRootDirectory = DefaultDbRootDirectory;
            cfg = cfg.WithUnifiedDbDirectory();
        }

        if (DirectoryPathConflictsWithFile(cfg.NgRateDbSaveDirectory))
            cfg.NgRateDbSaveDirectory = string.IsNullOrWhiteSpace(cfg.DbRootDirectory)
                ? DefaultDbRootDirectory
                : GetNgRateDirectory(cfg.DbRootDirectory);

        if (DirectoryPathConflictsWithFile(cfg.NgRateSettingsDbDirectory))
            cfg.NgRateSettingsDbDirectory = string.IsNullOrWhiteSpace(cfg.DbRootDirectory)
                ? BootstrapNgRateSettingsDbDirectory
                : GetNgRateDirectory(cfg.DbRootDirectory);

        if (DirectoryPathConflictsWithFile(cfg.ModelBmesJsonFolder))
            cfg.ModelBmesJsonFolder = string.IsNullOrWhiteSpace(cfg.DbRootDirectory)
                ? Path.Combine(DefaultDbRootDirectory, "ModelBmes")
                : Path.Combine(GetNgRateDirectory(cfg.DbRootDirectory), "ModelBmes");

        string? mainDir = Path.GetDirectoryName(cfg.MainDbPath);
        if (DirectoryPathConflictsWithFile(mainDir))
            cfg.MainDbPath = Path.Combine(DefaultDbRootDirectory, "process-review.db");

        return cfg;
    }

    private static string? DbFilePath(string? directory, string fileName)
    {
        if (string.IsNullOrWhiteSpace(directory)) return null;
        return Path.Combine(directory.Trim(), fileName);
    }

    private static void EnsureConfiguredDirectories(AppPathsConfig cfg)
    {
        CreateIfSet(cfg.DbRootDirectory);
        CreateIfSet(Path.GetDirectoryName(cfg.MainDbPath));
        CreateIfSet(cfg.NgRateDbSaveDirectory);
        CreateIfSet(cfg.NgRateSettingsDbDirectory);
        CreateIfSet(GetEffectiveFCostDirectory(cfg));
        CreateIfSet(GetEffectiveWorkerStatusDirectory(cfg));
    }

    private static void CreateIfSet(string? directory)
    {
        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory.Trim());
    }

    private static void MigrateSettingsDbFile(string? oldDirectory, string? newDirectory, string fileName)
    {
        string? oldPath = DbFilePath(oldDirectory, fileName);
        string? newPath = DbFilePath(newDirectory, fileName);
        MigrateDbFile(oldPath, newPath);
    }

    private static void MigrateNgRateStorage(string? oldDirectory, string? newDirectory)
    {
        if (!CanMigrateDirectory(oldDirectory, newDirectory)) return;

        CopyMatchingFiles(oldDirectory!, newDirectory!, "*.db", IsNgRateDataDb);

        string oldDaily = Path.Combine(oldDirectory!.Trim(), "daily");
        string newDaily = Path.Combine(newDirectory!.Trim(), "daily");
        if (Directory.Exists(oldDaily))
            CopyMatchingFiles(oldDaily, newDaily, "*.db", IsNgRateDataDb);

        string oldMonthly = Path.Combine(oldDirectory!.Trim(), "monthly");
        string newMonthly = Path.Combine(newDirectory!.Trim(), "monthly");
        if (Directory.Exists(oldMonthly))
            CopyMatchingFiles(oldMonthly, newMonthly, "*.db", IsNgRateDataDb);
    }

    private static void MigrateFCostStorage(string? oldDirectory, string? newDirectory)
    {
        if (!CanMigrateDirectory(oldDirectory, newDirectory)) return;

        CopyMatchingFiles(oldDirectory!, newDirectory!, "fcost_????????_??????.db");
        MigrateFileIfMissing(Path.Combine(oldDirectory!.Trim(), "fcost_raw.db"), Path.Combine(newDirectory!.Trim(), "fcost_raw.db"));
        MigrateFileIfMissing(Path.Combine(oldDirectory.Trim(), "fcost_raw.db-wal"), Path.Combine(newDirectory.Trim(), "fcost_raw.db-wal"));
        MigrateFileIfMissing(Path.Combine(oldDirectory.Trim(), "fcost_raw.db-shm"), Path.Combine(newDirectory.Trim(), "fcost_raw.db-shm"));
    }

    private static void MigrateWorkerStatusStorage(string? oldDirectory, string? newDirectory)
    {
        if (!CanMigrateDirectory(oldDirectory, newDirectory)) return;
        CopyMatchingFiles(oldDirectory!, newDirectory!, "*.json");
    }

    private static void MigrateKnownLegacyLayouts(AppPathsConfig cfg)
    {
        if (string.IsNullOrWhiteSpace(cfg.DbRootDirectory)) return;

        string root = cfg.DbRootDirectory.Trim();
        string ngRateDir = cfg.NgRateDbSaveDirectory;
        string fCostDir = GetEffectiveFCostDirectory(cfg);
        string workerStatusDir = GetEffectiveWorkerStatusDirectory(cfg);

        // Previous common-root layout.
        MigrateSettingsDbFile(root, cfg.NgRateSettingsDbDirectory, "ngrate_settings.db");
        MigrateSettingsDbFile(root, cfg.NgRateSettingsDbDirectory, "bmes_routing_raw.db");
        MigrateNgRateStorage(root, ngRateDir);
        MigrateFCostStorage(root, fCostDir);
        MigrateFileIfMissing(Path.Combine(root, "Routing.txt"), cfg.NgRateRoutingFilePath);
        MigrateFileIfMissing(Path.Combine(root, "routing.txt"), cfg.NgRateRoutingFilePath);
        MigrateFileIfMissing(Path.Combine(root, "reason.txt"), cfg.NgRateReasonFilePath);
        MigrateWorkerStatusStorage(Path.Combine(root, LegacyWorkerStatusFolderName), workerStatusDir);

        // Older split layout.
        string legacyNgRateDir = Path.Combine(root, LegacyNgRateFolderName);
        string legacyModelBmesDir = Path.Combine(legacyNgRateDir, "ModelBmes");
        MigrateSettingsDbFile(legacyNgRateDir, cfg.NgRateSettingsDbDirectory, "ngrate_settings.db");
        MigrateSettingsDbFile(legacyModelBmesDir, cfg.NgRateSettingsDbDirectory, "ngrate_settings.db");
        MigrateSettingsDbFile(legacyModelBmesDir, cfg.NgRateSettingsDbDirectory, "bmes_routing_raw.db");
        MigrateNgRateStorage(legacyNgRateDir, ngRateDir);
        MigrateFCostStorage(legacyNgRateDir, fCostDir);
        MigrateFileIfMissing(Path.Combine(legacyNgRateDir, "Routing.txt"), cfg.NgRateRoutingFilePath);
        MigrateFileIfMissing(Path.Combine(legacyNgRateDir, "routing.txt"), cfg.NgRateRoutingFilePath);
        MigrateFileIfMissing(Path.Combine(legacyNgRateDir, "reason.txt"), cfg.NgRateReasonFilePath);
        MigrateWorkerStatusStorage(Path.Combine(legacyNgRateDir, LegacyWorkerStatusFolderName), workerStatusDir);

        string legacyFCostDir = Path.Combine(root, FCostFolderName);
        MigrateFCostStorage(legacyFCostDir, fCostDir);
    }

    private static bool CanMigrateDirectory(string? oldDirectory, string? newDirectory)
    {
        if (string.IsNullOrWhiteSpace(oldDirectory) || string.IsNullOrWhiteSpace(newDirectory))
            return false;
        oldDirectory = oldDirectory.Trim();
        newDirectory = newDirectory.Trim();
        return !PathEquals(oldDirectory, newDirectory) && Directory.Exists(oldDirectory);
    }

    private static void CopyMatchingFiles(
        string oldDirectory,
        string newDirectory,
        string searchPattern,
        Func<string, bool>? include = null)
    {
        foreach (string oldPath in Directory.GetFiles(oldDirectory.Trim(), searchPattern, SearchOption.TopDirectoryOnly))
        {
            if (include is not null && !include(oldPath)) continue;
            string newPath = Path.Combine(newDirectory.Trim(), Path.GetFileName(oldPath));
            MigrateFileIfMissing(oldPath, newPath);
        }
    }

    private static bool IsNgRateDataDb(string path)
    {
        string name = Path.GetFileName(path);
        if (string.Equals(name, "ngrate_settings.db", StringComparison.OrdinalIgnoreCase)) return false;
        if (string.Equals(name, "bmes_routing_raw.db", StringComparison.OrdinalIgnoreCase)) return false;
        if (string.Equals(name, "process-review.db", StringComparison.OrdinalIgnoreCase)) return false;
        if (string.Equals(name, "fcost_raw.db", StringComparison.OrdinalIgnoreCase)) return false;
        if (name.StartsWith("fcost_", StringComparison.OrdinalIgnoreCase)) return false;
        return name.StartsWith("temp_", StringComparison.OrdinalIgnoreCase) ||
               IsMonthlyDbFileName(name) ||
               IsDailyDbFileName(name);
    }

    private static bool IsMonthlyDbFileName(string name)
    {
        if (name.Length != "yyyyMM.db".Length ||
            !name.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
            return false;

        for (int i = 0; i < 6; i++)
        {
            if (!char.IsDigit(name[i])) return false;
        }
        return true;
    }

    private static bool IsDailyDbFileName(string name)
    {
        if (name.Length != "yyyyMMdd.db".Length ||
            !name.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
            return false;

        for (int i = 0; i < 8; i++)
        {
            if (!char.IsDigit(name[i])) return false;
        }
        return true;
    }

    private static void MigrateFileIfMissing(string? oldPath, string? newPath)
    {
        MigrateDbFile(oldPath, newPath);
    }

    private static void MigrateDbFile(string? oldPath, string? newPath)
    {
        if (string.IsNullOrWhiteSpace(oldPath) || string.IsNullOrWhiteSpace(newPath)) return;
        oldPath = oldPath.Trim();
        newPath = newPath.Trim();
        if (PathEquals(oldPath, newPath) || !File.Exists(oldPath) || File.Exists(newPath)) return;
        string? dir = Path.GetDirectoryName(newPath);
        if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
        File.Copy(oldPath, newPath);
    }

    private static bool PathEquals(string? a, string? b) =>
        string.Equals((a ?? string.Empty).Trim(), (b ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);

    private static string GetEffectiveFCostDirectory(AppPathsConfig cfg) =>
        string.IsNullOrWhiteSpace(cfg.DbRootDirectory)
            ? cfg.NgRateDbSaveDirectory
            : GetFCostDirectory(cfg.DbRootDirectory);

    private static string GetEffectiveWorkerStatusDirectory(AppPathsConfig cfg)
    {
        if (!string.IsNullOrWhiteSpace(cfg.DbRootDirectory))
            return GetWorkerStatusDirectory(cfg.DbRootDirectory);

        return string.IsNullOrWhiteSpace(cfg.NgRateDbSaveDirectory)
            ? string.Empty
            : Path.Combine(cfg.NgRateDbSaveDirectory.Trim(), LegacyWorkerStatusFolderName);
    }

    private static LegacyNgRatePaths ReadLegacyNgRatePaths(string settingsDbDirectory)
    {
        string dbPath = Path.Combine(settingsDbDirectory, "ngrate_settings.db");
        if (!File.Exists(dbPath))
            return new LegacyNgRatePaths(string.Empty, string.Empty, string.Empty, string.Empty);

        try
        {
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            var cs = new SqliteConnectionStringBuilder
            {
                DataSource = dbPath,
                Mode = SqliteOpenMode.ReadOnly,
            }.ToString();
            using var conn = new SqliteConnection(cs);
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT Key, Value
                FROM NgRateSettings
                WHERE Key IN (@ngDir, @fcostDir, @routing, @reason);
                """;
            cmd.Parameters.AddWithValue("@ngDir", NgRateSettingsService.KeyDbSaveDirectory);
            cmd.Parameters.AddWithValue("@fcostDir", NgRateSettingsService.KeyFCostDbSaveDirectory);
            cmd.Parameters.AddWithValue("@routing", NgRateSettingsService.KeyRoutingFilePath);
            cmd.Parameters.AddWithValue("@reason", NgRateSettingsService.KeyReasonFilePath);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                string key = r.IsDBNull(0) ? string.Empty : r.GetString(0);
                string value = r.IsDBNull(1) ? string.Empty : r.GetString(1);
                if (key.Length > 0) values[key] = value;
            }

            values.TryGetValue(NgRateSettingsService.KeyDbSaveDirectory, out string? dbDir);
            values.TryGetValue(NgRateSettingsService.KeyFCostDbSaveDirectory, out string? fcostDir);
            values.TryGetValue(NgRateSettingsService.KeyRoutingFilePath, out string? routing);
            values.TryGetValue(NgRateSettingsService.KeyReasonFilePath, out string? reason);
            return new LegacyNgRatePaths(dbDir ?? string.Empty, fcostDir ?? string.Empty, routing ?? string.Empty, reason ?? string.Empty);
        }
        catch
        {
            return new LegacyNgRatePaths(string.Empty, string.Empty, string.Empty, string.Empty);
        }
    }

    public static bool HasNgRateStorageConfigured(AppPathsConfig cfg) =>
        GetMissingRequiredNgRatePathLabels(cfg).Count == 0;

    public static IReadOnlyList<string> GetMissingRequiredNgRatePathLabels(AppPathsConfig cfg)
    {
        cfg = ApplyUnifiedDbDirectory(cfg);
        List<string> missing = [];
        foreach (string propertyName in RequiredNgRatePathProperties)
        {
            string? value = propertyName switch
            {
                nameof(AppPathsConfig.NgRateDbSaveDirectory) => cfg.NgRateDbSaveDirectory,
                nameof(AppPathsConfig.NgRateSettingsDbDirectory) => cfg.NgRateSettingsDbDirectory,
                _ => string.Empty,
            };
            if (string.IsNullOrWhiteSpace(value))
                missing.Add(RequiredNgRatePathLabels[propertyName]);
        }

        return missing;
    }
}

internal static class AppPathsConfigExtensions
{
    public static AppPathsConfig WithUnifiedDbDirectory(this AppPathsConfig cfg)
    {
        string root = (cfg.DbRootDirectory ?? string.Empty).Trim();
        if (root.Length == 0) return cfg;
        string ngRateDir = AppPathsService.GetNgRateDirectory(root);

        cfg.DbRootDirectory           = root;
        cfg.MainDbPath                = Path.Combine(root, "process-review.db");
        cfg.NgRateDbSaveDirectory     = ngRateDir;
        cfg.NgRateRoutingFilePath     = Path.Combine(ngRateDir, "Routing.txt");
        cfg.NgRateReasonFilePath      = Path.Combine(ngRateDir, "reason.txt");
        cfg.NgRateSettingsDbDirectory = ngRateDir;
        cfg.ModelBmesJsonFolder       = Path.Combine(ngRateDir, "ModelBmes");
        return cfg;
    }
}
