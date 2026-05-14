using System.Text.Json;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Configuration;

namespace JinoSupporter.Web.Services;

public sealed class UserRecord
{
    public long   Id           { get; init; }
    public string Username     { get; init; } = string.Empty;
    public string DisplayName  { get; init; } = string.Empty;
    public string PasswordHash { get; init; } = string.Empty;
    public string Role         { get; init; } = string.Empty;
    public string CreatedAt    { get; init; } = string.Empty;
}

public sealed record ScheduleItem(
    long         Id,
    string       Title,
    string       Description,
    DateOnly     StartDate,
    DateOnly     EndDate,
    TimeOnly?    StartTime,
    TimeOnly?    EndTime,
    string       Color,
    List<string> Tags,
    string       CreatedAt)
{
    public bool   IsAllDay   => StartTime is null;
    public bool   IsMultiDay => EndDate > StartDate;
    public string TimeDisplay => IsAllDay
        ? "All day"
        : EndTime.HasValue
            ? $"{StartTime!.Value:HH\\:mm}–{EndTime.Value:HH\\:mm}"
            : StartTime!.Value.ToString("HH:mm");
}

public sealed class WebRepository
{
    private static readonly JsonSerializerOptions JsonOpts = new() { PropertyNameCaseInsensitive = true };

    private readonly string _dbPath;
    private readonly string _scheduleDbPath;

    public WebRepository(IConfiguration config)
    {
        // Priority 1: appsettings.json → Database:Path (explicit override)
        // Priority 2: WPF app settings file (DataInference.DatabasePath)
        // Priority 3: default path (%LOCALAPPDATA%\JinoWorkHost\process-review.db)
        string? configured = config["Database:Path"];
        _dbPath = !string.IsNullOrWhiteSpace(configured)
            ? configured
            : WpfSettingsReader.TryGetDatabasePath()
              ?? Path.Combine(
                  Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                  "JinoWorkHost", "process-review.db");

        // Schedule DB: shared with WPF app (schedule.db)
        // Priority 1: appsettings.json → Schedule:Path
        // Priority 2: WPF app settings file (Schedule.DatabasePath)
        // Priority 3: default %LOCALAPPDATA%\JinoWorkHost\schedule.db
        string? schedulePath = config["Schedule:Path"];
        _scheduleDbPath = !string.IsNullOrWhiteSpace(schedulePath)
            ? schedulePath
            : WpfSettingsReader.TryGetScheduleDatabasePath()
              ?? Path.Combine(
                  Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                  "JinoWorkHost", "schedule.db");

        EnsureDatabase();
        EnsureScheduleDatabase();
    }

    public string GetDbPath() => _dbPath;

    // ── Connection ────────────────────────────────────────────────────────────

    private SqliteConnection OpenConnection()
    {
        var conn = new SqliteConnection($"Data Source={_dbPath}");
        conn.Open();
        return conn;
    }

    private SqliteConnection OpenScheduleConnection()
    {
        string? dir = Path.GetDirectoryName(_scheduleDbPath);
        if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
        var conn = new SqliteConnection($"Data Source={_scheduleDbPath}");
        conn.Open();
        return conn;
    }

    private void EnsureScheduleDatabase()
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS Schedules (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                Title       TEXT    NOT NULL DEFAULT '',
                Description TEXT    NOT NULL DEFAULT '',
                StartDate   TEXT    NOT NULL,
                EndDate     TEXT    NOT NULL,
                Color       TEXT    NOT NULL DEFAULT '#4A90D9',
                Tags        TEXT    NOT NULL DEFAULT '[]',
                CreatedAt   TEXT    NOT NULL,
                StartTime   TEXT    NULL,
                EndTime     TEXT    NULL
            );
            CREATE INDEX IF NOT EXISTS idx_sched_start ON Schedules(StartDate);
            CREATE INDEX IF NOT EXISTS idx_sched_end   ON Schedules(EndDate);
            """;
        cmd.ExecuteNonQuery();

        // Migrate: ensure StartTime/EndTime columns exist (legacy DBs may not have them)
        using SqliteCommand check = conn.CreateCommand();
        check.CommandText = "PRAGMA table_info(Schedules);";
        bool hasStartTime = false, hasEndTime = false;
        using (SqliteDataReader r = check.ExecuteReader())
            while (r.Read())
            {
                string col = r.GetString(1);
                if (col.Equals("StartTime", StringComparison.OrdinalIgnoreCase)) hasStartTime = true;
                if (col.Equals("EndTime",   StringComparison.OrdinalIgnoreCase)) hasEndTime   = true;
            }
        if (!hasStartTime) { using SqliteCommand a = conn.CreateCommand(); a.CommandText = "ALTER TABLE Schedules ADD COLUMN StartTime TEXT NULL;"; a.ExecuteNonQuery(); }
        if (!hasEndTime)   { using SqliteCommand a = conn.CreateCommand(); a.CommandText = "ALTER TABLE Schedules ADD COLUMN EndTime TEXT NULL;";   a.ExecuteNonQuery(); }
    }

    private void EnsureDatabase()
    {
        string? dir = Path.GetDirectoryName(_dbPath);
        if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);

        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS DataTables (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL,
                TableName   TEXT    NOT NULL DEFAULT '',
                Columns     TEXT    NOT NULL DEFAULT '[]',
                CreatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_dt_dataset ON DataTables(DatasetName);

            CREATE TABLE IF NOT EXISTS DataTableRows (
                Id      INTEGER PRIMARY KEY AUTOINCREMENT,
                TableId INTEGER NOT NULL REFERENCES DataTables(Id),
                RowData TEXT    NOT NULL DEFAULT '{}'
            );
            CREATE INDEX IF NOT EXISTS idx_dtr_table ON DataTableRows(TableId);

            CREATE TABLE IF NOT EXISTS DatasetTags (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL UNIQUE,
                Tags        TEXT    NOT NULL DEFAULT '[]',
                CreatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_dtag_dataset ON DatasetTags(DatasetName);

            CREATE TABLE IF NOT EXISTS DatasetMemo (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL UNIQUE,
                Memo        TEXT    NOT NULL DEFAULT '',
                UpdatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_dmemo_dataset ON DatasetMemo(DatasetName);

            CREATE TABLE IF NOT EXISTS Reports (
                Id           INTEGER PRIMARY KEY AUTOINCREMENT,
                Title        TEXT    NOT NULL DEFAULT '',
                DatasetNames TEXT    NOT NULL DEFAULT '',
                HtmlContent  TEXT    NOT NULL DEFAULT '',
                CreatedAt    TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS Users (
                Id           INTEGER PRIMARY KEY AUTOINCREMENT,
                Username     TEXT    NOT NULL UNIQUE COLLATE NOCASE,
                PasswordHash TEXT    NOT NULL,
                Role         TEXT    NOT NULL DEFAULT 'Viewer',
                CreatedAt    TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS AppSettings (
                Key   TEXT PRIMARY KEY NOT NULL,
                Value TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS DatasetImages (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL,
                FileName    TEXT    NOT NULL DEFAULT '',
                ImageData   BLOB    NOT NULL,
                CreatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_di_dataset ON DatasetImages(DatasetName);

            CREATE TABLE IF NOT EXISTS RawReports (
                Id            INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName   TEXT    NOT NULL UNIQUE,
                ProductType   TEXT    NOT NULL DEFAULT '',
                ReportDate    TEXT    NOT NULL DEFAULT '',
                CreatedAt     TEXT    NOT NULL,
                BatchExcluded INTEGER NOT NULL DEFAULT 0
            );
            CREATE INDEX IF NOT EXISTS idx_rr_name ON RawReports(DatasetName);

            CREATE TABLE IF NOT EXISTS RawReportImages (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL,
                FileName    TEXT    NOT NULL DEFAULT '',
                SortOrder   INTEGER NOT NULL DEFAULT 0,
                MediaType   TEXT    NOT NULL DEFAULT 'image/png',
                ImageData   BLOB    NOT NULL,
                CreatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_rri_dataset ON RawReportImages(DatasetName);

            CREATE TABLE IF NOT EXISTS RawReportFiles (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL,
                FileName    TEXT    NOT NULL DEFAULT '',
                MediaType   TEXT    NOT NULL DEFAULT 'application/octet-stream',
                FileSize    INTEGER NOT NULL DEFAULT 0,
                FileData    BLOB    NOT NULL,
                CreatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_rrf_dataset ON RawReportFiles(DatasetName);

            CREATE TABLE IF NOT EXISTS NormalizedMeasurements (
                Id             INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName    TEXT    NOT NULL,
                ProductType    TEXT    NOT NULL DEFAULT '',
                TestDate       TEXT    NOT NULL DEFAULT '',
                Line           TEXT    NOT NULL DEFAULT '',
                CheckType      TEXT    NOT NULL DEFAULT '',
                Variable       TEXT    NOT NULL DEFAULT '',
                VariableDetail TEXT    NOT NULL DEFAULT '',
                VariableGroup  TEXT    NOT NULL DEFAULT '',
                Intervention   TEXT    NOT NULL DEFAULT '',
                InputQty       INTEGER NOT NULL DEFAULT 0,
                OkQty          INTEGER NOT NULL DEFAULT 0,
                NgTotal        INTEGER NOT NULL DEFAULT 0,
                NgRate         REAL    NOT NULL DEFAULT 0,
                DefectCategory TEXT    NOT NULL DEFAULT '',
                DefectType     TEXT    NOT NULL DEFAULT '',
                DefectCount    INTEGER NOT NULL DEFAULT 0,
                CreatedAt      TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_nm_dataset ON NormalizedMeasurements(DatasetName);

            CREATE TABLE IF NOT EXISTS DatasetSummary (
                Id                INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName       TEXT    NOT NULL UNIQUE,
                ProductType       TEXT    NOT NULL DEFAULT '',
                Summary           TEXT    NOT NULL DEFAULT '',
                KeyFindings       TEXT    NOT NULL DEFAULT '',
                CreatedAt         TEXT    NOT NULL,
                Purpose           TEXT    NOT NULL DEFAULT '',
                TestConditions    TEXT    NOT NULL DEFAULT '',
                RootCause         TEXT    NOT NULL DEFAULT '',
                Decision          TEXT    NOT NULL DEFAULT '',
                RecommendedAction TEXT    NOT NULL DEFAULT '',
                Verdict           TEXT    NOT NULL DEFAULT '',
                Headline          TEXT    NOT NULL DEFAULT '',
                EvidenceJson      TEXT    NOT NULL DEFAULT '',
                ActionsJson       TEXT    NOT NULL DEFAULT '',
                ContextJson       TEXT    NOT NULL DEFAULT '',
                ReportType        TEXT    NOT NULL DEFAULT '',
                DoeGridJson       TEXT    NOT NULL DEFAULT '',
                TrendJson         TEXT    NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_dsum_name ON DatasetSummary(DatasetName);

            -- One row per (dataset, language). The base DatasetSummary row
            -- is treated as the original/AI-output language; additional
            -- rows here hold translations (ko / vi) so the Detail page can
            -- switch between them without re-calling the API.
            CREATE TABLE IF NOT EXISTS DatasetSummaryTranslations (
                DatasetName       TEXT NOT NULL,
                Lang              TEXT NOT NULL,
                Summary           TEXT NOT NULL DEFAULT '',
                KeyFindings       TEXT NOT NULL DEFAULT '',
                Purpose           TEXT NOT NULL DEFAULT '',
                TestConditions    TEXT NOT NULL DEFAULT '',
                RootCause         TEXT NOT NULL DEFAULT '',
                Decision          TEXT NOT NULL DEFAULT '',
                RecommendedAction TEXT NOT NULL DEFAULT '',
                Headline          TEXT NOT NULL DEFAULT '',
                ActionsJson       TEXT NOT NULL DEFAULT '',
                ContextJson       TEXT NOT NULL DEFAULT '',
                UpdatedAt         TEXT NOT NULL,
                PRIMARY KEY (DatasetName, Lang)
            );

            CREATE TABLE IF NOT EXISTS RawReportText (
                DatasetName   TEXT NOT NULL,
                Kind          TEXT NOT NULL DEFAULT 'ocr',
                ExtractedText TEXT NOT NULL DEFAULT '',
                CreatedAt     TEXT NOT NULL,
                PRIMARY KEY (DatasetName, Kind)
            );

            CREATE TABLE IF NOT EXISTS AskAiHistory (
                Id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                Question           TEXT    NOT NULL,
                ProductTypeFilter  TEXT    NOT NULL DEFAULT '',
                Overall            TEXT    NOT NULL DEFAULT '',
                PerDatasetJson     TEXT    NOT NULL DEFAULT '[]',
                CreatedAt          TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_askai_created ON AskAiHistory(CreatedAt DESC);

            -- Maps an arbitrary filename token (e.g. "BRS-161014", "BRS-2015")
            -- to a Mid-group Material name (e.g. "BRS-161016S08ZZ"). Used by
            -- the Data Input page to resolve Product Type for files whose
            -- token doesn't match any registered Material via fuzzy rules.
            -- Populated by the user through the Review Tokens panel.
            CREATE TABLE IF NOT EXISTS DataInputAliases (
                Token     TEXT NOT NULL PRIMARY KEY,
                Material  TEXT NOT NULL,
                CreatedAt TEXT NOT NULL,
                UpdatedAt TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS MenuPermissions (
                Role   TEXT NOT NULL,
                MenuId TEXT NOT NULL,
                PRIMARY KEY (Role, MenuId)
            );

            CREATE TABLE IF NOT EXISTS MtypeCategories (
                Code TEXT PRIMARY KEY,
                Name TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS ModelGroups (
                Id           INTEGER PRIMARY KEY AUTOINCREMENT,
                Name         TEXT    NOT NULL DEFAULT '',
                ProductGroup TEXT    NOT NULL DEFAULT 'ETC',
                SortOrder    INTEGER NOT NULL DEFAULT 0,
                UpdatedAt    TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS ModelGroupItems (
                Id        INTEGER PRIMARY KEY AUTOINCREMENT,
                GroupId   INTEGER NOT NULL REFERENCES ModelGroups(Id) ON DELETE CASCADE,
                LineShift TEXT    NOT NULL DEFAULT '',
                Material  TEXT    NOT NULL DEFAULT '',
                SubGroup  TEXT    NOT NULL DEFAULT '',
                SortOrder INTEGER NOT NULL DEFAULT 0
            );
            CREATE INDEX IF NOT EXISTS idx_mgi_group ON ModelGroupItems(GroupId);

            CREATE TABLE IF NOT EXISTS BmesMaterials (
                Matnr     TEXT PRIMARY KEY,
                Maktx     TEXT NOT NULL DEFAULT '',
                Meins     TEXT NOT NULL DEFAULT '',
                Injtp     TEXT NOT NULL DEFAULT '',
                Mtype     TEXT NOT NULL DEFAULT '',
                Btype     TEXT NOT NULL DEFAULT '',
                MngCode   TEXT NOT NULL DEFAULT '',
                ModNameB  TEXT NOT NULL DEFAULT '',
                LotQt     TEXT NOT NULL DEFAULT '',
                Bunch     TEXT NOT NULL DEFAULT '',
                NgTar     TEXT NOT NULL DEFAULT '',
                McLv1Tx   TEXT NOT NULL DEFAULT '',
                McLv2Tx   TEXT NOT NULL DEFAULT '',
                McLv3Tx   TEXT NOT NULL DEFAULT '',
                McLv4Tx   TEXT NOT NULL DEFAULT '',
                McLv5Tx   TEXT NOT NULL DEFAULT '',
                McLv6Tx   TEXT NOT NULL DEFAULT '',
                Ernam     TEXT NOT NULL DEFAULT '',
                Erdat     TEXT NOT NULL DEFAULT '',
                Grcod     TEXT NOT NULL DEFAULT '',
                Grnam     TEXT NOT NULL DEFAULT '',
                MfPhi     TEXT NOT NULL DEFAULT '',
                FetchedAt TEXT NOT NULL
            );

            -- ── AI_EXCEL_PROC.md schema (Batch AI inference output) ─────────────
            -- One row per normalized Excel report. RawJson preserves the full
            -- agent output for re-ingestion / debugging. Source_dataset links
            -- back to RawReports.DatasetName so the row can be located in the
            -- existing Data Input UI.
            CREATE TABLE IF NOT EXISTS AiDocuments (
                DocumentId          TEXT PRIMARY KEY,
                SourceDataset       TEXT NOT NULL DEFAULT '',
                SourceFile          TEXT NOT NULL DEFAULT '',
                Title               TEXT NOT NULL DEFAULT '',
                Model               TEXT NOT NULL DEFAULT '',
                ReportDate          TEXT NOT NULL DEFAULT '',
                Department          TEXT NOT NULL DEFAULT '',
                Marker              TEXT NOT NULL DEFAULT '',
                Line                TEXT NOT NULL DEFAULT '',
                ReportType          TEXT NOT NULL DEFAULT '',
                PrimaryDefect       TEXT NOT NULL DEFAULT '',
                PrimaryDefectJson   TEXT NOT NULL DEFAULT '',  -- {canonical_name, aliases_in_document}
                RelatedDefectsJson  TEXT NOT NULL DEFAULT '',
                PartsJson           TEXT NOT NULL DEFAULT '',
                ProcessesJson       TEXT NOT NULL DEFAULT '',
                Purpose             TEXT NOT NULL DEFAULT '',
                ContentJson         TEXT NOT NULL DEFAULT '',
                SourceCellsJson     TEXT NOT NULL DEFAULT '',
                Confidence          REAL NOT NULL DEFAULT 0,
                SchemaVersion       TEXT NOT NULL DEFAULT '',
                RawJson             TEXT NOT NULL DEFAULT '',  -- full JSON dump from agent
                CreatedAt           TEXT NOT NULL,
                UpdatedAt           TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_aidoc_dataset ON AiDocuments(SourceDataset);

            CREATE TABLE IF NOT EXISTS AiTestConditions (
                ConditionId      TEXT PRIMARY KEY,
                DocumentId       TEXT NOT NULL REFERENCES AiDocuments(DocumentId) ON DELETE CASCADE,
                ConditionGroup   TEXT NOT NULL DEFAULT '',
                Line             TEXT NOT NULL DEFAULT '',
                Process          TEXT NOT NULL DEFAULT '',
                ChangedFactor    TEXT NOT NULL DEFAULT '',
                BeforeValue      TEXT,
                AfterValue       TEXT,
                Unit             TEXT,
                Machine          TEXT,
                Jig              TEXT,
                MaterialLot      TEXT,
                Supplier         TEXT,
                DryTimeSec       REAL,
                Temperature      TEXT,
                Pressure         TEXT,
                BondAmount       TEXT,
                UvEnergy         TEXT,
                SourceFile       TEXT NOT NULL DEFAULT '',
                SheetName        TEXT NOT NULL DEFAULT '',
                SourceCellsJson  TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_aitc_doc ON AiTestConditions(DocumentId);

            CREATE TABLE IF NOT EXISTS AiResults (
                ResultId         TEXT PRIMARY KEY,
                DocumentId       TEXT NOT NULL REFERENCES AiDocuments(DocumentId) ON DELETE CASCADE,
                ConditionId      TEXT,
                MeasurementType  TEXT NOT NULL DEFAULT '',
                ConditionGroup   TEXT NOT NULL DEFAULT '',
                ResultDate       TEXT NOT NULL DEFAULT '',
                Line             TEXT NOT NULL DEFAULT '',
                InputCount       REAL,
                OkCount          REAL,
                NgCount          REAL,
                NgRateDecimal    REAL,
                NgRatePercent    REAL,
                MetricName       TEXT NOT NULL DEFAULT '',
                MetricValue      REAL,
                Unit             TEXT,
                Judgement        TEXT,
                SourceFile       TEXT NOT NULL DEFAULT '',
                SheetName        TEXT NOT NULL DEFAULT '',
                SourceCellsJson  TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_airesult_doc ON AiResults(DocumentId);
            CREATE INDEX IF NOT EXISTS idx_airesult_cond ON AiResults(ConditionId);

            CREATE TABLE IF NOT EXISTS AiNgBreakdowns (
                BreakdownId  TEXT PRIMARY KEY,
                ResultId     TEXT NOT NULL REFERENCES AiResults(ResultId) ON DELETE CASCADE,
                DefectName   TEXT NOT NULL DEFAULT '',
                DefectCount  REAL,
                DefectRate   REAL
            );
            CREATE INDEX IF NOT EXISTS idx_aing_result ON AiNgBreakdowns(ResultId);

            CREATE TABLE IF NOT EXISTS AiConclusions (
                ConclusionId             TEXT PRIMARY KEY,
                DocumentId               TEXT NOT NULL REFERENCES AiDocuments(DocumentId) ON DELETE CASCADE,
                Topic                    TEXT NOT NULL DEFAULT '',
                StatementFromReport      TEXT NOT NULL DEFAULT '',
                NormalizedInterpretation TEXT NOT NULL DEFAULT '',
                SourceFile               TEXT NOT NULL DEFAULT '',
                SheetName                TEXT NOT NULL DEFAULT '',
                SourceCellsJson          TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_aiconcl_doc ON AiConclusions(DocumentId);

            CREATE TABLE IF NOT EXISTS AiTroubleshootingHints (
                HintId           TEXT PRIMARY KEY,
                DocumentId       TEXT NOT NULL REFERENCES AiDocuments(DocumentId) ON DELETE CASCADE,
                DefectName       TEXT NOT NULL DEFAULT '',
                CheckItem        TEXT NOT NULL DEFAULT '',
                Reason           TEXT NOT NULL DEFAULT '',
                EvidenceStrength TEXT NOT NULL DEFAULT '',
                RelatedProcess   TEXT NOT NULL DEFAULT '',
                RelatedPart      TEXT NOT NULL DEFAULT '',
                SourceFile       TEXT NOT NULL DEFAULT '',
                SheetName        TEXT NOT NULL DEFAULT '',
                SourceCellsJson  TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_aihint_doc ON AiTroubleshootingHints(DocumentId);

            CREATE TABLE IF NOT EXISTS AiExtractionLogs (
                LogId             TEXT PRIMARY KEY,
                DocumentId        TEXT NOT NULL REFERENCES AiDocuments(DocumentId) ON DELETE CASCADE,
                Confidence        REAL NOT NULL DEFAULT 0,
                AssumptionsJson   TEXT NOT NULL DEFAULT '',
                WarningsJson      TEXT NOT NULL DEFAULT '',
                DecisionRationale TEXT NOT NULL DEFAULT '',
                CreatedAt         TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_ailog_doc ON AiExtractionLogs(DocumentId);

            -- ── ko/en/vi translations for narrative fields ──────────────────────
            -- Agent writes one row per (record, lang) — Lang in ('ko','en','vi').
            -- Numerics + measurements live in the base tables; only the
            -- AI-narrative text needs translating.
            CREATE TABLE IF NOT EXISTS AiDocumentTranslations (
                DocumentId  TEXT NOT NULL REFERENCES AiDocuments(DocumentId) ON DELETE CASCADE,
                Lang        TEXT NOT NULL,           -- 'ko' | 'en' | 'vi'
                Title       TEXT NOT NULL DEFAULT '',
                Purpose     TEXT NOT NULL DEFAULT '',
                ContentJson TEXT NOT NULL DEFAULT '',-- translated content[] array
                UpdatedAt   TEXT NOT NULL,
                PRIMARY KEY (DocumentId, Lang)
            );

            CREATE TABLE IF NOT EXISTS AiConclusionTranslations (
                ConclusionId             TEXT NOT NULL REFERENCES AiConclusions(ConclusionId) ON DELETE CASCADE,
                Lang                     TEXT NOT NULL,
                Topic                    TEXT NOT NULL DEFAULT '',
                StatementFromReport      TEXT NOT NULL DEFAULT '',
                NormalizedInterpretation TEXT NOT NULL DEFAULT '',
                UpdatedAt                TEXT NOT NULL,
                PRIMARY KEY (ConclusionId, Lang)
            );

            CREATE TABLE IF NOT EXISTS AiHintTranslations (
                HintId    TEXT NOT NULL REFERENCES AiTroubleshootingHints(HintId) ON DELETE CASCADE,
                Lang      TEXT NOT NULL,
                CheckItem TEXT NOT NULL DEFAULT '',
                Reason    TEXT NOT NULL DEFAULT '',
                UpdatedAt TEXT NOT NULL,
                PRIMARY KEY (HintId, Lang)
            );

            CREATE TABLE IF NOT EXISTS AiLogTranslations (
                LogId             TEXT NOT NULL REFERENCES AiExtractionLogs(LogId) ON DELETE CASCADE,
                Lang              TEXT NOT NULL,
                AssumptionsJson   TEXT NOT NULL DEFAULT '',
                WarningsJson      TEXT NOT NULL DEFAULT '',
                DecisionRationale TEXT NOT NULL DEFAULT '',
                UpdatedAt         TEXT NOT NULL,
                PRIMARY KEY (LogId, Lang)
            );
            """;
        cmd.ExecuteNonQuery();
        MigrateSchema(conn);
        SeedMtypeCategories(conn);
    }

    /// <summary>
    /// Idempotent seed of the Mtype → Category-name mapping. Uses INSERT OR IGNORE so
    /// that any user renames done later are preserved on subsequent starts.
    /// </summary>
    private static void SeedMtypeCategories(SqliteConnection conn)
    {
        (string Code, string Name)[] seed =
        [
            ("D001", "RA1(L)"),       ("D002", "RA2(L)"),       ("D003", "RM(L)"),
            ("D004", "RA1(R)"),       ("D005", "RA2(R)"),       ("D006", "RM(R)"),
            ("D007", "BUDASSY(L)"),   ("D008", "BUDASSY(R)"),   ("D009", "INSPECTION(L)"),
            ("D010", "INSPECTION(R)"),("D011", "SRVC"),         ("D012", "SPEAKER"),
            ("D013", "MODULE"),       ("D014", "SPK(ZZ)"),      ("D015", "UNIT(ZZ)"),
            ("D016", "FPBA"),         ("D017", "EXCITER"),      ("D018", "HEADSET"),
            ("D019", "RECEIVER"),     ("D020", "SFPRECEIVER"),  ("D021", "ACCESSORY"),
            ("D022", "ACCESSORYSUB"), ("D023", "FRONT(L)"),     ("D024", "FRONT(R)"),
            ("D025", "CKD"),          ("D026", "TAG"),          ("D028", "FA2(L)"),
            ("D029", "FA2(R)"),       ("D030", "SUB2(L)"),      ("D031", "SUB2(R)"),
            ("D032", "RA(L)"),        ("D033", "RA(R)"),        ("D034", "FA1(L)"),
            ("D035", "FA1(R)"),       ("D036", "SUB3FRONT"),    ("D037", "SUB3REAR"),
            ("D038", "BUZZERASSY"),   ("D039", "KITTING"),      ("D040", "CRADLE"),
            ("D041", "PACKING"),      ("D042", "CRADLE-SRVC"),  ("D043", "POGOASSY(L)"),
            ("D044", "POGOASSY(R)"),  ("D046", "RM(L)"),        ("D047", "RM(R)"),
            ("D048", "FRONTASSY(L)"), ("D049", "FRONTASSY(R)"), ("D052", "REARASSY(L)"),
            ("D053", "REARASSY(R)"),  ("D054", "FRONTASSY(L)"), ("D055", "FRONTASSY(R)"),
            ("D999", "OTHER"),
        ];

        using SqliteTransaction tx = conn.BeginTransaction();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = "INSERT OR IGNORE INTO MtypeCategories (Code, Name) VALUES (@c, @n);";
        SqliteParameter pC = cmd.Parameters.Add("@c", SqliteType.Text);
        SqliteParameter pN = cmd.Parameters.Add("@n", SqliteType.Text);
        foreach (var (code, name) in seed)
        {
            pC.Value = code;
            pN.Value = name;
            cmd.ExecuteNonQuery();
        }
        tx.Commit();
    }

    private static void MigrateSchema(SqliteConnection conn)
    {
        // ModelGroupItems.Material (for 중그룹 Material)
        bool hasMaterial = false;
        using (SqliteCommand ck = conn.CreateCommand())
        {
            ck.CommandText = "PRAGMA table_info(ModelGroupItems);";
            using SqliteDataReader r = ck.ExecuteReader();
            while (r.Read())
                if (r.GetString(1).Equals("Material", StringComparison.OrdinalIgnoreCase))
                    hasMaterial = true;
        }
        if (!hasMaterial)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE ModelGroupItems ADD COLUMN Material TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        // ModelGroupItems.SubGroup (optional sub-grouping inside a mid-group)
        bool hasSubGroup = false;
        using (SqliteCommand ck = conn.CreateCommand())
        {
            ck.CommandText = "PRAGMA table_info(ModelGroupItems);";
            using SqliteDataReader r = ck.ExecuteReader();
            while (r.Read())
                if (r.GetString(1).Equals("SubGroup", StringComparison.OrdinalIgnoreCase))
                    hasSubGroup = true;
        }
        if (!hasSubGroup)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE ModelGroupItems ADD COLUMN SubGroup TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        // ModelGroups.ProductGroup (SPK/UNIT/MODULE/TWS/ETC — replaces per-JSON-file attribute)
        bool hasProductGroup = false;
        using (SqliteCommand ck = conn.CreateCommand())
        {
            ck.CommandText = "PRAGMA table_info(ModelGroups);";
            using SqliteDataReader r = ck.ExecuteReader();
            while (r.Read())
                if (r.GetString(1).Equals("ProductGroup", StringComparison.OrdinalIgnoreCase))
                    hasProductGroup = true;
        }
        if (!hasProductGroup)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE ModelGroups ADD COLUMN ProductGroup TEXT NOT NULL DEFAULT 'ETC';";
            alter.ExecuteNonQuery();
        }

        // One-shot tracker for data migrations that should only run once.
        using (SqliteCommand mig = conn.CreateCommand())
        {
            mig.CommandText = """
                CREATE TABLE IF NOT EXISTS AppMigrations (
                    Name      TEXT PRIMARY KEY,
                    AppliedAt TEXT NOT NULL
                );
                """;
            mig.ExecuteNonQuery();
        }

        bool hasDatasetNames = false;
        bool hasPurpose      = false;
        using SqliteCommand check = conn.CreateCommand();
        check.CommandText = "PRAGMA table_info(Reports);";
        using (SqliteDataReader r = check.ExecuteReader())
            while (r.Read())
                if (r.GetString(1).Equals("DatasetNames", StringComparison.OrdinalIgnoreCase))
                    hasDatasetNames = true;

        if (!hasDatasetNames)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE Reports ADD COLUMN DatasetNames TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        using SqliteCommand checkMemo = conn.CreateCommand();
        checkMemo.CommandText = "PRAGMA table_info(DatasetMemo);";
        using (SqliteDataReader r = checkMemo.ExecuteReader())
            while (r.Read())
                if (r.GetString(1).Equals("Purpose", StringComparison.OrdinalIgnoreCase))
                    hasPurpose = true;

        if (!hasPurpose)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE DatasetMemo ADD COLUMN Purpose TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        bool hasDisplayName = false;
        using SqliteCommand checkUsers = conn.CreateCommand();
        checkUsers.CommandText = "PRAGMA table_info(Users);";
        using (SqliteDataReader r = checkUsers.ExecuteReader())
            while (r.Read())
                if (r.GetString(1).Equals("DisplayName", StringComparison.OrdinalIgnoreCase))
                    hasDisplayName = true;

        if (!hasDisplayName)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE Users ADD COLUMN DisplayName TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        bool hasEditorHtml  = false;
        bool hasProductType = false;
        using SqliteCommand checkEh = conn.CreateCommand();
        checkEh.CommandText = "PRAGMA table_info(DatasetMemo);";
        using (SqliteDataReader r = checkEh.ExecuteReader())
            while (r.Read())
            {
                string col = r.GetString(1);
                if (col.Equals("EditorHtml",  StringComparison.OrdinalIgnoreCase)) hasEditorHtml  = true;
                if (col.Equals("ProductType", StringComparison.OrdinalIgnoreCase)) hasProductType = true;
            }

        if (!hasEditorHtml)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE DatasetMemo ADD COLUMN EditorHtml TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        if (!hasProductType)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE DatasetMemo ADD COLUMN ProductType TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        bool hasTags = false;
        using SqliteCommand checkTags = conn.CreateCommand();
        checkTags.CommandText = "PRAGMA table_info(DatasetSummary);";
        using (SqliteDataReader r = checkTags.ExecuteReader())
            while (r.Read())
                if (r.GetString(1).Equals("Tags", StringComparison.OrdinalIgnoreCase))
                    hasTags = true;

        if (!hasTags)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE DatasetSummary ADD COLUMN Tags TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        // DatasetSummary: add structured context columns (Purpose/TestConditions/RootCause/Decision/RecommendedAction)
        var existingDsumCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (SqliteCommand checkDsum = conn.CreateCommand())
        {
            checkDsum.CommandText = "PRAGMA table_info(DatasetSummary);";
            using SqliteDataReader r = checkDsum.ExecuteReader();
            while (r.Read()) existingDsumCols.Add(r.GetString(1));
        }
        foreach (string newCol in new[] { "Purpose", "TestConditions", "RootCause", "Decision", "RecommendedAction" })
        {
            if (existingDsumCols.Contains(newCol)) continue;
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = $"ALTER TABLE DatasetSummary ADD COLUMN {newCol} TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        // v2 (verdict-first) columns on DatasetSummary
        foreach (string newCol in new[] { "Verdict", "Headline", "EvidenceJson", "ActionsJson", "ContextJson" })
        {
            if (existingDsumCols.Contains(newCol)) continue;
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = $"ALTER TABLE DatasetSummary ADD COLUMN {newCol} TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        // v7 (reportType + DOE/trend payloads) columns
        foreach (string newCol in new[] { "ReportType", "DoeGridJson", "TrendJson" })
        {
            if (existingDsumCols.Contains(newCol)) continue;
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = $"ALTER TABLE DatasetSummary ADD COLUMN {newCol} TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        // v2 columns on DatasetSummaryTranslations (headline + actions + context only;
        // numbers/labels in evidence stay verbatim and are not translated)
        var existingTrCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (SqliteCommand checkTr = conn.CreateCommand())
        {
            checkTr.CommandText = "PRAGMA table_info(DatasetSummaryTranslations);";
            using SqliteDataReader r = checkTr.ExecuteReader();
            while (r.Read()) existingTrCols.Add(r.GetString(1));
        }
        foreach (string newCol in new[] { "Headline", "ActionsJson", "ContextJson" })
        {
            if (existingTrCols.Contains(newCol)) continue;
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = $"ALTER TABLE DatasetSummaryTranslations ADD COLUMN {newCol} TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        bool hasBatchExcluded = false;
        using SqliteCommand checkRr = conn.CreateCommand();
        checkRr.CommandText = "PRAGMA table_info(RawReports);";
        using (SqliteDataReader r = checkRr.ExecuteReader())
            while (r.Read())
                if (r.GetString(1).Equals("BatchExcluded", StringComparison.OrdinalIgnoreCase))
                    hasBatchExcluded = true;

        if (!hasBatchExcluded)
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE RawReports ADD COLUMN BatchExcluded INTEGER NOT NULL DEFAULT 0;";
            alter.ExecuteNonQuery();
        }

        using SqliteCommand ensureFiles = conn.CreateCommand();
        ensureFiles.CommandText = """
            CREATE TABLE IF NOT EXISTS RawReportFiles (
                Id          INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName TEXT    NOT NULL,
                FileName    TEXT    NOT NULL DEFAULT '',
                MediaType   TEXT    NOT NULL DEFAULT 'application/octet-stream',
                FileSize    INTEGER NOT NULL DEFAULT 0,
                FileData    BLOB    NOT NULL,
                CreatedAt   TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_rrf_dataset ON RawReportFiles(DatasetName);
            """;
        ensureFiles.ExecuteNonQuery();

        using SqliteCommand ensureText = conn.CreateCommand();
        ensureText.CommandText = """
            CREATE TABLE IF NOT EXISTS RawReportText (
                DatasetName   TEXT NOT NULL,
                Kind          TEXT NOT NULL DEFAULT 'ocr',
                ExtractedText TEXT NOT NULL DEFAULT '',
                CreatedAt     TEXT NOT NULL,
                PRIMARY KEY (DatasetName, Kind)
            );
            """;
        ensureText.ExecuteNonQuery();

        // Migration: RawReportText originally had only DatasetName as PK with no Kind column.
        // Add Kind + rebuild table with composite PK, classifying existing rows as
        // 'ocr' (structured markdown transcript) vs 'excel_paste' (raw tab-separated
        // text pasted from Excel) by content heuristic.
        bool hasKind = false;
        using (SqliteCommand ck = conn.CreateCommand())
        {
            ck.CommandText = "PRAGMA table_info(RawReportText);";
            using SqliteDataReader r = ck.ExecuteReader();
            while (r.Read())
                if (r.GetString(1).Equals("Kind", StringComparison.OrdinalIgnoreCase))
                    hasKind = true;
        }
        if (!hasKind)
        {
            using SqliteCommand migrate = conn.CreateCommand();
            migrate.CommandText = """
                ALTER TABLE RawReportText ADD COLUMN Kind TEXT NOT NULL DEFAULT 'ocr';

                UPDATE RawReportText
                   SET Kind = 'excel_paste'
                 WHERE ExtractedText NOT LIKE '%### Table:%'
                   AND ExtractedText NOT LIKE '%## I. Purpose%'
                   AND ExtractedText NOT LIKE '%Columns:%';

                CREATE TABLE RawReportText_new (
                    DatasetName   TEXT NOT NULL,
                    Kind          TEXT NOT NULL DEFAULT 'ocr',
                    ExtractedText TEXT NOT NULL DEFAULT '',
                    CreatedAt     TEXT NOT NULL,
                    PRIMARY KEY (DatasetName, Kind)
                );
                INSERT INTO RawReportText_new (DatasetName, Kind, ExtractedText, CreatedAt)
                SELECT DatasetName, Kind, ExtractedText, CreatedAt FROM RawReportText;
                DROP TABLE RawReportText;
                ALTER TABLE RawReportText_new RENAME TO RawReportText;
                """;
            migrate.ExecuteNonQuery();
        }
    }

    // ── Raw report text ───────────────────────────────────────────────────────
    // Two kinds live side-by-side per dataset:
    //   "ocr"         — structured markdown transcript produced by Vision OCR.
    //                   Batch normalize-from-text flow consumes this.
    //   "excel_paste" — raw tab-separated text pasted from Excel at Input time.
    //                   Passed as auxiliary rawText to NormalizeFromImagesAsync;
    //                   NOT a substitute for OCR markdown.

    public const string TextKindOcr        = "ocr";
    public const string TextKindExcelPaste = "excel_paste";

    public void SaveExtractedText(string datasetName, string text, string kind = TextKindOcr)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO RawReportText (DatasetName, Kind, ExtractedText, CreatedAt)
            VALUES (@n, @k, @t, @at)
            ON CONFLICT(DatasetName, Kind) DO UPDATE SET ExtractedText=@t, CreatedAt=@at;
            """;
        cmd.Parameters.AddWithValue("@n",  datasetName);
        cmd.Parameters.AddWithValue("@k",  kind);
        cmd.Parameters.AddWithValue("@t",  text ?? "");
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        cmd.ExecuteNonQuery();
    }

    public string? GetExtractedText(string datasetName, string kind = TextKindOcr)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT ExtractedText FROM RawReportText WHERE DatasetName=@n AND Kind=@k;";
        cmd.Parameters.AddWithValue("@n", datasetName);
        cmd.Parameters.AddWithValue("@k", kind);
        object? r = cmd.ExecuteScalar();
        return r is string s ? s : null;
    }

    /// <summary>
    /// Returns the set of DatasetNames that have non-empty text of the given
    /// <paramref name="kind"/> cached. Default kind is "ocr" so callers asking
    /// "which datasets have OCR cache" need no change.
    /// </summary>
    public HashSet<string> GetDatasetsWithExtractedText(string kind = TextKindOcr)
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DatasetName FROM RawReportText WHERE Kind=@k AND LENGTH(TRIM(ExtractedText)) > 0;";
        cmd.Parameters.AddWithValue("@k", kind);
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) set.Add(r.GetString(0));
        return set;
    }

    /// <summary>Delete a specific kind (default "ocr") or all kinds when <paramref name="kind"/> is null.</summary>
    public void DeleteExtractedText(string datasetName, string? kind = TextKindOcr)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        if (kind is null)
        {
            cmd.CommandText = "DELETE FROM RawReportText WHERE DatasetName=@n;";
            cmd.Parameters.AddWithValue("@n", datasetName);
        }
        else
        {
            cmd.CommandText = "DELETE FROM RawReportText WHERE DatasetName=@n AND Kind=@k;";
            cmd.Parameters.AddWithValue("@n", datasetName);
            cmd.Parameters.AddWithValue("@k", kind);
        }
        cmd.ExecuteNonQuery();
    }

    // ── App Settings ──────────────────────────────────────────────────────────

    public string? GetSetting(string key)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Value FROM AppSettings WHERE Key = @k;";
        cmd.Parameters.AddWithValue("@k", key);
        object? result = cmd.ExecuteScalar();
        return result is string s && s.Length > 0 ? s : null;
    }

    public void SetSetting(string key, string value)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO AppSettings (Key, Value) VALUES (@k, @v)
            ON CONFLICT(Key) DO UPDATE SET Value = excluded.Value;
            """;
        cmd.Parameters.AddWithValue("@k", key);
        cmd.Parameters.AddWithValue("@v", value);
        cmd.ExecuteNonQuery();
    }

    // ── Datasets ──────────────────────────────────────────────────────────────

    public List<string> GetAllDatasets()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT DatasetName FROM DataTables ORDER BY DatasetName;";
        var list = new List<string>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    public List<string> GetAllDistinctTags()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Tags FROM DatasetTags;";

        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            try
            {
                var tags = JsonSerializer.Deserialize<List<string>>(r.GetString(0)) ?? [];
                foreach (string t in tags)
                    if (!string.IsNullOrWhiteSpace(t)) set.Add(t);
            }
            catch { }
        }
        return [.. set.Order(StringComparer.OrdinalIgnoreCase)];
    }

    public List<string> GetDatasetsByTags(IReadOnlyList<string> tags)
    {
        if (tags.Count == 0) return GetAllDatasets();

        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DatasetName, Tags FROM DatasetTags;";

        var tagSet   = new HashSet<string>(tags, StringComparer.OrdinalIgnoreCase);
        var matchSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            try
            {
                var dsTags = JsonSerializer.Deserialize<List<string>>(r.GetString(1)) ?? [];
                if (dsTags.Any(tagSet.Contains))
                    matchSet.Add(r.GetString(0));
            }
            catch { }
        }
        return [.. matchSet.Order(StringComparer.OrdinalIgnoreCase)];
    }

    // ── Tables ────────────────────────────────────────────────────────────────

    public List<DataTableInfo> GetTablesForDataset(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT dt.Id, dt.DatasetName, dt.TableName, dt.Columns, dt.CreatedAt,
                   COUNT(dtr.Id) AS RowCount
            FROM DataTables dt
            LEFT JOIN DataTableRows dtr ON dtr.TableId = dt.Id
            WHERE dt.DatasetName = @d
            GROUP BY dt.Id
            ORDER BY dt.Id;
            """;
        cmd.Parameters.AddWithValue("@d", datasetName);

        var list = new List<DataTableInfo>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            List<ColumnDef> cols = [];
            try { cols = JsonSerializer.Deserialize<List<ColumnDef>>(r.GetString(3), JsonOpts) ?? []; } catch { }

            list.Add(new DataTableInfo
            {
                Id          = r.GetInt64(0),
                DatasetName = r.GetString(1),
                TableName   = r.GetString(2),
                Columns     = cols,
                CreatedAt   = r.GetString(4),
                RowCount    = r.GetInt32(5)
            });
        }
        return list;
    }

    public List<Dictionary<string, string>> GetTableRows(long tableId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT RowData FROM DataTableRows WHERE TableId = @id ORDER BY Id;";
        cmd.Parameters.AddWithValue("@id", tableId);

        var list = new List<Dictionary<string, string>>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            try
            {
                var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(r.GetString(0), JsonOpts) ?? [];
                list.Add(dict);
            }
            catch { }
        }
        return list;
    }

    public long SaveTable(string datasetName, string tableName,
                          List<ColumnDef> columns, List<Dictionary<string, string>> rows)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        using SqliteCommand ins = conn.CreateCommand();
        ins.Transaction = tx;
        ins.CommandText = """
            INSERT INTO DataTables (DatasetName, TableName, Columns, CreatedAt)
            VALUES (@d, @n, @c, @at);
            SELECT last_insert_rowid();
            """;
        ins.Parameters.AddWithValue("@d",  datasetName);
        ins.Parameters.AddWithValue("@n",  tableName);
        ins.Parameters.AddWithValue("@c",  JsonSerializer.Serialize(columns, JsonOpts));
        ins.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        long tableId = Convert.ToInt64(ins.ExecuteScalar() ?? 0);

        foreach (Dictionary<string, string> row in rows)
        {
            using SqliteCommand insRow = conn.CreateCommand();
            insRow.Transaction = tx;
            insRow.CommandText = "INSERT INTO DataTableRows (TableId, RowData) VALUES (@t, @r);";
            insRow.Parameters.AddWithValue("@t", tableId);
            insRow.Parameters.AddWithValue("@r", JsonSerializer.Serialize(row, JsonOpts));
            insRow.ExecuteNonQuery();
        }

        tx.Commit();
        return tableId;
    }

    // ── Tags & Memo ───────────────────────────────────────────────────────────

    public void SaveTags(string datasetName, List<string> tags)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DatasetTags (DatasetName, Tags, CreatedAt)
            VALUES (@d, @t, @at)
            ON CONFLICT(DatasetName) DO UPDATE SET Tags=@t, CreatedAt=@at;
            """;
        cmd.Parameters.AddWithValue("@d",  datasetName);
        cmd.Parameters.AddWithValue("@t",  JsonSerializer.Serialize(tags));
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        cmd.ExecuteNonQuery();
    }

    public List<string> GetTags(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Tags FROM DatasetTags WHERE DatasetName=@d;";
        cmd.Parameters.AddWithValue("@d", datasetName);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (r.Read())
            try { return JsonSerializer.Deserialize<List<string>>(r.GetString(0)) ?? []; } catch { }
        return [];
    }

    public void SaveMemo(string datasetName, string memo)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DatasetMemo (DatasetName, Memo, UpdatedAt)
            VALUES (@d, @m, @at)
            ON CONFLICT(DatasetName) DO UPDATE SET Memo=@m, UpdatedAt=@at;
            """;
        cmd.Parameters.AddWithValue("@d",  datasetName);
        cmd.Parameters.AddWithValue("@m",  memo);
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        cmd.ExecuteNonQuery();
    }

    public string GetMemo(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Memo FROM DatasetMemo WHERE DatasetName=@d;";
        cmd.Parameters.AddWithValue("@d", datasetName);
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? r.GetString(0) : string.Empty;
    }

    public void SavePurpose(string datasetName, string purpose)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DatasetMemo (DatasetName, Memo, Purpose, UpdatedAt)
            VALUES (@d, '', @p, @at)
            ON CONFLICT(DatasetName) DO UPDATE SET Purpose=@p, UpdatedAt=@at;
            """;
        cmd.Parameters.AddWithValue("@d",  datasetName);
        cmd.Parameters.AddWithValue("@p",  purpose);
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        cmd.ExecuteNonQuery();
    }

    public string GetPurpose(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Purpose FROM DatasetMemo WHERE DatasetName=@d;";
        cmd.Parameters.AddWithValue("@d", datasetName);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return string.Empty;
        return r.IsDBNull(0) ? string.Empty : r.GetString(0);
    }

    public void SaveEditorHtml(string datasetName, string html)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DatasetMemo (DatasetName, Memo, Purpose, EditorHtml, UpdatedAt)
            VALUES (@d, '', '', @h, @at)
            ON CONFLICT(DatasetName) DO UPDATE SET EditorHtml=@h, UpdatedAt=@at;
            """;
        cmd.Parameters.AddWithValue("@d",  datasetName);
        cmd.Parameters.AddWithValue("@h",  html);
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        cmd.ExecuteNonQuery();
    }

    public string GetEditorHtml(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT EditorHtml FROM DatasetMemo WHERE DatasetName=@d;";
        cmd.Parameters.AddWithValue("@d", datasetName);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return string.Empty;
        return r.IsDBNull(0) ? string.Empty : r.GetString(0);
    }

    // ── RawReports ────────────────────────────────────────────────────────────

    public void SaveRawReport(string name, string productType, string date,
                               List<(string MediaType, byte[] Data, string FileName)> images)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        using var meta = conn.CreateCommand();
        meta.Transaction = tx;
        meta.CommandText = """
            INSERT INTO RawReports (DatasetName, ProductType, ReportDate, CreatedAt)
            VALUES (@n, @p, @d, @at)
            ON CONFLICT(DatasetName) DO UPDATE SET ProductType=@p, ReportDate=@d;
            """;
        meta.Parameters.AddWithValue("@n",  name);
        meta.Parameters.AddWithValue("@p",  productType);
        meta.Parameters.AddWithValue("@d",  date);
        meta.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        meta.ExecuteNonQuery();

        using var del = conn.CreateCommand();
        del.Transaction = tx;
        del.CommandText = "DELETE FROM RawReportImages WHERE DatasetName=@n;";
        del.Parameters.AddWithValue("@n", name);
        del.ExecuteNonQuery();

        for (int i = 0; i < images.Count; i++)
        {
            using var ins = conn.CreateCommand();
            ins.Transaction = tx;
            ins.CommandText = """
                INSERT INTO RawReportImages (DatasetName, FileName, SortOrder, MediaType, ImageData, CreatedAt)
                VALUES (@n, @f, @s, @m, @d, @at);
                """;
            ins.Parameters.AddWithValue("@n",  name);
            ins.Parameters.AddWithValue("@f",  images[i].FileName);
            ins.Parameters.AddWithValue("@s",  i);
            ins.Parameters.AddWithValue("@m",  images[i].MediaType);
            ins.Parameters.AddWithValue("@d",  images[i].Data);
            ins.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
            ins.ExecuteNonQuery();
        }

        tx.Commit();
    }

    public List<RawReportInfo> GetAllRawReports()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        // BatchedAt is a derived value. Prefer the newest of:
        //   AiDocuments.UpdatedAt   (AI_EXCEL_PROC.md schema — new CLI flow)
        //   DatasetSummary.CreatedAt (legacy v2 narrative card)
        //   NormalizedMeasurements.MAX(CreatedAt) (oldest fallback)
        // Without including AiDocuments here, rows processed by the new CLI
        // appear "never batched" even though AiDocuments has fresh rows.
        cmd.CommandText = """
            SELECT r.Id, r.DatasetName, r.ProductType, r.ReportDate, r.CreatedAt,
                   (SELECT COUNT(*) FROM RawReportImages WHERE DatasetName=r.DatasetName) AS ImgCnt,
                   (SELECT COUNT(*) FROM NormalizedMeasurements WHERE DatasetName=r.DatasetName) AS MeasCnt,
                   r.BatchExcluded,
                   COALESCE(
                     (SELECT MAX(UpdatedAt) FROM AiDocuments      WHERE SourceDataset=r.DatasetName),
                     (SELECT CreatedAt      FROM DatasetSummary   WHERE DatasetName  =r.DatasetName),
                     (SELECT MAX(CreatedAt) FROM NormalizedMeasurements WHERE DatasetName=r.DatasetName),
                     ''
                   ) AS BatchedAt
            FROM RawReports r
            ORDER BY r.CreatedAt DESC;
            """;
        var list = new List<RawReportInfo>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new RawReportInfo(r.GetInt64(0), r.GetString(1), r.GetString(2),
                                       r.GetString(3), r.GetInt32(5), r.GetInt32(6),
                                       r.GetString(4), r.GetInt32(7) != 0,
                                       r.IsDBNull(8) ? "" : r.GetString(8)));
        return list;
    }

    public void SetRawReportBatchExcluded(string datasetName, bool excluded)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE RawReports SET BatchExcluded=@e WHERE DatasetName=@n;";
        cmd.Parameters.AddWithValue("@e", excluded ? 1 : 0);
        cmd.Parameters.AddWithValue("@n", datasetName);
        cmd.ExecuteNonQuery();
    }

    /// <summary>Re-stamp the ProductType column for an existing RawReport.
    /// Used when the Data Input alias map changes and previously-saved rows
    /// should pick up the new mapping. Returns 1 when a row updated, 0 when
    /// the name didn't exist.</summary>
    public int UpdateRawReportProductType(string datasetName, string productType)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE RawReports SET ProductType=@p WHERE DatasetName=@n;";
        cmd.Parameters.AddWithValue("@p", productType ?? "");
        cmd.Parameters.AddWithValue("@n", datasetName);
        return cmd.ExecuteNonQuery();
    }

    public List<(string MediaType, byte[] Data, string FileName)> GetRawReportImages(string name)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT MediaType, ImageData, FileName FROM RawReportImages WHERE DatasetName=@n ORDER BY SortOrder;";
        cmd.Parameters.AddWithValue("@n", name);
        var list = new List<(string, byte[], string)>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.GetString(0), (byte[])r["ImageData"], r.GetString(2)));
        return list;
    }

    public void SaveNormalizedMeasurements(string name, List<NormalizedMeasurement> measurements)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        using var del = conn.CreateCommand();
        del.Transaction = tx;
        del.CommandText = "DELETE FROM NormalizedMeasurements WHERE DatasetName=@n;";
        del.Parameters.AddWithValue("@n", name);
        del.ExecuteNonQuery();

        foreach (NormalizedMeasurement m in measurements)
        {
            using var ins = conn.CreateCommand();
            ins.Transaction = tx;
            ins.CommandText = """
                INSERT INTO NormalizedMeasurements (
                    DatasetName, ProductType, TestDate, Line, CheckType,
                    Variable, VariableDetail, VariableGroup, Intervention,
                    InputQty, OkQty, NgTotal, NgRate,
                    DefectCategory, DefectType, DefectCount, CreatedAt)
                VALUES (@dn,@pt,@td,@li,@ct,@va,@vd,@vg,@iv,@iq,@oq,@nt,@nr,@dc,@dt,@dct,@at);
                """;
            ins.Parameters.AddWithValue("@dn",  name);
            ins.Parameters.AddWithValue("@pt",  m.ProductType);
            ins.Parameters.AddWithValue("@td",  m.TestDate);
            ins.Parameters.AddWithValue("@li",  m.Line);
            ins.Parameters.AddWithValue("@ct",  m.CheckType);
            ins.Parameters.AddWithValue("@va",  m.Variable);
            ins.Parameters.AddWithValue("@vd",  m.VariableDetail);
            ins.Parameters.AddWithValue("@vg",  m.VariableGroup);
            ins.Parameters.AddWithValue("@iv",  m.Intervention);
            ins.Parameters.AddWithValue("@iq",  m.InputQty);
            ins.Parameters.AddWithValue("@oq",  m.OkQty);
            ins.Parameters.AddWithValue("@nt",  m.NgTotal);
            ins.Parameters.AddWithValue("@nr",  m.NgRate);
            ins.Parameters.AddWithValue("@dc",  m.DefectCategory);
            ins.Parameters.AddWithValue("@dt",  m.DefectType);
            ins.Parameters.AddWithValue("@dct", m.DefectCount);
            ins.Parameters.AddWithValue("@at",  DateTime.UtcNow.ToString("O"));
            ins.ExecuteNonQuery();
        }

        tx.Commit();
    }

    public List<NormalizedMeasurement> GetNormalizedMeasurements(string name)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, ProductType, TestDate, Line, CheckType, Variable, VariableDetail,
                   VariableGroup, Intervention, InputQty, OkQty, NgTotal, NgRate,
                   DefectCategory, DefectType, DefectCount
            FROM NormalizedMeasurements WHERE DatasetName=@n ORDER BY Id;
            """;
        cmd.Parameters.AddWithValue("@n", name);
        var list = new List<NormalizedMeasurement>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new NormalizedMeasurement
            {
                Id             = r.GetInt64(0),
                ProductType    = r.GetString(1),
                TestDate       = r.GetString(2),
                Line           = r.GetString(3),
                CheckType      = r.GetString(4),
                Variable       = r.GetString(5),
                VariableDetail = r.GetString(6),
                VariableGroup  = r.GetString(7),
                Intervention   = r.GetString(8),
                InputQty       = r.GetInt32(9),
                OkQty          = r.GetInt32(10),
                NgTotal        = r.GetInt32(11),
                NgRate         = r.GetDouble(12),
                DefectCategory = r.GetString(13),
                DefectType     = r.GetString(14),
                DefectCount    = r.GetInt32(15),
            });
        return list;
    }

    private static readonly HashSet<string> _editableMeasurementFields = new(StringComparer.Ordinal)
    {
        "Variable", "VariableDetail", "VariableGroup", "Line", "CheckType",
        "InputQty", "OkQty", "NgTotal", "NgRate",
        "DefectType", "DefectCategory", "DefectCount", "Intervention",
    };

    public void UpdateNormalizedMeasurementField(long id, string field, string value)
    {
        if (!_editableMeasurementFields.Contains(field))
            throw new ArgumentException($"Field '{field}' is not editable.");

        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = $"UPDATE NormalizedMeasurements SET {field} = @v WHERE Id = @id;";
        cmd.Parameters.AddWithValue("@id", id);

        // Type-aware binding
        switch (field)
        {
            case "InputQty" or "OkQty" or "NgTotal" or "DefectCount":
                cmd.Parameters.AddWithValue("@v",
                    int.TryParse(value, out int i) ? i : 0);
                break;
            case "NgRate":
                string v = value.Trim().TrimEnd('%').Trim();
                cmd.Parameters.AddWithValue("@v",
                    double.TryParse(v, System.Globalization.NumberStyles.Float,
                                    System.Globalization.CultureInfo.InvariantCulture, out double d) ? d : 0.0);
                break;
            default:
                cmd.Parameters.AddWithValue("@v", value ?? "");
                break;
        }
        cmd.ExecuteNonQuery();
    }

    /// <summary>Convenience overload that serialises the v7 structured fields
    /// (tags/evidence/actions/context/reportType/doeGrid/trendPoints) from a
    /// NormalizeResult, so callers don't have to repeat the JSON.Serialize boilerplate.</summary>
    public void SaveDatasetSummaryRecord(string name, string productType, NormalizeResult r)
    {
        string tagsJson     = (r.Tags?.Count     ?? 0) > 0 ? System.Text.Json.JsonSerializer.Serialize(r.Tags)     : "";
        string evidenceJson = (r.Evidence?.Count ?? 0) > 0 ? System.Text.Json.JsonSerializer.Serialize(r.Evidence) : "";
        string actionsJson  = (r.Actions?.Count  ?? 0) > 0 ? System.Text.Json.JsonSerializer.Serialize(r.Actions)  : "";
        string contextJson  = r.Context is not null     ? System.Text.Json.JsonSerializer.Serialize(r.Context)  : "";
        string doeJson      = r.DoeGrid is not null     ? System.Text.Json.JsonSerializer.Serialize(r.DoeGrid)  : "";
        string trendJson    = (r.TrendPoints?.Count ?? 0) > 0 ? System.Text.Json.JsonSerializer.Serialize(r.TrendPoints) : "";
        SaveDatasetSummaryRecord(name, productType,
            r.Summary, r.KeyFindings, tagsJson,
            r.Purpose, r.TestConditions, r.RootCause, r.Decision, r.RecommendedAction,
            r.Verdict, r.Headline, evidenceJson, actionsJson, contextJson,
            r.ReportType, doeJson, trendJson);
    }

    public void SaveDatasetSummaryRecord(string name, string productType, string summary, string keyFindings, string tagsJson = "",
        string purpose = "", string testConditions = "", string rootCause = "",
        string decision = "", string recommendedAction = "",
        string verdict = "", string headline = "",
        string evidenceJson = "", string actionsJson = "", string contextJson = "",
        string reportType = "", string doeGridJson = "", string trendJson = "")
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DatasetSummary
                (DatasetName, ProductType, Summary, KeyFindings, Tags, CreatedAt,
                 Purpose, TestConditions, RootCause, Decision, RecommendedAction,
                 Verdict, Headline, EvidenceJson, ActionsJson, ContextJson,
                 ReportType, DoeGridJson, TrendJson)
            VALUES (@n, @p, @s, @k, @t, @at, @pu, @tc, @rc, @de, @ra,
                    @vd, @hl, @ev, @ac, @cx, @rt, @dg, @tr)
            ON CONFLICT(DatasetName) DO UPDATE SET
                ProductType=@p, Summary=@s, KeyFindings=@k, Tags=@t, CreatedAt=@at,
                Purpose=@pu, TestConditions=@tc, RootCause=@rc, Decision=@de, RecommendedAction=@ra,
                Verdict=@vd, Headline=@hl, EvidenceJson=@ev, ActionsJson=@ac, ContextJson=@cx,
                ReportType=@rt, DoeGridJson=@dg, TrendJson=@tr;
            """;
        cmd.Parameters.AddWithValue("@n",  name);
        cmd.Parameters.AddWithValue("@p",  productType);
        cmd.Parameters.AddWithValue("@s",  summary);
        cmd.Parameters.AddWithValue("@k",  keyFindings);
        cmd.Parameters.AddWithValue("@t",  tagsJson ?? "");
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        cmd.Parameters.AddWithValue("@pu", purpose           ?? "");
        cmd.Parameters.AddWithValue("@tc", testConditions    ?? "");
        cmd.Parameters.AddWithValue("@rc", rootCause         ?? "");
        cmd.Parameters.AddWithValue("@de", decision          ?? "");
        cmd.Parameters.AddWithValue("@ra", recommendedAction ?? "");
        cmd.Parameters.AddWithValue("@vd", verdict      ?? "");
        cmd.Parameters.AddWithValue("@hl", headline     ?? "");
        cmd.Parameters.AddWithValue("@ev", evidenceJson ?? "");
        cmd.Parameters.AddWithValue("@ac", actionsJson  ?? "");
        cmd.Parameters.AddWithValue("@cx", contextJson  ?? "");
        cmd.Parameters.AddWithValue("@rt", reportType   ?? "");
        cmd.Parameters.AddWithValue("@dg", doeGridJson  ?? "");
        cmd.Parameters.AddWithValue("@tr", trendJson    ?? "");
        cmd.ExecuteNonQuery();
    }

    // ── Raw attached files (any type — Excel, PDF, etc.) ─────────────────────

    public void AppendRawReportFiles(string name, List<(string MediaType, byte[] Data, string FileName)> files)
    {
        if (files.Count == 0) return;

        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        foreach (var (mediaType, data, fileName) in files)
        {
            using SqliteCommand ins = conn.CreateCommand();
            ins.Transaction = tx;
            ins.CommandText = """
                INSERT INTO RawReportFiles (DatasetName, FileName, MediaType, FileSize, FileData, CreatedAt)
                VALUES (@n, @fn, @mt, @sz, @d, @at);
                """;
            ins.Parameters.AddWithValue("@n",  name);
            ins.Parameters.AddWithValue("@fn", fileName ?? "");
            ins.Parameters.AddWithValue("@mt", string.IsNullOrEmpty(mediaType) ? "application/octet-stream" : mediaType);
            ins.Parameters.AddWithValue("@sz", (long)data.Length);
            ins.Parameters.AddWithValue("@d",  data);
            ins.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
            ins.ExecuteNonQuery();
        }
        tx.Commit();
    }

    public List<RawFileInfo> GetRawReportFileInfos(string name)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, FileName, MediaType, FileSize, CreatedAt
            FROM RawReportFiles WHERE DatasetName=@n ORDER BY Id;
            """;
        cmd.Parameters.AddWithValue("@n", name);
        List<RawFileInfo> list = [];
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new RawFileInfo(
                r.GetInt64(0),
                r.GetString(1),
                r.GetString(2),
                r.GetInt64(3),
                r.GetString(4)));
        }
        return list;
    }

    public (string FileName, string MediaType, byte[] Data)? GetRawReportFile(long fileId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT FileName, MediaType, FileData FROM RawReportFiles WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", fileId);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return null;

        string fn = r.GetString(0);
        string mt = r.GetString(1);
        using var ms = new MemoryStream();
        using var stream = r.GetStream(2);
        stream.CopyTo(ms);
        return (fn, mt, ms.ToArray());
    }

    public void DeleteRawReportFile(long fileId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM RawReportFiles WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", fileId);
        cmd.ExecuteNonQuery();
    }

    /// <summary>Distinct tags across all DatasetSummary.Tags JSON blobs.</summary>
    public List<string> GetAllDataInferenceTags()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Tags FROM DatasetSummary WHERE Tags IS NOT NULL AND Tags != '';";
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string json = r.GetString(0);
            if (string.IsNullOrWhiteSpace(json)) continue;
            try
            {
                var tags = System.Text.Json.JsonSerializer.Deserialize<List<string>>(json) ?? [];
                foreach (string t in tags)
                    if (!string.IsNullOrWhiteSpace(t)) set.Add(t.Trim());
            }
            catch { }
        }
        return [.. set.Order(StringComparer.OrdinalIgnoreCase)];
    }

    /// <summary>DataInference datasets whose Tags include ALL of the given tags.</summary>
    public List<string> GetDataInferenceDatasetsByTags(IReadOnlyList<string> filterTags)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT r.DatasetName, s.Tags
            FROM RawReports r
            LEFT JOIN DatasetSummary s ON s.DatasetName = r.DatasetName
            ORDER BY r.CreatedAt DESC;
            """;

        var list = new List<string>();
        using SqliteDataReader rd = cmd.ExecuteReader();
        while (rd.Read())
        {
            string name = rd.GetString(0);
            string json = rd.IsDBNull(1) ? "" : rd.GetString(1);

            if (filterTags.Count == 0) { list.Add(name); continue; }
            if (string.IsNullOrWhiteSpace(json)) continue;

            List<string> tags;
            try { tags = System.Text.Json.JsonSerializer.Deserialize<List<string>>(json) ?? []; }
            catch { continue; }

            bool hasAll = filterTags.All(f =>
                tags.Any(t => string.Equals(t, f, StringComparison.OrdinalIgnoreCase)));
            if (hasAll) list.Add(name);
        }
        return list;
    }

    public void UpdateDatasetTags(string name, string tagsJson)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE DatasetSummary SET Tags=@t WHERE DatasetName=@n;";
        cmd.Parameters.AddWithValue("@n", name);
        cmd.Parameters.AddWithValue("@t", tagsJson ?? "");
        cmd.ExecuteNonQuery();
    }

    public DatasetSummaryRecord? GetDatasetSummaryRecord(string name)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Summary, KeyFindings, Tags,
                   Purpose, TestConditions, RootCause, Decision, RecommendedAction,
                   Verdict, Headline, EvidenceJson, ActionsJson, ContextJson,
                   ReportType, DoeGridJson, TrendJson
            FROM DatasetSummary WHERE DatasetName=@n;
            """;
        cmd.Parameters.AddWithValue("@n", name);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return null;

        string summary           = r.GetString(0);
        string keyFindings       = r.GetString(1);
        string tagsJson          = r.IsDBNull(2) ? "" : r.GetString(2);
        string purpose           = r.IsDBNull(3) ? "" : r.GetString(3);
        string testConditions    = r.IsDBNull(4) ? "" : r.GetString(4);
        string rootCause         = r.IsDBNull(5) ? "" : r.GetString(5);
        string decision          = r.IsDBNull(6) ? "" : r.GetString(6);
        string recommendedAction = r.IsDBNull(7) ? "" : r.GetString(7);
        string verdict           = r.IsDBNull(8)  ? "" : r.GetString(8);
        string headline          = r.IsDBNull(9)  ? "" : r.GetString(9);
        string evidenceJson      = r.IsDBNull(10) ? "" : r.GetString(10);
        string actionsJson       = r.IsDBNull(11) ? "" : r.GetString(11);
        string contextJson       = r.IsDBNull(12) ? "" : r.GetString(12);
        string reportType        = r.IsDBNull(13) ? "" : r.GetString(13);
        string doeGridJson       = r.IsDBNull(14) ? "" : r.GetString(14);
        string trendJson         = r.IsDBNull(15) ? "" : r.GetString(15);

        var jsonOpts = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };

        List<string> tags = [];
        if (!string.IsNullOrWhiteSpace(tagsJson))
        {
            try { tags = System.Text.Json.JsonSerializer.Deserialize<List<string>>(tagsJson, jsonOpts) ?? []; }
            catch { }
        }

        List<EvidenceRow> evidence = [];
        if (!string.IsNullOrWhiteSpace(evidenceJson))
        {
            try { evidence = System.Text.Json.JsonSerializer.Deserialize<List<EvidenceRow>>(evidenceJson, jsonOpts) ?? []; }
            catch { }
        }

        List<ActionItem> actions = [];
        if (!string.IsNullOrWhiteSpace(actionsJson))
        {
            try { actions = System.Text.Json.JsonSerializer.Deserialize<List<ActionItem>>(actionsJson, jsonOpts) ?? []; }
            catch { }
        }

        AnalysisContext? context = null;
        if (!string.IsNullOrWhiteSpace(contextJson))
        {
            try { context = System.Text.Json.JsonSerializer.Deserialize<AnalysisContext>(contextJson, jsonOpts); }
            catch { }
        }

        DoeGrid? doeGrid = null;
        if (!string.IsNullOrWhiteSpace(doeGridJson))
        {
            try { doeGrid = System.Text.Json.JsonSerializer.Deserialize<DoeGrid>(doeGridJson, jsonOpts); }
            catch { }
        }

        List<TrendPoint>? trendPoints = null;
        if (!string.IsNullOrWhiteSpace(trendJson))
        {
            try { trendPoints = System.Text.Json.JsonSerializer.Deserialize<List<TrendPoint>>(trendJson, jsonOpts); }
            catch { }
        }

        var rec = new DatasetSummaryRecord
        {
            Summary           = summary,
            KeyFindings       = keyFindings,
            Tags              = tags,
            Purpose           = purpose,
            TestConditions    = testConditions,
            RootCause         = rootCause,
            Decision          = decision,
            RecommendedAction = recommendedAction,
            Verdict           = verdict,
            Headline          = headline,
            Evidence          = evidence,
            Actions           = actions,
            Context           = context,
            ReportType        = reportType,
            DoeGrid           = doeGrid,
            TrendPoints       = trendPoints,
        };

        // Attach all available translations (ko/vi/...) so the UI can switch
        // languages without an extra round-trip.
        foreach (var (lang, tr) in GetDatasetSummaryTranslations(name))
            rec.Translations[lang] = tr;
        return rec;
    }

    // ── AI_EXCEL_PROC.md schema read path ─────────────────────────────────────
    // Used by DataInferenceDbPage to render the new-CLI bundle when a row was
    // processed via AI_EXCEL_PROC.md (writes AiDocuments + children). Returns
    // null when the dataset has no AiDocuments row — the page falls back to
    // the old DatasetSummary card in that case.
    public AiDocBundle? GetAiDocBundle(string sourceDataset)
    {
        if (string.IsNullOrEmpty(sourceDataset)) return null;
        using SqliteConnection conn = OpenConnection();

        // 1) Base document
        AiDocBundle? bundle;
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT DocumentId, SourceDataset, SourceFile, Title, Purpose, PrimaryDefect,
                       ReportType, ReportDate, Confidence,
                       ContentJson, RelatedDefectsJson, PartsJson, ProcessesJson
                FROM AiDocuments WHERE SourceDataset=@n
                ORDER BY UpdatedAt DESC LIMIT 1;
                """;
            cmd.Parameters.AddWithValue("@n", sourceDataset);
            using var rd = cmd.ExecuteReader();
            if (!rd.Read()) return null;
            bundle = new AiDocBundle
            {
                DocumentId    = rd.GetString(0),
                SourceDataset = rd.GetString(1),
                SourceFile    = rd.GetString(2),
                Title         = rd.GetString(3),
                Purpose       = rd.GetString(4),
                PrimaryDefect = rd.GetString(5),
                ReportType    = rd.GetString(6),
                ReportDate    = rd.GetString(7),
                Confidence    = rd.IsDBNull(8) ? 0 : rd.GetDouble(8),
                Content        = ReadJsonStringList(rd.IsDBNull(9)  ? "" : rd.GetString(9)),
                RelatedDefects = ReadJsonStringList(rd.IsDBNull(10) ? "" : rd.GetString(10)),
                Parts          = ReadJsonStringList(rd.IsDBNull(11) ? "" : rd.GetString(11)),
                Processes      = ReadJsonStringList(rd.IsDBNull(12) ? "" : rd.GetString(12)),
            };
        }
        string docId = bundle.DocumentId;

        // 1b) Counts and measurement rows for dashboard-style rendering
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT COUNT(*) FROM AiTestConditions WHERE DocumentId=@d;";
            cmd.Parameters.AddWithValue("@d", docId);
            bundle.ConditionsCount = Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
        }

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT r.ResultId, COALESCE(r.ConditionId,''), r.MeasurementType, r.ConditionGroup,
                       COALESCE(c.Process,''), COALESCE(c.ChangedFactor,''),
                       COALESCE(c.BeforeValue,''), COALESCE(c.AfterValue,''),
                       r.ResultDate, r.Line, r.InputCount, r.OkCount, r.NgCount, r.NgRateDecimal,
                       r.NgRatePercent, r.MetricName, r.MetricValue, r.Unit, r.Judgement,
                       r.SourceFile, r.SheetName, r.SourceCellsJson
                FROM AiResults r
                LEFT JOIN AiTestConditions c ON c.ConditionId = r.ConditionId
                WHERE r.DocumentId=@d
                ORDER BY r.ResultDate, r.ResultId;
                """;
            cmd.Parameters.AddWithValue("@d", docId);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                bundle.Results.Add(new AiResultRow
                {
                    ResultId        = rd.GetString(0),
                    ConditionId     = rd.GetString(1),
                    MeasurementType = rd.GetString(2),
                    ConditionGroup  = rd.GetString(3),
                    ConditionProcess = rd.GetString(4),
                    ChangedFactor   = rd.GetString(5),
                    BeforeValue     = rd.GetString(6),
                    AfterValue      = rd.GetString(7),
                    ResultDate      = rd.GetString(8),
                    Line            = rd.GetString(9),
                    InputCount      = ReadNullableDouble(rd, 10),
                    OkCount         = ReadNullableDouble(rd, 11),
                    NgCount         = ReadNullableDouble(rd, 12),
                    NgRateDecimal   = ReadNullableDouble(rd, 13),
                    NgRatePercent   = ReadNullableDouble(rd, 14),
                    MetricName      = rd.GetString(15),
                    MetricValue     = ReadNullableDouble(rd, 16),
                    Unit            = rd.IsDBNull(17) ? "" : rd.GetString(17),
                    Judgement       = rd.IsDBNull(18) ? "" : rd.GetString(18),
                    SourceFile      = rd.GetString(19),
                    SheetName       = rd.GetString(20),
                    SourceCellsJson = rd.IsDBNull(21) ? "" : rd.GetString(21),
                });
            }
        }

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT b.DefectName,
                       COUNT(*) AS RowCount,
                       SUM(COALESCE(b.DefectCount,0)) AS TotalCount,
                       AVG(b.DefectRate) AS AvgRate
                FROM AiNgBreakdowns b
                JOIN AiResults r ON r.ResultId = b.ResultId
                WHERE r.DocumentId=@d
                GROUP BY b.DefectName
                HAVING TotalCount > 0
                ORDER BY TotalCount DESC, RowCount DESC
                LIMIT 12;
                """;
            cmd.Parameters.AddWithValue("@d", docId);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                bundle.NgBreakdowns.Add(new AiNgBreakdownSummaryRow
                {
                    DefectName = rd.GetString(0),
                    RowCount   = rd.IsDBNull(1) ? 0 : rd.GetInt32(1),
                    TotalCount = rd.IsDBNull(2) ? 0 : rd.GetDouble(2),
                    AvgRate    = ReadNullableDouble(rd, 3),
                });
            }
        }

        // 2) Conclusions
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT ConclusionId, Topic, StatementFromReport, NormalizedInterpretation,
                       SourceFile, SheetName
                FROM AiConclusions WHERE DocumentId=@d ORDER BY ConclusionId;
                """;
            cmd.Parameters.AddWithValue("@d", docId);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                bundle.Conclusions.Add(new AiConclusionRow
                {
                    ConclusionId             = rd.GetString(0),
                    Topic                    = rd.GetString(1),
                    StatementFromReport      = rd.GetString(2),
                    NormalizedInterpretation = rd.GetString(3),
                    SourceFile               = rd.GetString(4),
                    SheetName                = rd.GetString(5),
                });
            }
        }

        // 3) Troubleshooting hints
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT HintId, DefectName, CheckItem, Reason, EvidenceStrength,
                       RelatedProcess, RelatedPart, SourceFile, SheetName
                FROM AiTroubleshootingHints WHERE DocumentId=@d ORDER BY HintId;
                """;
            cmd.Parameters.AddWithValue("@d", docId);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                bundle.Hints.Add(new AiHintRow
                {
                    HintId           = rd.GetString(0),
                    DefectName       = rd.GetString(1),
                    CheckItem        = rd.GetString(2),
                    Reason           = rd.GetString(3),
                    EvidenceStrength = rd.GetString(4),
                    RelatedProcess   = rd.GetString(5),
                    RelatedPart      = rd.GetString(6),
                    SourceFile       = rd.GetString(7),
                    SheetName        = rd.GetString(8),
                });
            }
        }

        // 4) DecisionRationale from the (single) AiExtractionLogs row
        string logId = "";
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT LogId, AssumptionsJson, WarningsJson, DecisionRationale FROM AiExtractionLogs
                WHERE DocumentId=@d ORDER BY CreatedAt DESC LIMIT 1;
                """;
            cmd.Parameters.AddWithValue("@d", docId);
            using var rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                logId = rd.GetString(0);
                bundle.Assumptions = ReadJsonStringList(rd.GetString(1));
                bundle.Warnings = ReadJsonStringList(rd.GetString(2));
                bundle.DecisionRationale = rd.GetString(3);
            }
        }

        // 5) Translations (ko/en/vi) — narrative-only
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT Lang, Title, Purpose, ContentJson FROM AiDocumentTranslations WHERE DocumentId=@d;";
            cmd.Parameters.AddWithValue("@d", docId);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                bundle.DocTranslations[rd.GetString(0)] = new AiDocTranslationRow
                {
                    Title   = rd.GetString(1),
                    Purpose = rd.GetString(2),
                    Content = ReadJsonStringList(rd.IsDBNull(3) ? "" : rd.GetString(3)),
                };
            }
        }

        if (bundle.Conclusions.Count > 0)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT t.ConclusionId, t.Lang, t.Topic, t.StatementFromReport, t.NormalizedInterpretation
                FROM AiConclusionTranslations t
                JOIN AiConclusions c ON c.ConclusionId = t.ConclusionId
                WHERE c.DocumentId = @d;
                """;
            cmd.Parameters.AddWithValue("@d", docId);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string cid  = rd.GetString(0);
                string lang = rd.GetString(1);
                if (!bundle.ConclusionTranslations.TryGetValue(cid, out var byLang))
                    bundle.ConclusionTranslations[cid] = byLang = new();
                byLang[lang] = new AiConclusionTranslationRow
                {
                    Topic                    = rd.GetString(2),
                    StatementFromReport      = rd.GetString(3),
                    NormalizedInterpretation = rd.GetString(4),
                };
            }
        }

        if (bundle.Hints.Count > 0)
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT t.HintId, t.Lang, t.CheckItem, t.Reason
                FROM AiHintTranslations t
                JOIN AiTroubleshootingHints h ON h.HintId = t.HintId
                WHERE h.DocumentId = @d;
                """;
            cmd.Parameters.AddWithValue("@d", docId);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                string hid  = rd.GetString(0);
                string lang = rd.GetString(1);
                if (!bundle.HintTranslations.TryGetValue(hid, out var byLang))
                    bundle.HintTranslations[hid] = byLang = new();
                byLang[lang] = new AiHintTranslationRow
                {
                    CheckItem = rd.GetString(2),
                    Reason    = rd.GetString(3),
                };
            }
        }

        if (!string.IsNullOrEmpty(logId))
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT Lang, AssumptionsJson, WarningsJson, DecisionRationale FROM AiLogTranslations WHERE LogId=@l;";
            cmd.Parameters.AddWithValue("@l", logId);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                bundle.LogTranslations[rd.GetString(0)] = new AiLogTranslationRow
                {
                    Assumptions       = ReadJsonStringList(rd.GetString(1)),
                    Warnings          = ReadJsonStringList(rd.GetString(2)),
                    DecisionRationale = rd.GetString(3),
                };
            }
        }

        return bundle;
    }

    private static double? ReadNullableDouble(SqliteDataReader r, int ordinal)
        => r.IsDBNull(ordinal) ? null : r.GetDouble(ordinal);

    private static List<string> ReadJsonStringList(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return [];
        try { return JsonSerializer.Deserialize<List<string>>(json, JsonOpts) ?? []; }
        catch { return []; }
    }

    // ── Dataset summary translations (multi-language) ─────────────────────────

    public void SaveDatasetSummaryTranslation(
        string datasetName, string lang, DatasetSummaryTranslation tr)
    {
        string actionsJson = "";
        try
        {
            if (tr.Actions is { Count: > 0 })
                actionsJson = System.Text.Json.JsonSerializer.Serialize(tr.Actions);
        }
        catch { }

        string contextJson = "";
        try
        {
            if (tr.Context is not null)
                contextJson = System.Text.Json.JsonSerializer.Serialize(tr.Context);
        }
        catch { }

        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DatasetSummaryTranslations
                (DatasetName, Lang, Summary, KeyFindings, Purpose, TestConditions,
                 RootCause, Decision, RecommendedAction,
                 Headline, ActionsJson, ContextJson, UpdatedAt)
            VALUES (@n, @l, @s, @k, @pu, @tc, @rc, @de, @ra,
                    @hl, @ac, @cx, @at)
            ON CONFLICT(DatasetName, Lang) DO UPDATE SET
                Summary=@s, KeyFindings=@k, Purpose=@pu, TestConditions=@tc,
                RootCause=@rc, Decision=@de, RecommendedAction=@ra,
                Headline=@hl, ActionsJson=@ac, ContextJson=@cx, UpdatedAt=@at;
            """;
        cmd.Parameters.AddWithValue("@n",  datasetName);
        cmd.Parameters.AddWithValue("@l",  lang);
        cmd.Parameters.AddWithValue("@s",  tr.Summary           ?? "");
        cmd.Parameters.AddWithValue("@k",  tr.KeyFindings       ?? "");
        cmd.Parameters.AddWithValue("@pu", tr.Purpose           ?? "");
        cmd.Parameters.AddWithValue("@tc", tr.TestConditions    ?? "");
        cmd.Parameters.AddWithValue("@rc", tr.RootCause         ?? "");
        cmd.Parameters.AddWithValue("@de", tr.Decision          ?? "");
        cmd.Parameters.AddWithValue("@ra", tr.RecommendedAction ?? "");
        cmd.Parameters.AddWithValue("@hl", tr.Headline          ?? "");
        cmd.Parameters.AddWithValue("@ac", actionsJson);
        cmd.Parameters.AddWithValue("@cx", contextJson);
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        cmd.ExecuteNonQuery();
    }

    public Dictionary<string, DatasetSummaryTranslation> GetDatasetSummaryTranslations(string datasetName)
    {
        var dict = new Dictionary<string, DatasetSummaryTranslation>(StringComparer.OrdinalIgnoreCase);
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Lang, Summary, KeyFindings, Purpose, TestConditions,
                   RootCause, Decision, RecommendedAction,
                   Headline, ActionsJson, ContextJson
            FROM DatasetSummaryTranslations WHERE DatasetName=@n;
            """;
        cmd.Parameters.AddWithValue("@n", datasetName);
        var jsonOpts = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string headline    = r.IsDBNull(8)  ? "" : r.GetString(8);
            string actionsJson = r.IsDBNull(9)  ? "" : r.GetString(9);
            string contextJson = r.IsDBNull(10) ? "" : r.GetString(10);

            List<ActionItem> actions = [];
            if (!string.IsNullOrWhiteSpace(actionsJson))
            {
                try { actions = System.Text.Json.JsonSerializer.Deserialize<List<ActionItem>>(actionsJson, jsonOpts) ?? []; }
                catch { }
            }
            AnalysisContext? context = null;
            if (!string.IsNullOrWhiteSpace(contextJson))
            {
                try { context = System.Text.Json.JsonSerializer.Deserialize<AnalysisContext>(contextJson, jsonOpts); }
                catch { }
            }

            dict[r.GetString(0)] = new DatasetSummaryTranslation
            {
                Summary           = r.IsDBNull(1) ? "" : r.GetString(1),
                KeyFindings       = r.IsDBNull(2) ? "" : r.GetString(2),
                Purpose           = r.IsDBNull(3) ? "" : r.GetString(3),
                TestConditions    = r.IsDBNull(4) ? "" : r.GetString(4),
                RootCause         = r.IsDBNull(5) ? "" : r.GetString(5),
                Decision          = r.IsDBNull(6) ? "" : r.GetString(6),
                RecommendedAction = r.IsDBNull(7) ? "" : r.GetString(7),
                Headline          = headline,
                Actions           = actions,
                Context           = context,
            };
        }
        return dict;
    }

    // ── AskAi history ─────────────────────────────────────────────────────────

    public long SaveAskAiHistory(string question, string productTypeFilter, string overall, string perDatasetJson)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO AskAiHistory (Question, ProductTypeFilter, Overall, PerDatasetJson, CreatedAt)
            VALUES (@q, @pt, @o, @p, @c);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@q",  question ?? "");
        cmd.Parameters.AddWithValue("@pt", productTypeFilter ?? "");
        cmd.Parameters.AddWithValue("@o",  overall ?? "");
        cmd.Parameters.AddWithValue("@p",  perDatasetJson ?? "[]");
        cmd.Parameters.AddWithValue("@c",  DateTime.UtcNow.ToString("o"));
        return (long)(cmd.ExecuteScalar() ?? 0L);
    }

    public List<AskAiHistoryRecord> GetAskAiHistory(int limit = 100)
    {
        var list = new List<AskAiHistoryRecord>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Question, ProductTypeFilter, Overall, PerDatasetJson, CreatedAt
            FROM AskAiHistory
            ORDER BY Id DESC
            LIMIT @lim;
            """;
        cmd.Parameters.AddWithValue("@lim", limit);
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new AskAiHistoryRecord(
                r.GetInt64(0), r.GetString(1), r.GetString(2),
                r.GetString(3), r.GetString(4), r.GetString(5)));
        }
        return list;
    }

    public AskAiHistoryRecord? GetAskAiHistoryById(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Question, ProductTypeFilter, Overall, PerDatasetJson, CreatedAt
            FROM AskAiHistory WHERE Id=@i;
            """;
        cmd.Parameters.AddWithValue("@i", id);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return null;
        return new AskAiHistoryRecord(
            r.GetInt64(0), r.GetString(1), r.GetString(2),
            r.GetString(3), r.GetString(4), r.GetString(5));
    }

    public void DeleteAskAiHistory(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM AskAiHistory WHERE Id=@i;";
        cmd.Parameters.AddWithValue("@i", id);
        cmd.ExecuteNonQuery();
    }

    public void DeleteAllAskAiHistory()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM AskAiHistory;";
        cmd.ExecuteNonQuery();
    }

    public void DeleteRawReport(string name)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();
        foreach (string table in new[] { "RawReportImages", "RawReportFiles", "NormalizedMeasurements", "DatasetSummary", "RawReports" })
        {
            using var del = conn.CreateCommand();
            del.Transaction = tx;
            del.CommandText = $"DELETE FROM {table} WHERE DatasetName=@n;";
            del.Parameters.AddWithValue("@n", name);
            del.ExecuteNonQuery();
        }
        tx.Commit();
    }

    public void DeleteAllDataInference()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();
        foreach (string table in new[] { "RawReportImages", "RawReportFiles", "NormalizedMeasurements", "DatasetSummary", "RawReports" })
        {
            using var del = conn.CreateCommand();
            del.Transaction = tx;
            del.CommandText = $"DELETE FROM {table};";
            del.ExecuteNonQuery();
        }
        tx.Commit();
    }

    public List<string> GetAllProductTypes()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT DISTINCT ProductType FROM RawReports WHERE ProductType != '' ORDER BY ProductType;";
        var list = new List<string>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    public List<string> GetRawReportDatasets(string? productType = null)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        if (string.IsNullOrEmpty(productType))
        {
            cmd.CommandText = "SELECT DatasetName FROM RawReports ORDER BY CreatedAt DESC;";
        }
        else
        {
            cmd.CommandText = "SELECT DatasetName FROM RawReports WHERE ProductType=@p ORDER BY CreatedAt DESC;";
            cmd.Parameters.AddWithValue("@p", productType);
        }
        var list = new List<string>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    public List<ImprovementRow> GetImprovementComparisons(string? productType = null, string? datasetName = null)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();

        var where = new List<string> { "VariableGroup IN ('normal','test')" };
        if (!string.IsNullOrEmpty(productType)) where.Add("ProductType=@pt");
        if (!string.IsNullOrEmpty(datasetName)) where.Add("DatasetName=@dn");

        cmd.CommandText = $"""
            SELECT DatasetName,
                   MAX(ProductType)                                    AS ProductType,
                   MAX(TestDate)                                       AS TestDate,
                   Line,
                   MAX(CheckType)                                      AS CheckType,
                   VariableDetail,
                   DefectCategory,
                   DefectType,
                   MAX(CASE WHEN VariableGroup='normal' THEN NgRate   END) AS NormalNgRate,
                   MAX(CASE WHEN VariableGroup='test'   THEN NgRate   END) AS TestNgRate,
                   MAX(CASE WHEN VariableGroup='normal' THEN InputQty END) AS NormalInputQty,
                   MAX(CASE WHEN VariableGroup='test'   THEN InputQty END) AS TestInputQty,
                   MAX(CASE WHEN VariableGroup='test'   THEN Intervention END) AS Intervention
            FROM NormalizedMeasurements
            WHERE {string.Join(" AND ", where)}
            GROUP BY DatasetName, Line, VariableDetail, DefectCategory, DefectType
            ORDER BY DatasetName, Line, VariableDetail, DefectCategory, DefectType;
            """;

        if (!string.IsNullOrEmpty(productType)) cmd.Parameters.AddWithValue("@pt", productType);
        if (!string.IsNullOrEmpty(datasetName)) cmd.Parameters.AddWithValue("@dn", datasetName);

        var list = new List<ImprovementRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            double? normalRate = r.IsDBNull(8)  ? null : r.GetDouble(8);
            double? testRate   = r.IsDBNull(9)  ? null : r.GetDouble(9);
            double? impPct     = (normalRate.HasValue && testRate.HasValue && normalRate.Value > 0)
                                  ? (normalRate.Value - testRate.Value) / normalRate.Value * 100.0
                                  : null;
            list.Add(new ImprovementRow(
                DatasetName    : r.GetString(0),
                ProductType    : r.IsDBNull(1)  ? "" : r.GetString(1),
                TestDate       : r.IsDBNull(2)  ? "" : r.GetString(2),
                Line           : r.GetString(3),
                CheckType      : r.IsDBNull(4)  ? "" : r.GetString(4),
                VariableDetail : r.GetString(5),
                DefectCategory : r.GetString(6),
                DefectType     : r.GetString(7),
                NormalNgRate   : normalRate,
                TestNgRate     : testRate,
                ImprovementPct : impPct,
                NormalInputQty : r.IsDBNull(10) ? 0 : r.GetInt32(10),
                TestInputQty   : r.IsDBNull(11) ? 0 : r.GetInt32(11),
                Intervention   : r.IsDBNull(12) ? "" : r.GetString(12)));
        }
        return list;
    }

    public void SaveProductType(string datasetName, string productType)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DatasetMemo (DatasetName, Memo, Purpose, ProductType, UpdatedAt)
            VALUES (@d, '', '', @p, @at)
            ON CONFLICT(DatasetName) DO UPDATE SET ProductType=@p, UpdatedAt=@at;
            """;
        cmd.Parameters.AddWithValue("@d",  datasetName);
        cmd.Parameters.AddWithValue("@p",  productType);
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        cmd.ExecuteNonQuery();
    }

    public string GetProductType(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT ProductType FROM DatasetMemo WHERE DatasetName=@d;";
        cmd.Parameters.AddWithValue("@d", datasetName);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return string.Empty;
        return r.IsDBNull(0) ? string.Empty : r.GetString(0);
    }

    public int RenameTag(string oldTag, string newTag)
    {
        if (string.IsNullOrWhiteSpace(oldTag) || string.IsNullOrWhiteSpace(newTag)) return 0;

        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Tags FROM DatasetTags;";

        var updates = new List<(long Id, string Json)>();
        using (SqliteDataReader r = cmd.ExecuteReader())
        {
            while (r.Read())
            {
                try
                {
                    var tags = JsonSerializer.Deserialize<List<string>>(r.GetString(1)) ?? [];
                    bool changed = false;
                    for (int i = 0; i < tags.Count; i++)
                    {
                        if (string.Equals(tags[i], oldTag, StringComparison.OrdinalIgnoreCase))
                        { tags[i] = newTag; changed = true; }
                    }
                    if (changed) updates.Add((r.GetInt64(0), JsonSerializer.Serialize(tags)));
                }
                catch { }
            }
        }

        using SqliteTransaction tx = conn.BeginTransaction();
        foreach ((long id, string json) in updates)
        {
            using SqliteCommand upd = conn.CreateCommand();
            upd.Transaction = tx;
            upd.CommandText = "UPDATE DatasetTags SET Tags=@t WHERE Id=@id;";
            upd.Parameters.AddWithValue("@t",  json);
            upd.Parameters.AddWithValue("@id", id);
            upd.ExecuteNonQuery();
        }
        tx.Commit();
        return updates.Count;
    }

    // ── Reports ───────────────────────────────────────────────────────────────

    public List<(long Id, string Title, string DatasetNames, string CreatedAt)> GetReports()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Title, DatasetNames, CreatedAt FROM Reports ORDER BY CreatedAt DESC;";
        var list = new List<(long, string, string, string)>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add((r.GetInt64(0), r.GetString(1), r.GetString(2), r.GetString(3)));
        return list;
    }

    public string GetReportHtml(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT HtmlContent FROM Reports WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? r.GetString(0) : string.Empty;
    }

    public long SaveReport(string title, string datasetNames, string htmlContent)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO Reports (Title, DatasetNames, HtmlContent, CreatedAt)
            VALUES (@t, @d, @h, @at);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@t",  title);
        cmd.Parameters.AddWithValue("@d",  datasetNames);
        cmd.Parameters.AddWithValue("@h",  htmlContent);
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        return Convert.ToInt64(cmd.ExecuteScalar() ?? 0);
    }

    public void DeleteReport(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Reports WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    // ── Dataset management ────────────────────────────────────────────────────

    public List<(string Name, int TableCount)> GetDatasetSummary()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT DatasetName, COUNT(*) AS TableCount
            FROM DataTables
            GROUP BY DatasetName
            ORDER BY DatasetName;
            """;
        var list = new List<(string, int)>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add((r.GetString(0), r.GetInt32(1)));
        return list;
    }

    public void DeleteTable(long tableId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();
        using SqliteCommand del1 = conn.CreateCommand();
        del1.Transaction = tx;
        del1.CommandText = "DELETE FROM DataTableRows WHERE TableId=@id;";
        del1.Parameters.AddWithValue("@id", tableId);
        del1.ExecuteNonQuery();

        using SqliteCommand del2 = conn.CreateCommand();
        del2.Transaction = tx;
        del2.CommandText = "DELETE FROM DataTables WHERE Id=@id;";
        del2.Parameters.AddWithValue("@id", tableId);
        del2.ExecuteNonQuery();
        tx.Commit();
    }

    public void RenameTable(long tableId, string newName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE DataTables SET TableName=@n WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@n",  newName);
        cmd.Parameters.AddWithValue("@id", tableId);
        cmd.ExecuteNonQuery();
    }

    public void DeleteTableRow(long tableId, int rowIndex)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id FROM DataTableRows WHERE TableId=@tid ORDER BY Id;";
        cmd.Parameters.AddWithValue("@tid", tableId);
        var ids = new List<long>();
        using (SqliteDataReader r = cmd.ExecuteReader()) while (r.Read()) ids.Add(r.GetInt64(0));
        if (rowIndex < 0 || rowIndex >= ids.Count) return;
        using SqliteCommand del = conn.CreateCommand();
        del.CommandText = "DELETE FROM DataTableRows WHERE Id=@id;";
        del.Parameters.AddWithValue("@id", ids[rowIndex]);
        del.ExecuteNonQuery();
    }

    public void DeleteTableColumn(long tableId, string colField)
    {
        using SqliteConnection conn = OpenConnection();

        // 1. Remove column from schema
        using SqliteCommand getSchema = conn.CreateCommand();
        getSchema.CommandText = "SELECT Columns FROM DataTables WHERE Id=@id;";
        getSchema.Parameters.AddWithValue("@id", tableId);
        string colsJson = (getSchema.ExecuteScalar() as string) ?? "[]";
        var cols = JsonSerializer.Deserialize<List<ColumnDef>>(colsJson, JsonOpts) ?? [];
        cols.RemoveAll(c => c.Field.Equals(colField, StringComparison.OrdinalIgnoreCase));
        using SqliteCommand updSchema = conn.CreateCommand();
        updSchema.CommandText = "UPDATE DataTables SET Columns=@c WHERE Id=@id;";
        updSchema.Parameters.AddWithValue("@c", JsonSerializer.Serialize(cols, JsonOpts));
        updSchema.Parameters.AddWithValue("@id", tableId);
        updSchema.ExecuteNonQuery();

        // 2. Remove field from every row
        using SqliteCommand getRows = conn.CreateCommand();
        getRows.CommandText = "SELECT Id, RowData FROM DataTableRows WHERE TableId=@tid;";
        getRows.Parameters.AddWithValue("@tid", tableId);
        var updates = new List<(long Id, string Json)>();
        using (SqliteDataReader r = getRows.ExecuteReader())
            while (r.Read())
            {
                try
                {
                    var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(r.GetString(1), JsonOpts) ?? [];
                    dict.Remove(colField);
                    updates.Add((r.GetInt64(0), JsonSerializer.Serialize(dict)));
                }
                catch { }
            }
        using SqliteTransaction tx = conn.BeginTransaction();
        foreach ((long id, string json) in updates)
        {
            using SqliteCommand upd = conn.CreateCommand();
            upd.Transaction = tx;
            upd.CommandText = "UPDATE DataTableRows SET RowData=@d WHERE Id=@id;";
            upd.Parameters.AddWithValue("@d", json);
            upd.Parameters.AddWithValue("@id", id);
            upd.ExecuteNonQuery();
        }
        tx.Commit();
    }

    public void UpdateTableRowData(long tableId, int rowIndex, Dictionary<string, string> data)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id FROM DataTableRows WHERE TableId=@tid ORDER BY Id;";
        cmd.Parameters.AddWithValue("@tid", tableId);
        var ids = new List<long>();
        using (SqliteDataReader r = cmd.ExecuteReader()) while (r.Read()) ids.Add(r.GetInt64(0));
        if (rowIndex < 0 || rowIndex >= ids.Count) return;
        using SqliteCommand upd = conn.CreateCommand();
        upd.CommandText = "UPDATE DataTableRows SET RowData=@d WHERE Id=@id;";
        upd.Parameters.AddWithValue("@d", JsonSerializer.Serialize(data));
        upd.Parameters.AddWithValue("@id", ids[rowIndex]);
        upd.ExecuteNonQuery();
    }

    public void DeleteDataset(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        // Get table IDs first
        using SqliteCommand getIds = conn.CreateCommand();
        getIds.Transaction = tx;
        getIds.CommandText = "SELECT Id FROM DataTables WHERE DatasetName=@d;";
        getIds.Parameters.AddWithValue("@d", datasetName);
        var ids = new List<long>();
        using (SqliteDataReader r = getIds.ExecuteReader())
            while (r.Read()) ids.Add(r.GetInt64(0));

        foreach (long id in ids)
        {
            using SqliteCommand dr = conn.CreateCommand();
            dr.Transaction = tx;
            dr.CommandText = "DELETE FROM DataTableRows WHERE TableId=@id;";
            dr.Parameters.AddWithValue("@id", id);
            dr.ExecuteNonQuery();
        }

        using SqliteCommand dt = conn.CreateCommand();
        dt.Transaction = tx;
        dt.CommandText = "DELETE FROM DataTables WHERE DatasetName=@d;";
        dt.Parameters.AddWithValue("@d", datasetName);
        dt.ExecuteNonQuery();

        using SqliteCommand dtag = conn.CreateCommand();
        dtag.Transaction = tx;
        dtag.CommandText = "DELETE FROM DatasetTags WHERE DatasetName=@d;";
        dtag.Parameters.AddWithValue("@d", datasetName);
        dtag.ExecuteNonQuery();

        using SqliteCommand dm = conn.CreateCommand();
        dm.Transaction = tx;
        dm.CommandText = "DELETE FROM DatasetMemo WHERE DatasetName=@d;";
        dm.Parameters.AddWithValue("@d", datasetName);
        dm.ExecuteNonQuery();

        tx.Commit();
    }

    public int GetTotalRowCount()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM DataTableRows;";
        return Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
    }

    // ── Users ─────────────────────────────────────────────────────────────────

    public UserRecord? GetUser(string username)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Username, DisplayName, PasswordHash, Role, CreatedAt FROM Users WHERE Username=@u COLLATE NOCASE LIMIT 1;";
        cmd.Parameters.AddWithValue("@u", username);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return null;
        return new UserRecord
        {
            Id           = r.GetInt64(0),
            Username     = r.GetString(1),
            DisplayName  = r.IsDBNull(2) ? string.Empty : r.GetString(2),
            PasswordHash = r.GetString(3),
            Role         = r.GetString(4),
            CreatedAt    = r.GetString(5)
        };
    }

    public List<UserRecord> GetAllUsers()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Username, DisplayName, PasswordHash, Role, CreatedAt FROM Users ORDER BY Id;";
        var list = new List<UserRecord>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new UserRecord
            {
                Id           = r.GetInt64(0),
                Username     = r.GetString(1),
                DisplayName  = r.IsDBNull(2) ? string.Empty : r.GetString(2),
                PasswordHash = r.GetString(3),
                Role         = r.GetString(4),
                CreatedAt    = r.GetString(5)
            });
        return list;
    }

    public void AddUser(string username, string passwordHash, string role, string displayName = "")
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "INSERT INTO Users (Username, DisplayName, PasswordHash, Role, CreatedAt) VALUES (@u, @dn, @p, @r, @t);";
        cmd.Parameters.AddWithValue("@u",  username);
        cmd.Parameters.AddWithValue("@dn", displayName);
        cmd.Parameters.AddWithValue("@p",  passwordHash);
        cmd.Parameters.AddWithValue("@r",  role);
        cmd.Parameters.AddWithValue("@t",  DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.ExecuteNonQuery();
    }

    public void UpdateUserDisplayName(long id, string displayName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE Users SET DisplayName=@dn WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@dn", displayName);
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public void DeleteUser(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Users WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public void UpdateUserRole(long id, string role)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE Users SET Role=@r WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@r", role);
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public void UpdateUserPassword(long id, string passwordHash)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE Users SET PasswordHash=@p WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@p", passwordHash);
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    // ── Menu permissions ──────────────────────────────────────────────────────

    public HashSet<string> GetMenuPermissionsForRole(string role)
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT MenuId FROM MenuPermissions WHERE Role=@r;";
        cmd.Parameters.AddWithValue("@r", role);
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) set.Add(r.GetString(0));
        return set;
    }

    public Dictionary<string, HashSet<string>> GetAllMenuPermissions()
    {
        var map = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Role, MenuId FROM MenuPermissions;";
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string role   = r.GetString(0);
            string menuId = r.GetString(1);
            if (!map.TryGetValue(role, out var set))
            {
                set = new HashSet<string>(StringComparer.Ordinal);
                map[role] = set;
            }
            set.Add(menuId);
        }
        return map;
    }

    public void SetMenuPermission(string role, string menuId, bool allowed)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        if (allowed)
        {
            cmd.CommandText = "INSERT OR IGNORE INTO MenuPermissions (Role, MenuId) VALUES (@r, @m);";
        }
        else
        {
            cmd.CommandText = "DELETE FROM MenuPermissions WHERE Role=@r AND MenuId=@m;";
        }
        cmd.Parameters.AddWithValue("@r", role);
        cmd.Parameters.AddWithValue("@m", menuId);
        cmd.ExecuteNonQuery();
    }

    public void SeedDefaultMenuPermissionsIfEmpty()
    {
        // Per-role seeding: if a role has NO entries in MenuPermissions, seed it
        // from DefaultsByRole. This preserves admin-customised permissions for
        // existing roles while auto-populating NEW roles added in later releases
        // (e.g., ManagerAi) without manual DB surgery.
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        using SqliteCommand check = conn.CreateCommand();
        check.Transaction = tx;
        check.CommandText = "SELECT COUNT(*) FROM MenuPermissions WHERE Role = @r;";
        var pcr = check.Parameters.Add("@r", SqliteType.Text);

        using SqliteCommand ins = conn.CreateCommand();
        ins.Transaction = tx;
        ins.CommandText = "INSERT OR IGNORE INTO MenuPermissions (Role, MenuId) VALUES (@r, @m);";
        var pr = ins.Parameters.Add("@r", SqliteType.Text);
        var pm = ins.Parameters.Add("@m", SqliteType.Text);

        foreach ((string role, string[] menus) in AppMenus.DefaultsByRole)
        {
            pcr.Value = role;
            long existing = (long)(check.ExecuteScalar() ?? 0L);
            if (existing > 0) continue;

            pr.Value = role;
            foreach (string menu in menus)
            {
                pm.Value = menu;
                ins.ExecuteNonQuery();
            }
        }
        tx.Commit();
    }

    public void ReplaceDatasetEditorImages(string datasetName, List<(string Slug, byte[] Data)> images)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx  = conn.BeginTransaction();
        using (SqliteCommand del = conn.CreateCommand())
        {
            del.Transaction = tx;
            del.CommandText = "DELETE FROM DatasetImages WHERE DatasetName=@d AND FileName LIKE 'di-img-%';";
            del.Parameters.AddWithValue("@d", datasetName);
            del.ExecuteNonQuery();
        }
        string now = DateTime.UtcNow.ToString("O");
        foreach (var (slug, data) in images)
        {
            using SqliteCommand ins = conn.CreateCommand();
            ins.Transaction = tx;
            ins.CommandText = "INSERT INTO DatasetImages (DatasetName, FileName, ImageData, CreatedAt) VALUES (@d, @f, @img, @at);";
            ins.Parameters.AddWithValue("@d",   datasetName);
            ins.Parameters.AddWithValue("@f",   slug);
            ins.Parameters.AddWithValue("@img", data);
            ins.Parameters.AddWithValue("@at",  now);
            ins.ExecuteNonQuery();
        }
        tx.Commit();
    }

    public Dictionary<string, string> GetEditorImageDataUrls(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT FileName, ImageData FROM DatasetImages WHERE DatasetName=@d AND FileName LIKE 'di-img-%' ORDER BY Id;";
        cmd.Parameters.AddWithValue("@d", datasetName);
        var result = new Dictionary<string, string>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string fileName = r.IsDBNull(0) ? "image" : r.GetString(0);
            long   byteLen  = r.GetBytes(1, 0, null, 0, 0);
            byte[] buf      = new byte[byteLen];
            r.GetBytes(1, 0, buf, 0, (int)byteLen);
            string ext = Path.GetExtension(fileName).TrimStart('.').ToLowerInvariant();
            string mediaType = ext switch
            {
                "jpg" or "jpeg" => "image/jpeg",
                "gif"           => "image/gif",
                "webp"          => "image/webp",
                _               => "image/png"
            };
            result[fileName] = $"data:{mediaType};base64,{Convert.ToBase64String(buf)}";
        }
        return result;
    }

    // Returns (FileName, MediaType, Base64) for all images attached to a dataset.
    public List<(string FileName, string MediaType, string Base64)> GetImages(string datasetName)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT FileName, ImageData FROM DatasetImages WHERE DatasetName=@n ORDER BY Id;";
        cmd.Parameters.AddWithValue("@n", datasetName);

        var result = new List<(string, string, string)>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string fileName = r.IsDBNull(0) ? "image" : r.GetString(0);
            long   byteLen  = r.GetBytes(1, 0, null, 0, 0);
            byte[] buf      = new byte[byteLen];
            r.GetBytes(1, 0, buf, 0, (int)byteLen);

            // Infer media type from file extension; default to image/png
            string ext = Path.GetExtension(fileName).TrimStart('.').ToLowerInvariant();
            string mediaType = ext switch
            {
                "jpg" or "jpeg" => "image/jpeg",
                "gif"           => "image/gif",
                "webp"          => "image/webp",
                _               => "image/png"
            };
            result.Add((fileName, mediaType, Convert.ToBase64String(buf)));
        }
        return result;
    }

    // ── Schedule ──────────────────────────────────────────────────────────────

    private static ScheduleItem ReadScheduleItem(SqliteDataReader r)
    {
        string? stStr = r.IsDBNull(5) ? null : r.GetString(5);
        string? etStr = r.IsDBNull(6) ? null : r.GetString(6);
        TimeOnly? st = stStr is not null && TimeOnly.TryParse(stStr, out var tp) ? tp : null;
        TimeOnly? et = etStr is not null && TimeOnly.TryParse(etStr, out var tp2) ? tp2 : null;
        List<string> tags;
        try { tags = JsonSerializer.Deserialize<List<string>>(r.GetString(8), JsonOpts) ?? []; }
        catch { tags = []; }
        return new ScheduleItem(
            r.GetInt64(0),
            r.GetString(1),
            r.GetString(2),
            DateOnly.Parse(r.GetString(3)),
            DateOnly.Parse(r.GetString(4)),
            st, et,
            r.GetString(7),
            tags,
            r.GetString(9));
    }

    public List<ScheduleItem> GetSchedulesInRange(DateOnly from, DateOnly to, List<string>? filterTags = null)
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Title, Description, StartDate, EndDate, StartTime, EndTime, Color, Tags, CreatedAt
            FROM Schedules
            WHERE EndDate >= @from AND StartDate <= @to
            ORDER BY StartDate, StartTime, EndDate;
            """;
        cmd.Parameters.AddWithValue("@from", from.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@to",   to.ToString("yyyy-MM-dd"));

        var items = new List<ScheduleItem>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) items.Add(ReadScheduleItem(r));

        if (filterTags is { Count: > 0 })
            items = items.Where(s => s.Tags.Any(t =>
                filterTags.Contains(t, StringComparer.OrdinalIgnoreCase))).ToList();
        return items;
    }

    public List<string> GetAllScheduleTags()
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Tags FROM Schedules;";
        var set = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            try
            {
                var tags = JsonSerializer.Deserialize<List<string>>(r.GetString(0), JsonOpts) ?? [];
                foreach (string t in tags) set.Add(t.Trim());
            }
            catch { }
        }
        return [.. set];
    }

    public long AddSchedule(string title, string description,
        DateOnly start, DateOnly end, string color, List<string> tags,
        TimeOnly? startTime = null, TimeOnly? endTime = null)
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO Schedules (Title, Description, StartDate, EndDate, StartTime, EndTime, Color, Tags, CreatedAt)
            VALUES (@t, @d, @sd, @ed, @st, @et, @c, @tg, @at);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@t",  title);
        cmd.Parameters.AddWithValue("@d",  description);
        cmd.Parameters.AddWithValue("@sd", start.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@ed", end.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@st", startTime.HasValue ? (object)startTime.Value.ToString("HH:mm") : DBNull.Value);
        cmd.Parameters.AddWithValue("@et", endTime.HasValue   ? (object)endTime.Value.ToString("HH:mm")   : DBNull.Value);
        cmd.Parameters.AddWithValue("@c",  color);
        cmd.Parameters.AddWithValue("@tg", JsonSerializer.Serialize(tags));
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("o"));
        return Convert.ToInt64(cmd.ExecuteScalar());
    }

    public void UpdateSchedule(long id, string title, string description,
        DateOnly start, DateOnly end, string color, List<string> tags,
        TimeOnly? startTime = null, TimeOnly? endTime = null)
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE Schedules SET Title=@t, Description=@d, StartDate=@sd, EndDate=@ed,
                StartTime=@st, EndTime=@et, Color=@c, Tags=@tg
            WHERE Id=@id;
            """;
        cmd.Parameters.AddWithValue("@t",  title);
        cmd.Parameters.AddWithValue("@d",  description);
        cmd.Parameters.AddWithValue("@sd", start.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@ed", end.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@st", startTime.HasValue ? (object)startTime.Value.ToString("HH:mm") : DBNull.Value);
        cmd.Parameters.AddWithValue("@et", endTime.HasValue   ? (object)endTime.Value.ToString("HH:mm")   : DBNull.Value);
        cmd.Parameters.AddWithValue("@c",  color);
        cmd.Parameters.AddWithValue("@tg", JsonSerializer.Serialize(tags));
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public void DeleteSchedule(long id)
    {
        using SqliteConnection conn = OpenScheduleConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Schedules WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    // ── BMES Materials ───────────────────────────────────────────────────────

    public void UpsertBmesMaterial(BmesMaterial m)
        => UpsertBmesMaterials(new[] { m });

    public int UpsertBmesMaterials(IEnumerable<BmesMaterial> materials)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = """
            INSERT INTO BmesMaterials
                (Matnr, Maktx, Meins, Injtp, Mtype, Btype, MngCode, ModNameB,
                 LotQt, Bunch, NgTar,
                 McLv1Tx, McLv2Tx, McLv3Tx, McLv4Tx, McLv5Tx, McLv6Tx,
                 Ernam, Erdat, Grcod, Grnam, MfPhi, FetchedAt)
            VALUES
                (@Matnr, @Maktx, @Meins, @Injtp, @Mtype, @Btype, @MngCode, @ModNameB,
                 @LotQt, @Bunch, @NgTar,
                 @McLv1Tx, @McLv2Tx, @McLv3Tx, @McLv4Tx, @McLv5Tx, @McLv6Tx,
                 @Ernam, @Erdat, @Grcod, @Grnam, @MfPhi, @FetchedAt)
            ON CONFLICT(Matnr) DO UPDATE SET
                Maktx     = excluded.Maktx,
                Meins     = excluded.Meins,
                Injtp     = excluded.Injtp,
                Mtype     = excluded.Mtype,
                Btype     = excluded.Btype,
                MngCode   = excluded.MngCode,
                ModNameB  = excluded.ModNameB,
                LotQt     = excluded.LotQt,
                Bunch     = excluded.Bunch,
                NgTar     = excluded.NgTar,
                McLv1Tx   = excluded.McLv1Tx,
                McLv2Tx   = excluded.McLv2Tx,
                McLv3Tx   = excluded.McLv3Tx,
                McLv4Tx   = excluded.McLv4Tx,
                McLv5Tx   = excluded.McLv5Tx,
                McLv6Tx   = excluded.McLv6Tx,
                Ernam     = excluded.Ernam,
                Erdat     = excluded.Erdat,
                Grcod     = excluded.Grcod,
                Grnam     = excluded.Grnam,
                MfPhi     = excluded.MfPhi,
                FetchedAt = excluded.FetchedAt;
            """;
        SqliteParameter pMatnr    = cmd.Parameters.Add("@Matnr",     SqliteType.Text);
        SqliteParameter pMaktx    = cmd.Parameters.Add("@Maktx",     SqliteType.Text);
        SqliteParameter pMeins    = cmd.Parameters.Add("@Meins",     SqliteType.Text);
        SqliteParameter pInjtp    = cmd.Parameters.Add("@Injtp",     SqliteType.Text);
        SqliteParameter pMtype    = cmd.Parameters.Add("@Mtype",     SqliteType.Text);
        SqliteParameter pBtype    = cmd.Parameters.Add("@Btype",     SqliteType.Text);
        SqliteParameter pMng      = cmd.Parameters.Add("@MngCode",   SqliteType.Text);
        SqliteParameter pMod      = cmd.Parameters.Add("@ModNameB",  SqliteType.Text);
        SqliteParameter pLot      = cmd.Parameters.Add("@LotQt",     SqliteType.Text);
        SqliteParameter pBunch    = cmd.Parameters.Add("@Bunch",     SqliteType.Text);
        SqliteParameter pNgTar    = cmd.Parameters.Add("@NgTar",     SqliteType.Text);
        SqliteParameter pLv1      = cmd.Parameters.Add("@McLv1Tx",   SqliteType.Text);
        SqliteParameter pLv2      = cmd.Parameters.Add("@McLv2Tx",   SqliteType.Text);
        SqliteParameter pLv3      = cmd.Parameters.Add("@McLv3Tx",   SqliteType.Text);
        SqliteParameter pLv4      = cmd.Parameters.Add("@McLv4Tx",   SqliteType.Text);
        SqliteParameter pLv5      = cmd.Parameters.Add("@McLv5Tx",   SqliteType.Text);
        SqliteParameter pLv6      = cmd.Parameters.Add("@McLv6Tx",   SqliteType.Text);
        SqliteParameter pErnam    = cmd.Parameters.Add("@Ernam",     SqliteType.Text);
        SqliteParameter pErdat    = cmd.Parameters.Add("@Erdat",     SqliteType.Text);
        SqliteParameter pGrcod    = cmd.Parameters.Add("@Grcod",     SqliteType.Text);
        SqliteParameter pGrnam    = cmd.Parameters.Add("@Grnam",     SqliteType.Text);
        SqliteParameter pMf       = cmd.Parameters.Add("@MfPhi",     SqliteType.Text);
        SqliteParameter pFetched  = cmd.Parameters.Add("@FetchedAt", SqliteType.Text);

        int n = 0;
        foreach (var m in materials)
        {
            if (string.IsNullOrWhiteSpace(m.Matnr)) continue;
            pMatnr.Value   = m.Matnr;
            pMaktx.Value   = m.Maktx;
            pMeins.Value   = m.Meins;
            pInjtp.Value   = m.Injtp;
            pMtype.Value   = m.Mtype;
            pBtype.Value   = m.Btype;
            pMng.Value     = m.MngCode;
            pMod.Value     = m.ModNameB;
            pLot.Value     = m.LotQt;
            pBunch.Value   = m.Bunch;
            pNgTar.Value   = m.NgTar;
            pLv1.Value     = m.McLv1Tx;
            pLv2.Value     = m.McLv2Tx;
            pLv3.Value     = m.McLv3Tx;
            pLv4.Value     = m.McLv4Tx;
            pLv5.Value     = m.McLv5Tx;
            pLv6.Value     = m.McLv6Tx;
            pErnam.Value   = m.Ernam;
            pErdat.Value   = m.Erdat;
            pGrcod.Value   = m.Grcod;
            pGrnam.Value   = m.Grnam;
            pMf.Value      = m.MfPhi;
            pFetched.Value = m.FetchedAt;
            cmd.ExecuteNonQuery();
            n++;
        }
        tx.Commit();
        return n;
    }

    public int GetBmesMaterialCount()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM BmesMaterials;";
        object? o = cmd.ExecuteScalar();
        return o is null ? 0 : Convert.ToInt32(o);
    }

    // ── Mtype → Category name map ────────────────────────────────────────────

    public Dictionary<string, string> GetMtypeCategoryMap()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Code, Name FROM MtypeCategories;";
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string code = r.IsDBNull(0) ? "" : r.GetString(0);
            string name = r.IsDBNull(1) ? "" : r.GetString(1);
            if (code.Length > 0) map[code] = name;
        }
        return map;
    }

    // ── Model Groups ─────────────────────────────────────────────────────────

    // ── DataInputAliases: user-curated token → Material map ───────────────────

    public Dictionary<string, string> GetDataInputAliases()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Token, Material FROM DataInputAliases;";
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) map[r.GetString(0)] = r.GetString(1);
        return map;
    }

    public void SaveDataInputAlias(string token, string material)
    {
        if (string.IsNullOrWhiteSpace(token) || string.IsNullOrWhiteSpace(material)) return;
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DataInputAliases (Token, Material, CreatedAt, UpdatedAt)
            VALUES (@t, @m, @at, @at)
            ON CONFLICT(Token) DO UPDATE SET Material=@m, UpdatedAt=@at;
            """;
        cmd.Parameters.AddWithValue("@t",  token.Trim());
        cmd.Parameters.AddWithValue("@m",  material.Trim());
        cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
        cmd.ExecuteNonQuery();
    }

    public void DeleteDataInputAlias(string token)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM DataInputAliases WHERE Token=@t;";
        cmd.Parameters.AddWithValue("@t", token);
        cmd.ExecuteNonQuery();
    }

    public List<ModelGroupRecord> GetModelGroups()
    {
        var groups = new List<ModelGroupRecord>();
        using SqliteConnection conn = OpenConnection();

        using (SqliteCommand g = conn.CreateCommand())
        {
            g.CommandText = "SELECT Id, Name, ProductGroup, SortOrder FROM ModelGroups ORDER BY SortOrder, Id;";
            using SqliteDataReader r = g.ExecuteReader();
            while (r.Read())
                groups.Add(new ModelGroupRecord
                {
                    Id           = r.GetInt64(0),
                    Name         = r.IsDBNull(1) ? "" : r.GetString(1),
                    ProductGroup = r.IsDBNull(2) || string.IsNullOrEmpty(r.GetString(2)) ? "ETC" : r.GetString(2),
                    SortOrder    = r.GetInt32(3),
                });
        }

        foreach (var grp in groups)
        {
            using SqliteCommand i = conn.CreateCommand();
            i.CommandText =
                "SELECT LineShift, Material, SubGroup FROM ModelGroupItems " +
                "WHERE GroupId=@gid ORDER BY SortOrder, Id;";
            i.Parameters.AddWithValue("@gid", grp.Id);
            using SqliteDataReader r = i.ExecuteReader();

            // Preserve insertion order; group first by Material, then rebuild the
            // (potentially-nested) sub-group tree from the SubGroup-path column.
            var midIdx = new Dictionary<string, int>(StringComparer.Ordinal);
            while (r.Read())
            {
                string ls       = r.IsDBNull(0) ? "" : r.GetString(0);
                string material = r.IsDBNull(1) ? "" : r.GetString(1);
                string subPath  = r.IsDBNull(2) ? "" : r.GetString(2);

                // Legacy rows without material → derive from LineShift (split at last '_').
                if (string.IsNullOrEmpty(material) && !string.IsNullOrEmpty(ls))
                {
                    int idx = ls.LastIndexOf('_');
                    material = idx > 0 ? ls.Substring(0, idx) : ls;
                }

                if (!midIdx.TryGetValue(material, out int mi))
                {
                    grp.MidGroups.Add(new MidGroupRecord { Material = material });
                    mi = grp.MidGroups.Count - 1;
                    midIdx[material] = mi;
                }

                var mid = grp.MidGroups[mi];
                var target = ResolveOrCreateSubPath(mid.SubGroups, subPath);
                if (!string.IsNullOrEmpty(ls))
                    target.LineShifts.Add(ls);
            }
        }
        return groups;
    }

    /// <summary>Sub-group path separator (unit separator control char — impossible in user input).</summary>
    private const char SubPathSep = '';

    /// <summary>Navigate the sub-group tree by name-path, creating nodes as needed. Returns the leaf.</summary>
    private static SubGroupRecord ResolveOrCreateSubPath(List<SubGroupRecord> rootList, string path)
    {
        // Empty path → the "default" (unnamed) sub-group at the top of this material's tree.
        var segments = string.IsNullOrEmpty(path)
            ? new[] { "" }
            : path.Split(SubPathSep);

        List<SubGroupRecord> list = rootList;
        SubGroupRecord? cur = null;
        foreach (string seg in segments)
        {
            cur = list.FirstOrDefault(s => string.Equals(s.Name, seg, StringComparison.Ordinal));
            if (cur is null)
            {
                cur = new SubGroupRecord { Name = seg };
                list.Add(cur);
            }
            list = cur.SubGroups;
        }
        return cur!;
    }

    /// <summary>Replaces all model groups with the provided list (atomic).</summary>
    public void SaveModelGroups(IReadOnlyList<ModelGroupRecord> groups)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        using (SqliteCommand del = conn.CreateCommand())
        {
            del.Transaction  = tx;
            del.CommandText  = "DELETE FROM ModelGroupItems; DELETE FROM ModelGroups;";
            del.ExecuteNonQuery();
        }

        string nowStr = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        using SqliteCommand gIns = conn.CreateCommand();
        gIns.Transaction = tx;
        gIns.CommandText = """
            INSERT INTO ModelGroups (Name, ProductGroup, SortOrder, UpdatedAt)
            VALUES (@name, @pg, @sort, @ts);
            SELECT last_insert_rowid();
            """;
        SqliteParameter pName = gIns.Parameters.Add("@name", SqliteType.Text);
        SqliteParameter pPg   = gIns.Parameters.Add("@pg",   SqliteType.Text);
        SqliteParameter pSort = gIns.Parameters.Add("@sort", SqliteType.Integer);
        SqliteParameter pTs   = gIns.Parameters.Add("@ts",   SqliteType.Text);

        using SqliteCommand iIns = conn.CreateCommand();
        iIns.Transaction = tx;
        iIns.CommandText = """
            INSERT INTO ModelGroupItems (GroupId, LineShift, Material, SubGroup, SortOrder)
            VALUES (@gid, @ls, @mat, @sub, @sort);
            """;
        SqliteParameter pGid   = iIns.Parameters.Add("@gid",  SqliteType.Integer);
        SqliteParameter pLs    = iIns.Parameters.Add("@ls",   SqliteType.Text);
        SqliteParameter pMat   = iIns.Parameters.Add("@mat",  SqliteType.Text);
        SqliteParameter pSub   = iIns.Parameters.Add("@sub",  SqliteType.Text);
        SqliteParameter pISort = iIns.Parameters.Add("@sort", SqliteType.Integer);

        for (int gi = 0; gi < groups.Count; gi++)
        {
            var grp = groups[gi];
            pName.Value = grp.Name ?? "";
            pPg.Value   = NormalizeProductGroup(grp.ProductGroup);
            pSort.Value = gi;
            pTs.Value   = nowStr;
            long newId  = Convert.ToInt64(gIns.ExecuteScalar());

            int sortCounter = 0;
            void WriteSubTree(MidGroupRecord ownerMid, SubGroupRecord node, string path)
            {
                foreach (var ls in node.LineShifts)
                {
                    pGid.Value   = newId;
                    pLs.Value    = ls ?? "";
                    pMat.Value   = ownerMid.Material ?? "";
                    pSub.Value   = path;
                    pISort.Value = sortCounter++;
                    iIns.ExecuteNonQuery();
                }
                foreach (var child in node.SubGroups)
                {
                    string childPath = string.IsNullOrEmpty(path)
                        ? (child.Name ?? "")
                        : path + SubPathSep + (child.Name ?? "");
                    WriteSubTree(ownerMid, child, childPath);
                }
            }

            foreach (var mid in grp.MidGroups)
            {
                foreach (var sub in mid.SubGroups)
                    WriteSubTree(mid, sub, path: sub.Name ?? "");
            }
        }

        tx.Commit();
    }

    private static readonly string[] _knownProductGroups = ["SPK", "UNIT", "MODULE", "TWS", "ETC"];

    private static string NormalizeProductGroup(string? pg)
    {
        if (string.IsNullOrWhiteSpace(pg)) return "ETC";
        foreach (string k in _knownProductGroups)
            if (string.Equals(k, pg, StringComparison.OrdinalIgnoreCase)) return k;
        return "ETC";
    }

    // ── One-shot JSON → ModelGroups import ──────────────────────────────────────

    /// <summary>
    /// One-time migration that merges the legacy ModelBmes/*.json definitions into the
    /// ModelGroups DB. Tracked in AppMigrations so it runs at most once.
    /// - Existing group with the same Name: ProductGroup is set (if still default 'ETC')
    ///   and missing LineShifts are appended under Material=''.
    /// - Otherwise: a new group is inserted with Material='' for all LineShifts.
    /// Safe to call repeatedly.
    /// </summary>
    public void ImportModelBmesJsonIfNeeded(string jsonFolderPath)
    {
        const string MigrationName = "import_modelbmes_json_v1";

        if (!Directory.Exists(jsonFolderPath)) return;

        using SqliteConnection conn = OpenConnection();

        using (SqliteCommand chk = conn.CreateCommand())
        {
            chk.CommandText = "SELECT 1 FROM AppMigrations WHERE Name=@n LIMIT 1;";
            chk.Parameters.AddWithValue("@n", MigrationName);
            if (chk.ExecuteScalar() is not null) return;
        }

        var jsonFiles = Directory.GetFiles(jsonFolderPath, "*.json", SearchOption.TopDirectoryOnly);
        if (jsonFiles.Length == 0)
        {
            RecordMigration(conn, MigrationName);
            return;
        }

        var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

        using SqliteTransaction tx = conn.BeginTransaction();
        string nowStr = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        foreach (string path in jsonFiles)
        {
            string productGroup;
            List<(string GroupName, List<string> ModelList)> groups;
            try
            {
                string text = File.ReadAllText(path);
                using JsonDocument doc = JsonDocument.Parse(text);
                var root = doc.RootElement;

                productGroup = "ETC";
                if (root.ValueKind == JsonValueKind.Object &&
                    root.TryGetProperty("ProductGroup", out var pgEl) &&
                    pgEl.ValueKind == JsonValueKind.String)
                {
                    productGroup = NormalizeProductGroup(pgEl.GetString());
                }

                groups = new();
                JsonElement? groupsEl = null;
                if (root.ValueKind == JsonValueKind.Object &&
                    root.TryGetProperty("Groups", out var gEl))
                {
                    groupsEl = gEl;
                }
                else if (root.ValueKind == JsonValueKind.Array)
                {
                    groupsEl = root;
                }

                if (groupsEl is null) continue;
                foreach (var g in groupsEl.Value.EnumerateArray())
                {
                    string name = "";
                    if (g.TryGetProperty("GroupName", out var n) && n.ValueKind == JsonValueKind.String)
                        name = n.GetString() ?? "";
                    var list = new List<string>();
                    if (g.TryGetProperty("ModelList", out var ml) && ml.ValueKind == JsonValueKind.Array)
                        foreach (var it in ml.EnumerateArray())
                            if (it.ValueKind == JsonValueKind.String)
                            {
                                string? s = it.GetString();
                                if (!string.IsNullOrWhiteSpace(s)) list.Add(s.Trim());
                            }
                    if (!string.IsNullOrWhiteSpace(name) && list.Count > 0)
                        groups.Add((name.Trim(), list));
                }
            }
            catch
            {
                continue;
            }

            foreach (var (gName, modelList) in groups)
            {
                long existingId = -1;
                string existingPg = "";
                using (SqliteCommand find = conn.CreateCommand())
                {
                    find.Transaction = tx;
                    find.CommandText = "SELECT Id, ProductGroup FROM ModelGroups WHERE Name=@n LIMIT 1;";
                    find.Parameters.AddWithValue("@n", gName);
                    using SqliteDataReader r = find.ExecuteReader();
                    if (r.Read())
                    {
                        existingId = r.GetInt64(0);
                        existingPg = r.IsDBNull(1) ? "" : r.GetString(1);
                    }
                }

                if (existingId < 0)
                {
                    long newId;
                    using (SqliteCommand ins = conn.CreateCommand())
                    {
                        ins.Transaction = tx;
                        ins.CommandText = """
                            INSERT INTO ModelGroups (Name, ProductGroup, SortOrder, UpdatedAt)
                            VALUES (@n, @pg, (SELECT COALESCE(MAX(SortOrder)+1, 0) FROM ModelGroups), @ts);
                            SELECT last_insert_rowid();
                            """;
                        ins.Parameters.AddWithValue("@n",  gName);
                        ins.Parameters.AddWithValue("@pg", productGroup);
                        ins.Parameters.AddWithValue("@ts", nowStr);
                        newId = Convert.ToInt64(ins.ExecuteScalar());
                    }
                    InsertLineShiftsForImport(conn, tx, newId, modelList);
                }
                else
                {
                    // If existing row still has default PG, backfill from JSON.
                    if (string.Equals(existingPg, "ETC", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(productGroup, "ETC", StringComparison.OrdinalIgnoreCase))
                    {
                        using SqliteCommand upd = conn.CreateCommand();
                        upd.Transaction = tx;
                        upd.CommandText = "UPDATE ModelGroups SET ProductGroup=@pg WHERE Id=@id;";
                        upd.Parameters.AddWithValue("@pg", productGroup);
                        upd.Parameters.AddWithValue("@id", existingId);
                        upd.ExecuteNonQuery();
                    }

                    var existingLs = new HashSet<string>(StringComparer.Ordinal);
                    using (SqliteCommand q = conn.CreateCommand())
                    {
                        q.Transaction = tx;
                        q.CommandText = "SELECT LineShift FROM ModelGroupItems WHERE GroupId=@id;";
                        q.Parameters.AddWithValue("@id", existingId);
                        using SqliteDataReader r = q.ExecuteReader();
                        while (r.Read())
                            if (!r.IsDBNull(0)) existingLs.Add(r.GetString(0));
                    }

                    var missing = modelList.Where(ls => !existingLs.Contains(ls)).ToList();
                    if (missing.Count > 0)
                        InsertLineShiftsForImport(conn, tx, existingId, missing);
                }
            }
        }

        RecordMigration(conn, MigrationName, tx);
        tx.Commit();
    }

    private static void InsertLineShiftsForImport(
        SqliteConnection conn, SqliteTransaction tx, long groupId, IEnumerable<string> lineShifts)
    {
        using SqliteCommand q = conn.CreateCommand();
        q.Transaction = tx;
        q.CommandText = "SELECT COALESCE(MAX(SortOrder)+1, 0) FROM ModelGroupItems WHERE GroupId=@g;";
        q.Parameters.AddWithValue("@g", groupId);
        int sortStart = Convert.ToInt32(q.ExecuteScalar());

        using SqliteCommand ins = conn.CreateCommand();
        ins.Transaction = tx;
        ins.CommandText = """
            INSERT INTO ModelGroupItems (GroupId, LineShift, Material, SubGroup, SortOrder)
            VALUES (@g, @ls, '', '', @s);
            """;
        SqliteParameter pG = ins.Parameters.Add("@g",  SqliteType.Integer);
        SqliteParameter pL = ins.Parameters.Add("@ls", SqliteType.Text);
        SqliteParameter pS = ins.Parameters.Add("@s",  SqliteType.Integer);
        pG.Value = groupId;
        int s = sortStart;
        foreach (string ls in lineShifts)
        {
            pL.Value = ls ?? "";
            pS.Value = s++;
            ins.ExecuteNonQuery();
        }
    }

    private static void RecordMigration(SqliteConnection conn, string name, SqliteTransaction? tx = null)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        if (tx is not null) cmd.Transaction = tx;
        cmd.CommandText = "INSERT OR IGNORE INTO AppMigrations (Name, AppliedAt) VALUES (@n, @ts);";
        cmd.Parameters.AddWithValue("@n",  name);
        cmd.Parameters.AddWithValue("@ts", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.ExecuteNonQuery();
    }

    /// <summary>Returns distinct (Maktx, Mtype) pairs from BmesMaterials, ordered by Maktx.</summary>
    public List<(string Maktx, string Mtype)> GetBmesMaktxMtypeDistinct()
    {
        var list = new List<(string Maktx, string Mtype)>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT DISTINCT Maktx, Mtype
            FROM BmesMaterials
            ORDER BY Maktx, Mtype;
            """;
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string maktx = r.IsDBNull(0) ? string.Empty : r.GetString(0);
            string mtype = r.IsDBNull(1) ? string.Empty : r.GetString(1);
            list.Add((maktx, mtype));
        }
        return list;
    }

    public BmesMaterial? GetLatestBmesMaterial()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT * FROM BmesMaterials ORDER BY FetchedAt DESC LIMIT 1;";
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return null;
        return ReadBmesMaterial(r);
    }

    private static BmesMaterial ReadBmesMaterial(SqliteDataReader r)
    {
        string G(string col)
        {
            int ord = r.GetOrdinal(col);
            return r.IsDBNull(ord) ? string.Empty : r.GetString(ord);
        }
        return new BmesMaterial
        {
            Matnr     = G("Matnr"),
            Maktx     = G("Maktx"),
            Meins     = G("Meins"),
            Injtp     = G("Injtp"),
            Mtype     = G("Mtype"),
            Btype     = G("Btype"),
            MngCode   = G("MngCode"),
            ModNameB  = G("ModNameB"),
            LotQt     = G("LotQt"),
            Bunch     = G("Bunch"),
            NgTar     = G("NgTar"),
            McLv1Tx   = G("McLv1Tx"),
            McLv2Tx   = G("McLv2Tx"),
            McLv3Tx   = G("McLv3Tx"),
            McLv4Tx   = G("McLv4Tx"),
            McLv5Tx   = G("McLv5Tx"),
            McLv6Tx   = G("McLv6Tx"),
            Ernam     = G("Ernam"),
            Erdat     = G("Erdat"),
            Grcod     = G("Grcod"),
            Grnam     = G("Grnam"),
            MfPhi     = G("MfPhi"),
            FetchedAt = G("FetchedAt"),
        };
    }
}
