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

public sealed class WebRepository
{
    private static readonly JsonSerializerOptions JsonOpts = new() { PropertyNameCaseInsensitive = true };

    private readonly string _dbPath;

    public WebRepository(IConfiguration config)
    {
        // Priority 1: appsettings.json ??Database:Path (explicit override)
        // Priority 2: WPF app settings file (DataInference.DatabasePath)
        // Priority 3: default path under AppStoragePaths.RootDirectory.
        string? configured = config["Database:Path"];
        _dbPath = !string.IsNullOrWhiteSpace(configured)
            ? configured
            : WpfSettingsReader.TryGetDatabasePath()
              ?? AppStoragePaths.Combine("process-review.db");

        EnsureDatabase();
    }

    public string GetDbPath() => _dbPath;

    public CurrentProblemDbStatus GetCurrentProblemDbStatus()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT
                (SELECT COUNT(*)
                   FROM RawReports
                  WHERE COALESCE(BatchExcluded, 0) = 0),
                (SELECT COUNT(*)
                   FROM DatasetSummary
                  WHERE ReportType = 'first_pass_index'),
                (SELECT COUNT(*)
                   FROM AiDocuments);
            """;

        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return new CurrentProblemDbStatus(0, 0, 0);

        return new CurrentProblemDbStatus(
            Convert.ToInt32(r.GetValue(0)),
            Convert.ToInt32(r.GetValue(1)),
            Convert.ToInt32(r.GetValue(2)));
    }

    // ?? Connection ????????????????????????????????????????????????????????????

    private SqliteConnection OpenConnection()
    {
        var conn = new SqliteConnection($"Data Source={_dbPath}");
        conn.Open();
        return conn;
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

            CREATE TABLE IF NOT EXISTS BmesLpaImages (
                Path      TEXT NOT NULL,
                Kind      TEXT NOT NULL,
                ImageData BLOB NOT NULL,
                CreatedAt TEXT NOT NULL,
                PRIMARY KEY (Path, Kind)
            );

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

            CREATE TABLE IF NOT EXISTS InputDataComWorkbooks (
                Id             INTEGER PRIMARY KEY AUTOINCREMENT,
                DatasetName    TEXT NOT NULL,
                SourcePath     TEXT NOT NULL,
                SourceFileName TEXT NOT NULL DEFAULT '',
                FileSize       INTEGER NOT NULL DEFAULT 0,
                MtimeNs        INTEGER NOT NULL DEFAULT 0,
                Fingerprint    TEXT NOT NULL DEFAULT '',
                Status         TEXT NOT NULL DEFAULT 'OK',
                Error          TEXT NOT NULL DEFAULT '',
                SheetCount     INTEGER NOT NULL DEFAULT 0,
                TotalRows      INTEGER NOT NULL DEFAULT 0,
                TotalCells     INTEGER NOT NULL DEFAULT 0,
                NonEmptyCells  INTEGER NOT NULL DEFAULT 0,
                MergeCount     INTEGER NOT NULL DEFAULT 0,
                RawJsonPath    TEXT NOT NULL DEFAULT '',
                ExtractedAt    TEXT NOT NULL,
                CreatedAt      TEXT NOT NULL,
                UNIQUE(DatasetName, SourcePath)
            );
            CREATE INDEX IF NOT EXISTS idx_idcw_dataset ON InputDataComWorkbooks(DatasetName);
            CREATE INDEX IF NOT EXISTS idx_idcw_created ON InputDataComWorkbooks(CreatedAt DESC);

            CREATE TABLE IF NOT EXISTS InputDataComSheets (
                Id             INTEGER PRIMARY KEY AUTOINCREMENT,
                WorkbookId     INTEGER NOT NULL REFERENCES InputDataComWorkbooks(Id) ON DELETE CASCADE,
                SheetIndex     INTEGER NOT NULL DEFAULT 0,
                SheetName      TEXT NOT NULL DEFAULT '',
                UsedTop        INTEGER NOT NULL DEFAULT 0,
                UsedLeft       INTEGER NOT NULL DEFAULT 0,
                UsedBottom     INTEGER NOT NULL DEFAULT 0,
                UsedRight      INTEGER NOT NULL DEFAULT 0,
                RowCount       INTEGER NOT NULL DEFAULT 0,
                ColumnCount    INTEGER NOT NULL DEFAULT 0,
                NonEmptyCells  INTEGER NOT NULL DEFAULT 0,
                MergeCount     INTEGER NOT NULL DEFAULT 0
            );
            CREATE INDEX IF NOT EXISTS idx_idcs_workbook ON InputDataComSheets(WorkbookId);

            CREATE TABLE IF NOT EXISTS InputDataComCells (
                WorkbookId     INTEGER NOT NULL REFERENCES InputDataComWorkbooks(Id) ON DELETE CASCADE,
                SheetName      TEXT NOT NULL,
                RowNumber      INTEGER NOT NULL,
                ColNumber      INTEGER NOT NULL,
                ColLabel       TEXT NOT NULL DEFAULT '',
                CellAddress    TEXT NOT NULL DEFAULT '',
                CellValue      TEXT NOT NULL DEFAULT '',
                RawValue       TEXT NOT NULL DEFAULT '',
                MergeRole      TEXT NOT NULL DEFAULT 'none',
                MergeAddress   TEXT NOT NULL DEFAULT '',
                MergeAnchorRow INTEGER,
                MergeAnchorCol INTEGER,
                PRIMARY KEY(WorkbookId, SheetName, RowNumber, ColNumber)
            );
            CREATE INDEX IF NOT EXISTS idx_idcc_lookup ON InputDataComCells(WorkbookId, SheetName, RowNumber);

            CREATE TABLE IF NOT EXISTS InputDataComMerges (
                Id             INTEGER PRIMARY KEY AUTOINCREMENT,
                WorkbookId     INTEGER NOT NULL REFERENCES InputDataComWorkbooks(Id) ON DELETE CASCADE,
                SheetName      TEXT NOT NULL DEFAULT '',
                Address        TEXT NOT NULL DEFAULT '',
                TopRow         INTEGER NOT NULL DEFAULT 0,
                LeftCol        INTEGER NOT NULL DEFAULT 0,
                BottomRow      INTEGER NOT NULL DEFAULT 0,
                RightCol       INTEGER NOT NULL DEFAULT 0,
                RowSpan        INTEGER NOT NULL DEFAULT 0,
                ColumnSpan     INTEGER NOT NULL DEFAULT 0,
                AnchorValue    TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_idcm_workbook ON InputDataComMerges(WorkbookId);

            CREATE TABLE IF NOT EXISTS InputDataReviewCandidates (
                Id                INTEGER PRIMARY KEY AUTOINCREMENT,
                WorkbookId        INTEGER NOT NULL REFERENCES InputDataComWorkbooks(Id) ON DELETE CASCADE,
                CandidateKind     TEXT NOT NULL DEFAULT '',
                SheetName         TEXT NOT NULL DEFAULT '',
                RowNumber         INTEGER NOT NULL DEFAULT 0,
                Label             TEXT NOT NULL DEFAULT '',
                Confidence        TEXT NOT NULL DEFAULT '',
                EvidenceCellsJson TEXT NOT NULL DEFAULT '[]',
                RawText           TEXT NOT NULL DEFAULT '',
                CreatedAt         TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_idrc_workbook ON InputDataReviewCandidates(WorkbookId);
            CREATE INDEX IF NOT EXISTS idx_idrc_kind ON InputDataReviewCandidates(CandidateKind);

            CREATE TABLE IF NOT EXISTS InputDataReviewCases (
                Id               INTEGER PRIMARY KEY AUTOINCREMENT,
                WorkbookId       INTEGER NOT NULL REFERENCES InputDataComWorkbooks(Id) ON DELETE CASCADE,
                ReviewCaseId     TEXT NOT NULL DEFAULT '',
                Status           TEXT NOT NULL DEFAULT '',
                ApprovedForAskAi INTEGER NOT NULL DEFAULT 0,
                GenerationJson   TEXT NOT NULL DEFAULT '{}',
                VerificationJson TEXT NOT NULL DEFAULT '{}',
                UserAnswerJson   TEXT NOT NULL DEFAULT '{}',
                CreatedAt        TEXT NOT NULL,
                VerifiedAt       TEXT NOT NULL DEFAULT '',
                UserReviewedAt   TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_idrv_workbook ON InputDataReviewCases(WorkbookId);
            CREATE INDEX IF NOT EXISTS idx_idrv_status ON InputDataReviewCases(Status);

            CREATE TABLE IF NOT EXISTS AskAiHistory (
                Id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                Question           TEXT    NOT NULL,
                ProductTypeFilter  TEXT    NOT NULL DEFAULT '',
                Overall            TEXT    NOT NULL DEFAULT '',
                PerDatasetJson     TEXT    NOT NULL DEFAULT '[]',
                TranslationsJson   TEXT    NOT NULL DEFAULT '{}',
                CreatedAt          TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_askai_created ON AskAiHistory(CreatedAt DESC);

            CREATE TABLE IF NOT EXISTS AiModelAnalyses (
                Id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                ProductType          TEXT    NOT NULL DEFAULT '',
                AnalysisMode         TEXT    NOT NULL DEFAULT '',
                Language             TEXT    NOT NULL DEFAULT '',
                ReportCount          INTEGER NOT NULL DEFAULT 0,
                IncludedDatasetsJson TEXT    NOT NULL DEFAULT '[]',
                AnalysisMarkdown     TEXT    NOT NULL DEFAULT '',
                AnalysisTableMarkdown TEXT   NOT NULL DEFAULT '',
                SourceContextHash    TEXT    NOT NULL DEFAULT '',
                CreatedAt            TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_aimodel_product_created ON AiModelAnalyses(ProductType, CreatedAt DESC);

            CREATE TABLE IF NOT EXISTS DailyTestDataItems (
                Id               INTEGER PRIMARY KEY AUTOINCREMENT,
                Name             TEXT    NOT NULL UNIQUE COLLATE NOCASE,
                DataText         TEXT    NOT NULL DEFAULT '',
                PromptText       TEXT    NOT NULL DEFAULT '',
                ParametersJson   TEXT    NOT NULL DEFAULT '{}',
                AnalysisMarkdown TEXT    NOT NULL DEFAULT '',
                AnalysisHtml     TEXT    NOT NULL DEFAULT '',
                CreatedAt        TEXT    NOT NULL,
                UpdatedAt        TEXT    NOT NULL,
                AnalyzedAt       TEXT    NOT NULL DEFAULT '',
                HtmlGeneratedAt  TEXT    NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_dailytest_updated ON DailyTestDataItems(UpdatedAt DESC);

            CREATE TABLE IF NOT EXISTS DailyTestDataHistory (
                Id               INTEGER PRIMARY KEY AUTOINCREMENT,
                ItemId           INTEGER NOT NULL,
                ItemName         TEXT    NOT NULL DEFAULT '',
                DataText         TEXT    NOT NULL DEFAULT '',
                PromptText       TEXT    NOT NULL DEFAULT '',
                ParametersJson   TEXT    NOT NULL DEFAULT '{}',
                AnalysisMarkdown TEXT    NOT NULL DEFAULT '',
                AnalysisHtml     TEXT    NOT NULL DEFAULT '',
                CreatedAt        TEXT    NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_dailytest_history_item_created ON DailyTestDataHistory(ItemId, CreatedAt DESC);

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

            CREATE TABLE IF NOT EXISTS WeeklyReportFormSettings (
                RowKey          TEXT    PRIMARY KEY NOT NULL,
                IsVisible       INTEGER NOT NULL DEFAULT 1,
                DisplayMode     TEXT    NOT NULL DEFAULT 'B-GROUP',
                ProductName     TEXT    NOT NULL DEFAULT '',
                BaselineDec2025 REAL    NULL,
                BaselineApr2026 REAL    NULL,
                Target          REAL    NULL,
                Action          TEXT    NOT NULL DEFAULT '',
                SortOrder       INTEGER NOT NULL DEFAULT -1,
                UpdatedAt       TEXT    NOT NULL
            );

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

            -- ?? AI_EXCEL_PROC schema (Batch AI inference output) ?????????????
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
                GeneratedReportMarkdown TEXT NOT NULL DEFAULT '',
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

            -- ?? ko/en/vi translations for narrative fields ??????????????????????
            -- Agent writes one row per (record, lang) ??Lang in ('ko','en','vi').
            -- Numerics + measurements live in the base tables; only the
            -- AI-narrative text needs translating.
            CREATE TABLE IF NOT EXISTS AiDocumentTranslations (
                DocumentId  TEXT NOT NULL REFERENCES AiDocuments(DocumentId) ON DELETE CASCADE,
                Lang        TEXT NOT NULL,           -- 'ko' | 'en' | 'vi'
                Title       TEXT NOT NULL DEFAULT '',
                Purpose     TEXT NOT NULL DEFAULT '',
                ContentJson TEXT NOT NULL DEFAULT '',-- translated content[] array
                GeneratedReportMarkdown TEXT NOT NULL DEFAULT '',
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
    /// Idempotent seed of the Mtype ??Category-name mapping. Uses INSERT OR IGNORE so
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
        var inputReviewCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (SqliteCommand ck = conn.CreateCommand())
        {
            ck.CommandText = "PRAGMA table_info(InputDataReviewCases);";
            using SqliteDataReader r = ck.ExecuteReader();
            while (r.Read()) inputReviewCols.Add(r.GetString(1));
        }
        if (!inputReviewCols.Contains("UserAnswerJson"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE InputDataReviewCases ADD COLUMN UserAnswerJson TEXT NOT NULL DEFAULT '{}';";
            alter.ExecuteNonQuery();
        }
        if (!inputReviewCols.Contains("UserReviewedAt"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE InputDataReviewCases ADD COLUMN UserReviewedAt TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }
        // ModelGroupItems.Material (for 以묎렇猷?Material)
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

        // ModelGroups.ProductGroup (SPK/UNIT/MODULE/TWS/ETC ??replaces per-JSON-file attribute)
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

        var aiDocCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (SqliteCommand ck = conn.CreateCommand())
        {
            ck.CommandText = "PRAGMA table_info(AiDocuments);";
            using SqliteDataReader r = ck.ExecuteReader();
            while (r.Read()) aiDocCols.Add(r.GetString(1));
        }
        if (!aiDocCols.Contains("GeneratedReportMarkdown"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE AiDocuments ADD COLUMN GeneratedReportMarkdown TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        var aiDocTrCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (SqliteCommand ck = conn.CreateCommand())
        {
            ck.CommandText = "PRAGMA table_info(AiDocumentTranslations);";
            using SqliteDataReader r = ck.ExecuteReader();
            while (r.Read()) aiDocTrCols.Add(r.GetString(1));
        }
        if (!aiDocTrCols.Contains("GeneratedReportMarkdown"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE AiDocumentTranslations ADD COLUMN GeneratedReportMarkdown TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        var aiModelCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (SqliteCommand ck = conn.CreateCommand())
        {
            ck.CommandText = "PRAGMA table_info(AiModelAnalyses);";
            using SqliteDataReader r = ck.ExecuteReader();
            while (r.Read()) aiModelCols.Add(r.GetString(1));
        }
        if (!aiModelCols.Contains("AnalysisTableMarkdown"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE AiModelAnalyses ADD COLUMN AnalysisTableMarkdown TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }

        var askAiCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (SqliteCommand ck = conn.CreateCommand())
        {
            ck.CommandText = "PRAGMA table_info(AskAiHistory);";
            using SqliteDataReader r = ck.ExecuteReader();
            while (r.Read()) askAiCols.Add(r.GetString(1));
        }
        if (!askAiCols.Contains("TranslationsJson"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE AskAiHistory ADD COLUMN TranslationsJson TEXT NOT NULL DEFAULT '{}';";
            alter.ExecuteNonQuery();
        }

        var dailyCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (SqliteCommand ck = conn.CreateCommand())
        {
            ck.CommandText = "PRAGMA table_info(DailyTestDataItems);";
            using SqliteDataReader r = ck.ExecuteReader();
            while (r.Read()) dailyCols.Add(r.GetString(1));
        }
        if (!dailyCols.Contains("ParametersJson"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE DailyTestDataItems ADD COLUMN ParametersJson TEXT NOT NULL DEFAULT '{}';";
            alter.ExecuteNonQuery();
        }
        if (!dailyCols.Contains("AnalysisMarkdown"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE DailyTestDataItems ADD COLUMN AnalysisMarkdown TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }
        if (!dailyCols.Contains("AnalysisHtml"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE DailyTestDataItems ADD COLUMN AnalysisHtml TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }
        if (!dailyCols.Contains("AnalyzedAt"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE DailyTestDataItems ADD COLUMN AnalyzedAt TEXT NOT NULL DEFAULT '';";
            alter.ExecuteNonQuery();
        }
        if (!dailyCols.Contains("HtmlGeneratedAt"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE DailyTestDataItems ADD COLUMN HtmlGeneratedAt TEXT NOT NULL DEFAULT '';";
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

    // ?? Raw report text ???????????????????????????????????????????????????????
    // Two kinds live side-by-side per dataset:
    //   "ocr"         ??structured markdown transcript produced by Vision OCR.
    //                   Batch normalize-from-text flow consumes this.
    //   "excel_paste" ??raw tab-separated text pasted from Excel at Input time.
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

    // ?? App Settings ??????????????????????????????????????????????????????????

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

    // ?? Datasets ??????????????????????????????????????????????????????????????

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

    // ?? Tables ????????????????????????????????????????????????????????????????

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

    // ?? Tags & Memo ???????????????????????????????????????????????????????????

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

    // ?? RawReports ????????????????????????????????????????????????????????????

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

    // ── BMES LPA result photos ────────────────────────────────────────────────────
    // The LPA viewer no longer bakes photos into its HTML; each <img> points at
    // /bmes/lpa/img, which downscales the BMES original once and stores the two sizes it
    // needs here — so a photo is fetched from BMES exactly once and reused across every
    // view, search and restart afterwards. Kind is 't' (thumbnail) or 'v' (popup view).

    public byte[]? GetLpaImage(string path, bool view)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT ImageData FROM BmesLpaImages WHERE Path=@p AND Kind=@k;";
        cmd.Parameters.AddWithValue("@p", path);
        cmd.Parameters.AddWithValue("@k", view ? "v" : "t");
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return null;
        long len = r.GetBytes(0, 0, null, 0, 0);
        var buf = new byte[len];
        r.GetBytes(0, 0, buf, 0, (int)len);
        return buf;
    }

    public void SaveLpaImagePair(string path, byte[] thumb, byte[] view)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        void Insert(string kind, byte[] data)
        {
            using SqliteCommand cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = """
                INSERT INTO BmesLpaImages (Path, Kind, ImageData, CreatedAt)
                VALUES (@p, @k, @d, @at)
                ON CONFLICT(Path, Kind) DO UPDATE SET ImageData=@d, CreatedAt=@at;
                """;
            cmd.Parameters.AddWithValue("@p",  path);
            cmd.Parameters.AddWithValue("@k",  kind);
            cmd.Parameters.AddWithValue("@d",  data);
            cmd.Parameters.AddWithValue("@at", DateTime.UtcNow.ToString("O"));
            cmd.ExecuteNonQuery();
        }

        Insert("t", thumb);
        Insert("v", view);
        tx.Commit();
    }

    public List<RawReportInfo> GetAllRawReports()
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        // BatchedAt is a derived value. Prefer the newest of:
        //   AiDocuments.UpdatedAt   (AI_EXCEL_PROC schema ??new CLI flow)
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
                     (SELECT MAX(d.UpdatedAt)
                        FROM AiDocuments d
                       WHERE d.SourceDataset=r.DatasetName
                         AND LOWER(COALESCE(d.PrimaryDefect,'')) NOT LIKE '%auto-extracted%'
                         AND LOWER(COALESCE(d.RawJson,'')) NOT LIKE '%_batch_auto.py%'
                         AND LOWER(COALESCE(d.RawJson,'')) NOT LIKE '%see workbook title/purpose%'
                         AND LOWER(COALESCE(d.RawJson,'')) NOT LIKE '%workbook stored but extraction surfaced narrative only%'),
                     (SELECT CreatedAt      FROM DatasetSummary   WHERE DatasetName  =r.DatasetName),
                     (SELECT MAX(CreatedAt) FROM NormalizedMeasurements WHERE DatasetName=r.DatasetName),
                     ''
                   ) AS BatchedAt,
                   COALESCE(
                     (SELECT CASE
                               WHEN d.SchemaVersion='input-data-test-batch-v1'
                                 OR d.RawJson LIKE '%INPUT_DATA_BATCH_FROM_INPUT_DATA_TEST%'
                               THEN 'new'
                               ELSE 'old'
                             END
                        FROM AiDocuments d
                       WHERE d.SourceDataset=r.DatasetName
                         AND LOWER(COALESCE(d.PrimaryDefect,'')) NOT LIKE '%auto-extracted%'
                         AND LOWER(COALESCE(d.RawJson,'')) NOT LIKE '%_batch_auto.py%'
                         AND LOWER(COALESCE(d.RawJson,'')) NOT LIKE '%see workbook title/purpose%'
                         AND LOWER(COALESCE(d.RawJson,'')) NOT LIKE '%workbook stored but extraction surfaced narrative only%'
                       ORDER BY d.UpdatedAt DESC LIMIT 1),
                     CASE
                       WHEN EXISTS(SELECT 1 FROM DatasetSummary s
                                    WHERE s.DatasetName=r.DatasetName
                                      AND s.ReportType='first_pass_index') THEN 'index'
                     END,
                     CASE
                       WHEN EXISTS(SELECT 1 FROM DatasetSummary s WHERE s.DatasetName=r.DatasetName) THEN 'old'
                       ELSE 'none'
                     END
                   ) AS AiResultKind,
                   COALESCE(
                     (SELECT d.SchemaVersion
                        FROM AiDocuments d
                       WHERE d.SourceDataset=r.DatasetName
                         AND LOWER(COALESCE(d.PrimaryDefect,'')) NOT LIKE '%auto-extracted%'
                         AND LOWER(COALESCE(d.RawJson,'')) NOT LIKE '%_batch_auto.py%'
                         AND LOWER(COALESCE(d.RawJson,'')) NOT LIKE '%see workbook title/purpose%'
                         AND LOWER(COALESCE(d.RawJson,'')) NOT LIKE '%workbook stored but extraction surfaced narrative only%'
                       ORDER BY d.UpdatedAt DESC LIMIT 1),
                     CASE
                       WHEN EXISTS(SELECT 1 FROM DatasetSummary s
                                    WHERE s.DatasetName=r.DatasetName
                                      AND s.ReportType='first_pass_index') THEN 'first-pass-index-v1'
                     END,
                     CASE
                       WHEN EXISTS(SELECT 1 FROM DatasetSummary s WHERE s.DatasetName=r.DatasetName) THEN 'dataset-summary'
                       ELSE ''
                     END
                   ) AS AiSchemaVersion
            FROM RawReports r
            ORDER BY r.CreatedAt DESC;
            """;
        var list = new List<RawReportInfo>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new RawReportInfo(r.GetInt64(0), r.GetString(1), r.GetString(2),
                                       r.GetString(3), r.GetInt32(5), r.GetInt32(6),
                                       r.GetString(4), r.GetInt32(7) != 0,
                                       r.IsDBNull(8) ? "" : r.GetString(8),
                                       r.IsDBNull(9) ? "none" : r.GetString(9),
                                       r.IsDBNull(10) ? "" : r.GetString(10)));
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

    public CurrentProblemApplyResult ApplyCurrentProblemFirstPassRows(IReadOnlyList<CurrentProblemFirstPassRow> rows)
    {
        if (rows.Count == 0)
            return new CurrentProblemApplyResult(0, 0, 0, 0, 0, 0);

        string now = DateTime.UtcNow.ToString("O");
        int matched = 0;
        int summaryRows = 0;
        int productTypesFilled = 0;
        int reportDatesFilled = 0;
        int missing = 0;

        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        foreach (CurrentProblemFirstPassRow row in rows)
        {
            string dataset = (row.DatasetName ?? "").Trim();
            if (dataset.Length == 0)
                continue;

            string existingProduct = "";
            string existingDate = "";
            using (SqliteCommand find = conn.CreateCommand())
            {
                find.Transaction = tx;
                find.CommandText = "SELECT COALESCE(ProductType,''), COALESCE(ReportDate,'') FROM RawReports WHERE DatasetName=@n;";
                find.Parameters.AddWithValue("@n", dataset);
                using SqliteDataReader rd = find.ExecuteReader();
                if (!rd.Read())
                {
                    missing++;
                    continue;
                }

                existingProduct = rd.GetString(0);
                existingDate = rd.GetString(1);
            }

            matched++;

            string productType = FirstNonEmpty(row.Model, row.DbProductType, row.AiModel);
            string reportDate = FirstDateToken(row.Date);
            if (string.IsNullOrWhiteSpace(reportDate))
                reportDate = FirstDateToken(row.DbReportDate);

            if (string.IsNullOrWhiteSpace(existingProduct) && !string.IsNullOrWhiteSpace(productType))
            {
                using SqliteCommand update = conn.CreateCommand();
                update.Transaction = tx;
                update.CommandText = "UPDATE RawReports SET ProductType=@p WHERE DatasetName=@n;";
                update.Parameters.AddWithValue("@p", productType);
                update.Parameters.AddWithValue("@n", dataset);
                productTypesFilled += update.ExecuteNonQuery();
            }

            if (string.IsNullOrWhiteSpace(existingDate) && !string.IsNullOrWhiteSpace(reportDate))
            {
                using SqliteCommand update = conn.CreateCommand();
                update.Transaction = tx;
                update.CommandText = "UPDATE RawReports SET ReportDate=@d WHERE DatasetName=@n;";
                update.Parameters.AddWithValue("@d", reportDate);
                update.Parameters.AddWithValue("@n", dataset);
                reportDatesFilled += update.ExecuteNonQuery();
            }

            string summary = FirstNonEmpty(row.ReviewPurpose, row.Purpose, row.EvidenceSummary);
            string keyFindings = BuildFirstPassKeyFindings(row);
            string tagsJson = JsonSerializer.Serialize(FirstPassTags(row), JsonOpts);
            string purpose = row.Purpose ?? "";
            string testConditions = BuildFirstPassConditions(row, productType, reportDate);
            string rootCause = JoinNonEmpty(row.TargetDefects);
            string decision = row.Uncertainty ?? "";
            string recommendedAction = row.NeedsDetailedAnalysis
                ? "Use this as a first-pass index and run detailed current-problem analysis before process action."
                : "Use this as a first-pass index and verify the source workbook before process action.";

            using SqliteCommand upsert = conn.CreateCommand();
            upsert.Transaction = tx;
            upsert.CommandText = """
                INSERT INTO DatasetSummary
                    (DatasetName, ProductType, Summary, KeyFindings, Tags, CreatedAt,
                     Purpose, TestConditions, RootCause, Decision, RecommendedAction,
                     Verdict, Headline, EvidenceJson, ActionsJson, ContextJson,
                     ReportType, DoeGridJson, TrendJson)
                VALUES
                    (@n, @p, @s, @k, @t, @at,
                     @pu, @tc, @rc, @de, @ra,
                     '', '', '', '', '',
                     'first_pass_index', '', '')
                ON CONFLICT(DatasetName) DO UPDATE SET
                    ProductType=@p,
                    Summary=@s,
                    KeyFindings=@k,
                    Tags=@t,
                    CreatedAt=@at,
                    Purpose=@pu,
                    TestConditions=@tc,
                    RootCause=@rc,
                    Decision=@de,
                    RecommendedAction=@ra,
                    Verdict='',
                    Headline='',
                    EvidenceJson='',
                    ActionsJson='',
                    ContextJson='',
                    ReportType='first_pass_index',
                    DoeGridJson='',
                    TrendJson='';
                """;
            upsert.Parameters.AddWithValue("@n", dataset);
            upsert.Parameters.AddWithValue("@p", productType);
            upsert.Parameters.AddWithValue("@s", TruncateForDb(summary, 2000));
            upsert.Parameters.AddWithValue("@k", TruncateForDb(keyFindings, 4000));
            upsert.Parameters.AddWithValue("@t", tagsJson);
            upsert.Parameters.AddWithValue("@at", now);
            upsert.Parameters.AddWithValue("@pu", TruncateForDb(purpose, 2000));
            upsert.Parameters.AddWithValue("@tc", TruncateForDb(testConditions, 2000));
            upsert.Parameters.AddWithValue("@rc", TruncateForDb(rootCause, 1500));
            upsert.Parameters.AddWithValue("@de", TruncateForDb(decision, 2000));
            upsert.Parameters.AddWithValue("@ra", recommendedAction);
            upsert.ExecuteNonQuery();
            summaryRows++;
        }

        tx.Commit();
        return new CurrentProblemApplyResult(rows.Count, matched, summaryRows, productTypesFilled, reportDatesFilled, missing);
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

    public InputDataComExtractionStoreResult SaveInputDataComExtraction(InputDataComExtractionSave save)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        DeleteInputDataComWorkbookNoCommit(conn, tx, save.DatasetName, save.SourcePath);

        string createdAt = DateTime.UtcNow.ToString("O");
        long workbookId;
        using (SqliteCommand cmd = conn.CreateCommand())
        {
            cmd.Transaction = tx;
            cmd.CommandText = """
                INSERT INTO InputDataComWorkbooks
                    (DatasetName, SourcePath, SourceFileName, FileSize, MtimeNs, Fingerprint,
                     Status, Error, SheetCount, TotalRows, TotalCells, NonEmptyCells,
                     MergeCount, RawJsonPath, ExtractedAt, CreatedAt)
                VALUES
                    (@dataset, @path, @file, @size, @mtime, @fingerprint,
                     @status, @error, @sheetCount, @rows, @cells, @nonEmpty,
                     @merges, @rawJsonPath, @extractedAt, @createdAt);
                SELECT last_insert_rowid();
                """;
            cmd.Parameters.AddWithValue("@dataset", save.DatasetName);
            cmd.Parameters.AddWithValue("@path", save.SourcePath);
            cmd.Parameters.AddWithValue("@file", save.SourceFileName);
            cmd.Parameters.AddWithValue("@size", save.FileSize);
            cmd.Parameters.AddWithValue("@mtime", save.MtimeNs);
            cmd.Parameters.AddWithValue("@fingerprint", save.Fingerprint);
            cmd.Parameters.AddWithValue("@status", save.Status);
            cmd.Parameters.AddWithValue("@error", save.Error);
            cmd.Parameters.AddWithValue("@sheetCount", save.SheetCount);
            cmd.Parameters.AddWithValue("@rows", save.TotalRows);
            cmd.Parameters.AddWithValue("@cells", save.TotalCells);
            cmd.Parameters.AddWithValue("@nonEmpty", save.NonEmptyCells);
            cmd.Parameters.AddWithValue("@merges", save.MergeCount);
            cmd.Parameters.AddWithValue("@rawJsonPath", save.RawJsonPath);
            cmd.Parameters.AddWithValue("@extractedAt", save.ExtractedAt);
            cmd.Parameters.AddWithValue("@createdAt", createdAt);
            workbookId = Convert.ToInt64(cmd.ExecuteScalar());
        }

        using (SqliteCommand cmd = conn.CreateCommand())
        {
            cmd.Transaction = tx;
            cmd.CommandText = """
                INSERT INTO InputDataComSheets
                    (WorkbookId, SheetIndex, SheetName, UsedTop, UsedLeft, UsedBottom, UsedRight,
                     RowCount, ColumnCount, NonEmptyCells, MergeCount)
                VALUES
                    (@workbookId, @sheetIndex, @sheetName, @top, @left, @bottom, @right,
                     @rowCount, @colCount, @nonEmpty, @mergeCount);
                """;
            AddCommonParameter(cmd, "@workbookId", workbookId);
            SqliteParameter pSheetIndex = cmd.Parameters.Add("@sheetIndex", SqliteType.Integer);
            SqliteParameter pSheetName  = cmd.Parameters.Add("@sheetName", SqliteType.Text);
            SqliteParameter pTop        = cmd.Parameters.Add("@top", SqliteType.Integer);
            SqliteParameter pLeft       = cmd.Parameters.Add("@left", SqliteType.Integer);
            SqliteParameter pBottom     = cmd.Parameters.Add("@bottom", SqliteType.Integer);
            SqliteParameter pRight      = cmd.Parameters.Add("@right", SqliteType.Integer);
            SqliteParameter pRowCount   = cmd.Parameters.Add("@rowCount", SqliteType.Integer);
            SqliteParameter pColCount   = cmd.Parameters.Add("@colCount", SqliteType.Integer);
            SqliteParameter pNonEmpty   = cmd.Parameters.Add("@nonEmpty", SqliteType.Integer);
            SqliteParameter pMergeCount = cmd.Parameters.Add("@mergeCount", SqliteType.Integer);

            foreach (InputDataComSheetSave sheet in save.Sheets)
            {
                pSheetIndex.Value = sheet.SheetIndex;
                pSheetName.Value = sheet.SheetName;
                pTop.Value = sheet.UsedTop;
                pLeft.Value = sheet.UsedLeft;
                pBottom.Value = sheet.UsedBottom;
                pRight.Value = sheet.UsedRight;
                pRowCount.Value = sheet.RowCount;
                pColCount.Value = sheet.ColumnCount;
                pNonEmpty.Value = sheet.NonEmptyCells;
                pMergeCount.Value = sheet.MergeCount;
                cmd.ExecuteNonQuery();
            }
        }

        using (SqliteCommand cmd = conn.CreateCommand())
        {
            cmd.Transaction = tx;
            cmd.CommandText = """
                INSERT INTO InputDataComMerges
                    (WorkbookId, SheetName, Address, TopRow, LeftCol, BottomRow, RightCol,
                     RowSpan, ColumnSpan, AnchorValue)
                VALUES
                    (@workbookId, @sheetName, @address, @top, @left, @bottom, @right,
                     @rowSpan, @colSpan, @anchorValue);
                """;
            AddCommonParameter(cmd, "@workbookId", workbookId);
            SqliteParameter pSheetName = cmd.Parameters.Add("@sheetName", SqliteType.Text);
            SqliteParameter pAddress   = cmd.Parameters.Add("@address", SqliteType.Text);
            SqliteParameter pTop       = cmd.Parameters.Add("@top", SqliteType.Integer);
            SqliteParameter pLeft      = cmd.Parameters.Add("@left", SqliteType.Integer);
            SqliteParameter pBottom    = cmd.Parameters.Add("@bottom", SqliteType.Integer);
            SqliteParameter pRight     = cmd.Parameters.Add("@right", SqliteType.Integer);
            SqliteParameter pRowSpan   = cmd.Parameters.Add("@rowSpan", SqliteType.Integer);
            SqliteParameter pColSpan   = cmd.Parameters.Add("@colSpan", SqliteType.Integer);
            SqliteParameter pAnchor    = cmd.Parameters.Add("@anchorValue", SqliteType.Text);

            foreach (InputDataComMergeSave merge in save.Merges)
            {
                pSheetName.Value = merge.SheetName;
                pAddress.Value = merge.Address;
                pTop.Value = merge.TopRow;
                pLeft.Value = merge.LeftCol;
                pBottom.Value = merge.BottomRow;
                pRight.Value = merge.RightCol;
                pRowSpan.Value = merge.RowSpan;
                pColSpan.Value = merge.ColumnSpan;
                pAnchor.Value = merge.AnchorValue;
                cmd.ExecuteNonQuery();
            }
        }

        using (SqliteCommand cmd = conn.CreateCommand())
        {
            cmd.Transaction = tx;
            cmd.CommandText = """
                INSERT INTO InputDataComCells
                    (WorkbookId, SheetName, RowNumber, ColNumber, ColLabel, CellAddress,
                     CellValue, RawValue, MergeRole, MergeAddress, MergeAnchorRow, MergeAnchorCol)
                VALUES
                    (@workbookId, @sheetName, @rowNumber, @colNumber, @colLabel, @cellAddress,
                     @cellValue, @rawValue, @mergeRole, @mergeAddress, @anchorRow, @anchorCol);
                """;
            AddCommonParameter(cmd, "@workbookId", workbookId);
            SqliteParameter pSheetName    = cmd.Parameters.Add("@sheetName", SqliteType.Text);
            SqliteParameter pRowNumber    = cmd.Parameters.Add("@rowNumber", SqliteType.Integer);
            SqliteParameter pColNumber    = cmd.Parameters.Add("@colNumber", SqliteType.Integer);
            SqliteParameter pColLabel     = cmd.Parameters.Add("@colLabel", SqliteType.Text);
            SqliteParameter pCellAddress  = cmd.Parameters.Add("@cellAddress", SqliteType.Text);
            SqliteParameter pCellValue    = cmd.Parameters.Add("@cellValue", SqliteType.Text);
            SqliteParameter pRawValue     = cmd.Parameters.Add("@rawValue", SqliteType.Text);
            SqliteParameter pMergeRole    = cmd.Parameters.Add("@mergeRole", SqliteType.Text);
            SqliteParameter pMergeAddress = cmd.Parameters.Add("@mergeAddress", SqliteType.Text);
            SqliteParameter pAnchorRow    = cmd.Parameters.Add("@anchorRow", SqliteType.Integer);
            SqliteParameter pAnchorCol    = cmd.Parameters.Add("@anchorCol", SqliteType.Integer);

            foreach (InputDataComCellSave cell in save.Cells)
            {
                pSheetName.Value = cell.SheetName;
                pRowNumber.Value = cell.RowNumber;
                pColNumber.Value = cell.ColNumber;
                pColLabel.Value = cell.ColLabel;
                pCellAddress.Value = cell.CellAddress;
                pCellValue.Value = cell.CellValue;
                pRawValue.Value = cell.RawValue;
                pMergeRole.Value = cell.MergeRole;
                pMergeAddress.Value = cell.MergeAddress;
                pAnchorRow.Value = cell.MergeAnchorRow.HasValue ? cell.MergeAnchorRow.Value : DBNull.Value;
                pAnchorCol.Value = cell.MergeAnchorCol.HasValue ? cell.MergeAnchorCol.Value : DBNull.Value;
                cmd.ExecuteNonQuery();
            }
        }

        using (SqliteCommand cmd = conn.CreateCommand())
        {
            cmd.Transaction = tx;
            cmd.CommandText = """
                INSERT INTO InputDataReviewCandidates
                    (WorkbookId, CandidateKind, SheetName, RowNumber, Label, Confidence,
                     EvidenceCellsJson, RawText, CreatedAt)
                VALUES
                    (@workbookId, @kind, @sheetName, @rowNumber, @label, @confidence,
                     @evidence, @rawText, @createdAt);
                """;
            AddCommonParameter(cmd, "@workbookId", workbookId);
            SqliteParameter pKind       = cmd.Parameters.Add("@kind", SqliteType.Text);
            SqliteParameter pSheetName  = cmd.Parameters.Add("@sheetName", SqliteType.Text);
            SqliteParameter pRowNumber  = cmd.Parameters.Add("@rowNumber", SqliteType.Integer);
            SqliteParameter pLabel      = cmd.Parameters.Add("@label", SqliteType.Text);
            SqliteParameter pConfidence = cmd.Parameters.Add("@confidence", SqliteType.Text);
            SqliteParameter pEvidence   = cmd.Parameters.Add("@evidence", SqliteType.Text);
            SqliteParameter pRawText    = cmd.Parameters.Add("@rawText", SqliteType.Text);
            SqliteParameter pCreatedAt  = cmd.Parameters.Add("@createdAt", SqliteType.Text);

            foreach (InputDataReviewCandidateSave candidate in save.Candidates)
            {
                pKind.Value = candidate.CandidateKind;
                pSheetName.Value = candidate.SheetName;
                pRowNumber.Value = candidate.RowNumber;
                pLabel.Value = TruncateForDb(candidate.Label, 500);
                pConfidence.Value = candidate.Confidence;
                pEvidence.Value = JsonSerializer.Serialize(candidate.EvidenceCells, JsonOpts);
                pRawText.Value = TruncateForDb(candidate.RawText, 2000);
                pCreatedAt.Value = createdAt;
                cmd.ExecuteNonQuery();
            }
        }

        tx.Commit();

        return new InputDataComExtractionStoreResult(
            workbookId,
            save.DatasetName,
            save.SourceFileName,
            save.SheetCount,
            save.TotalRows,
            save.TotalCells,
            save.NonEmptyCells,
            save.MergeCount,
            save.Candidates.Count);
    }

    public void SaveInputDataReviewCase(InputDataReviewCaseSave save)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO InputDataReviewCases
                (WorkbookId, ReviewCaseId, Status, ApprovedForAskAi,
                 GenerationJson, VerificationJson, CreatedAt, VerifiedAt)
            VALUES
                (@workbookId, @reviewCaseId, @status, @approved,
                 @generation, @verification, @createdAt, @verifiedAt);
            """;
        cmd.Parameters.AddWithValue("@workbookId", save.WorkbookId);
        cmd.Parameters.AddWithValue("@reviewCaseId", save.ReviewCaseId);
        cmd.Parameters.AddWithValue("@status", save.Status);
        cmd.Parameters.AddWithValue("@approved", save.ApprovedForAskAi ? 1 : 0);
        cmd.Parameters.AddWithValue("@generation", save.GenerationJson);
        cmd.Parameters.AddWithValue("@verification", save.VerificationJson);
        cmd.Parameters.AddWithValue("@createdAt", save.CreatedAt);
        cmd.Parameters.AddWithValue("@verifiedAt", save.VerifiedAt);
        cmd.ExecuteNonQuery();
    }

    public List<InputDataReviewCaseRecord> GetInputDataReviewCases(long workbookId)
    {
        var rows = new List<InputDataReviewCaseRecord>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, WorkbookId, ReviewCaseId, Status, ApprovedForAskAi,
                   GenerationJson, VerificationJson, UserAnswerJson, CreatedAt, VerifiedAt, UserReviewedAt
              FROM InputDataReviewCases
             WHERE WorkbookId=@workbookId
             ORDER BY Id DESC;
            """;
        cmd.Parameters.AddWithValue("@workbookId", workbookId);

        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            rows.Add(new InputDataReviewCaseRecord(
                r.GetInt64(0),
                r.GetInt64(1),
                r.IsDBNull(2) ? "" : r.GetString(2),
                r.IsDBNull(3) ? "" : r.GetString(3),
                !r.IsDBNull(4) && r.GetInt32(4) != 0,
                r.IsDBNull(5) ? "{}" : r.GetString(5),
                r.IsDBNull(6) ? "{}" : r.GetString(6),
                r.IsDBNull(7) ? "{}" : r.GetString(7),
                r.IsDBNull(8) ? "" : r.GetString(8),
                r.IsDBNull(9) ? "" : r.GetString(9),
                r.IsDBNull(10) ? "" : r.GetString(10)));
        }

        return rows;
    }

    public List<InputDataComRowPreview> GetInputDataComRows(long workbookId, IReadOnlyList<string> rowRefs)
    {
        var parsedRefs = rowRefs
            .Select(TryParseInputDataRowRef)
            .Where(x => x.HasValue)
            .Select(x => x!.Value)
            .Distinct()
            .ToList();
        if (parsedRefs.Count == 0) return [];

        var order = parsedRefs
            .Select((row, index) => (row, index))
            .ToDictionary(x => x.row, x => x.index);
        var cellsByRow = new Dictionary<(string SheetName, int RowNumber), List<InputDataComCellPreview>>();

        using SqliteConnection conn = OpenConnection();
        foreach (IGrouping<string, (string SheetName, int RowNumber)> group in parsedRefs.GroupBy(x => x.SheetName))
        {
            List<int> rowNumbers = group.Select(x => x.RowNumber).Distinct().OrderBy(x => x).ToList();
            if (rowNumbers.Count == 0) continue;

            using SqliteCommand cmd = conn.CreateCommand();
            cmd.Parameters.AddWithValue("@workbookId", workbookId);
            cmd.Parameters.AddWithValue("@sheet", group.Key);

            var rowParams = new List<string>();
            for (int i = 0; i < rowNumbers.Count; i++)
            {
                string name = "@r" + i.ToString(System.Globalization.CultureInfo.InvariantCulture);
                rowParams.Add(name);
                cmd.Parameters.AddWithValue(name, rowNumbers[i]);
            }

            cmd.CommandText = $"""
                SELECT SheetName, RowNumber, ColNumber, ColLabel, CellAddress,
                       CellValue, MergeRole, MergeAddress
                  FROM InputDataComCells
                 WHERE WorkbookId=@workbookId
                   AND SheetName=@sheet
                   AND RowNumber IN ({string.Join(",", rowParams)})
                 ORDER BY RowNumber, ColNumber;
                """;

            using SqliteDataReader r = cmd.ExecuteReader();
            while (r.Read())
            {
                string sheetName = r.IsDBNull(0) ? "" : r.GetString(0);
                int rowNumber = r.GetInt32(1);
                var key = (sheetName, rowNumber);
                if (!cellsByRow.TryGetValue(key, out List<InputDataComCellPreview>? cells))
                {
                    cells = [];
                    cellsByRow[key] = cells;
                }

                cells.Add(new InputDataComCellPreview(
                    r.GetInt32(2),
                    r.IsDBNull(3) ? "" : r.GetString(3),
                    r.IsDBNull(4) ? "" : r.GetString(4),
                    r.IsDBNull(5) ? "" : r.GetString(5),
                    r.IsDBNull(6) ? "" : r.GetString(6),
                    r.IsDBNull(7) ? "" : r.GetString(7)));
            }
        }

        return cellsByRow
            .Select(pair => new InputDataComRowPreview(pair.Key.SheetName, pair.Key.RowNumber, pair.Value))
            .OrderBy(row => order.TryGetValue((row.SheetName, row.RowNumber), out int index) ? index : int.MaxValue)
            .ThenBy(row => row.SheetName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.RowNumber)
            .ToList();
    }

    public void SaveInputDataReviewCaseUserReview(InputDataReviewCaseUserReviewSave save)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE InputDataReviewCases
               SET UserAnswerJson=@answers,
                   UserReviewedAt=@reviewedAt,
                   ApprovedForAskAi=@approved,
                   Status=CASE WHEN @approved = 1 THEN 'user_verified' ELSE 'needs_review' END
             WHERE Id=@id;
            """;
        cmd.Parameters.AddWithValue("@answers", save.UserAnswerJson);
        cmd.Parameters.AddWithValue("@reviewedAt", save.ReviewedAt);
        cmd.Parameters.AddWithValue("@approved", save.ApprovedForAskAi ? 1 : 0);
        cmd.Parameters.AddWithValue("@id", save.ReviewCaseDbId);
        cmd.ExecuteNonQuery();
    }

    private static (string SheetName, int RowNumber)? TryParseInputDataRowRef(string rowRef)
    {
        string text = (rowRef ?? "").Trim();
        int bang = text.LastIndexOf('!');
        if (bang <= 0 || bang >= text.Length - 1) return null;

        string sheetName = text[..bang].Trim();
        string rowText = text[(bang + 1)..].Trim();
        return int.TryParse(rowText, out int rowNumber) && rowNumber > 0 && !string.IsNullOrWhiteSpace(sheetName)
            ? (sheetName, rowNumber)
            : null;
    }

    public bool DeleteInputDataReviewCase(long reviewCaseDbId)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM InputDataReviewCases WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", reviewCaseDbId);
        return cmd.ExecuteNonQuery() > 0;
    }

    public List<InputDataReviewCaseListRecord> GetRecentInputDataReviewCases(string datasetName, int limit = 20)
    {
        var rows = new List<InputDataReviewCaseListRecord>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT rv.Id, rv.WorkbookId, wb.DatasetName, wb.SourceFileName, wb.SourcePath,
                   rv.ReviewCaseId, rv.Status, rv.ApprovedForAskAi, rv.CreatedAt, rv.UserReviewedAt
              FROM InputDataReviewCases rv
              JOIN InputDataComWorkbooks wb ON wb.Id = rv.WorkbookId
             WHERE (@dataset = '' OR wb.DatasetName = @dataset)
             ORDER BY rv.Id DESC
             LIMIT @limit;
            """;
        cmd.Parameters.AddWithValue("@dataset", datasetName.Trim());
        cmd.Parameters.AddWithValue("@limit", Math.Clamp(limit, 1, 100));

        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            rows.Add(new InputDataReviewCaseListRecord(
                r.GetInt64(0),
                r.GetInt64(1),
                r.IsDBNull(2) ? "" : r.GetString(2),
                r.IsDBNull(3) ? "" : r.GetString(3),
                r.IsDBNull(4) ? "" : r.GetString(4),
                r.IsDBNull(5) ? "" : r.GetString(5),
                r.IsDBNull(6) ? "" : r.GetString(6),
                !r.IsDBNull(7) && r.GetInt32(7) != 0,
                r.IsDBNull(8) ? "" : r.GetString(8),
                r.IsDBNull(9) ? "" : r.GetString(9)));
        }

        return rows;
    }

    private static void DeleteInputDataComWorkbookNoCommit(SqliteConnection conn, SqliteTransaction tx, string datasetName, string sourcePath)
    {
        var ids = new List<long>();
        using (SqliteCommand find = conn.CreateCommand())
        {
            find.Transaction = tx;
            find.CommandText = "SELECT Id FROM InputDataComWorkbooks WHERE DatasetName=@dataset AND SourcePath=@path;";
            find.Parameters.AddWithValue("@dataset", datasetName);
            find.Parameters.AddWithValue("@path", sourcePath);
            using SqliteDataReader r = find.ExecuteReader();
            while (r.Read()) ids.Add(r.GetInt64(0));
        }

        foreach (long id in ids)
        {
            foreach (string table in new[]
                     {
                         "InputDataReviewCases",
                         "InputDataReviewCandidates",
                         "InputDataComMerges",
                         "InputDataComCells",
                         "InputDataComSheets",
                         "InputDataComWorkbooks"
                     })
            {
                using SqliteCommand del = conn.CreateCommand();
                del.Transaction = tx;
                del.CommandText = $"DELETE FROM {table} WHERE {(table == "InputDataComWorkbooks" ? "Id" : "WorkbookId")}=@id;";
                del.Parameters.AddWithValue("@id", id);
                del.ExecuteNonQuery();
            }
        }
    }

    private static void AddCommonParameter(SqliteCommand cmd, string name, object value)
    {
        SqliteParameter parameter = cmd.Parameters.Add(name, SqliteType.Integer);
        parameter.Value = value;
    }

    // ?? Raw attached files (any type ??Excel, PDF, etc.) ?????????????????????

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

    public List<InputDataBatchDbWorkbook> GetInputDataBatchWorkbookFiles(bool includeExcluded = false)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT f.Id,
                   f.DatasetName,
                   f.FileName,
                   f.MediaType,
                   f.FileSize,
                   COALESCE(r.ProductType, '') AS ProductType,
                   COALESCE(r.ReportDate, '') AS ReportDate,
                   f.CreatedAt,
                   COALESCE(r.BatchExcluded, 0) AS BatchExcluded,
                   COALESCE((SELECT MAX(d.UpdatedAt)
                               FROM AiDocuments d
                              WHERE d.SourceDataset=f.DatasetName), '') AS LatestAiUpdatedAt,
                   COALESCE((SELECT CASE
                                       WHEN d.SchemaVersion='input-data-test-batch-v1'
                                         OR d.RawJson LIKE '%INPUT_DATA_BATCH_FROM_INPUT_DATA_TEST%'
                                       THEN 'new'
                                       ELSE 'old'
                                     END
                               FROM AiDocuments d
                              WHERE d.SourceDataset=f.DatasetName
                              ORDER BY d.UpdatedAt DESC
                              LIMIT 1), 'none') AS AiResultKind,
                   COALESCE((SELECT d.SchemaVersion
                               FROM AiDocuments d
                              WHERE d.SourceDataset=f.DatasetName
                              ORDER BY d.UpdatedAt DESC
                              LIMIT 1), '') AS AiSchemaVersion
            FROM RawReportFiles f
            LEFT JOIN RawReports r ON r.DatasetName=f.DatasetName
            WHERE (
                    LOWER(f.FileName) LIKE '%.xlsx'
                 OR LOWER(f.FileName) LIKE '%.xlsm'
                 OR LOWER(f.FileName) LIKE '%.xlsb'
                 OR LOWER(f.FileName) LIKE '%.xls'
                 OR f.MediaType LIKE '%spreadsheet%'
                  )
              AND (@includeExcluded = 1 OR COALESCE(r.BatchExcluded, 0) = 0)
            ORDER BY f.DatasetName COLLATE NOCASE, f.Id;
            """;
        cmd.Parameters.AddWithValue("@includeExcluded", includeExcluded ? 1 : 0);

        var list = new List<InputDataBatchDbWorkbook>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new InputDataBatchDbWorkbook(
                r.GetInt64(0),
                r.GetString(1),
                r.GetString(2),
                r.GetString(3),
                r.GetInt64(4),
                r.IsDBNull(5) ? "" : r.GetString(5),
                r.IsDBNull(6) ? "" : r.GetString(6),
                r.GetString(7),
                r.GetInt32(8) != 0,
                r.IsDBNull(9) ? "" : r.GetString(9),
                r.IsDBNull(10) ? "none" : r.GetString(10),
                r.IsDBNull(11) ? "" : r.GetString(11)));
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

    public int DeleteAiAnalysisForDataset(string datasetName)
    {
        string dataset = (datasetName ?? "").Trim();
        if (dataset.Length == 0) return 0;

        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        using SqliteCommand count = conn.CreateCommand();
        count.Transaction = tx;
        count.CommandText = "SELECT COUNT(*) FROM AiDocuments WHERE SourceDataset=@dataset;";
        count.Parameters.AddWithValue("@dataset", dataset);
        int deletedDocuments = Convert.ToInt32(count.ExecuteScalar() ?? 0);

        DeleteAiAnalysisForDataset(conn, tx, dataset);
        tx.Commit();
        return deletedDocuments;
    }

    public void SaveInputDataBatchAnalysisOverwrite(
        string datasetName,
        string sourceFile,
        string productType,
        string reportDate,
        string analysisText,
        string analysisHtml,
        InputDataTestAnalysisParameters? parameters,
        string sessionPath)
    {
        if (string.IsNullOrWhiteSpace(datasetName))
            throw new ArgumentException("Dataset name is required.", nameof(datasetName));

        string now = DateTime.UtcNow.ToString("O");
        string docId = "input-data-test-" + Guid.NewGuid().ToString("N");
        InputDataTestAnalysisParameters p = parameters ?? InputDataTestAnalysisParameters.Empty;
        string title = FirstMarkdownHeading(analysisText);
        if (string.IsNullOrWhiteSpace(title))
            title = string.IsNullOrWhiteSpace(sourceFile) ? datasetName : Path.GetFileNameWithoutExtension(sourceFile);

        string purpose = FirstNonEmpty(p.ReviewPurpose, p.Purpose);
        if (!string.IsNullOrWhiteSpace(p.Purpose) && !purpose.Contains(p.Purpose, StringComparison.OrdinalIgnoreCase))
            purpose = string.IsNullOrWhiteSpace(purpose) ? p.Purpose : $"{purpose} / {p.Purpose}";
        purpose = TruncateForDb(CleanInputDataAnalysisLine(purpose), 1200);
        if (string.IsNullOrWhiteSpace(purpose))
            purpose = ExtractInputDataAnalysisPurpose(analysisText, analysisHtml, title);
        if (string.IsNullOrWhiteSpace(purpose))
            purpose = CleanInputDataAnalysisLine(FirstMeaningfulLine(analysisText, title));
        List<string> aiTargetDefects = ExtractInputDataAnalysisList(
            analysisText,
            analysisHtml,
            "대상불량",
            "대상 불량",
            "Target defect",
            "Target defects");
        List<string> aiReviewItems = ExtractInputDataAnalysisList(
            analysisText,
            analysisHtml,
            "검토사항",
            "검토 사항",
            "검토 항목",
            "확인 항목",
            "Review item",
            "Review items",
            "Check item",
            "Check items");
        List<string> parameterTargetDefects = CleanInputDataParameterList(p.TargetDefects);
        if (parameterTargetDefects.Count > 0)
            aiTargetDefects = parameterTargetDefects;

        List<string> parameterReviewItems = CleanInputDataParameterList(p.ReviewItems);
        if (parameterReviewItems.Count > 0)
            aiReviewItems = parameterReviewItems;

        string model = FirstNonEmpty(p.Model, productType);
        string date = FirstNonEmpty(p.Date, reportDate);
        double confidence = p.Confidence is >= 0 and <= 1 ? p.Confidence.Value : 0.90;

        string primaryDefect = aiTargetDefects.FirstOrDefault() ?? "";
        string primaryDefectJson = string.IsNullOrWhiteSpace(primaryDefect)
            ? "{}"
            : JsonSerializer.Serialize(new { canonical_name = primaryDefect, aliases_in_document = Array.Empty<string>() }, JsonOpts);
        string relatedDefectsJson = JsonSerializer.Serialize(aiTargetDefects, JsonOpts);
        string processesJson = JsonSerializer.Serialize(aiReviewItems, JsonOpts);
        string markdown = string.IsNullOrWhiteSpace(analysisText)
            ? "Analysis HTML was generated. See RawJson.analysisHtml."
            : analysisText.Trim();
        string rawJson = JsonSerializer.Serialize(new
        {
            pipeline = "INPUT_DATA_BATCH_FROM_INPUT_DATA_TEST",
            source = "INPUT DATA(BATCH)",
            sessionPath,
            parameters = InputDataParametersPayload(p),
            analysisText = analysisText ?? "",
            analysisHtml = analysisHtml ?? ""
        }, JsonOpts);
        string contentJson = JsonSerializer.Serialize(SplitAnalysisContent(markdown), JsonOpts);
        string logId = docId + "-log";
        string conclusionId = docId + "-conclusion-1";

        using SqliteConnection conn = OpenConnection();
        using SqliteTransaction tx = conn.BeginTransaction();

        DeleteAiAnalysisForDataset(conn, tx, datasetName);
        if (!string.IsNullOrWhiteSpace(productType))
        {
            using SqliteCommand delModel = conn.CreateCommand();
            delModel.Transaction = tx;
            delModel.CommandText = "DELETE FROM AiModelAnalyses WHERE ProductType=@p;";
            delModel.Parameters.AddWithValue("@p", productType);
            delModel.ExecuteNonQuery();
        }

        using (SqliteCommand ins = conn.CreateCommand())
        {
            ins.Transaction = tx;
            ins.CommandText = """
                INSERT INTO AiDocuments
                    (DocumentId, SourceDataset, SourceFile, Title, Model, ReportDate,
                     Department, Marker, Line, ReportType, PrimaryDefect,
                     PrimaryDefectJson, RelatedDefectsJson, PartsJson, ProcessesJson,
                     Purpose, ContentJson, GeneratedReportMarkdown, SourceCellsJson,
                     Confidence, SchemaVersion, RawJson, CreatedAt, UpdatedAt)
                VALUES
                    (@id, @dataset, @file, @title, @model, @date,
                     '', '', '', @type, @primaryDefect,
                     @primaryDefectJson, @relatedDefects, '[]', @processes,
                     @purpose, @content, @markdown, '{}',
                     @confidence, @schema, @raw, @created, @updated);
                """;
            ins.Parameters.AddWithValue("@id", docId);
            ins.Parameters.AddWithValue("@dataset", datasetName);
            ins.Parameters.AddWithValue("@file", sourceFile ?? "");
            ins.Parameters.AddWithValue("@title", title);
            ins.Parameters.AddWithValue("@model", model);
            ins.Parameters.AddWithValue("@date", date);
            ins.Parameters.AddWithValue("@type", "input_data_test_analysis");
            ins.Parameters.AddWithValue("@primaryDefect", primaryDefect);
            ins.Parameters.AddWithValue("@primaryDefectJson", primaryDefectJson);
            ins.Parameters.AddWithValue("@relatedDefects", relatedDefectsJson);
            ins.Parameters.AddWithValue("@processes", processesJson);
            ins.Parameters.AddWithValue("@purpose", purpose);
            ins.Parameters.AddWithValue("@content", contentJson);
            ins.Parameters.AddWithValue("@markdown", markdown);
            ins.Parameters.AddWithValue("@confidence", confidence);
            ins.Parameters.AddWithValue("@schema", "input-data-test-batch-v1");
            ins.Parameters.AddWithValue("@raw", rawJson);
            ins.Parameters.AddWithValue("@created", now);
            ins.Parameters.AddWithValue("@updated", now);
            ins.ExecuteNonQuery();
        }

        using (SqliteCommand concl = conn.CreateCommand())
        {
            concl.Transaction = tx;
            concl.CommandText = """
                INSERT INTO AiConclusions
                    (ConclusionId, DocumentId, Topic, StatementFromReport,
                     NormalizedInterpretation, SourceFile, SheetName, SourceCellsJson)
                VALUES
                    (@id, @doc, @topic, @statement, @interp, @file, '', '{}');
                """;
            concl.Parameters.AddWithValue("@id", conclusionId);
            concl.Parameters.AddWithValue("@doc", docId);
            concl.Parameters.AddWithValue("@topic", "INPUT DATA(TEST) analysis");
            concl.Parameters.AddWithValue("@statement", TruncateForDb(CleanInputDataAnalysisLine(FirstMeaningfulLine(markdown, "")), 1200));
            concl.Parameters.AddWithValue("@interp", TruncateForDb(markdown, 3000));
            concl.Parameters.AddWithValue("@file", sourceFile ?? "");
            concl.ExecuteNonQuery();
        }

        using (SqliteCommand log = conn.CreateCommand())
        {
            log.Transaction = tx;
            log.CommandText = """
                INSERT INTO AiExtractionLogs
                    (LogId, DocumentId, Confidence, AssumptionsJson, WarningsJson,
                     DecisionRationale, CreatedAt)
                VALUES
                    (@id, @doc, @confidence, '[]', '[]', @rationale, @created);
                """;
            log.Parameters.AddWithValue("@id", logId);
            log.Parameters.AddWithValue("@doc", docId);
            log.Parameters.AddWithValue("@confidence", confidence);
            log.Parameters.AddWithValue("@rationale", "Generated by INPUT DATA(BATCH) using INPUT DATA(TEST) Extract + Analysis flow.");
            log.Parameters.AddWithValue("@created", now);
            log.ExecuteNonQuery();
        }

        tx.Commit();
    }

    private static object InputDataParametersPayload(InputDataTestAnalysisParameters p)
        => new
        {
            reviewPurpose = p.ReviewPurpose,
            tags = p.Tags,
            purpose = p.Purpose,
            purposeCode = p.PurposeCode,
            targetDefects = p.TargetDefects,
            reviewItems = p.ReviewItems,
            model = p.Model,
            date = p.Date,
            confidence = p.Confidence
        };

    private static List<string> CleanInputDataParameterList(IEnumerable<string> values)
        => values
            .Select(CleanInputDataAnalysisLine)
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Where(v => !v.Equals("unknown", StringComparison.OrdinalIgnoreCase))
            .Where(v => !v.Equals("n/a", StringComparison.OrdinalIgnoreCase))
            .Where(v => !LooksLikeInputDataSourceFile(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(10)
            .Select(v => TruncateForDb(v, 180))
            .ToList();

    private static string BuildFirstPassKeyFindings(CurrentProblemFirstPassRow row)
    {
        var lines = new List<string>();
        if (!string.IsNullOrWhiteSpace(row.EvidenceSummary))
            lines.Add("- Evidence: " + row.EvidenceSummary.Trim());
        if (row.TargetDefects.Count > 0)
            lines.Add("- Target defects: " + JoinNonEmpty(row.TargetDefects));
        if (row.ReviewItems.Count > 0)
            lines.Add("- Review items: " + JoinNonEmpty(row.ReviewItems));
        if (row.Tags.Count > 0)
            lines.Add("- Tags: " + JoinNonEmpty(row.Tags));
        if (row.EvidenceCells.Count > 0)
            lines.Add("- Evidence cells: " + JoinNonEmpty(row.EvidenceCells.Take(20)));
        if (!string.IsNullOrWhiteSpace(row.Uncertainty))
            lines.Add("- Uncertainty: " + row.Uncertainty.Trim());
        return string.Join(Environment.NewLine, lines);
    }

    private static string BuildFirstPassConditions(CurrentProblemFirstPassRow row, string productType, string reportDate)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(productType)) parts.Add("Model: " + productType.Trim());
        if (!string.IsNullOrWhiteSpace(reportDate)) parts.Add("Date: " + reportDate.Trim());
        if (!string.IsNullOrWhiteSpace(row.AiModel)) parts.Add("AI model: " + row.AiModel.Trim());
        if (!string.IsNullOrWhiteSpace(row.ModelMappingSource)) parts.Add("Model mapping source: " + row.ModelMappingSource.Trim());
        if (!string.IsNullOrWhiteSpace(row.PurposeCode)) parts.Add("Purpose code: " + row.PurposeCode.Trim());
        if (!string.IsNullOrWhiteSpace(row.FileNames)) parts.Add("Source file: " + row.FileNames.Trim());
        parts.Add("Confidence: " + Math.Clamp(row.Confidence, 0, 1).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture));
        return string.Join(Environment.NewLine, parts);
    }

    private static List<string> FirstPassTags(CurrentProblemFirstPassRow row)
    {
        return row.Tags
            .Concat(row.TargetDefects)
            .Concat(row.ReviewItems)
            .Select(CleanInputDataAnalysisLine)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(24)
            .ToList();
    }

    private static string JoinNonEmpty(IEnumerable<string> values)
        => string.Join(", ", values
            .Select(x => (x ?? "").Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase));

    private static string FirstDateToken(string? value)
    {
        string text = (value ?? "").Trim();
        if (text.Length == 0) return "";

        var match = System.Text.RegularExpressions.Regex.Match(text, @"\b\d{4}-\d{2}-\d{2}\b");
        if (match.Success)
            return match.Value;

        string first = text.Split(';', ',', '|').Select(x => x.Trim()).FirstOrDefault(x => x.Length > 0) ?? text;
        return DateTime.TryParse(first, null, System.Globalization.DateTimeStyles.AssumeLocal, out DateTime dt)
            ? dt.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)
            : "";
    }

    private static string FirstNonEmpty(params string?[] values)
    {
        foreach (string? value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
                return value.Trim();
        }
        return "";
    }

    private static void DeleteAiAnalysisForDataset(SqliteConnection conn, SqliteTransaction tx, string datasetName)
    {
        string docIds = "SELECT DocumentId FROM AiDocuments WHERE SourceDataset=@dataset";
        string[] statements =
        [
            $"DELETE FROM AiNgBreakdowns WHERE ResultId IN (SELECT ResultId FROM AiResults WHERE DocumentId IN ({docIds}));",
            $"DELETE FROM AiResults WHERE DocumentId IN ({docIds});",
            $"DELETE FROM AiTestConditions WHERE DocumentId IN ({docIds});",
            $"DELETE FROM AiConclusionTranslations WHERE ConclusionId IN (SELECT ConclusionId FROM AiConclusions WHERE DocumentId IN ({docIds}));",
            $"DELETE FROM AiConclusions WHERE DocumentId IN ({docIds});",
            $"DELETE FROM AiHintTranslations WHERE HintId IN (SELECT HintId FROM AiTroubleshootingHints WHERE DocumentId IN ({docIds}));",
            $"DELETE FROM AiTroubleshootingHints WHERE DocumentId IN ({docIds});",
            $"DELETE FROM AiLogTranslations WHERE LogId IN (SELECT LogId FROM AiExtractionLogs WHERE DocumentId IN ({docIds}));",
            $"DELETE FROM AiExtractionLogs WHERE DocumentId IN ({docIds});",
            $"DELETE FROM AiDocumentTranslations WHERE DocumentId IN ({docIds});",
            $"DELETE FROM AiDocuments WHERE SourceDataset=@dataset;"
        ];

        foreach (string sql in statements)
        {
            using SqliteCommand cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@dataset", datasetName);
            cmd.ExecuteNonQuery();
        }
    }

    private static List<string> SplitAnalysisContent(string markdown)
    {
        var lines = new List<string>();
        foreach (string rawLine in (markdown ?? "").Replace("\r\n", "\n").Split('\n'))
        {
            string line = rawLine.Trim();
            if (line.Length == 0) continue;
            line = line.TrimStart('#').Trim();
            line = line.TrimStart('-', '*').Trim();
            if (line.Length == 0) continue;
            lines.Add(TruncateForDb(line, 700));
            if (lines.Count >= 12) break;
        }
        return lines;
    }

    private static string FirstMarkdownHeading(string? markdown)
    {
        foreach (string rawLine in (markdown ?? "").Replace("\r\n", "\n").Split('\n'))
        {
            string line = rawLine.Trim();
            if (line.StartsWith("#", StringComparison.Ordinal))
                return TruncateForDb(line.TrimStart('#').Trim(), 180);
        }
        return "";
    }

    private static string FirstMeaningfulLine(string? markdown, string skip)
    {
        foreach (string rawLine in (markdown ?? "").Replace("\r\n", "\n").Split('\n'))
        {
            string line = rawLine.Trim();
            if (line.Length == 0) continue;
            line = line.TrimStart('#').Trim();
            line = line.TrimStart('-', '*').Trim();
            if (line.Length == 0) continue;
            if (!string.IsNullOrWhiteSpace(skip) && line.Equals(skip, StringComparison.OrdinalIgnoreCase)) continue;
            return TruncateForDb(line, 1200);
        }
        return "";
    }

    private static string ExtractInputDataAnalysisPurpose(string? analysisText, string? analysisHtml, string title)
    {
        string htmlText = StripHtmlForAnalysisText(analysisHtml);
        string combined = (analysisText ?? "") + "\n" + htmlText;
        foreach (string label in new[]
                 {
                     "검토 목적", "목적", "시험 의도",
                     "Review purpose", "Purpose", "Test intent", "Goal"
                 })
        {
            string? value = ExtractLabeledLineValue(combined, label);
            if (string.IsNullOrWhiteSpace(value)) continue;
            if (!string.IsNullOrWhiteSpace(title)
                && value.Equals(title, StringComparison.OrdinalIgnoreCase)) continue;
            if (LooksLikeInputDataSourceFile(value)) continue;
            return TruncateForDb(CleanInputDataAnalysisLine(value), 1200);
        }

        string? malformedValue = ExtractMalformedInputDataLabelValue(combined, 3)
            ?? ExtractMalformedInputDataLabelValue(combined, 1);
        if (!string.IsNullOrWhiteSpace(malformedValue)
            && !malformedValue.Equals(title, StringComparison.OrdinalIgnoreCase)
            && !LooksLikeInputDataSourceFile(malformedValue))
        {
            return TruncateForDb(malformedValue, 1200);
        }
        return "";
    }

    private static List<string> ExtractInputDataAnalysisList(string? analysisText, string? analysisHtml, params string[] labels)
    {
        string htmlText = StripHtmlForAnalysisText(analysisHtml);
        string combined = (analysisText ?? "") + "\n" + htmlText;
        foreach (string label in labels)
        {
            string? value = ExtractLabeledLineValue(combined, label);
            if (string.IsNullOrWhiteSpace(value)) continue;

            return SplitInputDataAnalysisList(value)
                .Take(10)
                .ToList();
        }

        int malformedOrdinal = MalformedInputDataOrdinalForLabels(labels);
        if (malformedOrdinal > 0)
        {
            string? value = ExtractMalformedInputDataLabelValue(combined, malformedOrdinal);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return SplitInputDataAnalysisList(value)
                    .Take(10)
                    .ToList();
            }
        }
        return [];
    }

    private static IEnumerable<string> SplitInputDataAnalysisList(string value)
    {
        foreach (string item in System.Text.RegularExpressions.Regex.Split(value ?? "", @"[,\n;|、，]+"))
        {
            string clean = System.Text.RegularExpressions.Regex.Replace(item, @"\s+", " ")
                .Trim(' ', '-', '*', '·', ':', '：');
            clean = CleanInputDataAnalysisLine(clean);
            if (clean.Length < 2) continue;
            if (clean.Equals("확인 필요", StringComparison.OrdinalIgnoreCase)) continue;
            if (clean.Equals("unknown", StringComparison.OrdinalIgnoreCase)) continue;
            if (clean.Equals("n/a", StringComparison.OrdinalIgnoreCase)) continue;
            if (LooksLikeInputDataSourceFile(clean)) continue;

            yield return TruncateForDb(clean, 180);
        }
    }

    private static string? ExtractLabeledLineValue(string text, string label)
    {
        string[] lines = (text ?? "").Replace("\r\n", "\n").Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            string line = System.Text.RegularExpressions.Regex.Replace(lines[i], @"\s+", " ").Trim();
            if (line.Length == 0) continue;

            System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(
                line,
                @"^\s*" + System.Text.RegularExpressions.Regex.Escape(label) + @"\s*[:：]\s*(.{2,220})$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (!match.Success) continue;

            string value = match.Groups[1].Value.Trim(' ', '-', '*', '·', ':', '：');
            return value.Length == 0 ? null : value;
        }

        for (int i = 0; i < lines.Length - 1; i++)
        {
            string line = System.Text.RegularExpressions.Regex.Replace(lines[i], @"\s+", " ").Trim(' ', '#', '-', '*', '·', ':', '：');
            if (!line.Equals(label, StringComparison.OrdinalIgnoreCase)) continue;

            for (int j = i + 1; j < lines.Length && j <= i + 3; j++)
            {
                string value = System.Text.RegularExpressions.Regex.Replace(lines[j], @"\s+", " ").Trim(' ', '-', '*', '·', ':', '：');
                if (value.Length >= 2) return value;
            }
        }
        return null;
    }

    private static string CleanInputDataAnalysisLine(string? value)
    {
        string text = System.Text.RegularExpressions.Regex.Replace(value ?? "", @"\s+", " ").Trim(' ', '-', '*', '·', ':', '：');
        if (TrySplitMalformedInputDataLabelRow(text, out _, out string rowValue))
            return rowValue;

        int colon = FirstInputDataColonIndex(text);
        if (colon <= 0 || colon > 40) return text;

        string prefix = text[..colon].Trim();
        string remainder = text[(colon + 1)..].Trim(' ', '-', '*', '·', ':', '：');
        if (remainder.Length < 2) return text;
        return LooksLikeMalformedInputDataLabel(prefix) ? remainder : text;
    }

    private static string? ExtractMalformedInputDataLabelValue(string text, int ordinal)
    {
        int index = 0;
        foreach (string rawLine in (text ?? "").Replace("\r\n", "\n").Split('\n').Take(10))
        {
            string line = System.Text.RegularExpressions.Regex.Replace(rawLine, @"\s+", " ").Trim(' ', '-', '*', '·');
            if (!TryGetMalformedInputDataLabelValue(line, out string value)) continue;
            index++;
            if (index != ordinal) continue;
            if (value.Length >= 2 && value.Length <= 240)
                return value;
        }
        return null;
    }

    private static int MalformedInputDataOrdinalForLabels(string[] labels)
    {
        if (labels.Any(l => l.Contains("대상", StringComparison.OrdinalIgnoreCase)
                            || l.Contains("target", StringComparison.OrdinalIgnoreCase)))
            return 4;
        if (labels.Any(l => l.Contains("검토", StringComparison.OrdinalIgnoreCase)
                            || l.Contains("확인", StringComparison.OrdinalIgnoreCase)
                            || l.Contains("review", StringComparison.OrdinalIgnoreCase)
                            || l.Contains("check", StringComparison.OrdinalIgnoreCase)))
            return 5;
        return 0;
    }

    private static bool TryGetMalformedInputDataLabelValue(string line, out string value)
    {
        value = "";
        string text = System.Text.RegularExpressions.Regex.Replace(line ?? "", @"\s+", " ").Trim(' ', '-', '*', '·');
        if (TrySplitMalformedInputDataLabelRow(text, out _, out value))
            return true;

        int colon = FirstInputDataColonIndex(text);
        if (colon <= 0 || colon > 40) return false;
        if (!LooksLikeMalformedInputDataLabel(text[..colon])) return false;

        value = text[(colon + 1)..].Trim(' ', '-', '*', '·', ':', '：');
        return value.Length >= 2;
    }

    private static bool TrySplitMalformedInputDataLabelRow(string line, out string label, out string value)
    {
        label = "";
        value = "";
        string text = System.Text.RegularExpressions.Regex.Replace(line ?? "", @"\s+", " ").Trim(' ', '-', '*', '·');
        if (text.Length == 0) return false;

        System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(text, @"^(\S{2,40})\s+(.{2,240})$");
        if (!m.Success) return false;
        string first = m.Groups[1].Value.Trim();
        if (!LooksLikeMalformedInputDataLabel(first)) return false;

        label = first;
        value = m.Groups[2].Value.Trim(' ', '-', '*', '·', ':', '：');
        return value.Length >= 2;
    }

    private static int FirstInputDataColonIndex(string text)
    {
        int ascii = (text ?? "").IndexOf(':');
        int wide = (text ?? "").IndexOf('：');
        if (ascii < 0) return wide;
        if (wide < 0) return ascii;
        return Math.Min(ascii, wide);
    }

    private static bool LooksLikeMalformedInputDataLabel(string value)
    {
        string text = System.Text.RegularExpressions.Regex.Replace(value ?? "", @"\s+", "");
        if (text.Length == 0 || text.Length > 40) return false;
        if (text.Contains('?') || text.Contains('\uFFFD')) return true;

        foreach (char ch in text)
        {
            if (char.IsControl(ch)) return true;
            if (ch is >= '\u0080' and <= '\u009F') return true;
            if (ch is >= '\u4E00' and <= '\u9FFF') return true;
            if (ch is >= '\uF900' and <= '\uFAFF') return true;
        }
        return false;
    }

    private static string StripHtmlForAnalysisText(string? html)
    {
        if (string.IsNullOrWhiteSpace(html)) return "";
        string text = System.Text.RegularExpressions.Regex.Replace(
            html,
            @"<\s*(br|p|div|li|tr|h[1-6])\b[^>]*>",
            "\n",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        text = System.Text.RegularExpressions.Regex.Replace(
            text,
            "<[^>]+>",
            " ",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        return System.Net.WebUtility.HtmlDecode(text);
    }

    private static bool LooksLikeInputDataSourceFile(string value)
    {
        string s = (value ?? "").Trim();
        if (s.Contains(".xlsx", StringComparison.OrdinalIgnoreCase)
            || s.Contains(".xlsm", StringComparison.OrdinalIgnoreCase)
            || s.Contains("_clean_textonly", StringComparison.OrdinalIgnoreCase)
            || s.Contains("_textonly", StringComparison.OrdinalIgnoreCase))
            return true;

        if (System.Text.RegularExpressions.Regex.IsMatch(s, @"^[0-9a-f]{12,16}[_-]", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            return true;

        return s.Length > 45
               && System.Text.RegularExpressions.Regex.IsMatch(s, @"\b(report|result|test|summary)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
               && System.Text.RegularExpressions.Regex.IsMatch(s, @"\b(clean|textonly|copy|date|\d{4}[.-]\d{1,2}[.-]\d{1,2})\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    }

    private static string TruncateForDb(string? value, int maxLength)
    {
        string text = (value ?? "").Trim();
        if (text.Length <= maxLength) return text;
        return text[..maxLength].TrimEnd();
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

    // ?? AI_EXCEL_PROC schema read path ?????????????????????????????????????
    // Used by DataInferenceDbPage to render the new-CLI bundle when a row was
    // processed via the AI_EXCEL_PROC schema (writes AiDocuments + children). Returns
    // null when the dataset has no AiDocuments row ??the page falls back to
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
                       ReportType, ReportDate, Confidence, SchemaVersion, RawJson,
                       ContentJson, GeneratedReportMarkdown, RelatedDefectsJson, PartsJson, ProcessesJson
                FROM AiDocuments
                WHERE SourceDataset=@n
                  AND LOWER(COALESCE(PrimaryDefect,'')) NOT LIKE '%auto-extracted%'
                  AND LOWER(COALESCE(RawJson,'')) NOT LIKE '%_batch_auto.py%'
                  AND LOWER(COALESCE(RawJson,'')) NOT LIKE '%see workbook title/purpose%'
                  AND LOWER(COALESCE(RawJson,'')) NOT LIKE '%workbook stored but extraction surfaced narrative only%'
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
                SchemaVersion  = rd.IsDBNull(9) ? "" : rd.GetString(9),
                RawJson        = rd.IsDBNull(10) ? "" : rd.GetString(10),
                Content        = ReadJsonStringList(rd.IsDBNull(11) ? "" : rd.GetString(11)),
                GeneratedReportMarkdown = rd.IsDBNull(12) ? "" : rd.GetString(12),
                RelatedDefects = ReadJsonStringList(rd.IsDBNull(13) ? "" : rd.GetString(13)),
                Parts          = ReadJsonStringList(rd.IsDBNull(14) ? "" : rd.GetString(14)),
                Processes      = ReadJsonStringList(rd.IsDBNull(15) ? "" : rd.GetString(15)),
            };
            bundle.AnalysisHtml = JsonStringFromRaw(bundle.RawJson, "analysisHtml");
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

        // 5) Translations (ko/en/vi) ??narrative-only
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT Lang, Title, Purpose, ContentJson, GeneratedReportMarkdown FROM AiDocumentTranslations WHERE DocumentId=@d;";
            cmd.Parameters.AddWithValue("@d", docId);
            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                bundle.DocTranslations[rd.GetString(0)] = new AiDocTranslationRow
                {
                    Title   = rd.GetString(1),
                    Purpose = rd.GetString(2),
                    Content = ReadJsonStringList(rd.IsDBNull(3) ? "" : rd.GetString(3)),
                    GeneratedReportMarkdown = rd.IsDBNull(4) ? "" : rd.GetString(4),
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

    private static string JsonStringFromRaw(string? json, string propertyName)
    {
        if (string.IsNullOrWhiteSpace(json) || string.IsNullOrWhiteSpace(propertyName)) return "";
        try
        {
            using JsonDocument doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty(propertyName, out JsonElement value)
                && value.ValueKind == JsonValueKind.String)
            {
                return value.GetString() ?? "";
            }
        }
        catch { }
        return "";
    }

    // ?? Dataset summary translations (multi-language) ?????????????????????????

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

    // ?? AskAi history ?????????????????????????????????????????????????????????

    public long SaveAskAiHistory(
        string question,
        string productTypeFilter,
        string overall,
        string perDatasetJson,
        string translationsJson = "{}")
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO AskAiHistory (Question, ProductTypeFilter, Overall, PerDatasetJson, TranslationsJson, CreatedAt)
            VALUES (@q, @pt, @o, @p, @tr, @c);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@q",  question ?? "");
        cmd.Parameters.AddWithValue("@pt", productTypeFilter ?? "");
        cmd.Parameters.AddWithValue("@o",  overall ?? "");
        cmd.Parameters.AddWithValue("@p",  perDatasetJson ?? "[]");
        cmd.Parameters.AddWithValue("@tr", translationsJson ?? "{}");
        cmd.Parameters.AddWithValue("@c",  DateTime.UtcNow.ToString("o"));
        return (long)(cmd.ExecuteScalar() ?? 0L);
    }

    public void UpdateAskAiHistoryTranslations(long id, string translationsJson)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE AskAiHistory SET TranslationsJson=@tr WHERE Id=@i;";
        cmd.Parameters.AddWithValue("@tr", translationsJson ?? "{}");
        cmd.Parameters.AddWithValue("@i", id);
        cmd.ExecuteNonQuery();
    }

    public List<AskAiHistoryRecord> GetAskAiHistory(int limit = 100)
    {
        var list = new List<AskAiHistoryRecord>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Question, ProductTypeFilter, Overall, PerDatasetJson, TranslationsJson, CreatedAt
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
                r.GetString(3), r.GetString(4), r.GetString(5), r.GetString(6)));
        }
        return list;
    }

    public AskAiHistoryRecord? GetAskAiHistoryById(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Question, ProductTypeFilter, Overall, PerDatasetJson, TranslationsJson, CreatedAt
            FROM AskAiHistory WHERE Id=@i;
            """;
        cmd.Parameters.AddWithValue("@i", id);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return null;
        return new AskAiHistoryRecord(
            r.GetInt64(0), r.GetString(1), r.GetString(2),
            r.GetString(3), r.GetString(4), r.GetString(5), r.GetString(6));
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

    // Daily Test Data projects

    public List<DailyTestDataItemRecord> GetDailyTestDataItems()
    {
        var list = new List<DailyTestDataItemRecord>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Name, DataText, PromptText, ParametersJson, AnalysisMarkdown, AnalysisHtml, CreatedAt, UpdatedAt, AnalyzedAt, HtmlGeneratedAt
            FROM DailyTestDataItems
            ORDER BY UpdatedAt DESC, Id DESC;
            """;
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(ReadDailyTestDataItem(r));
        return list;
    }

    public DailyTestDataItemRecord? GetDailyTestDataItem(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Name, DataText, PromptText, ParametersJson, AnalysisMarkdown, AnalysisHtml, CreatedAt, UpdatedAt, AnalyzedAt, HtmlGeneratedAt
            FROM DailyTestDataItems
            WHERE Id=@id;
            """;
        cmd.Parameters.AddWithValue("@id", id);
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? ReadDailyTestDataItem(r) : null;
    }

    public long CreateDailyTestDataItem(string name)
    {
        string now = DateTime.UtcNow.ToString("o");
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DailyTestDataItems (Name, CreatedAt, UpdatedAt)
            VALUES (@n, @c, @u);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@n", name?.Trim() ?? "");
        cmd.Parameters.AddWithValue("@c", now);
        cmd.Parameters.AddWithValue("@u", now);
        return (long)(cmd.ExecuteScalar() ?? 0L);
    }

    public void UpdateDailyTestDataItem(long id, string name, string dataText, string promptText)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE DailyTestDataItems
            SET Name=@n, DataText=@d, PromptText=@p, UpdatedAt=@u
            WHERE Id=@id;
            """;
        cmd.Parameters.AddWithValue("@n", name?.Trim() ?? "");
        cmd.Parameters.AddWithValue("@d", dataText ?? "");
        cmd.Parameters.AddWithValue("@p", promptText ?? "");
        cmd.Parameters.AddWithValue("@u", DateTime.UtcNow.ToString("o"));
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public void StartDailyTestDataNewInput(long id, string name)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE DailyTestDataItems
            SET Name=@n,
                DataText='',
                UpdatedAt=@u
            WHERE Id=@id;
            """;
        cmd.Parameters.AddWithValue("@n", name?.Trim() ?? "");
        cmd.Parameters.AddWithValue("@u", DateTime.UtcNow.ToString("o"));
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public void SaveDailyTestDataAnalysis(
        long id,
        string dataText,
        string promptText,
        string parametersJson,
        string analysisMarkdown)
    {
        string now = DateTime.UtcNow.ToString("o");
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE DailyTestDataItems
            SET DataText=@d,
                PromptText=@p,
                ParametersJson=@j,
                AnalysisMarkdown=@a,
                AnalysisHtml='',
                UpdatedAt=@u,
                AnalyzedAt=@u,
                HtmlGeneratedAt=''
            WHERE Id=@id;
            """;
        cmd.Parameters.AddWithValue("@d", dataText ?? "");
        cmd.Parameters.AddWithValue("@p", promptText ?? "");
        cmd.Parameters.AddWithValue("@j", string.IsNullOrWhiteSpace(parametersJson) ? "{}" : parametersJson);
        cmd.Parameters.AddWithValue("@a", analysisMarkdown ?? "");
        cmd.Parameters.AddWithValue("@u", now);
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public void SaveDailyTestDataAnalysisHtml(long id, string analysisHtml)
    {
        string now = DateTime.UtcNow.ToString("o");
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE DailyTestDataItems
            SET AnalysisHtml=@h,
                UpdatedAt=@u,
                HtmlGeneratedAt=@u
            WHERE Id=@id;
            """;
        cmd.Parameters.AddWithValue("@h", analysisHtml ?? "");
        cmd.Parameters.AddWithValue("@u", now);
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public void SaveDailyTestDataAnalysisResult(
        long id,
        string dataText,
        string promptText,
        string parametersJson,
        string analysisMarkdown,
        string analysisHtml)
    {
        string now = DateTime.UtcNow.ToString("o");
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            UPDATE DailyTestDataItems
            SET DataText=@d,
                PromptText=@p,
                ParametersJson=@j,
                AnalysisMarkdown=@a,
                AnalysisHtml=@h,
                UpdatedAt=@u,
                AnalyzedAt=@u,
                HtmlGeneratedAt=@u
            WHERE Id=@id;
            """;
        cmd.Parameters.AddWithValue("@d", dataText ?? "");
        cmd.Parameters.AddWithValue("@p", promptText ?? "");
        cmd.Parameters.AddWithValue("@j", string.IsNullOrWhiteSpace(parametersJson) ? "{}" : parametersJson);
        cmd.Parameters.AddWithValue("@a", analysisMarkdown ?? "");
        cmd.Parameters.AddWithValue("@h", analysisHtml ?? "");
        cmd.Parameters.AddWithValue("@u", now);
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    public List<DailyTestDataHistoryRecord> GetDailyTestDataHistory(long itemId)
    {
        var list = new List<DailyTestDataHistoryRecord>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, ItemId, ItemName, DataText, PromptText, ParametersJson, AnalysisMarkdown, AnalysisHtml, CreatedAt
            FROM DailyTestDataHistory
            WHERE ItemId=@id
            ORDER BY CreatedAt DESC, Id DESC;
            """;
        cmd.Parameters.AddWithValue("@id", itemId);
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(ReadDailyTestDataHistory(r));
        return list;
    }

    public long SaveDailyTestDataHistory(
        long itemId,
        string itemName,
        string dataText,
        string promptText,
        string parametersJson,
        string analysisMarkdown,
        string analysisHtml,
        string createdAt)
    {
        string at = string.IsNullOrWhiteSpace(createdAt) ? DateTime.UtcNow.ToString("o") : createdAt;
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO DailyTestDataHistory
                (ItemId, ItemName, DataText, PromptText, ParametersJson, AnalysisMarkdown, AnalysisHtml, CreatedAt)
            VALUES
                (@item, @name, @data, @prompt, @params, @md, @html, @created);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@item", itemId);
        cmd.Parameters.AddWithValue("@name", itemName ?? "");
        cmd.Parameters.AddWithValue("@data", dataText ?? "");
        cmd.Parameters.AddWithValue("@prompt", promptText ?? "");
        cmd.Parameters.AddWithValue("@params", string.IsNullOrWhiteSpace(parametersJson) ? "{}" : parametersJson);
        cmd.Parameters.AddWithValue("@md", analysisMarkdown ?? "");
        cmd.Parameters.AddWithValue("@html", analysisHtml ?? "");
        cmd.Parameters.AddWithValue("@created", at);
        return (long)(cmd.ExecuteScalar() ?? 0L);
    }

    public void DeleteDailyTestDataItem(long id)
    {
        using SqliteConnection conn = OpenConnection();
        using (SqliteCommand hist = conn.CreateCommand())
        {
            hist.CommandText = "DELETE FROM DailyTestDataHistory WHERE ItemId=@id;";
            hist.Parameters.AddWithValue("@id", id);
            hist.ExecuteNonQuery();
        }
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM DailyTestDataItems WHERE Id=@id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    private static DailyTestDataItemRecord ReadDailyTestDataItem(SqliteDataReader r)
        => new(
            r.GetInt64(0),
            r.GetString(1),
            r.GetString(2),
            r.GetString(3),
            r.GetString(4),
            r.GetString(5),
            r.GetString(6),
            r.GetString(7),
            r.GetString(8),
            r.GetString(9),
            r.GetString(10));

    private static DailyTestDataHistoryRecord ReadDailyTestDataHistory(SqliteDataReader r)
        => new(
            r.GetInt64(0),
            r.GetInt64(1),
            r.GetString(2),
            r.GetString(3),
            r.GetString(4),
            r.GetString(5),
            r.GetString(6),
            r.GetString(7),
            r.GetString(8));

    public List<ModelAnalysisReportRecord> GetAskAiReviewRecords(string productTypeFilter = "", int limit = 300)
    {
        var list = new List<ModelAnalysisReportRecord>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT r.DatasetName, r.ProductType, r.ReportDate,
                   COALESCE(d.DocumentId, '') AS DocumentId,
                   COALESCE(d.Title, '') AS Title,
                   COALESCE(d.Purpose, '') AS Purpose,
                   COALESCE(d.ReportType, '') AS ReportType,
                   COALESCE(d.PrimaryDefect, '') AS PrimaryDefect,
                   COALESCE(d.GeneratedReportMarkdown, '') AS GeneratedReportMarkdown,
                   COALESCE(d.UpdatedAt, '') AS AiUpdatedAt,
                   COALESCE((SELECT COUNT(*) FROM AiResults ar WHERE ar.DocumentId=d.DocumentId), 0) AS ResultCount,
                   COALESCE((SELECT COUNT(*) FROM AiConclusions ac WHERE ac.DocumentId=d.DocumentId), 0) AS ConclusionCount
            FROM RawReports r
            JOIN AiDocuments d
              ON d.DocumentId = (
                    SELECT d2.DocumentId
                    FROM AiDocuments d2
                    WHERE d2.SourceDataset = r.DatasetName
                      AND LOWER(COALESCE(d2.PrimaryDefect,'')) NOT LIKE '%auto-extracted%'
                      AND LOWER(COALESCE(d2.RawJson,'')) NOT LIKE '%_batch_auto.py%'
                      AND LOWER(COALESCE(d2.RawJson,'')) NOT LIKE '%see workbook title/purpose%'
                      AND LOWER(COALESCE(d2.RawJson,'')) NOT LIKE '%workbook stored but extraction surfaced narrative only%'
                      AND LENGTH(TRIM(COALESCE(d2.GeneratedReportMarkdown,''))) > 0
                    ORDER BY d2.UpdatedAt DESC
                    LIMIT 1
              )
            WHERE r.BatchExcluded = 0
              AND (@p = '' OR r.ProductType = @p)
            ORDER BY d.UpdatedAt DESC, r.ReportDate DESC, r.DatasetName COLLATE NOCASE
            LIMIT @l;
            """;
        cmd.Parameters.AddWithValue("@p", productTypeFilter ?? "");
        cmd.Parameters.AddWithValue("@l", Math.Max(1, limit));

        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new ModelAnalysisReportRecord(
                r.GetString(0),
                r.GetString(1),
                r.GetString(2),
                r.GetString(3),
                r.GetString(4),
                r.GetString(5),
                r.GetString(6),
                r.GetString(7),
                r.GetString(8),
                r.GetString(9),
                Convert.ToInt32(r.GetValue(10)),
                Convert.ToInt32(r.GetValue(11))));
        }
        return list;
    }

    // Model-level AI analysis built from per-report AI markdown.

    public List<ModelAnalysisGroupRecord> GetModelAnalysisGroups()
    {
        var list = new List<ModelAnalysisGroupRecord>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT r.ProductType,
                   COUNT(*) AS ReportCount,
                   SUM(CASE WHEN d.DocumentId IS NOT NULL
                             AND LENGTH(TRIM(COALESCE(d.GeneratedReportMarkdown,''))) > 0
                            THEN 1 ELSE 0 END) AS AiReportCount,
                   COALESCE(MAX(r.ReportDate), '') AS LatestReportDate,
                   COALESCE(MAX(d.UpdatedAt), '') AS LatestAiUpdatedAt,
                   COALESCE((SELECT MAX(a.CreatedAt)
                               FROM AiModelAnalyses a
                              WHERE a.ProductType = r.ProductType), '') AS LatestAnalysisAt
            FROM RawReports r
            LEFT JOIN AiDocuments d
              ON d.DocumentId = (
                    SELECT d2.DocumentId
                    FROM AiDocuments d2
                    WHERE d2.SourceDataset = r.DatasetName
                      AND LOWER(COALESCE(d2.PrimaryDefect,'')) NOT LIKE '%auto-extracted%'
                      AND LOWER(COALESCE(d2.RawJson,'')) NOT LIKE '%_batch_auto.py%'
                      AND LOWER(COALESCE(d2.RawJson,'')) NOT LIKE '%see workbook title/purpose%'
                      AND LOWER(COALESCE(d2.RawJson,'')) NOT LIKE '%workbook stored but extraction surfaced narrative only%'
                    ORDER BY d2.UpdatedAt DESC
                    LIMIT 1
              )
            WHERE r.BatchExcluded = 0
              AND LENGTH(TRIM(COALESCE(r.ProductType,''))) > 0
            GROUP BY r.ProductType
            ORDER BY r.ProductType COLLATE NOCASE;
            """;

        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new ModelAnalysisGroupRecord(
                r.GetString(0),
                Convert.ToInt32(r.GetValue(1)),
                Convert.ToInt32(r.GetValue(2)),
                r.IsDBNull(3) ? "" : r.GetString(3),
                r.IsDBNull(4) ? "" : r.GetString(4),
                r.IsDBNull(5) ? "" : r.GetString(5)));
        }
        return list;
    }

    public List<ModelAnalysisReportRecord> GetModelAnalysisReports(string productType)
    {
        var list = new List<ModelAnalysisReportRecord>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT r.DatasetName, r.ProductType, r.ReportDate,
                   COALESCE(d.DocumentId, '') AS DocumentId,
                   COALESCE(d.Title, '') AS Title,
                   COALESCE(d.Purpose, '') AS Purpose,
                   COALESCE(d.ReportType, '') AS ReportType,
                   COALESCE(d.PrimaryDefect, '') AS PrimaryDefect,
                   COALESCE(d.GeneratedReportMarkdown, '') AS GeneratedReportMarkdown,
                   COALESCE(d.UpdatedAt, '') AS AiUpdatedAt,
                   COALESCE((SELECT COUNT(*) FROM AiResults ar WHERE ar.DocumentId=d.DocumentId), 0) AS ResultCount,
                   COALESCE((SELECT COUNT(*) FROM AiConclusions ac WHERE ac.DocumentId=d.DocumentId), 0) AS ConclusionCount
            FROM RawReports r
            LEFT JOIN AiDocuments d
              ON d.DocumentId = (
                    SELECT d2.DocumentId
                    FROM AiDocuments d2
                    WHERE d2.SourceDataset = r.DatasetName
                      AND LOWER(COALESCE(d2.PrimaryDefect,'')) NOT LIKE '%auto-extracted%'
                      AND LOWER(COALESCE(d2.RawJson,'')) NOT LIKE '%_batch_auto.py%'
                      AND LOWER(COALESCE(d2.RawJson,'')) NOT LIKE '%see workbook title/purpose%'
                      AND LOWER(COALESCE(d2.RawJson,'')) NOT LIKE '%workbook stored but extraction surfaced narrative only%'
                    ORDER BY d2.UpdatedAt DESC
                    LIMIT 1
              )
            WHERE r.BatchExcluded = 0
              AND r.ProductType = @p
            ORDER BY r.ReportDate DESC, r.DatasetName COLLATE NOCASE;
            """;
        cmd.Parameters.AddWithValue("@p", productType ?? "");

        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new ModelAnalysisReportRecord(
                r.GetString(0),
                r.GetString(1),
                r.GetString(2),
                r.GetString(3),
                r.GetString(4),
                r.GetString(5),
                r.GetString(6),
                r.GetString(7),
                r.GetString(8),
                r.GetString(9),
                Convert.ToInt32(r.GetValue(10)),
                Convert.ToInt32(r.GetValue(11))));
        }
        return list;
    }

    public long SaveModelAnalysis(string productType, string analysisMode, string language,
                                  int reportCount, string includedDatasetsJson,
                                  string analysisMarkdown, string analysisTableMarkdown,
                                  string sourceContextHash)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO AiModelAnalyses
                (ProductType, AnalysisMode, Language, ReportCount,
                 IncludedDatasetsJson, AnalysisMarkdown, AnalysisTableMarkdown, SourceContextHash, CreatedAt)
            VALUES (@p, @m, @l, @c, @d, @a, @ta, @h, @t);
            SELECT last_insert_rowid();
            """;
        cmd.Parameters.AddWithValue("@p", productType ?? "");
        cmd.Parameters.AddWithValue("@m", analysisMode ?? "");
        cmd.Parameters.AddWithValue("@l", language ?? "");
        cmd.Parameters.AddWithValue("@c", reportCount);
        cmd.Parameters.AddWithValue("@d", includedDatasetsJson ?? "[]");
        cmd.Parameters.AddWithValue("@a", analysisMarkdown ?? "");
        cmd.Parameters.AddWithValue("@ta", analysisTableMarkdown ?? "");
        cmd.Parameters.AddWithValue("@h", sourceContextHash ?? "");
        cmd.Parameters.AddWithValue("@t", DateTime.UtcNow.ToString("o"));
        return (long)(cmd.ExecuteScalar() ?? 0L);
    }

    public AiModelAnalysisRecord? GetLatestModelAnalysis(string productType)
    {
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, ProductType, AnalysisMode, Language, ReportCount,
                   IncludedDatasetsJson, AnalysisMarkdown, AnalysisTableMarkdown, SourceContextHash, CreatedAt
            FROM AiModelAnalyses
            WHERE ProductType=@p
            ORDER BY CreatedAt DESC, Id DESC
            LIMIT 1;
            """;
        cmd.Parameters.AddWithValue("@p", productType ?? "");
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read() ? ReadAiModelAnalysis(r) : null;
    }

    public List<AiModelAnalysisRecord> GetModelAnalysisHistory(string productType, int limit = 30)
    {
        var list = new List<AiModelAnalysisRecord>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, ProductType, AnalysisMode, Language, ReportCount,
                   IncludedDatasetsJson, AnalysisMarkdown, AnalysisTableMarkdown, SourceContextHash, CreatedAt
            FROM AiModelAnalyses
            WHERE ProductType=@p
            ORDER BY CreatedAt DESC, Id DESC
            LIMIT @lim;
            """;
        cmd.Parameters.AddWithValue("@p", productType ?? "");
        cmd.Parameters.AddWithValue("@lim", limit);
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read()) list.Add(ReadAiModelAnalysis(r));
        return list;
    }

    private static AiModelAnalysisRecord ReadAiModelAnalysis(SqliteDataReader r)
        => new(
            r.GetInt64(0),
            r.GetString(1),
            r.GetString(2),
            r.GetString(3),
            Convert.ToInt32(r.GetValue(4)),
            r.GetString(5),
            r.GetString(6),
            r.GetString(7),
            r.GetString(8),
            r.GetString(9));

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

    // ?? Reports ???????????????????????????????????????????????????????????????

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

    // ?? Dataset management ????????????????????????????????????????????????????

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

    // ?? Users ?????????????????????????????????????????????????????????????????

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

    // ?? Menu permissions ??????????????????????????????????????????????????????

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

        using SqliteCommand grantModelAnalysis = conn.CreateCommand();
        grantModelAnalysis.Transaction = tx;
        grantModelAnalysis.CommandText = """
            INSERT OR IGNORE INTO MenuPermissions (Role, MenuId)
            SELECT DISTINCT Role, @newMenu
            FROM MenuPermissions
            WHERE MenuId IN (@analysisMenu, @askMenu);
            """;
        grantModelAnalysis.Parameters.AddWithValue("@newMenu", AppMenus.DiModelAnalysis);
        grantModelAnalysis.Parameters.AddWithValue("@analysisMenu", AppMenus.DiAnalysis);
        grantModelAnalysis.Parameters.AddWithValue("@askMenu", AppMenus.DiAsk);
        grantModelAnalysis.ExecuteNonQuery();

        using SqliteCommand grantCurrentProblem = conn.CreateCommand();
        grantCurrentProblem.Transaction = tx;
        grantCurrentProblem.CommandText = """
            INSERT OR IGNORE INTO MenuPermissions (Role, MenuId)
            SELECT DISTINCT Role, @newMenu
            FROM MenuPermissions
            WHERE MenuId IN (@dbMenu, @modelMenu, @askMenu);
            """;
        grantCurrentProblem.Parameters.AddWithValue("@newMenu", AppMenus.DiCurrentProblem);
        grantCurrentProblem.Parameters.AddWithValue("@dbMenu", AppMenus.DiDb);
        grantCurrentProblem.Parameters.AddWithValue("@modelMenu", AppMenus.DiModelAnalysis);
        grantCurrentProblem.Parameters.AddWithValue("@askMenu", AppMenus.DiAsk);
        grantCurrentProblem.ExecuteNonQuery();

        using SqliteCommand grantInputBatch = conn.CreateCommand();
        grantInputBatch.Transaction = tx;
        grantInputBatch.CommandText = """
            INSERT OR IGNORE INTO MenuPermissions (Role, MenuId)
            SELECT DISTINCT Role, @newMenu
            FROM MenuPermissions
            WHERE MenuId = @inputMenu;
            """;
        grantInputBatch.Parameters.AddWithValue("@newMenu", AppMenus.DiInputBatch);
        grantInputBatch.Parameters.AddWithValue("@inputMenu", AppMenus.DiInputTest);
        grantInputBatch.ExecuteNonQuery();

        using SqliteCommand grantAiPrompt = conn.CreateCommand();
        grantAiPrompt.Transaction = tx;
        grantAiPrompt.CommandText = """
            INSERT OR IGNORE INTO MenuPermissions (Role, MenuId)
            SELECT DISTINCT Role, @newMenu
            FROM MenuPermissions
            WHERE MenuId = @inputBatchMenu;
            """;
        grantAiPrompt.Parameters.AddWithValue("@newMenu", AppMenus.DiAiPrompt);
        grantAiPrompt.Parameters.AddWithValue("@inputBatchMenu", AppMenus.DiInputBatch);
        grantAiPrompt.ExecuteNonQuery();

        using SqliteCommand grantDailyTest = conn.CreateCommand();
        grantDailyTest.Transaction = tx;
        grantDailyTest.CommandText = """
            INSERT OR IGNORE INTO MenuPermissions (Role, MenuId)
            SELECT DISTINCT Role, @newMenu
            FROM MenuPermissions
            WHERE MenuId IN (@inputMenu, @inputBatchMenu, @dbMenu, @askMenu, @modelMenu);
            """;
        grantDailyTest.Parameters.AddWithValue("@newMenu", AppMenus.DailyTestInput);
        grantDailyTest.Parameters.AddWithValue("@inputMenu", AppMenus.DiInputTest);
        grantDailyTest.Parameters.AddWithValue("@inputBatchMenu", AppMenus.DiInputBatch);
        grantDailyTest.Parameters.AddWithValue("@dbMenu", AppMenus.DiDb);
        grantDailyTest.Parameters.AddWithValue("@askMenu", AppMenus.DiAsk);
        grantDailyTest.Parameters.AddWithValue("@modelMenu", AppMenus.DiModelAnalysis);
        grantDailyTest.ExecuteNonQuery();

        using SqliteCommand grantBmesTest3 = conn.CreateCommand();
        grantBmesTest3.Transaction = tx;
        grantBmesTest3.CommandText = """
            INSERT OR IGNORE INTO MenuPermissions (Role, MenuId)
            SELECT DISTINCT Role, @newMenu
            FROM MenuPermissions
            WHERE MenuId IN (@modelGroupMenu, @fcostMenu, @routingMenu);
            """;
        grantBmesTest3.Parameters.AddWithValue("@newMenu", AppMenus.BmesTest3);
        grantBmesTest3.Parameters.AddWithValue("@modelGroupMenu", AppMenus.BmesMakeModelGroup);
        grantBmesTest3.Parameters.AddWithValue("@fcostMenu", AppMenus.BmesFCost);
        grantBmesTest3.Parameters.AddWithValue("@routingMenu", AppMenus.BmesRoutingTable);
        grantBmesTest3.ExecuteNonQuery();

        using SqliteCommand grantBmesTest4 = conn.CreateCommand();
        grantBmesTest4.Transaction = tx;
        grantBmesTest4.CommandText = """
            INSERT OR IGNORE INTO MenuPermissions (Role, MenuId)
            SELECT DISTINCT Role, @newMenu
            FROM MenuPermissions
            WHERE MenuId IN (@test3Menu, @ngRateMenu, @fcostMenu);
            """;
        grantBmesTest4.Parameters.AddWithValue("@newMenu", AppMenus.BmesTest4);
        grantBmesTest4.Parameters.AddWithValue("@test3Menu", AppMenus.BmesTest3);
        grantBmesTest4.Parameters.AddWithValue("@ngRateMenu", AppMenus.NgRate);
        grantBmesTest4.Parameters.AddWithValue("@fcostMenu", AppMenus.BmesFCost);
        grantBmesTest4.ExecuteNonQuery();

        using SqliteCommand grantBmesTest5 = conn.CreateCommand();
        grantBmesTest5.Transaction = tx;
        grantBmesTest5.CommandText = """
            INSERT OR IGNORE INTO MenuPermissions (Role, MenuId)
            SELECT DISTINCT Role, @newMenu
            FROM MenuPermissions
            WHERE MenuId IN (@test3Menu, @test4Menu, @modelGroupMenu);
            """;
        grantBmesTest5.Parameters.AddWithValue("@newMenu", AppMenus.BmesTest5);
        grantBmesTest5.Parameters.AddWithValue("@test3Menu", AppMenus.BmesTest3);
        grantBmesTest5.Parameters.AddWithValue("@test4Menu", AppMenus.BmesTest4);
        grantBmesTest5.Parameters.AddWithValue("@modelGroupMenu", AppMenus.BmesMakeModelGroup);
        grantBmesTest5.ExecuteNonQuery();

        using SqliteCommand grantQrBakoData = conn.CreateCommand();
        grantQrBakoData.Transaction = tx;
        grantQrBakoData.CommandText = "INSERT OR IGNORE INTO MenuPermissions (Role, MenuId) VALUES (@role, @menu);";
        grantQrBakoData.Parameters.AddWithValue("@menu", AppMenus.QrBakoData);
        var qrRole = grantQrBakoData.Parameters.Add("@role", SqliteType.Text);
        foreach (string role in AppRoles.All)
        {
            qrRole.Value = role;
            grantQrBakoData.ExecuteNonQuery();
        }

        // DAILY REPORT doubles as the post-login landing page, so every role needs it
        // — otherwise "/" would dead-end for anyone whose permissions were customised
        // before this menu existed.
        using SqliteCommand grantDailyReport = conn.CreateCommand();
        grantDailyReport.Transaction = tx;
        grantDailyReport.CommandText = "INSERT OR IGNORE INTO MenuPermissions (Role, MenuId) VALUES (@role, @menu);";
        grantDailyReport.Parameters.AddWithValue("@menu", AppMenus.BmesDailyReport);
        var dailyReportRole = grantDailyReport.Parameters.Add("@role", SqliteType.Text);
        foreach (string role in AppRoles.All)
        {
            dailyReportRole.Value = role;
            grantDailyReport.ExecuteNonQuery();
        }

        // PC Download only hands out the standalone desktop installer, so every role gets it
        // — including roles whose permissions were customised before this menu existed.
        using SqliteCommand grantPcDownload = conn.CreateCommand();
        grantPcDownload.Transaction = tx;
        grantPcDownload.CommandText = "INSERT OR IGNORE INTO MenuPermissions (Role, MenuId) VALUES (@role, @menu);";
        grantPcDownload.Parameters.AddWithValue("@menu", AppMenus.PcDownload);
        var pcDownloadRole = grantPcDownload.Parameters.Add("@role", SqliteType.Text);
        foreach (string role in AppRoles.All)
        {
            pcDownloadRole.Value = role;
            grantPcDownload.ExecuteNonQuery();
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

    // ?? BMES Materials ???????????????????????????????????????????????????????

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

    public List<BmesMaterial> GetBmesMaterials()
    {
        var list = new List<BmesMaterial>();
        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT *
            FROM BmesMaterials
            ORDER BY Maktx, Mtype, Matnr;
            """;
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(ReadBmesMaterial(r));
        return list;
    }

    // ?? Mtype ??Category name map ????????????????????????????????????????????

    public Dictionary<string, string> GetBmesMaterialNames(IEnumerable<string> matnrs)
    {
        var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var keys = matnrs
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (keys.Count == 0)
            return names;

        using SqliteConnection conn = OpenConnection();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Maktx FROM BmesMaterials WHERE Matnr = @matnr LIMIT 1;";
        SqliteParameter pMatnr = cmd.Parameters.Add("@matnr", SqliteType.Text);

        foreach (string key in keys)
        {
            pMatnr.Value = key;
            object? value = cmd.ExecuteScalar();
            if (value is string name && !string.IsNullOrWhiteSpace(name))
                names[key] = name.Trim();
        }

        return names;
    }

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

    // ?? Model Groups ?????????????????????????????????????????????????????????

    // ?? DataInputAliases: user-curated token ??Material map ???????????????????

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

    private static void EnsureWeeklyReportFormSettingsTable(SqliteConnection conn)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS WeeklyReportFormSettings (
                RowKey          TEXT    PRIMARY KEY NOT NULL,
                IsVisible       INTEGER NOT NULL DEFAULT 1,
                DisplayMode     TEXT    NOT NULL DEFAULT 'B-GROUP',
                ProductName     TEXT    NOT NULL DEFAULT '',
                BaselineDec2025 REAL    NULL,
                BaselineApr2026 REAL    NULL,
                Target          REAL    NULL,
                Action          TEXT    NOT NULL DEFAULT '',
                SortOrder       INTEGER NOT NULL DEFAULT -1,
                UpdatedAt       TEXT    NOT NULL
            );
            """;
        cmd.ExecuteNonQuery();

        using SqliteCommand check = conn.CreateCommand();
        check.CommandText = "PRAGMA table_info(WeeklyReportFormSettings);";
        var columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (SqliteDataReader r = check.ExecuteReader())
        {
            while (r.Read())
            {
                string col = r.IsDBNull(1) ? string.Empty : r.GetString(1);
                if (!string.IsNullOrEmpty(col))
                    columns.Add(col);
            }
        }

        if (!columns.Contains("SortOrder"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE WeeklyReportFormSettings ADD COLUMN SortOrder INTEGER NOT NULL DEFAULT -1;";
            alter.ExecuteNonQuery();
        }

        if (!columns.Contains("DisplayMode"))
        {
            using SqliteCommand alter = conn.CreateCommand();
            alter.CommandText = "ALTER TABLE WeeklyReportFormSettings ADD COLUMN DisplayMode TEXT NOT NULL DEFAULT 'B-GROUP';";
            alter.ExecuteNonQuery();
        }
    }

    private static string NormalizeWeeklyReportDisplayMode(string? value)
    {
        string mode = (value ?? string.Empty).Trim();
        return mode.Equals("Group", StringComparison.OrdinalIgnoreCase)
            ? "Group"
            : "B-GROUP";
    }

    public Dictionary<string, WeeklyReportFormSettingRecord> GetWeeklyReportFormSettings()
    {
        var map = new Dictionary<string, WeeklyReportFormSettingRecord>(StringComparer.Ordinal);
        using SqliteConnection conn = OpenConnection();
        EnsureWeeklyReportFormSettingsTable(conn);

        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT RowKey, IsVisible, DisplayMode, ProductName, BaselineDec2025, BaselineApr2026, Target, Action, SortOrder, UpdatedAt
            FROM WeeklyReportFormSettings;
            """;
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string rowKey = r.IsDBNull(0) ? string.Empty : r.GetString(0);
            if (rowKey.Length == 0) continue;

            map[rowKey] = new WeeklyReportFormSettingRecord
            {
                RowKey          = rowKey,
                IsVisible       = !r.IsDBNull(1) && r.GetInt64(1) != 0,
                DisplayMode     = NormalizeWeeklyReportDisplayMode(r.IsDBNull(2) ? string.Empty : r.GetString(2)),
                ProductName     = r.IsDBNull(3) ? string.Empty : r.GetString(3),
                BaselineDec2025 = r.IsDBNull(4) ? null : r.GetDouble(4),
                BaselineApr2026 = r.IsDBNull(5) ? null : r.GetDouble(5),
                Target          = r.IsDBNull(6) ? null : r.GetDouble(6),
                Action          = r.IsDBNull(7) ? string.Empty : r.GetString(7),
                SortOrder       = r.IsDBNull(8) ? -1 : r.GetInt32(8),
                UpdatedAt       = r.IsDBNull(9) ? string.Empty : r.GetString(9),
            };
        }
        return map;
    }

    public void SaveWeeklyReportFormSettings(IReadOnlyList<WeeklyReportFormSettingRecord> settings)
    {
        using SqliteConnection conn = OpenConnection();
        EnsureWeeklyReportFormSettingsTable(conn);
        using SqliteTransaction tx = conn.BeginTransaction();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = """
            INSERT INTO WeeklyReportFormSettings
                (RowKey, IsVisible, DisplayMode, ProductName, BaselineDec2025, BaselineApr2026, Target, Action, SortOrder, UpdatedAt)
            VALUES
                (@key, @visible, @displayMode, @name, @dec, @apr, @target, @action, @sort, @updated)
            ON CONFLICT(RowKey) DO UPDATE SET
                IsVisible       = excluded.IsVisible,
                DisplayMode     = excluded.DisplayMode,
                ProductName     = excluded.ProductName,
                BaselineDec2025 = excluded.BaselineDec2025,
                BaselineApr2026 = excluded.BaselineApr2026,
                Target          = excluded.Target,
                Action          = excluded.Action,
                SortOrder       = excluded.SortOrder,
                UpdatedAt       = excluded.UpdatedAt;
            """;
        SqliteParameter pKey     = cmd.Parameters.Add("@key",     SqliteType.Text);
        SqliteParameter pVisible = cmd.Parameters.Add("@visible", SqliteType.Integer);
        SqliteParameter pMode    = cmd.Parameters.Add("@displayMode", SqliteType.Text);
        SqliteParameter pName    = cmd.Parameters.Add("@name",    SqliteType.Text);
        SqliteParameter pDec     = cmd.Parameters.Add("@dec",     SqliteType.Real);
        SqliteParameter pApr     = cmd.Parameters.Add("@apr",     SqliteType.Real);
        SqliteParameter pTarget  = cmd.Parameters.Add("@target",  SqliteType.Real);
        SqliteParameter pAction  = cmd.Parameters.Add("@action",  SqliteType.Text);
        SqliteParameter pSort    = cmd.Parameters.Add("@sort",    SqliteType.Integer);
        SqliteParameter pUpdated = cmd.Parameters.Add("@updated", SqliteType.Text);

        string now = DateTime.UtcNow.ToString("O");
        foreach (var setting in settings)
        {
            if (string.IsNullOrWhiteSpace(setting.RowKey)) continue;

            pKey.Value     = setting.RowKey;
            pVisible.Value = setting.IsVisible ? 1 : 0;
            pMode.Value    = NormalizeWeeklyReportDisplayMode(setting.DisplayMode);
            pName.Value    = setting.ProductName?.Trim() ?? string.Empty;
            pDec.Value     = setting.BaselineDec2025.HasValue ? (object)setting.BaselineDec2025.Value : DBNull.Value;
            pApr.Value     = setting.BaselineApr2026.HasValue ? (object)setting.BaselineApr2026.Value : DBNull.Value;
            pTarget.Value  = setting.Target.HasValue ? (object)setting.Target.Value : DBNull.Value;
            pAction.Value  = setting.Action?.Trim() ?? string.Empty;
            pSort.Value    = setting.SortOrder;
            pUpdated.Value = now;
            cmd.ExecuteNonQuery();
        }
        tx.Commit();
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

                // Legacy rows without material ??derive from LineShift (split at last '_').
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

    /// <summary>Sub-group path separator (unit separator control char ??impossible in user input).</summary>
    private const char SubPathSep = '';

    /// <summary>Navigate the sub-group tree by name-path, creating nodes as needed. Returns the leaf.</summary>
    private static SubGroupRecord ResolveOrCreateSubPath(List<SubGroupRecord> rootList, string path)
    {
        // Empty path ??the "default" (unnamed) sub-group at the top of this material's tree.
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

    // ?? One-shot JSON ??ModelGroups import ??????????????????????????????????????

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

