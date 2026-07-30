using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

public sealed class MicroSpeakerInputDataService
{
    private const string DefaultMicroSpeakerRoot = @"D:\000. MyWorks\005. Program\Repository\MicroSpeaker_ProductTech_DB";
    private const string DefaultExcelFolder = @"D:\000. MyWorks\test\result\InputDataFinish";
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = false };

    private readonly IConfiguration _config;
    private readonly IWebHostEnvironment _env;
    private readonly WebRepository _repo;

    public MicroSpeakerInputDataService(IConfiguration config, IWebHostEnvironment env, WebRepository repo)
    {
        _config = config;
        _env = env;
        _repo = repo;
    }

    public MicroSpeakerPaths ResolvePaths()
    {
        string? configuredRoot = _config["MicroSpeaker:ProjectRoot"];
        string siblingRoot = Path.GetFullPath(Path.Combine(_env.ContentRootPath, "..", "..", "MicroSpeaker_ProductTech_DB"));
        string projectRoot =
            !string.IsNullOrWhiteSpace(configuredRoot) ? configuredRoot.Trim() :
            Directory.Exists(siblingRoot) ? siblingRoot :
            DefaultMicroSpeakerRoot;

        projectRoot = Path.GetFullPath(projectRoot);
        string dbPath = _config["MicroSpeaker:DatabasePath"] ?? Path.Combine(projectRoot, "db", "InputDataFinish.sqlite");
        string dashboardPath = _config["MicroSpeaker:DashboardPath"] ?? Path.Combine(projectRoot, "db", "InputDataFinish_dashboard.html");
        string cliExePath = _config["MicroSpeaker:CliExePath"] ?? Path.Combine(projectRoot, "dist", "MicroSpeakerAnalysis.exe");
        string cliScriptPath = Path.Combine(projectRoot, "tools", "run_microspeaker_analysis.py");

        return new MicroSpeakerPaths(
            projectRoot,
            Path.GetFullPath(dbPath),
            Path.GetFullPath(dashboardPath),
            Path.GetFullPath(cliExePath),
            Path.GetFullPath(cliScriptPath));
    }

    public MicroSpeakerSourceFile? FindSourceFile(long fileId)
    {
        MicroSpeakerPaths paths = ResolvePaths();
        if (!File.Exists(paths.DatabasePath)) return null;

        using SqliteConnection conn = OpenReadOnly(paths.DatabasePath);
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT path, file_name FROM files WHERE file_id=@id LIMIT 1;";
        cmd.Parameters.AddWithValue("@id", fileId);
        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return null;

        string rawPath = S(r, 0);
        string fileName = S(r, 1);
        if (string.IsNullOrWhiteSpace(rawPath)) return null;

        string fullPath = Path.IsPathRooted(rawPath)
            ? rawPath
            : Path.Combine(paths.ProjectRoot, rawPath);
        fullPath = Path.GetFullPath(fullPath);
        string safeName = string.IsNullOrWhiteSpace(fileName) ? Path.GetFileName(fullPath) : fileName;
        return new MicroSpeakerSourceFile(fileId, fullPath, safeName);
    }

    public MicroSpeakerSummary GetSummary()
    {
        MicroSpeakerPaths paths = ResolvePaths();
        if (!File.Exists(paths.DatabasePath))
        {
            return new MicroSpeakerSummary(
                paths,
                DatabaseExists: false,
                LatestRun: null,
                FileCount: 0,
                OkFileCount: 0,
                FailedFileCount: 0,
                MetricCandidateCount: 0,
                MeasurementStatCount: 0,
                PairCount: 0,
                TermHitCount: 0,
                Structures: [],
                TopWorsenedPairs: [],
                TopImprovedPairs: []);
        }

        using SqliteConnection conn = OpenReadOnly(paths.DatabasePath);
        MicroSpeakerRunInfo? latestRun = ReadLatestRun(conn);
        long fileCount = ScalarLong(conn, "SELECT COUNT(*) FROM files;");
        long okFileCount = ScalarLong(conn, "SELECT COUNT(*) FROM files WHERE status='OK';");
        long failedFileCount = ScalarLong(conn, "SELECT COUNT(*) FROM files WHERE status<>'OK';");
        long metricCandidateCount = ScalarLong(conn, "SELECT COUNT(*) FROM metric_candidates;");
        long measurementStatCount = ScalarLong(conn, "SELECT COUNT(*) FROM measurement_stats;");
        long pairCount = ScalarLong(conn, "SELECT COUNT(*) FROM comparison_pairs;");
        long termHitCount = ScalarLong(conn, "SELECT COUNT(*) FROM term_hits;");

        return new MicroSpeakerSummary(
            paths,
            DatabaseExists: true,
            LatestRun: latestRun,
            FileCount: fileCount,
            OkFileCount: okFileCount,
            FailedFileCount: failedFileCount,
            MetricCandidateCount: metricCandidateCount,
            MeasurementStatCount: measurementStatCount,
            PairCount: pairCount,
            TermHitCount: termHitCount,
            Structures: ReadStructures(conn),
            TopWorsenedPairs: ReadTopPairs(conn, worsened: true),
            TopImprovedPairs: ReadTopPairs(conn, worsened: false));
    }

    public async Task<MicroSpeakerCliRunResult> RunIndexAsync(
        MicroSpeakerIndexRequest request,
        Action<string>? onOutput,
        CancellationToken cancellationToken)
    {
        MicroSpeakerPaths paths = ResolvePaths();
        if (!Directory.Exists(paths.ProjectRoot))
            return new MicroSpeakerCliRunResult(false, -1, "", $"Project root not found: {paths.ProjectRoot}");
        if (!Directory.Exists(request.InputDir))
            return new MicroSpeakerCliRunResult(false, -1, "", $"Excel folder not found: {request.InputDir}");

        string dataset = string.IsNullOrWhiteSpace(request.Dataset)
            ? new DirectoryInfo(request.InputDir).Name
            : request.Dataset.Trim();

        var psi = new ProcessStartInfo
        {
            WorkingDirectory = paths.ProjectRoot,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };

        if (File.Exists(paths.CliExePath))
        {
            psi.FileName = paths.CliExePath;
        }
        else if (File.Exists(paths.CliScriptPath))
        {
            psi.FileName = "python";
            psi.ArgumentList.Add(paths.CliScriptPath);
        }
        else
        {
            return new MicroSpeakerCliRunResult(false, -1, "", $"CLI not found: {paths.CliExePath}");
        }

        psi.ArgumentList.Add("index");
        psi.ArgumentList.Add("--input-dir");
        psi.ArgumentList.Add(request.InputDir);
        psi.ArgumentList.Add("--dataset");
        psi.ArgumentList.Add(dataset);
        if (request.Force)
            psi.ArgumentList.Add("--force");
        if (request.NoHtml)
            psi.ArgumentList.Add("--no-html");
        if (request.Limit > 0)
        {
            psi.ArgumentList.Add("--limit");
            psi.ArgumentList.Add(request.Limit.ToString());
        }

        var output = new StringBuilder();
        var error = new StringBuilder();
        using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };
        process.OutputDataReceived += (_, e) =>
        {
            if (e.Data is null) return;
            output.AppendLine(e.Data);
            onOutput?.Invoke(e.Data);
        };
        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data is null) return;
            error.AppendLine(e.Data);
            onOutput?.Invoke(e.Data);
        };

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();
        await process.WaitForExitAsync(cancellationToken);
        return new MicroSpeakerCliRunResult(process.ExitCode == 0, process.ExitCode, output.ToString(), error.ToString());
    }

    public MicroSpeakerImportResult ImportCurrentDatabase()
    {
        MicroSpeakerPaths paths = ResolvePaths();
        if (!File.Exists(paths.DatabasePath))
            throw new FileNotFoundException("MicroSpeaker database not found.", paths.DatabasePath);

        string jinoDbPath = _repo.GetDbPath();
        if (!File.Exists(jinoDbPath))
            throw new FileNotFoundException("JinoSupporter database not found.", jinoDbPath);

        string backupPath = BackupAiTables(jinoDbPath);
        using SqliteConnection micro = OpenReadOnly(paths.DatabasePath);
        using SqliteConnection jino = new($"Data Source={jinoDbPath}");
        jino.Open();

        using SqliteTransaction tx = jino.BeginTransaction();
        ClearAiAnalysisTables(jino, tx);

        var docMap = new Dictionary<long, MicroSpeakerDocRef>();
        long documents = ImportDocuments(micro, jino, tx, docMap);
        long datasetSummaries = ImportDatasetSummaries(micro, jino, tx);
        long pairRows = ImportPairs(micro, jino, tx, docMap);
        long metricRows = ImportMetricCandidates(micro, jino, tx, docMap);
        long measurementRows = ImportMeasurementStats(micro, jino, tx, docMap);
        tx.Commit();

        string firstAnalysisIndexPath = WriteFirstAnalysisArtifacts(micro);

        return new MicroSpeakerImportResult(
            jinoDbPath,
            paths.DatabasePath,
            backupPath,
            documents,
            datasetSummaries,
            pairRows,
            metricRows,
            measurementRows,
            firstAnalysisIndexPath);
    }

    private static SqliteConnection OpenReadOnly(string path)
    {
        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = path,
            Mode = SqliteOpenMode.ReadOnly,
            Cache = SqliteCacheMode.Shared
        };
        SqliteConnection conn = new(builder.ToString());
        conn.Open();
        return conn;
    }

    private static MicroSpeakerRunInfo? ReadLatestRun(SqliteConnection conn)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT run_id, dataset, input_dir, started_at, finished_at,
                   files_seen, files_processed, files_skipped, files_failed
            FROM runs
            ORDER BY run_id DESC
            LIMIT 1;
            """;
        using SqliteDataReader r = cmd.ExecuteReader();
        return r.Read()
            ? new MicroSpeakerRunInfo(
                r.GetInt64(0), S(r, 1), S(r, 2), S(r, 3), S(r, 4),
                r.GetInt64(5), r.GetInt64(6), r.GetInt64(7), r.GetInt64(8))
            : null;
    }

    private static List<MicroSpeakerStructureRow> ReadStructures(SqliteConnection conn)
    {
        var rows = new List<MicroSpeakerStructureRow>();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT structure_family, structure_confidence,
                   COUNT(*) AS file_count,
                   COALESCE(SUM(comparison_pair_count), 0) AS pair_count,
                   COALESCE(SUM(metric_candidate_count), 0) AS metric_count,
                   COALESCE(SUM(measurement_stat_count), 0) AS measurement_count
            FROM files
            GROUP BY structure_family, structure_confidence
            ORDER BY file_count DESC, structure_family;
            """;
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            rows.Add(new MicroSpeakerStructureRow(
                S(r, 0), S(r, 1), r.GetInt64(2), r.GetInt64(3), r.GetInt64(4), r.GetInt64(5)));
        }
        return rows;
    }

    private static List<MicroSpeakerPairPreview> ReadTopPairs(SqliteConnection conn, bool worsened)
    {
        var rows = new List<MicroSpeakerPairPreview>();
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = $"""
            SELECT f.file_name, p.compare_item, p.control_condition, p.test_condition,
                   p.control_rate, p.test_rate, p.delta_rate, p.effect_direction, p.evidence
            FROM comparison_pairs p
            JOIN files f ON f.file_id = p.file_id
            WHERE p.effect_direction = @direction
            ORDER BY p.delta_rate {(worsened ? "DESC" : "ASC")}
            LIMIT 10;
            """;
        cmd.Parameters.AddWithValue("@direction", worsened ? "WORSENED" : "IMPROVED");
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            rows.Add(new MicroSpeakerPairPreview(
                S(r, 0), S(r, 1), S(r, 2), S(r, 3),
                D(r, 4), D(r, 5), D(r, 6), S(r, 7), S(r, 8)));
        }
        return rows;
    }

    private string BackupAiTables(string jinoDbPath)
    {
        string backupDir = Path.Combine(Path.GetDirectoryName(jinoDbPath) ?? AppStoragePaths.RootDirectory, "backups");
        Directory.CreateDirectory(backupDir);
        string backupPath = Path.Combine(backupDir, $"process-review-ai-before-microspeaker-{DateTime.Now:yyyyMMdd_HHmmss_fff}.sqlite");
        using SqliteConnection conn = new($"Data Source={jinoDbPath}");
        conn.Open();
        using SqliteCommand attach = conn.CreateCommand();
        attach.CommandText = "ATTACH DATABASE @path AS backup;";
        attach.Parameters.AddWithValue("@path", backupPath);
        attach.ExecuteNonQuery();

        foreach (string table in AiBackupTables)
        {
            if (!TableExists(conn, table)) continue;
            using SqliteCommand copy = conn.CreateCommand();
            copy.CommandText = $"CREATE TABLE backup.{table} AS SELECT * FROM main.{table};";
            copy.ExecuteNonQuery();
        }

        using SqliteCommand detach = conn.CreateCommand();
        detach.CommandText = "DETACH DATABASE backup;";
        detach.ExecuteNonQuery();
        return backupPath;
    }

    private static void ClearAiAnalysisTables(SqliteConnection conn, SqliteTransaction tx)
    {
        foreach (string table in AiDeleteOrder)
        {
            if (!TableExists(conn, table)) continue;
            Execute(conn, tx, $"DELETE FROM {table};");
        }
    }

    private static long ImportDatasetSummaries(SqliteConnection micro, SqliteConnection jino, SqliteTransaction tx)
    {
        long count = 0;
        using SqliteCommand cmd = micro.CreateCommand();
        cmd.CommandText = """
            SELECT f.file_id, f.dataset, f.file_name, f.sheet_names, f.models, f.categories,
                   f.dates_found, f.structure_family, f.structure_confidence, f.term_summary,
                   f.metric_candidate_count, f.measurement_stat_count, f.comparison_pair_count,
                   f.processed_at,
                   COALESCE((SELECT COUNT(*) FROM comparison_pairs p WHERE p.file_id=f.file_id AND p.effect_direction='WORSENED'), 0) AS worsened_count,
                   COALESCE((SELECT COUNT(*) FROM comparison_pairs p WHERE p.file_id=f.file_id AND p.effect_direction='IMPROVED'), 0) AS improved_count,
                   COALESCE((SELECT COUNT(*) FROM comparison_pairs p WHERE p.file_id=f.file_id AND p.effect_direction='NO_CHANGE'), 0) AS nochange_count,
                   COALESCE((SELECT p.compare_item || ' | ' || p.control_condition || ' -> ' || p.test_condition
                               FROM comparison_pairs p
                              WHERE p.file_id=f.file_id
                              ORDER BY ABS(COALESCE(p.delta_rate, 0)) DESC
                              LIMIT 1), '') AS top_pair
            FROM files f
            ORDER BY f.file_id;
            """;

        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            long fileId = r.GetInt64(0);
            string dataset = S(r, 1);
            string fileName = S(r, 2);
            string sheetNames = S(r, 3);
            string models = S(r, 4);
            string categories = S(r, 5);
            string datesFound = S(r, 6);
            string structure = S(r, 7);
            string confidence = S(r, 8);
            string termSummary = S(r, 9);
            long metricCount = r.GetInt64(10);
            long measurementCount = r.GetInt64(11);
            long pairCount = r.GetInt64(12);
            string processedAt = S(r, 13);
            long worsened = r.GetInt64(14);
            long improved = r.GetInt64(15);
            long noChange = r.GetInt64(16);
            string topPair = S(r, 17);
            string sourceDataset = BuildSourceDataset(fileId, fileName);
            string productType = FirstToken(models, ';', "");
            string reportDate = FirstToken(datesFound, ';', "");
            string[] categoryTags = SplitSemi(categories);
            string[] termTags = SplitTermNames(termSummary);
            string[] tags = new[] { "microspeaker-cli", "microspeaker-db-v1", structure, confidence }
                .Concat(categoryTags.Take(8))
                .Concat(termTags.Take(8))
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            var evidence = new[]
            {
                new
                {
                    metric = string.IsNullOrWhiteSpace(topPair) ? "Structure" : "Top Pair",
                    baselineLabel = "MicroSpeaker",
                    baselineValue = structure,
                    variantLabel = "Evidence",
                    variantValue = topPair,
                    deltaText = $"WORSENED={worsened:N0}, IMPROVED={improved:N0}, NO_CHANGE={noChange:N0}",
                    deltaSign = worsened > 0 ? "up" : improved > 0 ? "down" : "no_change",
                    note = "Generated from MicroSpeaker CLI/GUI SQLite output."
                }
            };
            var actions = new[]
            {
                new { priority = 1, kind = "investigate", text = "Use MicroSpeaker Pair/Metric/Measurement rows as the source of truth. Do not use legacy AI analysis output." }
            };
            var context = new
            {
                process = structure,
                stage = FirstToken(categories, ';', structure),
                baselineReason = "Imported from MicroSpeaker CLI deterministic SQLite output."
            };

            Execute(jino, tx, """
                INSERT INTO DatasetSummary
                    (DatasetName, ProductType, Summary, KeyFindings, CreatedAt, Tags,
                     Purpose, TestConditions, RootCause, Decision, RecommendedAction,
                     Verdict, Headline, EvidenceJson, ActionsJson, ContextJson,
                     ReportType, DoeGridJson, TrendJson)
                VALUES
                    (@DatasetName, @ProductType, @Summary, @KeyFindings, @CreatedAt, @Tags,
                     @Purpose, @TestConditions, @RootCause, @Decision, @RecommendedAction,
                     'MICROSPEAKER_CLI_IMPORTED', @Headline, @EvidenceJson, @ActionsJson, @ContextJson,
                     'first_pass_index', '', '');
                """,
                ("@DatasetName", sourceDataset),
                ("@ProductType", productType),
                ("@Summary", $"MicroSpeaker CLI analysis. Structure={structure} ({confidence}), Pair={pairCount:N0}, Metric={metricCount:N0}, Measurement={measurementCount:N0}."),
                ("@KeyFindings", $"Sheets: {sheetNames}\nModels: {models}\nCategories: {categories}\nTerms: {termSummary}\nPair effect: WORSENED={worsened:N0}, IMPROVED={improved:N0}, NO_CHANGE={noChange:N0}\nTop pair: {topPair}"),
                ("@CreatedAt", string.IsNullOrWhiteSpace(processedAt) ? DateTime.UtcNow.ToString("O") : processedAt),
                ("@Tags", JsonSerializer.Serialize(tags, JsonOptions)),
                ("@Purpose", structure),
                ("@TestConditions", $"Sheets={sheetNames}; Models={models}; Dates={reportDate}"),
                ("@RootCause", string.IsNullOrWhiteSpace(categories) ? structure : categories),
                ("@Decision", $"PAIR WORSENED={worsened:N0}, IMPROVED={improved:N0}, NO_CHANGE={noChange:N0}"),
                ("@RecommendedAction", "Review only MicroSpeaker CLI/GUI Pair, Metric, and Measurement rows. Legacy AI analysis output has been removed."),
                ("@Headline", $"{structure} - Pair {pairCount:N0}"),
                ("@EvidenceJson", JsonSerializer.Serialize(evidence, JsonOptions)),
                ("@ActionsJson", JsonSerializer.Serialize(actions, JsonOptions)),
                ("@ContextJson", JsonSerializer.Serialize(context, JsonOptions)));

            count++;
        }
        return count;
    }

    private string WriteFirstAnalysisArtifacts(SqliteConnection micro)
    {
        string sampleDir = Path.Combine(_env.ContentRootPath, "App_Data", "ai-current-problem", "sample_ready");
        Directory.CreateDirectory(sampleDir);

        foreach (string fileName in new[] { "current_problem_search.html", "ai_batch_control.html", "ai_term_glossary.html" })
        {
            string stalePath = Path.Combine(sampleDir, fileName);
            if (File.Exists(stalePath))
                File.Delete(stalePath);
        }

        var indexRows = new List<object>();
        var structures = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        long totalPairs = 0;
        long totalMetrics = 0;
        long totalMeasurements = 0;

        using SqliteCommand cmd = micro.CreateCommand();
        cmd.CommandText = """
            SELECT file_id, file_name, sheet_names, models, categories, dates_found,
                   structure_family, structure_confidence, term_summary,
                   metric_candidate_count, measurement_stat_count, comparison_pair_count
            FROM files
            ORDER BY file_id;
            """;
        using SqliteDataReader r = cmd.ExecuteReader();
        int no = 0;
        while (r.Read())
        {
            no++;
            long fileId = r.GetInt64(0);
            string fileName = S(r, 1);
            string models = S(r, 3);
            string categories = S(r, 4);
            string datesFound = S(r, 5);
            string structure = S(r, 6);
            string confidence = S(r, 7);
            string termSummary = S(r, 8);
            long metricCount = r.GetInt64(9);
            long measurementCount = r.GetInt64(10);
            long pairCount = r.GetInt64(11);

            totalPairs += pairCount;
            totalMetrics += metricCount;
            totalMeasurements += measurementCount;
            structures[structure] = structures.GetValueOrDefault(structure) + 1;

            string[] targetDefects = SplitSemi(categories);
            string[] reviewItems = SplitTermNames(termSummary);
            string[] tags = new[] { "microspeaker-cli", "microspeaker-db-v1", structure, confidence }
                .Concat(targetDefects.Take(8))
                .Concat(reviewItems.Take(8))
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            indexRows.Add(new
            {
                datasetName = BuildSourceDataset(fileId, fileName),
                fileNames = fileName,
                dbProductType = FirstToken(models, ';', ""),
                aiModel = FirstToken(models, ';', ""),
                model = FirstToken(models, ';', ""),
                modelMappingSource = "MicroSpeaker DB files.models",
                date = FirstToken(datesFound, ';', ""),
                purposeCode = structure,
                reviewPurpose = structure,
                purpose = $"{structure}; Pair={pairCount:N0}; Metric={metricCount:N0}; Measurement={measurementCount:N0}",
                targetDefects,
                reviewItems = reviewItems.Length > 0 ? reviewItems.Take(12).ToArray() : targetDefects.Take(12).ToArray(),
                tags,
                confidence = ConfidenceScore(confidence),
                needsDetailedAnalysis = pairCount > 0,
                evidenceSummary = $"MicroSpeaker CLI/GUI output only. Pair={pairCount:N0}, Metric={metricCount:N0}, Measurement={measurementCount:N0}.",
                evidenceCells = Array.Empty<string>(),
                uncertainty = "Imported from MicroSpeaker CLI/GUI output only. Legacy AI analysis was removed."
            });
        }

        string indexPath = Path.Combine(sampleDir, "demo_index.json");
        File.WriteAllText(indexPath, JsonSerializer.Serialize(indexRows, new JsonSerializerOptions { WriteIndented = true }), Encoding.UTF8);

        string structureRows = string.Join("", structures
            .OrderByDescending(x => x.Value)
            .ThenBy(x => x.Key)
            .Select(x => $"<tr><td>{System.Net.WebUtility.HtmlEncode(x.Key)}</td><td>{x.Value:N0}</td></tr>"));
        string html = $$"""
            <!doctype html>
            <html>
            <head>
              <meta charset="utf-8">
              <title>MicroSpeaker CLI First Analysis</title>
              <style>
                body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #172033; }
                table { border-collapse: collapse; width: 100%; }
                th, td { border: 1px solid #d7deea; padding: 6px 8px; text-align: left; }
                th { background: #eef2f7; }
                .cards { display: flex; gap: 10px; flex-wrap: wrap; margin: 14px 0; }
                .card { border: 1px solid #d7deea; border-radius: 8px; padding: 10px 12px; background: #f8fafc; }
                .card strong { display: block; font-size: 20px; }
              </style>
            </head>
            <body>
              <h1>MicroSpeaker CLI First Analysis</h1>
              <p>Legacy AI analysis artifacts were removed. This report is generated from MicroSpeaker CLI/GUI SQLite only.</p>
              <div class="cards">
                <div class="card"><span>Files</span><strong>{{indexRows.Count:N0}}</strong></div>
                <div class="card"><span>Pairs</span><strong>{{totalPairs:N0}}</strong></div>
                <div class="card"><span>Metrics</span><strong>{{totalMetrics:N0}}</strong></div>
                <div class="card"><span>Measurements</span><strong>{{totalMeasurements:N0}}</strong></div>
              </div>
              <h2>Structure Families</h2>
              <table><thead><tr><th>Structure</th><th>Files</th></tr></thead><tbody>{{structureRows}}</tbody></table>
            </body>
            </html>
            """;
        File.WriteAllText(Path.Combine(sampleDir, "demo_report.html"), html, Encoding.UTF8);
        return indexPath;
    }

    private static long ImportDocuments(
        SqliteConnection micro,
        SqliteConnection jino,
        SqliteTransaction tx,
        Dictionary<long, MicroSpeakerDocRef> docMap)
    {
        long count = 0;
        using SqliteCommand cmd = micro.CreateCommand();
        cmd.CommandText = """
            SELECT file_id, dataset, path, file_name, status, error, sheet_count, sheet_names,
                   models, categories, dates_found, structure_family, structure_confidence,
                   term_summary, metric_candidate_count, measurement_stat_count,
                   comparison_pair_count, processed_at
            FROM files
            ORDER BY file_id;
            """;

        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            long fileId = r.GetInt64(0);
            string dataset = S(r, 1);
            string filePath = S(r, 2);
            string fileName = S(r, 3);
            string status = S(r, 4);
            string error = S(r, 5);
            string sheetNames = S(r, 7);
            string models = S(r, 8);
            string categories = S(r, 9);
            string datesFound = S(r, 10);
            string structure = S(r, 11);
            string confidence = S(r, 12);
            string termSummary = S(r, 13);
            long metricCount = r.GetInt64(14);
            long measurementCount = r.GetInt64(15);
            long pairCount = r.GetInt64(16);
            string processedAt = S(r, 17);

            string docId = $"microspeaker-file-{fileId}";
            string sourceDataset = BuildSourceDataset(fileId, fileName);
            string title = sourceDataset;
            var content = new[]
            {
                $"Structure: {structure} ({confidence})",
                $"Sheets: {sheetNames}",
                $"Pairs: {pairCount:N0}, Metrics: {metricCount:N0}, Measurements: {measurementCount:N0}",
                $"Terms: {termSummary}",
                string.IsNullOrWhiteSpace(error) ? $"Status: {status}" : $"Status: {status} - {error}",
            };
            var raw = new
            {
                source = "MicroSpeaker_ProductTech_DB",
                pipeline = "MicroSpeaker CLI deterministic indexer",
                fileId,
                dataset,
                filePath,
                status,
                sheetNames,
                structureFamily = structure,
                structureConfidence = confidence,
                pairCount,
                metricCount,
                measurementCount,
                termSummary,
            };

            Execute(jino, tx, """
                INSERT INTO AiDocuments
                    (DocumentId, SourceDataset, SourceFile, Title, Model, ReportDate,
                     Department, Marker, Line, ReportType, PrimaryDefect,
                     PrimaryDefectJson, RelatedDefectsJson, PartsJson, ProcessesJson,
                     Purpose, ContentJson, GeneratedReportMarkdown, SourceCellsJson,
                     Confidence, SchemaVersion, RawJson, CreatedAt, UpdatedAt)
                VALUES
                    (@DocumentId, @SourceDataset, @SourceFile, @Title, @Model, @ReportDate,
                     '', '', '', 'MicroSpeaker Input Data', @PrimaryDefect,
                     @PrimaryDefectJson, @RelatedDefectsJson, @PartsJson, @ProcessesJson,
                     @Purpose, @ContentJson, @GeneratedReportMarkdown, @SourceCellsJson,
                     @Confidence, 'microspeaker-db-v1', @RawJson, @CreatedAt, @UpdatedAt);
                """,
                ("@DocumentId", docId),
                ("@SourceDataset", sourceDataset),
                ("@SourceFile", fileName),
                ("@Title", title),
                ("@Model", models),
                ("@ReportDate", FirstToken(datesFound, ';', processedAt)),
                ("@PrimaryDefect", FirstToken(categories, ';', structure)),
                ("@PrimaryDefectJson", JsonSerializer.Serialize(new { canonical_name = FirstToken(categories, ';', structure), aliases_in_document = SplitSemi(categories) }, JsonOptions)),
                ("@RelatedDefectsJson", JsonSerializer.Serialize(SplitSemi(categories), JsonOptions)),
                ("@PartsJson", JsonSerializer.Serialize(SplitTermNames(termSummary), JsonOptions)),
                ("@ProcessesJson", JsonSerializer.Serialize(new[] { structure }.Concat(SplitSemi(categories)).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct().ToArray(), JsonOptions)),
                ("@Purpose", $"Imported from MicroSpeaker CLI. Structure={structure}; Pair={pairCount:N0}; Metric={metricCount:N0}; Measurement={measurementCount:N0}."),
                ("@ContentJson", JsonSerializer.Serialize(content, JsonOptions)),
                ("@GeneratedReportMarkdown", BuildMarkdown(fileName, dataset, structure, confidence, sheetNames, pairCount, metricCount, measurementCount, termSummary, status, error)),
                ("@SourceCellsJson", JsonSerializer.Serialize(new { file_id = fileId, path = filePath, sheets = sheetNames }, JsonOptions)),
                ("@Confidence", ConfidenceScore(confidence)),
                ("@RawJson", JsonSerializer.Serialize(raw, JsonOptions)),
                ("@CreatedAt", processedAt),
                ("@UpdatedAt", processedAt));

            Execute(jino, tx, """
                INSERT INTO AiConclusions
                    (ConclusionId, DocumentId, Topic, StatementFromReport,
                     NormalizedInterpretation, SourceFile, SheetName, SourceCellsJson)
                VALUES
                    (@ConclusionId, @DocumentId, 'MicroSpeaker import summary', @Statement,
                     @Interpretation, @SourceFile, '', @SourceCellsJson);
                """,
                ("@ConclusionId", $"microspeaker-conclusion-{fileId}"),
                ("@DocumentId", docId),
                ("@Statement", $"Structure={structure}; Pair={pairCount:N0}; Metric={metricCount:N0}; Measurement={measurementCount:N0}."),
                ("@Interpretation", "This row was imported from the deterministic MicroSpeaker CLI SQLite output, not from a new AI/Codex analysis."),
                ("@SourceFile", fileName),
                ("@SourceCellsJson", JsonSerializer.Serialize(new { file_id = fileId }, JsonOptions)));

            Execute(jino, tx, """
                INSERT INTO AiExtractionLogs
                    (LogId, DocumentId, Confidence, AssumptionsJson, WarningsJson,
                     DecisionRationale, CreatedAt)
                VALUES
                    (@LogId, @DocumentId, @Confidence, @AssumptionsJson, @WarningsJson,
                     @DecisionRationale, @CreatedAt);
                """,
                ("@LogId", $"microspeaker-log-{fileId}"),
                ("@DocumentId", docId),
                ("@Confidence", ConfidenceScore(confidence)),
                ("@AssumptionsJson", JsonSerializer.Serialize(new[] { "Imported from MicroSpeaker CLI SQLite output." }, JsonOptions)),
                ("@WarningsJson", JsonSerializer.Serialize(string.IsNullOrWhiteSpace(error) ? Array.Empty<string>() : new[] { error }, JsonOptions)),
                ("@DecisionRationale", "JinoSupporter legacy AI analysis rows were cleared and replaced by MicroSpeaker DB content."),
                ("@CreatedAt", processedAt));

            docMap[fileId] = new MicroSpeakerDocRef(docId, sourceDataset, fileName, processedAt);
            count++;
        }

        return count;
    }

    private static long ImportPairs(
        SqliteConnection micro,
        SqliteConnection jino,
        SqliteTransaction tx,
        IReadOnlyDictionary<long, MicroSpeakerDocRef> docMap)
    {
        long count = 0;
        using SqliteCommand cmd = micro.CreateCommand();
        cmd.CommandText = """
            SELECT pair_id, file_id, table_title, compare_item,
                   control_condition, test_condition,
                   control_input, control_ng, control_rate,
                   test_input, test_ng, test_rate,
                   delta_rate, improvement_rate, effect_direction,
                   evidence, pair_confidence
            FROM comparison_pairs
            ORDER BY pair_id;
            """;
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            long pairId = r.GetInt64(0);
            long fileId = r.GetInt64(1);
            if (!docMap.TryGetValue(fileId, out MicroSpeakerDocRef? doc)) continue;

            string conditionId = $"microspeaker-pair-cond-{pairId}";
            string resultId = $"microspeaker-pair-{pairId}";
            string tableTitle = S(r, 2);
            string compareItem = S(r, 3);
            string controlCondition = S(r, 4);
            string testCondition = S(r, 5);
            double? controlInput = D(r, 6);
            double? controlNg = D(r, 7);
            double? controlRate = D(r, 8);
            double? testInput = D(r, 9);
            double? testNg = D(r, 10);
            double? testRate = D(r, 11);
            double? deltaRate = D(r, 12);
            double? improvementRate = D(r, 13);
            string effect = S(r, 14);
            string evidence = S(r, 15);
            string pairConfidence = S(r, 16);
            string sourceJson = JsonSerializer.Serialize(new
            {
                pairId,
                tableTitle,
                compareItem,
                controlCondition,
                testCondition,
                controlInput,
                controlNg,
                controlRate,
                testInput,
                testNg,
                testRate,
                deltaRate,
                improvementRate,
                evidence,
                pairConfidence,
            }, JsonOptions);

            Execute(jino, tx, """
                INSERT INTO AiTestConditions
                    (ConditionId, DocumentId, ConditionGroup, Line, Process, ChangedFactor,
                     BeforeValue, AfterValue, SourceFile, SheetName, SourceCellsJson)
                VALUES
                    (@ConditionId, @DocumentId, @ConditionGroup, @Line, @Process, @ChangedFactor,
                     @BeforeValue, @AfterValue, @SourceFile, '', @SourceCellsJson);
                """,
                ("@ConditionId", conditionId),
                ("@DocumentId", doc.DocumentId),
                ("@ConditionGroup", effect),
                ("@Line", evidence),
                ("@Process", tableTitle),
                ("@ChangedFactor", compareItem),
                ("@BeforeValue", controlCondition),
                ("@AfterValue", testCondition),
                ("@SourceFile", doc.SourceFile),
                ("@SourceCellsJson", sourceJson));

            Execute(jino, tx, """
                INSERT INTO AiResults
                    (ResultId, DocumentId, ConditionId, MeasurementType, ConditionGroup,
                     ResultDate, Line, InputCount, OkCount, NgCount, NgRateDecimal,
                     NgRatePercent, MetricName, MetricValue, Unit, Judgement,
                     SourceFile, SheetName, SourceCellsJson)
                VALUES
                    (@ResultId, @DocumentId, @ConditionId, 'NG rate pair', @ConditionGroup,
                     @ResultDate, @Line, @InputCount, NULL, @NgCount, @NgRateDecimal,
                     @NgRatePercent, @MetricName, @MetricValue, '%p', @Judgement,
                     @SourceFile, '', @SourceCellsJson);
                """,
                ("@ResultId", resultId),
                ("@DocumentId", doc.DocumentId),
                ("@ConditionId", conditionId),
                ("@ConditionGroup", effect),
                ("@ResultDate", doc.ProcessedAt),
                ("@Line", $"{controlCondition} -> {testCondition}"),
                ("@InputCount", testInput),
                ("@NgCount", testNg),
                ("@NgRateDecimal", testRate),
                ("@NgRatePercent", Percent(testRate)),
                ("@MetricName", compareItem),
                ("@MetricValue", PercentPoint(deltaRate)),
                ("@Judgement", effect),
                ("@SourceFile", doc.SourceFile),
                ("@SourceCellsJson", sourceJson));

            if (testNg.GetValueOrDefault() > 0)
            {
                Execute(jino, tx, """
                    INSERT INTO AiNgBreakdowns
                        (BreakdownId, ResultId, DefectName, DefectCount, DefectRate)
                    VALUES
                        (@BreakdownId, @ResultId, @DefectName, @DefectCount, @DefectRate);
                    """,
                    ("@BreakdownId", $"microspeaker-pair-ng-{pairId}"),
                    ("@ResultId", resultId),
                    ("@DefectName", compareItem),
                    ("@DefectCount", testNg),
                    ("@DefectRate", Percent(testRate)));
            }

            count++;
        }
        return count;
    }

    private static long ImportMetricCandidates(
        SqliteConnection micro,
        SqliteConnection jino,
        SqliteTransaction tx,
        IReadOnlyDictionary<long, MicroSpeakerDocRef> docMap)
    {
        long count = 0;
        using SqliteCommand cmd = micro.CreateCommand();
        cmd.CommandText = """
            SELECT metric_id, file_id, sheet_name, row_number, table_title,
                   condition_label, input_qty, ok_qty, ng_qty, ng_rate,
                   detail, raw_row, parse_confidence
            FROM metric_candidates
            ORDER BY metric_id;
            """;
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            long metricId = r.GetInt64(0);
            long fileId = r.GetInt64(1);
            if (!docMap.TryGetValue(fileId, out MicroSpeakerDocRef? doc)) continue;

            string sheetName = S(r, 2);
            long rowNumber = r.GetInt64(3);
            string tableTitle = S(r, 4);
            string condition = S(r, 5);
            double? inputQty = D(r, 6);
            double? okQty = D(r, 7);
            double? ngQty = D(r, 8);
            double? ngRate = D(r, 9);
            string detail = S(r, 10);
            string rawRow = S(r, 11);
            string confidence = S(r, 12);
            string sourceJson = JsonSerializer.Serialize(new { metricId, sheetName, rowNumber, detail, rawRow, confidence }, JsonOptions);

            Execute(jino, tx, """
                INSERT INTO AiResults
                    (ResultId, DocumentId, ConditionId, MeasurementType, ConditionGroup,
                     ResultDate, Line, InputCount, OkCount, NgCount, NgRateDecimal,
                     NgRatePercent, MetricName, MetricValue, Unit, Judgement,
                     SourceFile, SheetName, SourceCellsJson)
                VALUES
                    (@ResultId, @DocumentId, '', 'NG metric candidate', @ConditionGroup,
                     @ResultDate, @Line, @InputCount, @OkCount, @NgCount, @NgRateDecimal,
                     @NgRatePercent, @MetricName, @MetricValue, '%', @Judgement,
                     @SourceFile, @SheetName, @SourceCellsJson);
                """,
                ("@ResultId", $"microspeaker-metric-{metricId}"),
                ("@DocumentId", doc.DocumentId),
                ("@ConditionGroup", condition),
                ("@ResultDate", doc.ProcessedAt),
                ("@Line", rowNumber.ToString()),
                ("@InputCount", inputQty),
                ("@OkCount", okQty),
                ("@NgCount", ngQty),
                ("@NgRateDecimal", ngRate),
                ("@NgRatePercent", Percent(ngRate)),
                ("@MetricName", tableTitle),
                ("@MetricValue", Percent(ngRate)),
                ("@Judgement", confidence),
                ("@SourceFile", doc.SourceFile),
                ("@SheetName", sheetName),
                ("@SourceCellsJson", sourceJson));

            if (ngQty.GetValueOrDefault() > 0)
            {
                Execute(jino, tx, """
                    INSERT INTO AiNgBreakdowns
                        (BreakdownId, ResultId, DefectName, DefectCount, DefectRate)
                    VALUES
                        (@BreakdownId, @ResultId, @DefectName, @DefectCount, @DefectRate);
                    """,
                    ("@BreakdownId", $"microspeaker-metric-ng-{metricId}"),
                    ("@ResultId", $"microspeaker-metric-{metricId}"),
                    ("@DefectName", string.IsNullOrWhiteSpace(detail) ? tableTitle : detail),
                    ("@DefectCount", ngQty),
                    ("@DefectRate", Percent(ngRate)));
            }

            count++;
        }
        return count;
    }

    private static long ImportMeasurementStats(
        SqliteConnection micro,
        SqliteConnection jino,
        SqliteTransaction tx,
        IReadOnlyDictionary<long, MicroSpeakerDocRef> docMap)
    {
        long count = 0;
        using SqliteCommand cmd = micro.CreateCommand();
        cmd.CommandText = """
            SELECT stat_id, file_id, sheet_name, row_number, item_label,
                   condition_label, spec, min_value, max_value, avg_value,
                   sample_count, violation_count, raw_row, parse_confidence
            FROM measurement_stats
            ORDER BY stat_id;
            """;
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            long statId = r.GetInt64(0);
            long fileId = r.GetInt64(1);
            if (!docMap.TryGetValue(fileId, out MicroSpeakerDocRef? doc)) continue;

            string sheetName = S(r, 2);
            long rowNumber = r.GetInt64(3);
            string item = S(r, 4);
            string condition = S(r, 5);
            string spec = S(r, 6);
            double? minValue = D(r, 7);
            double? maxValue = D(r, 8);
            double? avgValue = D(r, 9);
            long sampleCount = r.GetInt64(10);
            long violationCount = r.GetInt64(11);
            string rawRow = S(r, 12);
            string confidence = S(r, 13);
            string sourceJson = JsonSerializer.Serialize(new { statId, sheetName, rowNumber, spec, minValue, maxValue, avgValue, rawRow, confidence }, JsonOptions);

            Execute(jino, tx, """
                INSERT INTO AiResults
                    (ResultId, DocumentId, ConditionId, MeasurementType, ConditionGroup,
                     ResultDate, Line, InputCount, OkCount, NgCount, NgRateDecimal,
                     NgRatePercent, MetricName, MetricValue, Unit, Judgement,
                     SourceFile, SheetName, SourceCellsJson)
                VALUES
                    (@ResultId, @DocumentId, '', 'Measurement stat', @ConditionGroup,
                     @ResultDate, @Line, @InputCount, NULL, @NgCount, NULL,
                     NULL, @MetricName, @MetricValue, @Unit, @Judgement,
                     @SourceFile, @SheetName, @SourceCellsJson);
                """,
                ("@ResultId", $"microspeaker-stat-{statId}"),
                ("@DocumentId", doc.DocumentId),
                ("@ConditionGroup", condition),
                ("@ResultDate", doc.ProcessedAt),
                ("@Line", rowNumber.ToString()),
                ("@InputCount", sampleCount),
                ("@NgCount", violationCount),
                ("@MetricName", item),
                ("@MetricValue", avgValue),
                ("@Unit", spec),
                ("@Judgement", confidence),
                ("@SourceFile", doc.SourceFile),
                ("@SheetName", sheetName),
                ("@SourceCellsJson", sourceJson));

            count++;
        }
        return count;
    }

    private static bool TableExists(SqliteConnection conn, string tableName)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 1 FROM sqlite_master WHERE type='table' AND name=@name LIMIT 1;";
        cmd.Parameters.AddWithValue("@name", tableName);
        return cmd.ExecuteScalar() is not null;
    }

    private static void Execute(SqliteConnection conn, SqliteTransaction tx, string sql, params (string Name, object? Value)[] parameters)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = sql;
        foreach ((string name, object? value) in parameters)
            cmd.Parameters.AddWithValue(name, value ?? DBNull.Value);
        cmd.ExecuteNonQuery();
    }

    private static long ScalarLong(SqliteConnection conn, string sql)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        return Convert.ToInt64(cmd.ExecuteScalar() ?? 0);
    }

    private static string S(SqliteDataReader r, int index) => r.IsDBNull(index) ? "" : r.GetString(index);

    private static double? D(SqliteDataReader r, int index) => r.IsDBNull(index) ? null : r.GetDouble(index);

    private static double? Percent(double? decimalRate) => decimalRate.HasValue ? decimalRate.Value * 100.0 : null;

    private static double? PercentPoint(double? decimalRate) => decimalRate.HasValue ? decimalRate.Value * 100.0 : null;

    private static double ConfidenceScore(string confidence) => confidence.ToUpperInvariant() switch
    {
        "HIGH" => 0.95,
        "MEDIUM" => 0.75,
        "LOW" => 0.5,
        _ => 0.6
    };

    private static string BuildSourceDataset(long fileId, string fileName)
    {
        string stem = fileName;
        string extension = Path.GetExtension(stem);
        if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".xlsb", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
        {
            stem = stem[..^extension.Length];
        }
        if (stem.EndsWith("_clean", StringComparison.OrdinalIgnoreCase))
            stem = stem[..^"_clean".Length];
        stem = stem.Trim();
        if (string.IsNullOrWhiteSpace(stem))
            stem = $"MicroSpeaker file {fileId}";
        return stem;
    }

    private static string FirstToken(string text, char separator, string fallback)
    {
        if (string.IsNullOrWhiteSpace(text)) return fallback;
        string token = text.Split(separator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).FirstOrDefault() ?? "";
        return string.IsNullOrWhiteSpace(token) ? fallback : token;
    }

    private static string[] SplitSemi(string text) =>
        string.IsNullOrWhiteSpace(text)
            ? []
            : text.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

    private static string[] SplitTermNames(string termSummary) =>
        string.IsNullOrWhiteSpace(termSummary)
            ? []
            : termSummary.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(x =>
                {
                    int pos = x.IndexOf('(');
                    return pos > 0 ? x[..pos].Trim() : x.Trim();
                })
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

    private static string BuildMarkdown(
        string fileName,
        string dataset,
        string structure,
        string confidence,
        string sheetNames,
        long pairCount,
        long metricCount,
        long measurementCount,
        string termSummary,
        string status,
        string error)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"# {fileName}");
        sb.AppendLine();
        sb.AppendLine($"- Dataset: {dataset}");
        sb.AppendLine($"- Structure: {structure} ({confidence})");
        sb.AppendLine($"- Sheets: {sheetNames}");
        sb.AppendLine($"- Pair rows: {pairCount:N0}");
        sb.AppendLine($"- Metric candidate rows: {metricCount:N0}");
        sb.AppendLine($"- Measurement stat rows: {measurementCount:N0}");
        sb.AppendLine($"- Terms: {termSummary}");
        sb.AppendLine(string.IsNullOrWhiteSpace(error) ? $"- Status: {status}" : $"- Status: {status} - {error}");
        return sb.ToString();
    }

    private static readonly string[] AiDeleteOrder =
    [
        "AiNgBreakdowns",
        "AiResults",
        "AiTestConditions",
        "AiConclusionTranslations",
        "AiConclusions",
        "AiHintTranslations",
        "AiTroubleshootingHints",
        "AiLogTranslations",
        "AiExtractionLogs",
        "AiDocumentTranslations",
        "AiModelAnalyses",
        "AiDocuments",
        "AskAiHistory",
        "DatasetSummaryTranslations",
        "DatasetSummary",
        "NormalizedMeasurements",
    ];

    private static readonly string[] AiBackupTables =
    [
        "AiDocuments",
        "AiTestConditions",
        "AiResults",
        "AiNgBreakdowns",
        "AiConclusions",
        "AiTroubleshootingHints",
        "AiExtractionLogs",
        "AiDocumentTranslations",
        "AiConclusionTranslations",
        "AiHintTranslations",
        "AiLogTranslations",
        "AiModelAnalyses",
        "AskAiHistory",
        "DatasetSummary",
        "DatasetSummaryTranslations",
        "NormalizedMeasurements",
    ];
}

public sealed record MicroSpeakerPaths(
    string ProjectRoot,
    string DatabasePath,
    string DashboardPath,
    string CliExePath,
    string CliScriptPath);

public sealed record MicroSpeakerIndexRequest(
    string InputDir,
    string Dataset,
    bool Force,
    bool NoHtml,
    int Limit);

public sealed record MicroSpeakerCliRunResult(
    bool Success,
    int ExitCode,
    string Output,
    string Error);

public sealed record MicroSpeakerImportResult(
    string JinoDbPath,
    string MicroSpeakerDbPath,
    string BackupPath,
    long Documents,
    long DatasetSummaries,
    long PairRows,
    long MetricRows,
    long MeasurementRows,
    string FirstAnalysisIndexPath);

public sealed record MicroSpeakerSummary(
    MicroSpeakerPaths Paths,
    bool DatabaseExists,
    MicroSpeakerRunInfo? LatestRun,
    long FileCount,
    long OkFileCount,
    long FailedFileCount,
    long MetricCandidateCount,
    long MeasurementStatCount,
    long PairCount,
    long TermHitCount,
    IReadOnlyList<MicroSpeakerStructureRow> Structures,
    IReadOnlyList<MicroSpeakerPairPreview> TopWorsenedPairs,
    IReadOnlyList<MicroSpeakerPairPreview> TopImprovedPairs);

public sealed record MicroSpeakerRunInfo(
    long RunId,
    string Dataset,
    string InputDir,
    string StartedAt,
    string FinishedAt,
    long FilesSeen,
    long FilesProcessed,
    long FilesSkipped,
    long FilesFailed);

public sealed record MicroSpeakerStructureRow(
    string StructureFamily,
    string StructureConfidence,
    long FileCount,
    long PairCount,
    long MetricCount,
    long MeasurementCount);

public sealed record MicroSpeakerPairPreview(
    string FileName,
    string CompareItem,
    string ControlCondition,
    string TestCondition,
    double? ControlRate,
    double? TestRate,
    double? DeltaRate,
    string EffectDirection,
    string Evidence);

public sealed record MicroSpeakerSourceFile(
    long FileId,
    string FullPath,
    string FileName);

internal sealed record MicroSpeakerDocRef(
    string DocumentId,
    string SourceDataset,
    string SourceFile,
    string ProcessedAt);
