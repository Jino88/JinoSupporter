using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

public sealed class MicroSpeakerAskEvidenceService
{
    private const int CandidateLimit = 1500;
    private const int PairLimit = 120;
    private const int MetricLimit = 120;
    private const int MeasurementLimit = 120;
    private const int JinoResultLimit = 120;
    private const int JinoDocumentLimit = 80;
    private const int ModelCoverageFileLimit = 200;
    private const int VerifiedReviewCaseLimit = 20;
    private const int ReviewCaseChangedFactorLimit = 8;
    private const int ReviewCaseOutcomeLimit = 10;
    private const int ReviewCaseSubResultLimit = 10;
    private const int ReviewCaseEvidenceRowLimit = 40;

    private static readonly string[] FunctionSignalTerms =
    [
        "function", "functional", "Function total", "SPL", "THD", "Rub", "Buzz", "R&B",
        "Noise", "Hearing", "Sound", "Audio", "Acoustic", "DCR", "IMP"
    ];

    private static readonly string[] DefectSignalTerms =
    [
        "NG", "defect", "fail", "failure", "reject", "ppm", "yield"
    ];

    private static readonly string[] GenericProcessWords =
    [
        "process", "condition", "assembly", "bond", "bonding", "press", "plasma", "jig",
        "mold", "machine", "line", "lot", "supplier", "material", "laser", "cutting",
        "dry", "uv", "uc", "temperature", "pressure", "time", "speed"
    ];

    private static readonly string[] GenericOutcomeWords =
    [
        "function", "functional", "hearing", "sound", "audio", "acoustic", "spl", "thd",
        "noise", "rub", "buzz", "dcr", "imp", "ng", "defect", "fail", "failure",
        "ppm", "yield"
    ];

    private static readonly string[] StopWords =
    [
        "what", "when", "where", "which", "why", "how", "does", "with", "from", "that",
        "this", "into", "about", "effect", "impact", "influence", "review", "data",
        "the", "and", "for", "are", "was", "were", "have", "has",
        "무슨", "어떤", "어느", "얼마나", "영향", "영향을", "영향이", "끼치나", "미치나",
        "대해서", "대한", "관련", "검토", "분석", "자료", "데이터", "결과", "따른", "따라"
    ];

    private readonly WebRepository _repo;
    private readonly MicroSpeakerInputDataService _microSpeaker;

    public MicroSpeakerAskEvidenceService(WebRepository repo, MicroSpeakerInputDataService microSpeaker)
    {
        _repo = repo;
        _microSpeaker = microSpeaker;
    }

    public MicroSpeakerAskEvidencePack BuildEvidencePack(string question, string productTypeFilter)
    {
        string trimmedQuestion = (question ?? "").Trim();
        string trimmedProductType = (productTypeFilter ?? "").Trim();
        MicroSpeakerPaths paths = _microSpeaker.ResolvePaths();

        List<string> terms = BuildSearchTerms(trimmedQuestion);
        List<List<string>> requiredGroups = BuildRequiredTermGroups(trimmedQuestion);
        MicroSpeakerQuestionAnalysis questionAnalysis = AnalyzeQuestion(trimmedQuestion, terms, requiredGroups);

        var pack = new MicroSpeakerAskEvidencePack
        {
            CreatedAt = DateTime.UtcNow.ToString("O"),
            Question = trimmedQuestion,
            ProductTypeFilter = trimmedProductType,
            JinoDatabasePath = _repo.GetDbPath(),
            MicroSpeakerDatabasePath = paths.DatabasePath,
            SearchTerms = terms,
            RequiredTermGroups = requiredGroups,
            QuestionAnalysis = questionAnalysis,
            Notes =
            [
                "This evidence pack was built deterministically before Codex CLI reasoning.",
                "Rows marked matchesAllRequiredTerms=true matched every required question term group.",
                "QuestionAnalysis separates the user's factor/condition axis from the outcome/defect axis. Use those axes for generic report sections.",
                "Rate values from MicroSpeaker SQLite are decimals; ratePercent fields are normalized for report use.",
                "Display defect count basis explicitly as Input/NG. Defect rate = NG / Input * 100.",
                "Verified ReviewCases approved for Ask AI are exposed as microSpeaker.verifiedReviewCases and should be checked before raw rows.",
                "Do not aggregate unrelated rows. Sum NG counts only when source dataset/file/table, metric, and date/lot/line/model basis are compatible.",
                "Model names are evidence boundaries. If evidence spans multiple sourceModels/productTypes, group conclusions by model instead of merging them.",
                "Use pairConditionAggregates before individual pairRows when repeated rows share the same file/table/factor/condition.",
                "Do not combine process NG and function NG into one denominator or one total."
            ],
        };

        Dictionary<long, MicroSpeakerReviewCaseFileMetadata> reviewCaseMetadata = ReadMicroSpeakerFileMetadata(paths.DatabasePath);
        pack.MicroSpeaker.VerifiedReviewCases = ReadVerifiedReviewCases(terms, requiredGroups, reviewCaseMetadata, trimmedProductType);
        if (pack.MicroSpeaker.VerifiedReviewCases.Count > 0)
            pack.Notes.Add($"Loaded {pack.MicroSpeaker.VerifiedReviewCases.Count} approved verified ReviewCase evidence item(s).");

        if (terms.Count == 0)
        {
            pack.Notes.Add("No stable search terms were extracted from the question; Ask AI should fall back to registered report context.");
            return pack;
        }

        if (File.Exists(paths.DatabasePath))
        {
            using SqliteConnection micro = OpenReadOnly(paths.DatabasePath);
            pack.MicroSpeaker.DatabaseExists = true;
            pack.MicroSpeaker.TermHits = CountTermHits(micro, terms, new Dictionary<string, string[]>
            {
                ["files"] = ["dataset", "path", "file_name", "sheet_names", "models", "categories", "dates_found", "structure_family", "term_summary"],
                ["comparison_pairs"] = ["table_title", "compare_item", "control_condition", "test_condition", "effect_direction", "evidence", "pair_confidence"],
                ["metric_candidates"] = ["sheet_name", "table_title", "condition_label", "detail", "raw_row", "parse_confidence"],
                ["measurement_stats"] = ["sheet_name", "item_label", "condition_label", "spec", "raw_row", "parse_confidence"],
            });
            List<MicroSpeakerModelFileCoverageRow> modelFiles = ReadMicroSpeakerModelFileCoverage(micro, terms, requiredGroups, trimmedProductType);
            pack.MicroSpeaker.PairRows = ReadMicroSpeakerPairs(micro, terms, requiredGroups, trimmedProductType);
            pack.MicroSpeaker.MetricRows = ReadMicroSpeakerMetrics(micro, terms, requiredGroups, trimmedProductType);
            pack.MicroSpeaker.MeasurementRows = ReadMicroSpeakerMeasurements(micro, terms, requiredGroups, trimmedProductType);
            pack.MicroSpeaker.PairAggregates = BuildPairAggregates(pack.MicroSpeaker.PairRows);
            pack.MicroSpeaker.PairConditionAggregates = BuildPairConditionAggregates(pack.MicroSpeaker.PairRows);
            pack.MicroSpeaker.ModelCoverage = BuildModelCoverage(
                modelFiles,
                pack.MicroSpeaker.VerifiedReviewCases,
                pack.MicroSpeaker.PairRows,
                pack.MicroSpeaker.MetricRows,
                pack.MicroSpeaker.MeasurementRows);
        }
        else
        {
            pack.MicroSpeaker.DatabaseExists = false;
            pack.Notes.Add("MicroSpeaker SQLite database was not found.");
        }

        string jinoDbPath = _repo.GetDbPath();
        if (File.Exists(jinoDbPath))
        {
            using SqliteConnection jino = OpenReadOnly(jinoDbPath);
            pack.Jino.DatabaseExists = true;
            pack.Jino.TermHits = CountTermHits(jino, terms, new Dictionary<string, string[]>
            {
                ["AiDocuments"] = ["SourceDataset", "SourceFile", "Title", "Model", "ReportType", "PrimaryDefect", "RelatedDefectsJson", "PartsJson", "ProcessesJson", "Purpose", "ContentJson", "RawJson", "GeneratedReportMarkdown"],
                ["AiTestConditions"] = ["ConditionGroup", "Line", "Process", "ChangedFactor", "BeforeValue", "AfterValue", "SourceFile", "SheetName", "SourceCellsJson"],
                ["AiResults"] = ["MeasurementType", "ConditionGroup", "Line", "MetricName", "Judgement", "SourceFile", "SheetName", "SourceCellsJson"],
                ["AiNgBreakdowns"] = ["DefectName"],
                ["DatasetSummary"] = ["DatasetName", "ProductType", "Summary", "KeyFindings", "Tags", "Purpose", "TestConditions", "RootCause", "Decision", "RecommendedAction", "EvidenceJson", "ActionsJson", "ContextJson", "ReportType"],
            });
            pack.Jino.DocumentRows = ReadJinoDocuments(jino, terms, requiredGroups, trimmedProductType);
            pack.Jino.ResultRows = ReadJinoResults(jino, terms, requiredGroups, trimmedProductType);
        }
        else
        {
            pack.Jino.DatabaseExists = false;
            pack.Notes.Add("JinoSupporter process-review database was not found.");
        }

        return pack;
    }

    private static List<string> BuildSearchTerms(string question)
    {
        string q = question ?? "";
        var terms = new List<string>();
        List<string> explicitTerms = ExtractQuestionTerms(q);
        terms.AddRange(explicitTerms);
        terms.AddRange(BuildTechnicalPhraseVariants(explicitTerms));

        if (MentionsFunction(q, explicitTerms))
            terms.AddRange(FunctionSignalTerms);

        if (MentionsDefect(q, explicitTerms))
            terms.AddRange(DefectSignalTerms);

        if (ContainsAny(q, "process", "condition", "assembly") || q.Contains("\uC870\uB9BD", StringComparison.OrdinalIgnoreCase) || q.Contains("\uACF5\uC815", StringComparison.OrdinalIgnoreCase))
            terms.AddRange(GenericProcessWords);

        return NormalizeTerms(terms).Take(40).ToList();
    }

    private static List<List<string>> BuildRequiredTermGroups(string question)
    {
        string q = question ?? "";
        var groups = new List<List<string>>();
        List<string> explicitTerms = ExtractQuestionTerms(q);
        List<string> factorTerms = InferFactorTerms(q, explicitTerms);
        List<string> outcomeTerms = InferOutcomeTerms(q, explicitTerms);

        if (factorTerms.Count > 0)
            groups.Add(factorTerms);

        if (outcomeTerms.Count > 0)
            groups.Add(outcomeTerms);

        if (groups.Count == 0)
        {
            List<string> fallback = NormalizeTerms(explicitTerms.Concat(BuildTechnicalPhraseVariants(explicitTerms))).Take(12).ToList();
            if (fallback.Count > 0) groups.Add(fallback);
        }

        return groups;
    }

    private static MicroSpeakerQuestionAnalysis AnalyzeQuestion(
        string question,
        IReadOnlyList<string> searchTerms,
        IReadOnlyList<List<string>> requiredGroups)
    {
        List<string> explicitTerms = ExtractQuestionTerms(question);
        List<string> factorTerms = InferFactorTerms(question, explicitTerms);
        List<string> outcomeTerms = InferOutcomeTerms(question, explicitTerms);

        string factorLabel = factorTerms.Count > 0 ? string.Join(" / ", factorTerms.Take(4)) : "question factor/condition";
        string outcomeLabel = outcomeTerms.Count > 0 ? string.Join(" / ", outcomeTerms.Take(4)) : "question result/defect";

        return new MicroSpeakerQuestionAnalysis
        {
            FactorAxisLabel = factorLabel,
            OutcomeAxisLabel = outcomeLabel,
            FactorTerms = factorTerms,
            OutcomeTerms = outcomeTerms,
            SuggestedReviewSections =
            [
                $"{factorLabel} condition/process review",
                $"{outcomeLabel} defect/result review",
                $"{factorLabel} versus {outcomeLabel} linkage review",
                "data, denominator, and aggregation limits"
            ],
            SearchTermCount = searchTerms.Count,
            RequiredGroupCount = requiredGroups.Count,
        };
    }

    private static List<string> ExtractQuestionTerms(string question)
    {
        string q = question ?? "";
        var terms = new List<string>();

        foreach (Match m in Regex.Matches(q, @"[A-Za-z0-9][A-Za-z0-9+/\-_.&%]{1,}"))
        {
            string token = NormalizeQuestionToken(m.Value);
            if (token.Length < 2) continue;
            if (StopWords.Contains(token, StringComparer.OrdinalIgnoreCase)) continue;
            terms.Add(token);
        }

        foreach (Match m in Regex.Matches(q, @"[\uAC00-\uD7AF]{2,}"))
        {
            string token = NormalizeQuestionToken(m.Value);
            if (token.Length < 2) continue;
            if (StopWords.Contains(token, StringComparer.OrdinalIgnoreCase)) continue;
            terms.Add(token);
        }

        return NormalizeTerms(terms);
    }

    private static string NormalizeQuestionToken(string token)
    {
        string value = (token ?? "").Trim();
        if (value.Length < 3 || !Regex.IsMatch(value, @"^[\uAC00-\uD7AF]+$"))
            return value;

        string[] suffixes =
        [
            "으로는", "에서는", "에게는", "까지는", "부터는",
            "으로", "에서", "에게", "까지", "부터",
            "에는", "하고", "처럼", "보다",
            "가", "이", "은", "는", "을", "를", "에", "의", "와", "과", "로", "도", "만"
        ];

        foreach (string suffix in suffixes)
        {
            if (value.Length > suffix.Length + 1 && value.EndsWith(suffix, StringComparison.Ordinal))
                return value[..^suffix.Length];
        }

        return value;
    }

    private static List<string> BuildTechnicalPhraseVariants(IReadOnlyList<string> terms)
    {
        var result = new List<string>();
        List<string> candidates = terms
            .Where(IsCompactTechnicalToken)
            .Take(8)
            .ToList();

        for (int i = 0; i < candidates.Count - 1; i++)
        {
            string a = candidates[i];
            string b = candidates[i + 1];
            if (a.Length > 8 || b.Length > 8) continue;
            result.Add($"{a}+{b}");
            result.Add($"{a}/{b}");
            result.Add($"{a}-{b}");
            result.Add($"{a} {b}");
            result.Add(a + b);
        }

        return NormalizeTerms(result);
    }

    private static List<string> InferFactorTerms(string question, IReadOnlyList<string> explicitTerms)
    {
        var terms = new List<string>();
        terms.AddRange(explicitTerms.Where(t => !IsOutcomeTerm(t) && IsUsefulFactorTerm(t)));
        terms.AddRange(BuildTechnicalPhraseVariants(terms));

        if (ContainsAny(question, "process", "condition", "assembly", "bond", "press", "plasma", "jig", "mold")
            || question.Contains("\uC870\uB9BD", StringComparison.OrdinalIgnoreCase)
            || question.Contains("\uACF5\uC815", StringComparison.OrdinalIgnoreCase)
            || question.Contains("\uC870\uAC74", StringComparison.OrdinalIgnoreCase))
        {
            terms.AddRange(explicitTerms.Where(t => IsCompactTechnicalToken(t)));
        }

        return NormalizeTerms(terms).Take(16).ToList();
    }

    private static List<string> InferOutcomeTerms(string question, IReadOnlyList<string> explicitTerms)
    {
        var terms = new List<string>();
        terms.AddRange(explicitTerms.Where(IsOutcomeTerm));

        if (MentionsFunction(question, explicitTerms))
            terms.AddRange(FunctionSignalTerms);

        if (MentionsDefect(question, explicitTerms))
            terms.AddRange(DefectSignalTerms);

        return NormalizeTerms(terms).Take(20).ToList();
    }

    private static bool MentionsFunction(string question, IReadOnlyList<string> terms)
        => terms.Any(t => ContainsAny(t, FunctionSignalTerms))
           || question.Contains("\uAE30\uB2A5", StringComparison.OrdinalIgnoreCase);

    private static bool MentionsDefect(string question, IReadOnlyList<string> terms)
        => terms.Any(t => ContainsAny(t, DefectSignalTerms))
           || question.Contains("\uBD88\uB7C9", StringComparison.OrdinalIgnoreCase);

    private static bool IsOutcomeTerm(string term)
        => ContainsAny(term, GenericOutcomeWords)
           || term.Contains("\uBD88\uB7C9", StringComparison.OrdinalIgnoreCase)
           || term.Contains("\uAE30\uB2A5", StringComparison.OrdinalIgnoreCase);

    private static bool IsUsefulFactorTerm(string term)
        => !StopWords.Contains(term, StringComparer.OrdinalIgnoreCase)
           && !string.Equals(term, "NG", StringComparison.OrdinalIgnoreCase)
           && term.Length >= 2;

    private static bool IsCompactTechnicalToken(string? term)
    {
        string value = term ?? "";
        return Regex.IsMatch(value, @"^[A-Za-z0-9][A-Za-z0-9+/\-_.&%]{1,}$")
               && !StopWords.Contains(value, StringComparer.OrdinalIgnoreCase)
               && !IsOutcomeTerm(value);
    }

    private static List<string> NormalizeTerms(IEnumerable<string> terms)
        => terms
            .Select(t => (t ?? "").Trim())
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static List<TermHitEvidence> CountTermHits(
        SqliteConnection conn,
        IReadOnlyList<string> terms,
        IReadOnlyDictionary<string, string[]> tableColumns)
    {
        var hits = new List<TermHitEvidence>();
        foreach ((string table, string[] wantedColumns) in tableColumns)
        {
            if (!TableExists(conn, table)) continue;
            HashSet<string> existing = TableColumns(conn, table);
            string[] columns = wantedColumns.Where(existing.Contains).ToArray();
            if (columns.Length == 0) continue;

            foreach (string term in terms)
            {
                using SqliteCommand cmd = conn.CreateCommand();
                cmd.CommandText = $"SELECT COUNT(*) FROM {table} WHERE {BuildAnyLike(columns, [term], "p", cmd)};";
                long count = Convert.ToInt64(cmd.ExecuteScalar() ?? 0, CultureInfo.InvariantCulture);
                if (count > 0) hits.Add(new TermHitEvidence(table, term, count));
            }
        }
        return hits
            .OrderByDescending(x => x.Count)
            .ThenBy(x => x.Table, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.Term, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<MicroSpeakerVerifiedReviewCaseEvidence> ReadVerifiedReviewCases(
        IReadOnlyList<string> terms,
        IReadOnlyList<List<string>> groups,
        IReadOnlyDictionary<long, MicroSpeakerReviewCaseFileMetadata> fileMetadata,
        string productTypeFilter)
    {
        string repoRoot = AiPromptRegistry.FindRepositoryRoot();
        string manifestPath = Path.Combine(repoRoot, "REVIEWCASE_AI_DRAFTS", "verified", "reviewcase_ai_verification_manifest.json");
        if (!File.Exists(manifestPath)) return [];

        var rows = new List<MicroSpeakerVerifiedReviewCaseEvidence>();
        try
        {
            using JsonDocument manifestDoc = JsonDocument.Parse(File.ReadAllText(manifestPath));
            if (!TryGetJsonProperty(manifestDoc.RootElement, "entries", out JsonElement entries)
                || entries.ValueKind != JsonValueKind.Array)
                return [];

            foreach (JsonElement entry in entries.EnumerateArray())
            {
                if (!JsonBool(entry, "approvedForAskAi")) continue;
                if (!string.Equals(JsonString(entry, "status"), "verified", StringComparison.OrdinalIgnoreCase)) continue;

                string verificationPath = ResolveRepoPath(repoRoot, JsonString(entry, "path"));
                if (!File.Exists(verificationPath)) continue;

                rows.AddRange(ReadVerifiedReviewCaseFile(repoRoot, entry, verificationPath, terms, groups, fileMetadata, productTypeFilter));
            }
        }
        catch (IOException)
        {
            return [];
        }
        catch (JsonException)
        {
            return [];
        }

        List<MicroSpeakerVerifiedReviewCaseEvidence> candidates = terms.Count == 0
            ? rows
            : rows.Where(x => x.MatchScore > 0).ToList();

        return SelectRows(candidates, VerifiedReviewCaseLimit, r => r.MatchScore, ReviewCaseConfidenceScore);
    }

    private static List<MicroSpeakerModelFileCoverageRow> ReadMicroSpeakerModelFileCoverage(
        SqliteConnection conn,
        IReadOnlyList<string> terms,
        IReadOnlyList<List<string>> groups,
        string productTypeFilter)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        string[] columns =
        [
            "dataset", "path", "file_name", "sheet_names", "models", "categories", "dates_found", "structure_family", "term_summary"
        ];
        string where = BuildAnyLike(columns, terms, "mf", cmd);
        if (!string.IsNullOrWhiteSpace(productTypeFilter))
        {
            where = $"({where}) AND COALESCE(models, '') LIKE @modelCoverageFilter";
            cmd.Parameters.AddWithValue("@modelCoverageFilter", "%" + productTypeFilter.Trim() + "%");
        }

        cmd.CommandText = $"""
            SELECT file_id, dataset, file_name, models, categories, term_summary
            FROM files
            WHERE {where}
            ORDER BY file_id
            LIMIT {CandidateLimit};
            """;

        var rows = new List<MicroSpeakerModelFileCoverageRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            var row = new MicroSpeakerModelFileCoverageRow
            {
                FileId = L(r, 0),
                Dataset = S(r, 1),
                FileName = S(r, 2),
                Models = S(r, 3),
                Categories = S(r, 4),
                TermSummary = S(r, 5),
            };
            FinishMatch(row, terms, groups, row.SearchText);
            rows.Add(row);
        }

        return SelectRows(rows, ModelCoverageFileLimit, r => r.MatchScore, r => 0);
    }

    private static List<MicroSpeakerModelCoverage> BuildModelCoverage(
        IReadOnlyList<MicroSpeakerModelFileCoverageRow> modelFiles,
        IReadOnlyList<MicroSpeakerVerifiedReviewCaseEvidence> reviewCases,
        IReadOnlyList<MicroSpeakerPairEvidenceRow> pairRows,
        IReadOnlyList<MicroSpeakerMetricEvidenceRow> metricRows,
        IReadOnlyList<MicroSpeakerMeasurementEvidenceRow> measurementRows)
    {
        var map = new Dictionary<string, MicroSpeakerModelCoverage>(StringComparer.OrdinalIgnoreCase);

        foreach (MicroSpeakerModelFileCoverageRow row in modelFiles)
        {
            MicroSpeakerModelCoverage coverage = GetCoverage(map, row.Models);
            coverage.MatchedFileCount++;
            AddDistinct(coverage.SourceFileIds, row.FileId.HasValue ? new[] { row.FileId.Value.ToString(CultureInfo.InvariantCulture) } : Array.Empty<string>());
            AddDistinct(coverage.ExampleFiles, new[] { row.FileName });
        }

        foreach (MicroSpeakerVerifiedReviewCaseEvidence row in reviewCases)
        {
            MicroSpeakerModelCoverage coverage = GetCoverage(map, row.SourceModels);
            coverage.VerifiedReviewCaseCount++;
            AddDistinct(coverage.SourceFileIds, row.SourceFileId.HasValue ? new[] { row.SourceFileId.Value.ToString(CultureInfo.InvariantCulture) } : Array.Empty<string>());
            AddDistinct(coverage.ExampleFiles, new[] { row.SourceFile });
        }

        foreach (MicroSpeakerPairEvidenceRow row in pairRows)
        {
            MicroSpeakerModelCoverage coverage = GetCoverage(map, row.Models);
            coverage.PairRowCount++;
            AddDistinct(coverage.SourceFileIds, row.FileId.HasValue ? new[] { row.FileId.Value.ToString(CultureInfo.InvariantCulture) } : Array.Empty<string>());
            AddDistinct(coverage.ExampleFiles, new[] { row.FileName });
        }

        foreach (MicroSpeakerMetricEvidenceRow row in metricRows)
        {
            MicroSpeakerModelCoverage coverage = GetCoverage(map, row.Models);
            coverage.MetricRowCount++;
            AddDistinct(coverage.SourceFileIds, row.FileId.HasValue ? new[] { row.FileId.Value.ToString(CultureInfo.InvariantCulture) } : Array.Empty<string>());
            AddDistinct(coverage.ExampleFiles, new[] { row.FileName });
        }

        foreach (MicroSpeakerMeasurementEvidenceRow row in measurementRows)
        {
            MicroSpeakerModelCoverage coverage = GetCoverage(map, row.Models);
            coverage.MeasurementRowCount++;
            AddDistinct(coverage.SourceFileIds, row.FileId.HasValue ? new[] { row.FileId.Value.ToString(CultureInfo.InvariantCulture) } : Array.Empty<string>());
            AddDistinct(coverage.ExampleFiles, new[] { row.FileName });
        }

        return map.Values
            .Where(x => x.MatchedFileCount + x.VerifiedReviewCaseCount + x.PairRowCount + x.MetricRowCount + x.MeasurementRowCount > 0)
            .OrderByDescending(x => x.VerifiedReviewCaseCount)
            .ThenByDescending(x => x.PairRowCount + x.MetricRowCount + x.MeasurementRowCount)
            .ThenByDescending(x => x.MatchedFileCount)
            .ThenBy(x => x.Model, StringComparer.OrdinalIgnoreCase)
            .Take(20)
            .Select(x =>
            {
                x.SourceFileIds = x.SourceFileIds.Take(12).ToList();
                x.ExampleFiles = x.ExampleFiles.Take(6).ToList();
                x.EvidenceLevel = x.VerifiedReviewCaseCount > 0
                    ? "verified"
                    : x.PairRowCount + x.MetricRowCount + x.MeasurementRowCount > 0
                        ? "fallbackRows"
                        : "candidateFilesOnly";
                return x;
            })
            .ToList();
    }

    private static MicroSpeakerModelCoverage GetCoverage(
        Dictionary<string, MicroSpeakerModelCoverage> map,
        string models)
    {
        string model = PrimaryModel(models);
        if (!map.TryGetValue(model, out MicroSpeakerModelCoverage? coverage))
        {
            coverage = new MicroSpeakerModelCoverage { Model = model };
            map[model] = coverage;
        }

        AddDistinct(coverage.ModelAliases, SplitModelNames(models));
        return coverage;
    }

    private static Dictionary<long, MicroSpeakerReviewCaseFileMetadata> ReadMicroSpeakerFileMetadata(string databasePath)
    {
        var rows = new Dictionary<long, MicroSpeakerReviewCaseFileMetadata>();
        if (!File.Exists(databasePath)) return rows;

        try
        {
            using SqliteConnection conn = OpenReadOnly(databasePath);
            if (!TableExists(conn, "files")) return rows;

            using SqliteCommand cmd = conn.CreateCommand();
            cmd.CommandText = """
                SELECT file_id, dataset, file_name, models, categories
                FROM files
                WHERE file_id IS NOT NULL;
                """;
            using SqliteDataReader r = cmd.ExecuteReader();
            while (r.Read())
            {
                long? fileId = L(r, 0);
                if (!fileId.HasValue) continue;
                rows[fileId.Value] = new MicroSpeakerReviewCaseFileMetadata(
                    S(r, 1),
                    S(r, 2),
                    S(r, 3),
                    S(r, 4));
            }
        }
        catch (SqliteException)
        {
            return [];
        }
        catch (IOException)
        {
            return [];
        }

        return rows;
    }

    private static List<MicroSpeakerVerifiedReviewCaseEvidence> ReadVerifiedReviewCaseFile(
        string repoRoot,
        JsonElement entry,
        string verificationPath,
        IReadOnlyList<string> terms,
        IReadOnlyList<List<string>> groups,
        IReadOnlyDictionary<long, MicroSpeakerReviewCaseFileMetadata> fileMetadata,
        string productTypeFilter)
    {
        var rows = new List<MicroSpeakerVerifiedReviewCaseEvidence>();
        try
        {
            using JsonDocument verificationDoc = JsonDocument.Parse(File.ReadAllText(verificationPath));
            JsonElement verificationRoot = verificationDoc.RootElement;
            MicroSpeakerVerifiedReviewCaseVerification verification = ReadReviewCaseVerification(verificationRoot);
            if (!verification.ApprovedForAskAi) return [];
            if (!string.Equals(verification.AiReviewCaseStatus, "verified", StringComparison.OrdinalIgnoreCase)) return [];

            long? sourceFileId = JsonLong(verificationRoot, "sourceFileId") ?? JsonLong(entry, "fileId");
            MicroSpeakerReviewCaseFileMetadata metadata = sourceFileId.HasValue
                                                          && fileMetadata.TryGetValue(sourceFileId.Value, out MicroSpeakerReviewCaseFileMetadata? found)
                ? found
                : new MicroSpeakerReviewCaseFileMetadata("", "", "", "");
            if (!ReviewCaseModelMatches(metadata, productTypeFilter)) return [];

            string sourceFile = FirstNonBlank(JsonString(verificationRoot, "sourceFile"), JsonString(entry, "sourceFile"));
            string draftPath = ResolveReviewCaseDraftPath(repoRoot, JsonString(verificationRoot, "sourceDraftPath"), sourceFileId);
            if (!File.Exists(draftPath)) return [];

            using JsonDocument draftDoc = JsonDocument.Parse(File.ReadAllText(draftPath));
            foreach (JsonElement caseElement in EnumerateReviewCaseElements(draftDoc.RootElement))
            {
                var row = BuildVerifiedReviewCaseEvidence(
                    repoRoot,
                    draftDoc.RootElement,
                    caseElement,
                    sourceFileId,
                    sourceFile,
                    draftPath,
                    verificationPath,
                    JsonBool(verificationRoot, "manualDraftUsed"),
                    JsonString(verificationRoot, "verifiedAt"),
                    verification,
                    metadata);

                FinishMatch(row, terms, groups, row.SearchText);
                rows.Add(row);
            }
        }
        catch (IOException)
        {
            return [];
        }
        catch (JsonException)
        {
            return [];
        }

        return rows;
    }

    private static bool ReviewCaseModelMatches(MicroSpeakerReviewCaseFileMetadata metadata, string productTypeFilter)
    {
        if (string.IsNullOrWhiteSpace(productTypeFilter)) return true;
        string filter = productTypeFilter.Trim();
        return ContainsAny(metadata.Models, filter)
               || ContainsAny(metadata.Dataset, filter)
               || ContainsAny(metadata.FileName, filter);
    }

    private static MicroSpeakerVerifiedReviewCaseEvidence BuildVerifiedReviewCaseEvidence(
        string repoRoot,
        JsonElement draftRoot,
        JsonElement caseElement,
        long? sourceFileId,
        string sourceFile,
        string draftPath,
        string verificationPath,
        bool manualDraftUsed,
        string verifiedAt,
        MicroSpeakerVerifiedReviewCaseVerification verification,
        MicroSpeakerReviewCaseFileMetadata metadata)
    {
        string reviewCaseId = FirstNonBlank(
            JsonString(caseElement, "reviewCaseId"),
            JsonString(caseElement, "reviewCaseKey"),
            sourceFileId.HasValue ? $"ms-{sourceFileId.Value}-verified" : "verified-reviewcase");

        var row = new MicroSpeakerVerifiedReviewCaseEvidence
        {
            SourceFileId = sourceFileId,
            SourceDataset = metadata.Dataset,
            SourceFile = FirstNonBlank(sourceFile, JsonString(draftRoot, "sourceFile"), metadata.FileName),
            SourceModels = metadata.Models,
            SourceCategories = metadata.Categories,
            OriginalFileUrl = SourceFileUrl(sourceFileId),
            DraftPath = RepoRelativePath(repoRoot, draftPath),
            VerificationPath = RepoRelativePath(repoRoot, verificationPath),
            ManualDraftUsed = manualDraftUsed,
            VerifiedAt = verifiedAt,
            ReviewCaseId = reviewCaseId,
            ReviewTitle = FirstNonBlank(JsonString(caseElement, "reviewTitle"), JsonString(draftRoot, "reviewTitle"), sourceFile),
            ReviewPurpose = FirstNonBlank(JsonString(caseElement, "reviewPurpose"), JsonString(draftRoot, "reviewPurpose")),
            ReviewType = FirstNonBlank(JsonString(caseElement, "reviewType"), JsonString(draftRoot, "reviewType")),
            Verification = verification,
            ChangedFactors = ReadReviewCaseChangedFactors(caseElement),
            Outcomes = ReadReviewCaseOutcomes(caseElement),
            Limitations = JsonStringList(caseElement, "limitations").Take(ReviewCaseEvidenceRowLimit).ToList(),
        };

        if (TryGetJsonProperty(draftRoot, "verification", out JsonElement draftVerification)
            && draftVerification.ValueKind == JsonValueKind.Object)
        {
            var limitations = row.Limitations.ToList();
            AddDistinct(limitations, JsonStringList(draftVerification, "limitations"));
            row.Limitations = limitations.Take(ReviewCaseEvidenceRowLimit).ToList();
        }

        if (TryGetJsonProperty(draftRoot, "finalSourceDecision", out JsonElement finalDecision)
            && finalDecision.ValueKind == JsonValueKind.Object)
        {
            row.SourceDecision = JoinNonBlank(" | ", JsonString(finalDecision, "text"), JsonString(finalDecision, "interpretation"));
            row.SourceDecisionEvidenceRows = JsonStringList(finalDecision, "evidenceRows").Take(ReviewCaseEvidenceRowLimit).ToList();
        }

        var evidenceRows = new List<string>();
        AddDistinct(evidenceRows, JsonStringList(caseElement, "evidenceRows"));
        foreach (MicroSpeakerVerifiedReviewCaseChangedFactor factor in row.ChangedFactors)
            AddDistinct(evidenceRows, factor.EvidenceRows);
        foreach (MicroSpeakerVerifiedReviewCaseOutcome outcome in row.Outcomes)
        {
            AddDistinct(evidenceRows, outcome.ComparisonRows);
            AddDistinct(evidenceRows, outcome.EvidenceRows);
            foreach (MicroSpeakerVerifiedReviewCaseSubResult subResult in outcome.SubResults)
                AddDistinct(evidenceRows, subResult.EvidenceRows);
        }
        AddDistinct(evidenceRows, row.SourceDecisionEvidenceRows);
        row.EvidenceRows = evidenceRows.Take(ReviewCaseEvidenceRowLimit).ToList();

        return row;
    }

    private static MicroSpeakerVerifiedReviewCaseVerification ReadReviewCaseVerification(JsonElement root)
    {
        JsonElement source = root;
        if (TryGetJsonProperty(root, "aiVerification", out JsonElement aiVerification)
            && aiVerification.ValueKind == JsonValueKind.Object)
            source = aiVerification;

        return new MicroSpeakerVerifiedReviewCaseVerification
        {
            Model = JsonString(root, "model"),
            AiReviewCaseStatus = JsonString(source, "aiReviewCaseStatus", "status"),
            VerificationStatus = JsonString(source, "verificationStatus"),
            ApprovedForAskAi = JsonBool(source, "approvedForAskAi"),
            Confidence = JsonString(source, "confidence"),
            Summary = JsonString(source, "summary"),
            Issues = JsonStringList(source, "issues").Take(ReviewCaseEvidenceRowLimit).ToList(),
            RequiredUserQuestions = JsonStringList(source, "requiredUserQuestions").Take(ReviewCaseEvidenceRowLimit).ToList(),
            CorrectionPlan = JsonStringList(source, "correctionPlan").Take(ReviewCaseEvidenceRowLimit).ToList(),
            EvidencePolicy = JsonString(source, "evidencePolicy"),
        };
    }

    private static IEnumerable<JsonElement> EnumerateReviewCaseElements(JsonElement root)
    {
        if (TryGetJsonProperty(root, "reviewCases", out JsonElement reviewCases)
            && reviewCases.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement reviewCase in reviewCases.EnumerateArray())
                yield return reviewCase;
            yield break;
        }

        yield return root;
    }

    private static List<MicroSpeakerVerifiedReviewCaseChangedFactor> ReadReviewCaseChangedFactors(JsonElement root)
    {
        if (!TryGetJsonProperty(root, "changedFactors", out JsonElement factors)
            || factors.ValueKind != JsonValueKind.Array)
            return [];

        var rows = new List<MicroSpeakerVerifiedReviewCaseChangedFactor>();
        foreach (JsonElement factor in factors.EnumerateArray())
        {
            rows.Add(new MicroSpeakerVerifiedReviewCaseChangedFactor
            {
                ChangedFactorId = FirstNonBlank(JsonString(factor, "changedFactorId"), JsonString(factor, "changeKey")),
                ChangeDomain = JsonStringList(factor, "changeDomain"),
                ChangedFactor = JsonString(factor, "changedFactor"),
                BaselineCondition = FirstNonBlank(JsonString(factor, "baselineCondition"), JsonString(factor, "beforeCondition")),
                ChangedCondition = FirstNonBlank(JsonString(factor, "changedCondition"), JsonString(factor, "afterCondition")),
                ChangedConditions = JsonStringList(factor, "changedConditions", "afterConditions"),
                SubgroupKeys = JsonStringList(factor, "subgroupKeys", "subgroupDimensions"),
                EvidenceRows = JsonStringList(factor, "evidenceRows").Take(ReviewCaseEvidenceRowLimit).ToList(),
            });
        }

        return rows.Take(ReviewCaseChangedFactorLimit).ToList();
    }

    private static List<MicroSpeakerVerifiedReviewCaseOutcome> ReadReviewCaseOutcomes(JsonElement root)
    {
        if (!TryGetJsonProperty(root, "outcomes", out JsonElement outcomes)
            || outcomes.ValueKind != JsonValueKind.Array)
            return [];

        var rows = new List<MicroSpeakerVerifiedReviewCaseOutcome>();
        foreach (JsonElement outcome in outcomes.EnumerateArray())
        {
            rows.Add(new MicroSpeakerVerifiedReviewCaseOutcome
            {
                OutcomeId = FirstNonBlank(JsonString(outcome, "outcomeId"), JsonString(outcome, "outcomeKey")),
                ChangedFactorId = JsonString(outcome, "changedFactorId"),
                OutcomeDomain = JsonString(outcome, "outcomeDomain"),
                OutcomeMetric = JsonString(outcome, "outcomeMetric"),
                Judgement = JsonString(outcome, "judgement"),
                ResultSummary = JsonString(outcome, "resultSummary"),
                SourceJudgement = JsonString(outcome, "sourceJudgement"),
                ComparisonRows = JsonStringList(outcome, "comparisonRows").Take(ReviewCaseEvidenceRowLimit).ToList(),
                EvidenceRows = JsonStringList(outcome, "evidenceRows").Take(ReviewCaseEvidenceRowLimit).ToList(),
                Notes = JsonStringList(outcome, "notes").Take(ReviewCaseEvidenceRowLimit).ToList(),
                Limitations = JsonStringList(outcome, "limitations").Take(ReviewCaseEvidenceRowLimit).ToList(),
                SubResults = ReadReviewCaseSubResults(outcome),
            });
        }

        return rows.Take(ReviewCaseOutcomeLimit).ToList();
    }

    private static List<MicroSpeakerVerifiedReviewCaseSubResult> ReadReviewCaseSubResults(JsonElement root)
    {
        if (!TryGetJsonProperty(root, "subResults", out JsonElement subResults)
            || subResults.ValueKind != JsonValueKind.Array)
            return [];

        var rows = new List<MicroSpeakerVerifiedReviewCaseSubResult>();
        foreach (JsonElement subResult in subResults.EnumerateArray())
        {
            rows.Add(new MicroSpeakerVerifiedReviewCaseSubResult
            {
                PairId = JsonLong(subResult, "pairId"),
                StatId = JsonLong(subResult, "statId"),
                Date = JsonString(subResult, "date"),
                TestRound = JsonString(subResult, "testRound"),
                Condition = JsonString(subResult, "condition"),
                Spec = JsonString(subResult, "spec"),
                Summary = JsonString(subResult, "summary"),
                ControlCondition = JsonString(subResult, "controlCondition"),
                TestCondition = JsonString(subResult, "testCondition"),
                ControlInput = JsonNestedDouble(subResult, "control", "input"),
                ControlNg = JsonNestedDouble(subResult, "control", "ng"),
                ControlRatePercent = JsonNestedDouble(subResult, "control", "ratePercent"),
                TestInput = JsonNestedDouble(subResult, "test", "input"),
                TestNg = JsonNestedDouble(subResult, "test", "ng"),
                TestRatePercent = JsonNestedDouble(subResult, "test", "ratePercent"),
                DeltaRatePercentPoint = JsonDouble(subResult, "deltaRatePercentPoint"),
                EffectDirection = JsonString(subResult, "effectDirection"),
                MinValue = JsonDouble(subResult, "minValue"),
                MaxValue = JsonDouble(subResult, "maxValue"),
                AvgValue = JsonDouble(subResult, "avgValue"),
                SampleCount = JsonLong(subResult, "sampleCount"),
                ViolationCount = JsonLong(subResult, "violationCount"),
                EvidenceRows = JsonStringList(subResult, "evidenceRows", "rows").Take(ReviewCaseEvidenceRowLimit).ToList(),
            });
        }

        return rows.Take(ReviewCaseSubResultLimit).ToList();
    }

    private static List<MicroSpeakerPairEvidenceRow> ReadMicroSpeakerPairs(
        SqliteConnection conn,
        IReadOnlyList<string> terms,
        IReadOnlyList<List<string>> groups,
        string productTypeFilter)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        string[] columns =
        [
            "f.dataset", "f.path", "f.file_name", "f.sheet_names", "f.models", "f.categories", "f.term_summary",
            "p.table_title", "p.compare_item", "p.control_condition", "p.test_condition", "p.effect_direction", "p.evidence", "p.pair_confidence"
        ];
        string where = BuildAnyLike(columns, terms, "p", cmd);
        AddMicroProductFilter(cmd, ref where, productTypeFilter);

        cmd.CommandText = $"""
            SELECT p.pair_id, p.file_id, f.dataset, f.file_name, f.models, f.categories, f.term_summary,
                   p.table_title, p.compare_item, p.control_condition, p.test_condition,
                   p.control_input, p.control_ng, p.control_rate, p.test_input, p.test_ng, p.test_rate,
                   p.delta_rate, p.improvement_rate, p.effect_direction, p.evidence, p.pair_confidence
            FROM comparison_pairs p
            JOIN files f ON f.file_id = p.file_id
            WHERE {where}
            ORDER BY ABS(COALESCE(p.delta_rate, 0)) DESC, p.pair_id
            LIMIT {CandidateLimit};
            """;

        var rows = new List<MicroSpeakerPairEvidenceRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            var row = new MicroSpeakerPairEvidenceRow
            {
                PairId = L(r, 0),
                FileId = L(r, 1),
                Dataset = S(r, 2),
                FileName = S(r, 3),
                OriginalFileUrl = SourceFileUrl(L(r, 1)),
                Models = S(r, 4),
                Categories = S(r, 5),
                TermSummary = S(r, 6),
                TableTitle = S(r, 7),
                CompareItem = S(r, 8),
                ControlCondition = S(r, 9),
                TestCondition = S(r, 10),
                ControlInput = D(r, 11),
                ControlNg = D(r, 12),
                ControlRate = D(r, 13),
                ControlRatePercent = RatePercent(D(r, 13)),
                TestInput = D(r, 14),
                TestNg = D(r, 15),
                TestRate = D(r, 16),
                TestRatePercent = RatePercent(D(r, 16)),
                DeltaRate = D(r, 17),
                DeltaRatePercentPoint = RatePercent(D(r, 17)),
                ImprovementRate = D(r, 18),
                EffectDirection = S(r, 19),
                Evidence = S(r, 20),
                PairConfidence = S(r, 21),
            };
            row.RelativeChangePercent = RelativeChangePercent(row.ControlRate, row.TestRate);
            FinishMatch(row, terms, groups, row.SearchText);
            rows.Add(row);
        }

        return SelectRows(rows, PairLimit, r => r.MatchScore, r => Math.Abs(r.DeltaRate ?? 0));
    }

    private static List<MicroSpeakerMetricEvidenceRow> ReadMicroSpeakerMetrics(
        SqliteConnection conn,
        IReadOnlyList<string> terms,
        IReadOnlyList<List<string>> groups,
        string productTypeFilter)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        string[] columns =
        [
            "f.dataset", "f.path", "f.file_name", "f.sheet_names", "f.models", "f.categories", "f.term_summary",
            "m.sheet_name", "m.table_title", "m.condition_label", "m.detail", "m.raw_row", "m.parse_confidence"
        ];
        string where = BuildAnyLike(columns, terms, "m", cmd);
        AddMicroProductFilter(cmd, ref where, productTypeFilter);

        cmd.CommandText = $"""
            SELECT m.metric_id, m.file_id, f.dataset, f.file_name, f.models, f.categories, f.term_summary,
                   m.sheet_name, m.row_number, m.table_title, m.condition_label,
                   m.input_qty, m.ok_qty, m.ng_qty, m.ng_rate, m.detail, m.raw_row, m.parse_confidence
            FROM metric_candidates m
            JOIN files f ON f.file_id = m.file_id
            WHERE {where}
            ORDER BY COALESCE(m.ng_rate, 0) DESC, m.metric_id
            LIMIT {CandidateLimit};
            """;

        var rows = new List<MicroSpeakerMetricEvidenceRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            var row = new MicroSpeakerMetricEvidenceRow
            {
                MetricId = L(r, 0),
                FileId = L(r, 1),
                Dataset = S(r, 2),
                FileName = S(r, 3),
                OriginalFileUrl = SourceFileUrl(L(r, 1)),
                Models = S(r, 4),
                Categories = S(r, 5),
                TermSummary = S(r, 6),
                SheetName = S(r, 7),
                RowNumber = L(r, 8),
                TableTitle = S(r, 9),
                ConditionLabel = S(r, 10),
                InputQty = D(r, 11),
                OkQty = D(r, 12),
                NgQty = D(r, 13),
                NgRate = D(r, 14),
                NgRatePercent = RatePercent(D(r, 14)),
                Detail = S(r, 15),
                RawRow = S(r, 16),
                ParseConfidence = S(r, 17),
            };
            FinishMatch(row, terms, groups, row.SearchText);
            rows.Add(row);
        }

        return SelectRows(rows, MetricLimit, r => r.MatchScore, r => r.NgRate ?? 0);
    }

    private static List<MicroSpeakerMeasurementEvidenceRow> ReadMicroSpeakerMeasurements(
        SqliteConnection conn,
        IReadOnlyList<string> terms,
        IReadOnlyList<List<string>> groups,
        string productTypeFilter)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        string[] columns =
        [
            "f.dataset", "f.path", "f.file_name", "f.sheet_names", "f.models", "f.categories", "f.term_summary",
            "s.sheet_name", "s.item_label", "s.condition_label", "s.spec", "s.raw_row", "s.parse_confidence"
        ];
        string where = BuildAnyLike(columns, terms, "s", cmd);
        AddMicroProductFilter(cmd, ref where, productTypeFilter);

        cmd.CommandText = $"""
            SELECT s.stat_id, s.file_id, f.dataset, f.file_name, f.models, f.categories, f.term_summary,
                   s.sheet_name, s.row_number, s.item_label, s.condition_label, s.spec,
                   s.min_value, s.max_value, s.avg_value, s.sample_count, s.violation_count, s.raw_row, s.parse_confidence
            FROM measurement_stats s
            JOIN files f ON f.file_id = s.file_id
            WHERE {where}
            ORDER BY COALESCE(s.violation_count, 0) DESC, s.stat_id
            LIMIT {CandidateLimit};
            """;

        var rows = new List<MicroSpeakerMeasurementEvidenceRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            var row = new MicroSpeakerMeasurementEvidenceRow
            {
                StatId = L(r, 0),
                FileId = L(r, 1),
                Dataset = S(r, 2),
                FileName = S(r, 3),
                OriginalFileUrl = SourceFileUrl(L(r, 1)),
                Models = S(r, 4),
                Categories = S(r, 5),
                TermSummary = S(r, 6),
                SheetName = S(r, 7),
                RowNumber = L(r, 8),
                ItemLabel = S(r, 9),
                ConditionLabel = S(r, 10),
                Spec = S(r, 11),
                MinValue = D(r, 12),
                MaxValue = D(r, 13),
                AvgValue = D(r, 14),
                SampleCount = L(r, 15),
                ViolationCount = L(r, 16),
                RawRow = S(r, 17),
                ParseConfidence = S(r, 18),
            };
            FinishMatch(row, terms, groups, row.SearchText);
            rows.Add(row);
        }

        return SelectRows(rows, MeasurementLimit, r => r.MatchScore, r => r.ViolationCount ?? 0);
    }

    private static List<JinoDocumentEvidenceRow> ReadJinoDocuments(
        SqliteConnection conn,
        IReadOnlyList<string> terms,
        IReadOnlyList<List<string>> groups,
        string productTypeFilter)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        string[] columns =
        [
            "r.DatasetName", "r.ProductType", "d.SourceDataset", "d.SourceFile", "d.Title", "d.Model",
            "d.ReportType", "d.PrimaryDefect", "d.RelatedDefectsJson", "d.PartsJson", "d.ProcessesJson",
            "d.Purpose", "d.ContentJson", "d.GeneratedReportMarkdown"
        ];
        string where = BuildAnyLike(columns, terms, "d", cmd);
        AddJinoProductFilter(cmd, ref where, productTypeFilter);

        cmd.CommandText = $"""
            SELECT d.DocumentId, d.SourceDataset, d.SourceFile, COALESCE(r.ProductType, ''),
                   d.Title, d.Model, d.ReportType, d.PrimaryDefect, d.Purpose,
                   d.GeneratedReportMarkdown, d.RelatedDefectsJson, d.PartsJson, d.ProcessesJson,
                   COALESCE((SELECT COUNT(*) FROM AiResults ar WHERE ar.DocumentId=d.DocumentId), 0) AS ResultCount,
                   COALESCE((SELECT COUNT(*) FROM AiTestConditions ac WHERE ac.DocumentId=d.DocumentId), 0) AS ConditionCount,
                   d.UpdatedAt
            FROM AiDocuments d
            LEFT JOIN RawReports r ON r.DatasetName = d.SourceDataset
            WHERE {where}
              AND COALESCE(r.BatchExcluded, 0) = 0
            ORDER BY ResultCount DESC, d.UpdatedAt DESC
            LIMIT {CandidateLimit};
            """;

        var rows = new List<JinoDocumentEvidenceRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            var row = new JinoDocumentEvidenceRow
            {
                DocumentId = S(r, 0),
                SourceDataset = S(r, 1),
                SourceFile = S(r, 2),
                ProductType = S(r, 3),
                Title = S(r, 4),
                Model = S(r, 5),
                ReportType = S(r, 6),
                PrimaryDefect = S(r, 7),
                Purpose = S(r, 8),
                GeneratedReportPreview = Brief(S(r, 9), 900),
                RelatedDefectsJson = S(r, 10),
                PartsJson = S(r, 11),
                ProcessesJson = S(r, 12),
                ResultCount = L(r, 13),
                ConditionCount = L(r, 14),
                UpdatedAt = S(r, 15),
            };
            FinishMatch(row, terms, groups, row.SearchText);
            rows.Add(row);
        }

        return SelectRows(rows, JinoDocumentLimit, r => r.MatchScore, r => r.ResultCount ?? 0);
    }

    private static List<JinoResultEvidenceRow> ReadJinoResults(
        SqliteConnection conn,
        IReadOnlyList<string> terms,
        IReadOnlyList<List<string>> groups,
        string productTypeFilter)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        string[] columns =
        [
            "r.DatasetName", "r.ProductType", "d.SourceDataset", "d.SourceFile", "d.Title", "d.Model",
            "d.ReportType", "d.PrimaryDefect", "d.Purpose", "ar.MeasurementType", "ar.ConditionGroup",
            "ar.Line", "ar.MetricName", "ar.Judgement", "ar.SourceFile", "ar.SheetName", "ar.SourceCellsJson",
            "c.Process", "c.ChangedFactor", "c.BeforeValue", "c.AfterValue"
        ];
        string where = BuildAnyLike(columns, terms, "rj", cmd);
        AddJinoProductFilter(cmd, ref where, productTypeFilter);

        cmd.CommandText = $"""
            SELECT ar.ResultId, d.DocumentId, d.SourceDataset, d.SourceFile, COALESCE(r.ProductType, ''),
                   d.Title, d.ReportType, d.PrimaryDefect,
                   ar.MeasurementType, ar.ConditionGroup, COALESCE(c.Process, ''), COALESCE(c.ChangedFactor, ''),
                   COALESCE(c.BeforeValue, ''), COALESCE(c.AfterValue, ''),
                   ar.InputCount, ar.OkCount, ar.NgCount, ar.NgRateDecimal, ar.NgRatePercent,
                   ar.MetricName, ar.MetricValue, COALESCE(ar.Unit, ''), COALESCE(ar.Judgement, ''),
                   ar.SourceFile, ar.SheetName, ar.SourceCellsJson
            FROM AiResults ar
            JOIN AiDocuments d ON d.DocumentId = ar.DocumentId
            LEFT JOIN RawReports r ON r.DatasetName = d.SourceDataset
            LEFT JOIN AiTestConditions c ON c.ConditionId = ar.ConditionId
            WHERE {where}
              AND COALESCE(r.BatchExcluded, 0) = 0
            ORDER BY ABS(COALESCE(ar.NgRatePercent, ar.MetricValue, 0)) DESC, ar.ResultId
            LIMIT {CandidateLimit};
            """;

        var rows = new List<JinoResultEvidenceRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            var row = new JinoResultEvidenceRow
            {
                ResultId = S(r, 0),
                DocumentId = S(r, 1),
                SourceDataset = S(r, 2),
                SourceFile = S(r, 3),
                ProductType = S(r, 4),
                Title = S(r, 5),
                ReportType = S(r, 6),
                PrimaryDefect = S(r, 7),
                MeasurementType = S(r, 8),
                ConditionGroup = S(r, 9),
                Process = S(r, 10),
                ChangedFactor = S(r, 11),
                BeforeValue = S(r, 12),
                AfterValue = S(r, 13),
                InputCount = D(r, 14),
                OkCount = D(r, 15),
                NgCount = D(r, 16),
                NgRateDecimal = D(r, 17),
                NgRatePercent = D(r, 18) ?? RatePercent(D(r, 17)),
                MetricName = S(r, 19),
                MetricValue = D(r, 20),
                Unit = S(r, 21),
                Judgement = S(r, 22),
                ResultSourceFile = S(r, 23),
                SheetName = S(r, 24),
                SourceCellsJson = Brief(S(r, 25), 900),
            };
            FinishMatch(row, terms, groups, row.SearchText);
            rows.Add(row);
        }

        return SelectRows(rows, JinoResultLimit, r => r.MatchScore, r => Math.Abs(r.NgRatePercent ?? r.MetricValue ?? 0));
    }

    private static List<MicroSpeakerPairAggregate> BuildPairAggregates(IReadOnlyList<MicroSpeakerPairEvidenceRow> rows)
        => rows
            .GroupBy(r => string.IsNullOrWhiteSpace(r.CompareItem) ? "(blank)" : r.CompareItem.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => new MicroSpeakerPairAggregate
            {
                CompareItem = g.Key,
                RowCount = g.Count(),
                WorsenedCount = g.Count(x => string.Equals(x.EffectDirection, "WORSENED", StringComparison.OrdinalIgnoreCase)),
                ImprovedCount = g.Count(x => string.Equals(x.EffectDirection, "IMPROVED", StringComparison.OrdinalIgnoreCase)),
                NoChangeCount = g.Count(x => string.Equals(x.EffectDirection, "NO_CHANGE", StringComparison.OrdinalIgnoreCase)),
                MaxAbsRelativeChangePercent = g.Select(x => x.RelativeChangePercent).Where(x => x.HasValue).Select(x => Math.Abs(x!.Value)).DefaultIfEmpty(0).Max(),
                MaxAbsDeltaRatePercentPoint = g.Select(x => x.DeltaRatePercentPoint).Where(x => x.HasValue).Select(x => Math.Abs(x!.Value)).DefaultIfEmpty(0).Max(),
                ExampleDatasets = g.Select(x => x.Dataset).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Take(8).ToList(),
            })
            .OrderByDescending(x => x.RowCount)
            .ThenByDescending(x => x.MaxAbsRelativeChangePercent)
            .Take(30)
            .ToList();

    private static List<MicroSpeakerPairConditionAggregate> BuildPairConditionAggregates(IReadOnlyList<MicroSpeakerPairEvidenceRow> rows)
        => rows
            .Where(r => r.FileId.HasValue
                        && !string.IsNullOrWhiteSpace(r.TableTitle)
                        && !string.IsNullOrWhiteSpace(r.ControlCondition)
                        && !string.IsNullOrWhiteSpace(r.TestCondition))
            .GroupBy(r => new
            {
                FileId = r.FileId!.Value,
                Table = NormalizeAggregateKey(r.TableTitle),
                Compare = NormalizeAggregateKey(r.CompareItem),
                Control = NormalizeAggregateKey(r.ControlCondition),
                Test = NormalizeAggregateKey(r.TestCondition),
            })
            .Select(g => BuildPairConditionAggregate(g.ToList()))
            .Where(x => x.SourceRowCount > 1)
            .OrderByDescending(x => x.SourceRowCount)
            .ThenByDescending(x => Math.Abs(x.RelativeChangePercent ?? 0))
            .Take(80)
            .ToList();

    private static MicroSpeakerPairConditionAggregate BuildPairConditionAggregate(List<MicroSpeakerPairEvidenceRow> rows)
    {
        rows = rows.OrderBy(r => r.PairId ?? 0).ToList();
        MicroSpeakerPairEvidenceRow first = rows[0];
        MicroSpeakerPairEvidenceRow? totalRow = FindTotalEquivalentRow(rows);
        IReadOnlyList<MicroSpeakerPairEvidenceRow> usedRows = totalRow is null ? rows : [totalRow];

        double? controlInput = first.ControlInput;
        double? controlNg = first.ControlNg;
        double? controlRate = first.ControlRate ?? RateFromCounts(controlInput, controlNg);
        double? testInput = SumNullable(usedRows.Select(r => r.TestInput));
        double? testNg = SumNullable(usedRows.Select(r => r.TestNg));
        double? testRate = RateFromCounts(testInput, testNg) ?? (usedRows.Count == 1 ? usedRows[0].TestRate : null);
        double? delta = testRate.HasValue && controlRate.HasValue ? testRate.Value - controlRate.Value : null;
        string method = totalRow is not null ? "total_row" : "summed_rows";

        return new MicroSpeakerPairConditionAggregate
        {
            FileId = first.FileId,
            Dataset = first.Dataset,
            FileName = first.FileName,
            OriginalFileUrl = first.OriginalFileUrl,
            TableTitle = first.TableTitle,
            CompareItem = first.CompareItem,
            ControlCondition = first.ControlCondition,
            TestCondition = first.TestCondition,
            ControlInput = controlInput,
            ControlNg = controlNg,
            ControlRate = controlRate,
            ControlRatePercent = RatePercent(controlRate),
            TestInput = testInput,
            TestNg = testNg,
            TestRate = testRate,
            TestRatePercent = RatePercent(testRate),
            DeltaRate = delta,
            DeltaRatePercentPoint = RatePercent(delta),
            RelativeChangePercent = RelativeChangePercent(controlRate, testRate),
            EffectDirection = AggregateEffectDirection(delta),
            SourceRowCount = rows.Count,
            UsedRowCount = usedRows.Count,
            AggregationMethod = method,
            Evidence = string.Join("; ", usedRows.Select(r => r.Evidence).Where(x => !string.IsNullOrWhiteSpace(x)).Take(6)),
            Limit = totalRow is not null
                ? "Total-equivalent row selected; daily rows not double-counted."
                : "No total-equivalent row found; repeated same-condition rows were summed.",
        };
    }

    private static MicroSpeakerPairEvidenceRow? FindTotalEquivalentRow(IReadOnlyList<MicroSpeakerPairEvidenceRow> rows)
    {
        if (rows.Count < 3) return null;
        foreach (MicroSpeakerPairEvidenceRow candidate in rows)
        {
            double? input = candidate.TestInput;
            double? ng = candidate.TestNg;
            if (!input.HasValue || !ng.HasValue) continue;
            List<MicroSpeakerPairEvidenceRow> others = rows.Where(r => !ReferenceEquals(r, candidate)).ToList();
            double? otherInput = SumNullable(others.Select(r => r.TestInput));
            double? otherNg = SumNullable(others.Select(r => r.TestNg));
            if (NearlyEqual(input.Value, otherInput) && NearlyEqual(ng.Value, otherNg))
                return candidate;
        }
        return null;
    }

    private static string NormalizeAggregateKey(string value)
    {
        string normalized = Regex.Replace((value ?? "").Trim().ToLowerInvariant(), @"\s+", " ");
        normalized = Regex.Replace(normalized, @"^(total|subtotal)\s*\|\s*", "", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"^(total|subtotal)\s+", "", RegexOptions.IgnoreCase);
        return normalized;
    }

    private static double? SumNullable(IEnumerable<double?> values)
    {
        double total = 0;
        bool any = false;
        foreach (double? value in values)
        {
            if (!value.HasValue) continue;
            total += value.Value;
            any = true;
        }
        return any ? total : null;
    }

    private static double? RateFromCounts(double? input, double? ng)
    {
        if (!input.HasValue || !ng.HasValue || Math.Abs(input.Value) < 0.0000001) return null;
        return ng.Value / input.Value;
    }

    private static bool NearlyEqual(double expected, double? actual)
    {
        if (!actual.HasValue) return false;
        double tolerance = Math.Max(1.0, Math.Abs(expected) * 0.002);
        return Math.Abs(expected - actual.Value) <= tolerance;
    }

    private static string AggregateEffectDirection(double? delta)
    {
        if (!delta.HasValue || Math.Abs(delta.Value) < 0.0000001) return "NO_CHANGE";
        return delta.Value > 0 ? "WORSENED" : "IMPROVED";
    }

    private static void AddMicroProductFilter(SqliteCommand cmd, ref string where, string productTypeFilter)
    {
        if (string.IsNullOrWhiteSpace(productTypeFilter)) return;
        where = $"({where}) AND COALESCE(f.models, '') LIKE @modelFilter";
        cmd.Parameters.AddWithValue("@modelFilter", "%" + productTypeFilter.Trim() + "%");
    }

    private static void AddJinoProductFilter(SqliteCommand cmd, ref string where, string productTypeFilter)
    {
        if (string.IsNullOrWhiteSpace(productTypeFilter)) return;
        where = $"({where}) AND COALESCE(r.ProductType, '') = @productTypeFilter";
        cmd.Parameters.AddWithValue("@productTypeFilter", productTypeFilter.Trim());
    }

    private static string BuildAnyLike(string[] columns, IReadOnlyList<string> terms, string prefix, SqliteCommand cmd)
    {
        if (terms.Count == 0 || columns.Length == 0) return "1=0";

        var parts = new List<string>();
        for (int i = 0; i < terms.Count; i++)
        {
            string parameter = "@" + prefix + i.ToString(CultureInfo.InvariantCulture);
            cmd.Parameters.AddWithValue(parameter, "%" + terms[i] + "%");
            string columnSql = string.Join(" OR ", columns.Select(c => $"COALESCE({c}, '') LIKE {parameter}"));
            parts.Add("(" + columnSql + ")");
        }

        return "(" + string.Join(" OR ", parts) + ")";
    }

    private static List<T> SelectRows<T>(
        List<T> rows,
        int limit,
        Func<T, double> score,
        Func<T, double> secondary)
        where T : EvidenceMatchBase
    {
        List<T> strict = rows.Where(x => x.MatchesAllRequiredTerms).ToList();
        List<T> source = strict.Count > 0 ? strict : rows;
        bool fallback = strict.Count == 0 && rows.Count > 0;
        foreach (T row in source)
            row.StrictFallbackUsed = fallback;

        return source
            .OrderByDescending(score)
            .ThenByDescending(secondary)
            .Take(limit)
            .ToList();
    }

    private static void FinishMatch(EvidenceMatchBase row, IReadOnlyList<string> terms, IReadOnlyList<List<string>> groups, string searchText)
    {
        row.MatchesAllRequiredTerms = MatchesRequiredGroups(searchText, groups);
        row.MatchScore = MatchScore(searchText, terms, groups);
    }

    private static bool MatchesRequiredGroups(string text, IReadOnlyList<List<string>> groups)
    {
        if (groups.Count == 0) return true;
        return groups.All(group => group.Any(term => ContainsAny(text, term)));
    }

    private static double MatchScore(string text, IReadOnlyList<string> terms, IReadOnlyList<List<string>> groups)
    {
        double score = 0;
        foreach (string term in terms)
        {
            if (ContainsAny(text, term))
                score += term.Length >= 4 ? 2 : 1;
        }

        foreach (List<string> group in groups)
        {
            if (group.Any(term => ContainsAny(text, term)))
                score += 8;
        }

        return score;
    }

    private static bool ContainsAny(string text, params string[] terms)
    {
        string value = text ?? "";
        return terms.Any(term => !string.IsNullOrWhiteSpace(term)
                                 && value.Contains(term, StringComparison.OrdinalIgnoreCase));
    }

    private static double ReviewCaseConfidenceScore(MicroSpeakerVerifiedReviewCaseEvidence row)
    {
        double confidence = row.Verification.Confidence.ToLowerInvariant() switch
        {
            "high" => 3,
            "medium" => 2,
            "low" => 1,
            _ => 0,
        };
        return confidence + row.Outcomes.Count * 0.1 + row.ChangedFactors.Count * 0.05;
    }

    private static string ResolveRepoPath(string repoRoot, string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return "";
        string normalized = path
            .Replace('\\', Path.DirectorySeparatorChar)
            .Replace('/', Path.DirectorySeparatorChar);
        return Path.IsPathRooted(normalized)
            ? Path.GetFullPath(normalized)
            : Path.GetFullPath(Path.Combine(repoRoot, normalized));
    }

    private static string ResolveReviewCaseDraftPath(string repoRoot, string sourceDraftPath, long? sourceFileId)
    {
        string resolved = ResolveRepoPath(repoRoot, sourceDraftPath);
        if (File.Exists(resolved)) return resolved;

        if (sourceFileId.HasValue)
        {
            string fallback = Path.Combine(
                repoRoot,
                "REVIEWCASE_AI_DRAFTS",
                "batch",
                "files",
                $"{sourceFileId.Value}.reviewcase-draft.json");
            if (File.Exists(fallback)) return fallback;
        }

        return resolved;
    }

    private static string RepoRelativePath(string repoRoot, string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return "";
        try
        {
            string root = Path.GetFullPath(repoRoot);
            string fullPath = Path.GetFullPath(path);
            string relative = Path.GetRelativePath(root, fullPath);
            if (!relative.StartsWith("..", StringComparison.Ordinal))
                return relative.Replace('\\', '/');
        }
        catch (ArgumentException)
        {
        }
        catch (NotSupportedException)
        {
        }

        return path;
    }

    private static string FirstNonBlank(params string[] values)
        => values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v))?.Trim() ?? "";

    private static string JoinNonBlank(string separator, params string[] values)
        => string.Join(separator, values.Select(v => (v ?? "").Trim()).Where(v => !string.IsNullOrWhiteSpace(v)));

    private static void AddDistinct(List<string> target, IEnumerable<string> values)
    {
        foreach (string value in values)
        {
            string trimmed = (value ?? "").Trim();
            if (trimmed.Length == 0) continue;
            if (!target.Contains(trimmed, StringComparer.OrdinalIgnoreCase))
                target.Add(trimmed);
        }
    }

    private static string PrimaryModel(string models)
    {
        string first = SplitModelNames(models).FirstOrDefault() ?? "";
        return string.IsNullOrWhiteSpace(first) ? "model unspecified" : first;
    }

    private static List<string> SplitModelNames(string models)
        => (models ?? "")
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static bool TryGetJsonProperty(JsonElement root, string name, out JsonElement value)
    {
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty(name, out value))
            return true;

        if (root.ValueKind == JsonValueKind.Object)
        {
            foreach (JsonProperty property in root.EnumerateObject())
            {
                if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }

    private static string JsonString(JsonElement root, params string[] names)
    {
        foreach (string name in names)
        {
            if (!TryGetJsonProperty(root, name, out JsonElement value)) continue;
            string text = JsonValueText(value);
            if (!string.IsNullOrWhiteSpace(text)) return text;
        }

        return "";
    }

    private static List<string> JsonStringList(JsonElement root, params string[] names)
    {
        foreach (string name in names)
        {
            if (!TryGetJsonProperty(root, name, out JsonElement value)) continue;
            var rows = new List<string>();
            if (value.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement item in value.EnumerateArray())
                {
                    string text = JsonValueText(item);
                    if (!string.IsNullOrWhiteSpace(text))
                        rows.Add(text);
                }
            }
            else
            {
                string text = JsonValueText(value);
                if (!string.IsNullOrWhiteSpace(text))
                    rows.Add(text);
            }

            if (rows.Count > 0)
                return rows.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        return [];
    }

    private static bool JsonBool(JsonElement root, params string[] names)
    {
        foreach (string name in names)
        {
            if (!TryGetJsonProperty(root, name, out JsonElement value)) continue;
            if (value.ValueKind == JsonValueKind.True) return true;
            if (value.ValueKind == JsonValueKind.False) return false;
            if (value.ValueKind == JsonValueKind.String && bool.TryParse(value.GetString(), out bool result))
                return result;
        }

        return false;
    }

    private static long? JsonLong(JsonElement root, params string[] names)
    {
        foreach (string name in names)
        {
            if (!TryGetJsonProperty(root, name, out JsonElement value)) continue;
            if (value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out long number))
                return number;
            if (value.ValueKind == JsonValueKind.String
                && long.TryParse(value.GetString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out long parsed))
                return parsed;
        }

        return null;
    }

    private static double? JsonDouble(JsonElement root, params string[] names)
    {
        foreach (string name in names)
        {
            if (!TryGetJsonProperty(root, name, out JsonElement value)) continue;
            if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out double number))
                return number;
            if (value.ValueKind == JsonValueKind.String
                && double.TryParse(value.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
                return parsed;
        }

        return null;
    }

    private static double? JsonNestedDouble(JsonElement root, string objectName, params string[] names)
    {
        if (!TryGetJsonProperty(root, objectName, out JsonElement nested)
            || nested.ValueKind != JsonValueKind.Object)
            return null;

        return JsonDouble(nested, names);
    }

    private static string JsonValueText(JsonElement value)
        => value.ValueKind switch
        {
            JsonValueKind.String => value.GetString() ?? "",
            JsonValueKind.Number => value.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Array => string.Join("; ", value.EnumerateArray().Select(JsonValueText).Where(x => !string.IsNullOrWhiteSpace(x))),
            JsonValueKind.Object => Brief(value.GetRawText(), 900),
            _ => "",
        };

    private static double? RatePercent(double? value)
    {
        if (!value.HasValue) return null;
        double v = value.Value;
        return Math.Abs(v) <= 1.5 ? v * 100.0 : v;
    }

    private static double? RelativeChangePercent(double? normalRate, double? testRate)
    {
        if (!normalRate.HasValue || !testRate.HasValue) return null;
        if (Math.Abs(normalRate.Value) < 0.0000001) return null;
        return (testRate.Value / normalRate.Value - 1.0) * 100.0;
    }

    private static string SourceFileUrl(long? fileId)
        => fileId.HasValue ? $"/microspeaker/source-file/{fileId.Value}" : "";

    private static SqliteConnection OpenReadOnly(string path)
    {
        var cs = new SqliteConnectionStringBuilder
        {
            DataSource = path,
            Mode = SqliteOpenMode.ReadOnly,
            Cache = SqliteCacheMode.Shared,
        };
        var conn = new SqliteConnection(cs.ToString());
        conn.Open();
        return conn;
    }

    private static bool TableExists(SqliteConnection conn, string table)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 1 FROM sqlite_master WHERE type='table' AND name=@name LIMIT 1;";
        cmd.Parameters.AddWithValue("@name", table);
        return cmd.ExecuteScalar() is not null;
    }

    private static HashSet<string> TableColumns(SqliteConnection conn, string table)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = $"PRAGMA table_info({table});";
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
            set.Add(S(r, 1));
        return set;
    }

    private static string S(SqliteDataReader r, int i)
        => r.IsDBNull(i) ? "" : Convert.ToString(r.GetValue(i), CultureInfo.InvariantCulture) ?? "";

    private static long? L(SqliteDataReader r, int i)
        => r.IsDBNull(i) ? null : Convert.ToInt64(r.GetValue(i), CultureInfo.InvariantCulture);

    private static double? D(SqliteDataReader r, int i)
        => r.IsDBNull(i) ? null : Convert.ToDouble(r.GetValue(i), CultureInfo.InvariantCulture);

    private static string Brief(string text, int max)
    {
        string value = (text ?? "").Trim();
        if (value.Length <= max) return value;
        return value[..max] + "...";
    }
}

public sealed class MicroSpeakerAskEvidencePack
{
    [JsonPropertyName("createdAt")] public string CreatedAt { get; set; } = "";
    [JsonPropertyName("question")] public string Question { get; set; } = "";
    [JsonPropertyName("productTypeFilter")] public string ProductTypeFilter { get; set; } = "";
    [JsonPropertyName("jinoDatabasePath")] public string JinoDatabasePath { get; set; } = "";
    [JsonPropertyName("microSpeakerDatabasePath")] public string MicroSpeakerDatabasePath { get; set; } = "";
    [JsonPropertyName("searchTerms")] public List<string> SearchTerms { get; set; } = [];
    [JsonPropertyName("requiredTermGroups")] public List<List<string>> RequiredTermGroups { get; set; } = [];
    [JsonPropertyName("questionAnalysis")] public MicroSpeakerQuestionAnalysis QuestionAnalysis { get; set; } = new();
    [JsonPropertyName("microSpeaker")] public MicroSpeakerEvidenceSource MicroSpeaker { get; set; } = new();
    [JsonPropertyName("jino")] public JinoEvidenceSource Jino { get; set; } = new();
    [JsonPropertyName("notes")] public List<string> Notes { get; set; } = [];
}

public sealed class MicroSpeakerQuestionAnalysis
{
    [JsonPropertyName("factorAxisLabel")] public string FactorAxisLabel { get; set; } = "";
    [JsonPropertyName("outcomeAxisLabel")] public string OutcomeAxisLabel { get; set; } = "";
    [JsonPropertyName("factorTerms")] public List<string> FactorTerms { get; set; } = [];
    [JsonPropertyName("outcomeTerms")] public List<string> OutcomeTerms { get; set; } = [];
    [JsonPropertyName("suggestedReviewSections")] public List<string> SuggestedReviewSections { get; set; } = [];
    [JsonPropertyName("searchTermCount")] public int SearchTermCount { get; set; }
    [JsonPropertyName("requiredGroupCount")] public int RequiredGroupCount { get; set; }
}

public sealed class MicroSpeakerEvidenceSource
{
    [JsonPropertyName("databaseExists")] public bool DatabaseExists { get; set; }
    [JsonPropertyName("verifiedReviewCases")] public List<MicroSpeakerVerifiedReviewCaseEvidence> VerifiedReviewCases { get; set; } = [];
    [JsonPropertyName("modelCoverage")] public List<MicroSpeakerModelCoverage> ModelCoverage { get; set; } = [];
    [JsonPropertyName("termHits")] public List<TermHitEvidence> TermHits { get; set; } = [];
    [JsonPropertyName("pairAggregates")] public List<MicroSpeakerPairAggregate> PairAggregates { get; set; } = [];
    [JsonPropertyName("pairConditionAggregates")] public List<MicroSpeakerPairConditionAggregate> PairConditionAggregates { get; set; } = [];
    [JsonPropertyName("pairRows")] public List<MicroSpeakerPairEvidenceRow> PairRows { get; set; } = [];
    [JsonPropertyName("metricRows")] public List<MicroSpeakerMetricEvidenceRow> MetricRows { get; set; } = [];
    [JsonPropertyName("measurementRows")] public List<MicroSpeakerMeasurementEvidenceRow> MeasurementRows { get; set; } = [];
}

public sealed class JinoEvidenceSource
{
    [JsonPropertyName("databaseExists")] public bool DatabaseExists { get; set; }
    [JsonPropertyName("termHits")] public List<TermHitEvidence> TermHits { get; set; } = [];
    [JsonPropertyName("documentRows")] public List<JinoDocumentEvidenceRow> DocumentRows { get; set; } = [];
    [JsonPropertyName("resultRows")] public List<JinoResultEvidenceRow> ResultRows { get; set; } = [];
}

public sealed record TermHitEvidence(
    [property: JsonPropertyName("table")] string Table,
    [property: JsonPropertyName("term")] string Term,
    [property: JsonPropertyName("count")] long Count);

public sealed record MicroSpeakerReviewCaseFileMetadata(
    string Dataset,
    string FileName,
    string Models,
    string Categories);

public sealed class MicroSpeakerModelCoverage
{
    [JsonPropertyName("model")] public string Model { get; set; } = "";
    [JsonPropertyName("modelAliases")] public List<string> ModelAliases { get; set; } = [];
    [JsonPropertyName("evidenceLevel")] public string EvidenceLevel { get; set; } = "";
    [JsonPropertyName("matchedFileCount")] public int MatchedFileCount { get; set; }
    [JsonPropertyName("verifiedReviewCaseCount")] public int VerifiedReviewCaseCount { get; set; }
    [JsonPropertyName("pairRowCount")] public int PairRowCount { get; set; }
    [JsonPropertyName("metricRowCount")] public int MetricRowCount { get; set; }
    [JsonPropertyName("measurementRowCount")] public int MeasurementRowCount { get; set; }
    [JsonPropertyName("sourceFileIds")] public List<string> SourceFileIds { get; set; } = [];
    [JsonPropertyName("exampleFiles")] public List<string> ExampleFiles { get; set; } = [];
}

public sealed class MicroSpeakerModelFileCoverageRow : EvidenceMatchBase
{
    [JsonPropertyName("fileId")] public long? FileId { get; set; }
    [JsonPropertyName("dataset")] public string Dataset { get; set; } = "";
    [JsonPropertyName("fileName")] public string FileName { get; set; } = "";
    [JsonPropertyName("models")] public string Models { get; set; } = "";
    [JsonPropertyName("categories")] public string Categories { get; set; } = "";
    [JsonPropertyName("termSummary")] public string TermSummary { get; set; } = "";

    public override string SearchText => string.Join(" | ", Dataset, FileName, Models, Categories, TermSummary);
}

public abstract class EvidenceMatchBase
{
    [JsonPropertyName("matchesAllRequiredTerms")] public bool MatchesAllRequiredTerms { get; set; }
    [JsonPropertyName("strictFallbackUsed")] public bool StrictFallbackUsed { get; set; }
    [JsonPropertyName("matchScore")] public double MatchScore { get; set; }
    [JsonIgnore] public abstract string SearchText { get; }
}

public sealed class MicroSpeakerVerifiedReviewCaseEvidence : EvidenceMatchBase
{
    [JsonPropertyName("sourceFileId")] public long? SourceFileId { get; set; }
    [JsonPropertyName("sourceDataset")] public string SourceDataset { get; set; } = "";
    [JsonPropertyName("sourceFile")] public string SourceFile { get; set; } = "";
    [JsonPropertyName("sourceModels")] public string SourceModels { get; set; } = "";
    [JsonPropertyName("sourceCategories")] public string SourceCategories { get; set; } = "";
    [JsonPropertyName("originalFileUrl")] public string OriginalFileUrl { get; set; } = "";
    [JsonPropertyName("draftPath")] public string DraftPath { get; set; } = "";
    [JsonPropertyName("verificationPath")] public string VerificationPath { get; set; } = "";
    [JsonPropertyName("manualDraftUsed")] public bool ManualDraftUsed { get; set; }
    [JsonPropertyName("verifiedAt")] public string VerifiedAt { get; set; } = "";
    [JsonPropertyName("reviewCaseId")] public string ReviewCaseId { get; set; } = "";
    [JsonPropertyName("reviewTitle")] public string ReviewTitle { get; set; } = "";
    [JsonPropertyName("reviewPurpose")] public string ReviewPurpose { get; set; } = "";
    [JsonPropertyName("reviewType")] public string ReviewType { get; set; } = "";
    [JsonPropertyName("changedFactors")] public List<MicroSpeakerVerifiedReviewCaseChangedFactor> ChangedFactors { get; set; } = [];
    [JsonPropertyName("outcomes")] public List<MicroSpeakerVerifiedReviewCaseOutcome> Outcomes { get; set; } = [];
    [JsonPropertyName("evidenceRows")] public List<string> EvidenceRows { get; set; } = [];
    [JsonPropertyName("limitations")] public List<string> Limitations { get; set; } = [];
    [JsonPropertyName("sourceDecision")] public string SourceDecision { get; set; } = "";
    [JsonPropertyName("sourceDecisionEvidenceRows")] public List<string> SourceDecisionEvidenceRows { get; set; } = [];
    [JsonPropertyName("verification")] public MicroSpeakerVerifiedReviewCaseVerification Verification { get; set; } = new();

    public override string SearchText => string.Join(" | ",
        SourceDataset,
        SourceFile,
        SourceModels,
        SourceCategories,
        ReviewCaseId,
        ReviewTitle,
        ReviewPurpose,
        ReviewType,
        SourceDecision,
        string.Join(" | ", ChangedFactors.Select(x => x.SearchText)),
        string.Join(" | ", Outcomes.Select(x => x.SearchText)),
        string.Join(" | ", EvidenceRows),
        string.Join(" | ", Limitations),
        Verification.SearchText);
}

public sealed class MicroSpeakerVerifiedReviewCaseVerification
{
    [JsonPropertyName("model")] public string Model { get; set; } = "";
    [JsonPropertyName("aiReviewCaseStatus")] public string AiReviewCaseStatus { get; set; } = "";
    [JsonPropertyName("verificationStatus")] public string VerificationStatus { get; set; } = "";
    [JsonPropertyName("approvedForAskAi")] public bool ApprovedForAskAi { get; set; }
    [JsonPropertyName("confidence")] public string Confidence { get; set; } = "";
    [JsonPropertyName("summary")] public string Summary { get; set; } = "";
    [JsonPropertyName("issues")] public List<string> Issues { get; set; } = [];
    [JsonPropertyName("requiredUserQuestions")] public List<string> RequiredUserQuestions { get; set; } = [];
    [JsonPropertyName("correctionPlan")] public List<string> CorrectionPlan { get; set; } = [];
    [JsonPropertyName("evidencePolicy")] public string EvidencePolicy { get; set; } = "";

    [JsonIgnore]
    public string SearchText => string.Join(" | ",
        AiReviewCaseStatus,
        VerificationStatus,
        Confidence,
        Summary,
        string.Join(" | ", Issues),
        string.Join(" | ", RequiredUserQuestions),
        string.Join(" | ", CorrectionPlan),
        EvidencePolicy);
}

public sealed class MicroSpeakerVerifiedReviewCaseChangedFactor
{
    [JsonPropertyName("changedFactorId")] public string ChangedFactorId { get; set; } = "";
    [JsonPropertyName("changeDomain")] public List<string> ChangeDomain { get; set; } = [];
    [JsonPropertyName("changedFactor")] public string ChangedFactor { get; set; } = "";
    [JsonPropertyName("baselineCondition")] public string BaselineCondition { get; set; } = "";
    [JsonPropertyName("changedCondition")] public string ChangedCondition { get; set; } = "";
    [JsonPropertyName("changedConditions")] public List<string> ChangedConditions { get; set; } = [];
    [JsonPropertyName("subgroupKeys")] public List<string> SubgroupKeys { get; set; } = [];
    [JsonPropertyName("evidenceRows")] public List<string> EvidenceRows { get; set; } = [];

    [JsonIgnore]
    public string SearchText => string.Join(" | ",
        ChangedFactorId,
        string.Join(" | ", ChangeDomain),
        ChangedFactor,
        BaselineCondition,
        ChangedCondition,
        string.Join(" | ", ChangedConditions),
        string.Join(" | ", SubgroupKeys),
        string.Join(" | ", EvidenceRows));
}

public sealed class MicroSpeakerVerifiedReviewCaseOutcome
{
    [JsonPropertyName("outcomeId")] public string OutcomeId { get; set; } = "";
    [JsonPropertyName("changedFactorId")] public string ChangedFactorId { get; set; } = "";
    [JsonPropertyName("outcomeDomain")] public string OutcomeDomain { get; set; } = "";
    [JsonPropertyName("outcomeMetric")] public string OutcomeMetric { get; set; } = "";
    [JsonPropertyName("judgement")] public string Judgement { get; set; } = "";
    [JsonPropertyName("resultSummary")] public string ResultSummary { get; set; } = "";
    [JsonPropertyName("sourceJudgement")] public string SourceJudgement { get; set; } = "";
    [JsonPropertyName("comparisonRows")] public List<string> ComparisonRows { get; set; } = [];
    [JsonPropertyName("evidenceRows")] public List<string> EvidenceRows { get; set; } = [];
    [JsonPropertyName("notes")] public List<string> Notes { get; set; } = [];
    [JsonPropertyName("limitations")] public List<string> Limitations { get; set; } = [];
    [JsonPropertyName("subResults")] public List<MicroSpeakerVerifiedReviewCaseSubResult> SubResults { get; set; } = [];

    [JsonIgnore]
    public string SearchText => string.Join(" | ",
        OutcomeId,
        ChangedFactorId,
        OutcomeDomain,
        OutcomeMetric,
        Judgement,
        ResultSummary,
        SourceJudgement,
        string.Join(" | ", ComparisonRows),
        string.Join(" | ", EvidenceRows),
        string.Join(" | ", Notes),
        string.Join(" | ", Limitations),
        string.Join(" | ", SubResults.Select(x => x.SearchText)));
}

public sealed class MicroSpeakerVerifiedReviewCaseSubResult
{
    [JsonPropertyName("pairId")] public long? PairId { get; set; }
    [JsonPropertyName("statId")] public long? StatId { get; set; }
    [JsonPropertyName("date")] public string Date { get; set; } = "";
    [JsonPropertyName("testRound")] public string TestRound { get; set; } = "";
    [JsonPropertyName("condition")] public string Condition { get; set; } = "";
    [JsonPropertyName("spec")] public string Spec { get; set; } = "";
    [JsonPropertyName("summary")] public string Summary { get; set; } = "";
    [JsonPropertyName("controlCondition")] public string ControlCondition { get; set; } = "";
    [JsonPropertyName("testCondition")] public string TestCondition { get; set; } = "";
    [JsonPropertyName("controlInput")] public double? ControlInput { get; set; }
    [JsonPropertyName("controlNg")] public double? ControlNg { get; set; }
    [JsonPropertyName("controlRatePercent")] public double? ControlRatePercent { get; set; }
    [JsonPropertyName("testInput")] public double? TestInput { get; set; }
    [JsonPropertyName("testNg")] public double? TestNg { get; set; }
    [JsonPropertyName("testRatePercent")] public double? TestRatePercent { get; set; }
    [JsonPropertyName("deltaRatePercentPoint")] public double? DeltaRatePercentPoint { get; set; }
    [JsonPropertyName("effectDirection")] public string EffectDirection { get; set; } = "";
    [JsonPropertyName("minValue")] public double? MinValue { get; set; }
    [JsonPropertyName("maxValue")] public double? MaxValue { get; set; }
    [JsonPropertyName("avgValue")] public double? AvgValue { get; set; }
    [JsonPropertyName("sampleCount")] public long? SampleCount { get; set; }
    [JsonPropertyName("violationCount")] public long? ViolationCount { get; set; }
    [JsonPropertyName("evidenceRows")] public List<string> EvidenceRows { get; set; } = [];

    [JsonIgnore]
    public string SearchText => string.Join(" | ",
        PairId?.ToString(CultureInfo.InvariantCulture) ?? "",
        StatId?.ToString(CultureInfo.InvariantCulture) ?? "",
        Date,
        TestRound,
        Condition,
        Spec,
        Summary,
        ControlCondition,
        TestCondition,
        ControlInput?.ToString(CultureInfo.InvariantCulture) ?? "",
        ControlNg?.ToString(CultureInfo.InvariantCulture) ?? "",
        ControlRatePercent?.ToString(CultureInfo.InvariantCulture) ?? "",
        TestInput?.ToString(CultureInfo.InvariantCulture) ?? "",
        TestNg?.ToString(CultureInfo.InvariantCulture) ?? "",
        TestRatePercent?.ToString(CultureInfo.InvariantCulture) ?? "",
        DeltaRatePercentPoint?.ToString(CultureInfo.InvariantCulture) ?? "",
        EffectDirection,
        MinValue?.ToString(CultureInfo.InvariantCulture) ?? "",
        MaxValue?.ToString(CultureInfo.InvariantCulture) ?? "",
        AvgValue?.ToString(CultureInfo.InvariantCulture) ?? "",
        SampleCount?.ToString(CultureInfo.InvariantCulture) ?? "",
        ViolationCount?.ToString(CultureInfo.InvariantCulture) ?? "",
        string.Join(" | ", EvidenceRows));
}

public sealed class MicroSpeakerPairEvidenceRow : EvidenceMatchBase
{
    [JsonPropertyName("pairId")] public long? PairId { get; set; }
    [JsonPropertyName("fileId")] public long? FileId { get; set; }
    [JsonPropertyName("dataset")] public string Dataset { get; set; } = "";
    [JsonPropertyName("fileName")] public string FileName { get; set; } = "";
    [JsonPropertyName("originalFileUrl")] public string OriginalFileUrl { get; set; } = "";
    [JsonPropertyName("models")] public string Models { get; set; } = "";
    [JsonPropertyName("categories")] public string Categories { get; set; } = "";
    [JsonPropertyName("termSummary")] public string TermSummary { get; set; } = "";
    [JsonPropertyName("tableTitle")] public string TableTitle { get; set; } = "";
    [JsonPropertyName("compareItem")] public string CompareItem { get; set; } = "";
    [JsonPropertyName("controlCondition")] public string ControlCondition { get; set; } = "";
    [JsonPropertyName("testCondition")] public string TestCondition { get; set; } = "";
    [JsonPropertyName("controlInput")] public double? ControlInput { get; set; }
    [JsonPropertyName("controlNg")] public double? ControlNg { get; set; }
    [JsonPropertyName("controlRate")] public double? ControlRate { get; set; }
    [JsonPropertyName("controlRatePercent")] public double? ControlRatePercent { get; set; }
    [JsonPropertyName("testInput")] public double? TestInput { get; set; }
    [JsonPropertyName("testNg")] public double? TestNg { get; set; }
    [JsonPropertyName("testRate")] public double? TestRate { get; set; }
    [JsonPropertyName("testRatePercent")] public double? TestRatePercent { get; set; }
    [JsonPropertyName("deltaRate")] public double? DeltaRate { get; set; }
    [JsonPropertyName("deltaRatePercentPoint")] public double? DeltaRatePercentPoint { get; set; }
    [JsonPropertyName("relativeChangePercent")] public double? RelativeChangePercent { get; set; }
    [JsonPropertyName("improvementRate")] public double? ImprovementRate { get; set; }
    [JsonPropertyName("effectDirection")] public string EffectDirection { get; set; } = "";
    [JsonPropertyName("evidence")] public string Evidence { get; set; } = "";
    [JsonPropertyName("pairConfidence")] public string PairConfidence { get; set; } = "";

    public override string SearchText => string.Join(" | ", Dataset, FileName, Models, Categories, TermSummary,
        TableTitle, CompareItem, ControlCondition, TestCondition, EffectDirection, Evidence, PairConfidence);
}

public sealed class MicroSpeakerMetricEvidenceRow : EvidenceMatchBase
{
    [JsonPropertyName("metricId")] public long? MetricId { get; set; }
    [JsonPropertyName("fileId")] public long? FileId { get; set; }
    [JsonPropertyName("dataset")] public string Dataset { get; set; } = "";
    [JsonPropertyName("fileName")] public string FileName { get; set; } = "";
    [JsonPropertyName("originalFileUrl")] public string OriginalFileUrl { get; set; } = "";
    [JsonPropertyName("models")] public string Models { get; set; } = "";
    [JsonPropertyName("categories")] public string Categories { get; set; } = "";
    [JsonPropertyName("termSummary")] public string TermSummary { get; set; } = "";
    [JsonPropertyName("sheetName")] public string SheetName { get; set; } = "";
    [JsonPropertyName("rowNumber")] public long? RowNumber { get; set; }
    [JsonPropertyName("tableTitle")] public string TableTitle { get; set; } = "";
    [JsonPropertyName("conditionLabel")] public string ConditionLabel { get; set; } = "";
    [JsonPropertyName("inputQty")] public double? InputQty { get; set; }
    [JsonPropertyName("okQty")] public double? OkQty { get; set; }
    [JsonPropertyName("ngQty")] public double? NgQty { get; set; }
    [JsonPropertyName("ngRate")] public double? NgRate { get; set; }
    [JsonPropertyName("ngRatePercent")] public double? NgRatePercent { get; set; }
    [JsonPropertyName("detail")] public string Detail { get; set; } = "";
    [JsonPropertyName("rawRow")] public string RawRow { get; set; } = "";
    [JsonPropertyName("parseConfidence")] public string ParseConfidence { get; set; } = "";

    public override string SearchText => string.Join(" | ", Dataset, FileName, Models, Categories, TermSummary,
        SheetName, TableTitle, ConditionLabel, Detail, RawRow, ParseConfidence);
}

public sealed class MicroSpeakerMeasurementEvidenceRow : EvidenceMatchBase
{
    [JsonPropertyName("statId")] public long? StatId { get; set; }
    [JsonPropertyName("fileId")] public long? FileId { get; set; }
    [JsonPropertyName("dataset")] public string Dataset { get; set; } = "";
    [JsonPropertyName("fileName")] public string FileName { get; set; } = "";
    [JsonPropertyName("originalFileUrl")] public string OriginalFileUrl { get; set; } = "";
    [JsonPropertyName("models")] public string Models { get; set; } = "";
    [JsonPropertyName("categories")] public string Categories { get; set; } = "";
    [JsonPropertyName("termSummary")] public string TermSummary { get; set; } = "";
    [JsonPropertyName("sheetName")] public string SheetName { get; set; } = "";
    [JsonPropertyName("rowNumber")] public long? RowNumber { get; set; }
    [JsonPropertyName("itemLabel")] public string ItemLabel { get; set; } = "";
    [JsonPropertyName("conditionLabel")] public string ConditionLabel { get; set; } = "";
    [JsonPropertyName("spec")] public string Spec { get; set; } = "";
    [JsonPropertyName("minValue")] public double? MinValue { get; set; }
    [JsonPropertyName("maxValue")] public double? MaxValue { get; set; }
    [JsonPropertyName("avgValue")] public double? AvgValue { get; set; }
    [JsonPropertyName("sampleCount")] public long? SampleCount { get; set; }
    [JsonPropertyName("violationCount")] public long? ViolationCount { get; set; }
    [JsonPropertyName("rawRow")] public string RawRow { get; set; } = "";
    [JsonPropertyName("parseConfidence")] public string ParseConfidence { get; set; } = "";

    public override string SearchText => string.Join(" | ", Dataset, FileName, Models, Categories, TermSummary,
        SheetName, ItemLabel, ConditionLabel, Spec, RawRow, ParseConfidence);
}

public sealed class MicroSpeakerPairAggregate
{
    [JsonPropertyName("compareItem")] public string CompareItem { get; set; } = "";
    [JsonPropertyName("rowCount")] public int RowCount { get; set; }
    [JsonPropertyName("worsenedCount")] public int WorsenedCount { get; set; }
    [JsonPropertyName("improvedCount")] public int ImprovedCount { get; set; }
    [JsonPropertyName("noChangeCount")] public int NoChangeCount { get; set; }
    [JsonPropertyName("maxAbsRelativeChangePercent")] public double MaxAbsRelativeChangePercent { get; set; }
    [JsonPropertyName("maxAbsDeltaRatePercentPoint")] public double MaxAbsDeltaRatePercentPoint { get; set; }
    [JsonPropertyName("exampleDatasets")] public List<string> ExampleDatasets { get; set; } = [];
}

public sealed class MicroSpeakerPairConditionAggregate
{
    [JsonPropertyName("fileId")] public long? FileId { get; set; }
    [JsonPropertyName("dataset")] public string Dataset { get; set; } = "";
    [JsonPropertyName("fileName")] public string FileName { get; set; } = "";
    [JsonPropertyName("originalFileUrl")] public string OriginalFileUrl { get; set; } = "";
    [JsonPropertyName("tableTitle")] public string TableTitle { get; set; } = "";
    [JsonPropertyName("compareItem")] public string CompareItem { get; set; } = "";
    [JsonPropertyName("controlCondition")] public string ControlCondition { get; set; } = "";
    [JsonPropertyName("testCondition")] public string TestCondition { get; set; } = "";
    [JsonPropertyName("controlInput")] public double? ControlInput { get; set; }
    [JsonPropertyName("controlNg")] public double? ControlNg { get; set; }
    [JsonPropertyName("controlRate")] public double? ControlRate { get; set; }
    [JsonPropertyName("controlRatePercent")] public double? ControlRatePercent { get; set; }
    [JsonPropertyName("testInput")] public double? TestInput { get; set; }
    [JsonPropertyName("testNg")] public double? TestNg { get; set; }
    [JsonPropertyName("testRate")] public double? TestRate { get; set; }
    [JsonPropertyName("testRatePercent")] public double? TestRatePercent { get; set; }
    [JsonPropertyName("deltaRate")] public double? DeltaRate { get; set; }
    [JsonPropertyName("deltaRatePercentPoint")] public double? DeltaRatePercentPoint { get; set; }
    [JsonPropertyName("relativeChangePercent")] public double? RelativeChangePercent { get; set; }
    [JsonPropertyName("effectDirection")] public string EffectDirection { get; set; } = "";
    [JsonPropertyName("sourceRowCount")] public int SourceRowCount { get; set; }
    [JsonPropertyName("usedRowCount")] public int UsedRowCount { get; set; }
    [JsonPropertyName("aggregationMethod")] public string AggregationMethod { get; set; } = "";
    [JsonPropertyName("evidence")] public string Evidence { get; set; } = "";
    [JsonPropertyName("limit")] public string Limit { get; set; } = "";
}

public sealed class JinoDocumentEvidenceRow : EvidenceMatchBase
{
    [JsonPropertyName("documentId")] public string DocumentId { get; set; } = "";
    [JsonPropertyName("sourceDataset")] public string SourceDataset { get; set; } = "";
    [JsonPropertyName("sourceFile")] public string SourceFile { get; set; } = "";
    [JsonPropertyName("productType")] public string ProductType { get; set; } = "";
    [JsonPropertyName("title")] public string Title { get; set; } = "";
    [JsonPropertyName("model")] public string Model { get; set; } = "";
    [JsonPropertyName("reportType")] public string ReportType { get; set; } = "";
    [JsonPropertyName("primaryDefect")] public string PrimaryDefect { get; set; } = "";
    [JsonPropertyName("purpose")] public string Purpose { get; set; } = "";
    [JsonPropertyName("generatedReportPreview")] public string GeneratedReportPreview { get; set; } = "";
    [JsonPropertyName("relatedDefectsJson")] public string RelatedDefectsJson { get; set; } = "";
    [JsonPropertyName("partsJson")] public string PartsJson { get; set; } = "";
    [JsonPropertyName("processesJson")] public string ProcessesJson { get; set; } = "";
    [JsonPropertyName("resultCount")] public long? ResultCount { get; set; }
    [JsonPropertyName("conditionCount")] public long? ConditionCount { get; set; }
    [JsonPropertyName("updatedAt")] public string UpdatedAt { get; set; } = "";

    public override string SearchText => string.Join(" | ", SourceDataset, SourceFile, ProductType, Title, Model,
        ReportType, PrimaryDefect, Purpose, GeneratedReportPreview, RelatedDefectsJson, PartsJson, ProcessesJson);
}

public sealed class JinoResultEvidenceRow : EvidenceMatchBase
{
    [JsonPropertyName("resultId")] public string ResultId { get; set; } = "";
    [JsonPropertyName("documentId")] public string DocumentId { get; set; } = "";
    [JsonPropertyName("sourceDataset")] public string SourceDataset { get; set; } = "";
    [JsonPropertyName("sourceFile")] public string SourceFile { get; set; } = "";
    [JsonPropertyName("productType")] public string ProductType { get; set; } = "";
    [JsonPropertyName("title")] public string Title { get; set; } = "";
    [JsonPropertyName("reportType")] public string ReportType { get; set; } = "";
    [JsonPropertyName("primaryDefect")] public string PrimaryDefect { get; set; } = "";
    [JsonPropertyName("measurementType")] public string MeasurementType { get; set; } = "";
    [JsonPropertyName("conditionGroup")] public string ConditionGroup { get; set; } = "";
    [JsonPropertyName("process")] public string Process { get; set; } = "";
    [JsonPropertyName("changedFactor")] public string ChangedFactor { get; set; } = "";
    [JsonPropertyName("beforeValue")] public string BeforeValue { get; set; } = "";
    [JsonPropertyName("afterValue")] public string AfterValue { get; set; } = "";
    [JsonPropertyName("inputCount")] public double? InputCount { get; set; }
    [JsonPropertyName("okCount")] public double? OkCount { get; set; }
    [JsonPropertyName("ngCount")] public double? NgCount { get; set; }
    [JsonPropertyName("ngRateDecimal")] public double? NgRateDecimal { get; set; }
    [JsonPropertyName("ngRatePercent")] public double? NgRatePercent { get; set; }
    [JsonPropertyName("metricName")] public string MetricName { get; set; } = "";
    [JsonPropertyName("metricValue")] public double? MetricValue { get; set; }
    [JsonPropertyName("unit")] public string Unit { get; set; } = "";
    [JsonPropertyName("judgement")] public string Judgement { get; set; } = "";
    [JsonPropertyName("resultSourceFile")] public string ResultSourceFile { get; set; } = "";
    [JsonPropertyName("sheetName")] public string SheetName { get; set; } = "";
    [JsonPropertyName("sourceCellsJson")] public string SourceCellsJson { get; set; } = "";

    public override string SearchText => string.Join(" | ", SourceDataset, SourceFile, ProductType, Title, ReportType,
        PrimaryDefect, MeasurementType, ConditionGroup, Process, ChangedFactor, BeforeValue, AfterValue,
        MetricName, Judgement, ResultSourceFile, SheetName, SourceCellsJson);
}
