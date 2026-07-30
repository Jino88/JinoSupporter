using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

public sealed class MicroSpeakerReviewCaseService(
    MicroSpeakerInputDataService microSpeaker,
    IWebHostEnvironment env)
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public MicroSpeakerReviewCaseSample BuildSample(int limit = 80, string? query = null)
    {
        int safeLimit = Math.Clamp(limit, 1, 500);
        string trimmedQuery = (query ?? "").Trim();
        MicroSpeakerPaths paths = microSpeaker.ResolvePaths();

        var sample = new MicroSpeakerReviewCaseSample
        {
            CreatedAt = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            MicroSpeakerDatabasePath = paths.DatabasePath,
            Query = trimmedQuery,
            Limit = safeLimit,
        };

        if (!File.Exists(paths.DatabasePath))
        {
            sample.Notes.Add("MicroSpeaker SQLite database was not found.");
            return sample;
        }

        sample.DatabaseExists = true;

        using SqliteConnection conn = OpenReadOnly(paths.DatabasePath);
        List<ReviewPairRow> pairRows = ReadPairRows(conn, trimmedQuery);
        List<ReviewMeasurementRow> measurementRows = ReadMeasurementRows(conn, trimmedQuery);

        List<MicroSpeakerReviewCase> cases = [];
        cases.AddRange(BuildPairCases(pairRows));
        cases.AddRange(BuildMeasurementCases(measurementRows));
        List<MicroSpeakerReviewCaseGroup> reviewCases = BuildHierarchicalCases(cases);

        sample.SourcePairRows = pairRows.Count;
        sample.SourceMeasurementRows = measurementRows.Count;
        sample.PairCaseCount = cases.Count(c => string.Equals(c.ExtractionSource, "comparison_pairs", StringComparison.OrdinalIgnoreCase));
        sample.MeasurementCaseCount = cases.Count(c => string.Equals(c.ExtractionSource, "measurement_stats", StringComparison.OrdinalIgnoreCase));
        sample.FlatCandidates = cases
            .OrderByDescending(c => c.ConfidenceScore)
            .ThenByDescending(c => Math.Abs(c.RelativeChangePercent ?? c.DeltaRatePercentPoint ?? c.MeasurementRelativeChangePercent ?? 0))
            .ThenBy(c => c.SourceFile, StringComparer.OrdinalIgnoreCase)
            .Take(safeLimit)
            .ToList();
        sample.ReviewCases = reviewCases
            .OrderByDescending(c => c.ConfidenceScore)
            .ThenByDescending(c => c.Outcomes.Count)
            .ThenBy(c => c.SourceFile, StringComparer.OrdinalIgnoreCase)
            .Take(safeLimit)
            .ToList();
        sample.FlatCandidateCount = cases.Count;
        sample.ReviewCaseCount = reviewCases.Count;
        sample.OutcomeCount = reviewCases.Sum(c => c.Outcomes.Count);

        sample.Notes.Add("Diagnostic extraction only; no MicroSpeaker DB tables were created or modified.");
        sample.Notes.Add("Changed factors in this sample are heuristic candidates only, not final ReviewCase classification.");
        sample.Notes.Add("AI must create and verify final changedFactors, outcomes, and evidenceRows from extracted source rows/cells before saving or answering.");
        sample.Notes.Add("reviewCases and flatCandidates remain diagnostic inputs for AI analysis and rule tuning; downstream Ask AI must verify cited evidence rows before using them.");
        return sample;
    }

    public string WriteSampleJson(int limit = 80, string? query = null)
    {
        MicroSpeakerReviewCaseSample sample = BuildSample(limit, query);
        string dir = Path.Combine(env.ContentRootPath, "tmp");
        Directory.CreateDirectory(dir);
        string path = Path.Combine(dir, "review_cases_sample.json");
        File.WriteAllText(path, JsonSerializer.Serialize(sample, JsonOptions), Encoding.UTF8);
        return path;
    }

    public MicroSpeakerReviewCaseAiPacket BuildAiPacket(long fileId, int rowLimit = 1200, int candidateLimit = 300)
    {
        int safeRowLimit = Math.Clamp(rowLimit, 50, 5000);
        int safeCandidateLimit = Math.Clamp(candidateLimit, 10, 1000);
        MicroSpeakerPaths paths = microSpeaker.ResolvePaths();
        string repoRoot = AiPromptRegistry.FindRepositoryRoot();

        var packet = new MicroSpeakerReviewCaseAiPacket
        {
            CreatedAt = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            MicroSpeakerDatabasePath = paths.DatabasePath,
            SourceFileId = fileId,
            RowLimit = safeRowLimit,
            CandidateLimit = safeCandidateLimit,
            ReviewCasePromptPath = AiPromptRegistry.ResolvePath("data-inference/reviewcase-ai-analysis.md"),
            CalibrationReferencePath = Path.Combine(repoRoot, "ASK_AI_REVIEWCASE_NEXT_STEPS.md"),
            AuditDecisionPath = Path.Combine(repoRoot, "REVIEWCASE_AI_AUDIT_DECISIONS.md"),
        };

        packet.Notes.Add("This packet is AI input for ReviewCase generation; it is not a final ReviewCase.");
        packet.Notes.Add("AI must create changedFactors, outcomes, evidenceRows, and verification from cited source rows/cells.");
        packet.Notes.Add("Candidate pairs/metrics/measurements are hints only. Use sheetRows and cells as the source authority.");

        if (!File.Exists(paths.DatabasePath))
        {
            packet.Notes.Add("MicroSpeaker SQLite database was not found.");
            return packet;
        }

        packet.DatabaseExists = true;
        using SqliteConnection conn = OpenReadOnly(paths.DatabasePath);

        packet.File = ReadAiFile(conn, fileId);
        if (packet.File is null)
        {
            packet.Notes.Add("MicroSpeaker file_id was not found.");
            return packet;
        }

        packet.FileFound = true;
        packet.AuditDecision = ReadAuditDecision(fileId, packet.AuditDecisionPath);
        if (packet.AuditDecision is not null)
            packet.Notes.Add($"User audit decision is '{packet.AuditDecision.Decision}'. ReviewCase AI must honor it unless later evidence supersedes it.");

        packet.Sheets = ReadAiSheets(conn, fileId);
        packet.SheetRows = ReadAiSheetRows(conn, fileId, safeRowLimit);
        AttachAiCells(conn, fileId, packet.SheetRows, Math.Clamp(safeRowLimit * 64, 1000, 100000));
        packet.ContextRows = BuildAiContextRows(packet.SheetRows);
        packet.PairCandidates = ReadAiPairCandidates(conn, fileId, safeCandidateLimit);
        packet.MetricCandidates = ReadAiMetricCandidates(conn, fileId, safeCandidateLimit);
        packet.MeasurementCandidates = ReadAiMeasurementCandidates(conn, fileId, safeCandidateLimit);
        packet.TermHints = ReadAiTermHints(conn, fileId, 120);

        if (packet.File.SheetRowCount == 0 || packet.File.SheetCellCount == 0)
            packet.Notes.Add("No extracted sheet row/cell evidence is available. AI should exclude this file unless OCR/re-extraction is provided.");
        else if (packet.SheetRows.Count >= safeRowLimit)
            packet.Notes.Add("Sheet rows were truncated by rowLimit. Increase rowLimit before final ReviewCase generation if needed.");

        return packet;
    }

    private static List<MicroSpeakerReviewCase> BuildPairCases(IReadOnlyList<ReviewPairRow> rows)
    {
        var cases = new List<MicroSpeakerReviewCase>();
        var groups = rows
            .Where(r => !string.IsNullOrWhiteSpace(r.TableTitle)
                        && !string.IsNullOrWhiteSpace(r.ControlCondition)
                        && !string.IsNullOrWhiteSpace(r.TestCondition))
            .GroupBy(r => new
            {
                r.FileId,
                Table = NormalizeAggregateKey(r.TableTitle),
                Compare = NormalizeAggregateKey(r.CompareItem),
                Control = NormalizeAggregateKey(r.ControlCondition),
                Test = NormalizeAggregateKey(r.TestCondition),
            });

        foreach (var group in groups)
        {
            List<ReviewPairRow> groupRows = group.OrderBy(r => r.PairId).ToList();
            ReviewPairRow first = groupRows[0];
            ReviewPairRow? totalRow = FindTotalEquivalentRow(groupRows);
            IReadOnlyList<ReviewPairRow> usedRows = totalRow is null ? groupRows : [totalRow];

            double? normalInput = totalRow?.ControlInput ?? first.ControlInput;
            double? normalNg = totalRow?.ControlNg ?? first.ControlNg;
            double? normalRate = totalRow?.ControlRate ?? first.ControlRate ?? RateFromCounts(normalInput, normalNg);
            double? testInput = totalRow is null && groupRows.Count == 1 ? first.TestInput : SumNullable(usedRows.Select(r => r.TestInput));
            double? testNg = totalRow is null && groupRows.Count == 1 ? first.TestNg : SumNullable(usedRows.Select(r => r.TestNg));
            double? testRate = totalRow?.TestRate ?? RateFromCounts(testInput, testNg) ?? (usedRows.Count == 1 ? usedRows[0].TestRate : null);
            double? deltaRate = testRate.HasValue && normalRate.HasValue ? testRate.Value - normalRate.Value : null;
            string aggregationMethod =
                groupRows.Count == 1 ? "single_row" :
                totalRow is not null ? "total_row" :
                "summed_rows";
            string afterCondition = StripTotalPrefix(totalRow?.TestCondition ?? first.TestCondition);
            string evidence = string.Join("; ", usedRows.Select(r => r.Evidence).Where(x => !string.IsNullOrWhiteSpace(x)).Take(8));
            string sourceText = string.Join(" | ", first.FileName, first.TableTitle, first.CompareItem, first.ControlCondition, afterCondition, evidence);
            string confidence = InferConfidence(groupRows.Select(r => r.PairConfidence), aggregationMethod);

            cases.Add(new MicroSpeakerReviewCase
            {
                CaseId = StableId("ms-pair", first.FileId.ToString(CultureInfo.InvariantCulture), first.TableTitle, first.CompareItem, first.ControlCondition, afterCondition),
                SourceSystem = "MicroSpeaker",
                ExtractionSource = "comparison_pairs",
                FileId = first.FileId,
                Dataset = first.Dataset,
                SourceFile = first.FileName,
                OriginalFileUrl = SourceFileUrl(first.FileId),
                SheetName = ExtractSheetName(evidence),
                SourceRows = usedRows.Select(r => r.Evidence).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                ReviewTitle = Brief(first.TableTitle, 220),
                ReviewPurpose = InferReviewPurpose(first.FileName, first.TableTitle),
                ChangeDomain = InferChangeDomain(sourceText),
                ChangedFactor = InferChangedFactor(first.FileName, first.TableTitle, first.CompareItem, first.ControlCondition, afterCondition),
                BeforeCondition = first.ControlCondition,
                AfterCondition = afterCondition,
                ReviewedProcess = InferReviewedProcess(first.FileName, first.TableTitle),
                OutcomeDomain = InferOutcomeDomain(first.TableTitle + " | " + first.CompareItem),
                OutcomeMetric = InferOutcomeMetric(first.TableTitle, first.CompareItem, isMeasurement: false),
                NormalInput = normalInput,
                NormalNg = normalNg,
                NormalRate = normalRate,
                NormalRatePercent = RatePercent(normalRate),
                TestInput = testInput,
                TestNg = testNg,
                TestRate = testRate,
                TestRatePercent = RatePercent(testRate),
                DeltaRate = deltaRate,
                DeltaRatePercentPoint = RatePercent(deltaRate),
                RelativeChangePercent = RelativeChangePercent(normalRate, testRate),
                Judgement = JudgementFromEffect(AggregateEffectDirection(deltaRate, groupRows)),
                AggregationMethod = aggregationMethod,
                Confidence = confidence,
                ConfidenceScore = ConfidenceScore(confidence, aggregationMethod),
                LimitReason = PairLimitReason(aggregationMethod),
                Evidence = evidence,
            });
        }

        return cases;
    }

    private static List<MicroSpeakerReviewCaseGroup> BuildHierarchicalCases(IReadOnlyList<MicroSpeakerReviewCase> flatCases)
    {
        var groups = flatCases
            .GroupBy(c => new
            {
                c.FileId,
                ReviewKey = ReviewCaseGroupKey(c),
            });

        var result = new List<MicroSpeakerReviewCaseGroup>();
        foreach (var group in groups)
        {
            List<MicroSpeakerReviewCase> rows = group
                .OrderBy(c => OutcomeOrder(c.OutcomeDomain))
                .ThenBy(c => c.OutcomeMetric, StringComparer.OrdinalIgnoreCase)
                .ThenBy(c => c.SheetName, StringComparer.OrdinalIgnoreCase)
                .ToList();
            MicroSpeakerReviewCase first = rows[0];
            string reviewTitle = BestReviewTitle(rows);
            List<MicroSpeakerReviewChangedFactor> changedFactors = BuildChangedFactors(rows);
            List<MicroSpeakerReviewOutcome> outcomes = rows.Select(r => ToOutcome(r, ChangeSignatureKey(r))).ToList();
            List<string> evidenceRows = rows
                .SelectMany(r => r.SourceRows)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(80)
                .ToList();

            result.Add(new MicroSpeakerReviewCaseGroup
            {
                ReviewCaseId = StableId("ms-review", first.FileId.ToString(CultureInfo.InvariantCulture), group.Key.ReviewKey),
                SourceSystem = "MicroSpeaker",
                FileId = first.FileId,
                Dataset = first.Dataset,
                SourceFile = first.SourceFile,
                OriginalFileUrl = first.OriginalFileUrl,
                ReviewTitle = reviewTitle,
                ReviewPurpose = BestReviewPurpose(rows),
                ChangedFactors = changedFactors,
                Outcomes = outcomes,
                EvidenceRows = evidenceRows,
                Confidence = MergeConfidence(rows.Select(r => r.Confidence)),
                ConfidenceScore = rows.Select(r => r.ConfidenceScore).DefaultIfEmpty(0).Max(),
                LimitReason = string.Join("; ", rows.Select(r => r.LimitReason).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Take(4)),
            });
        }

        return result;
    }

    private static List<MicroSpeakerReviewChangedFactor> BuildChangedFactors(IReadOnlyList<MicroSpeakerReviewCase> rows)
    {
        return rows
            .GroupBy(ChangeSignatureKey)
            .Select(g =>
            {
                MicroSpeakerReviewCase first = g.First();
                string changeKey = ChangeSignatureKey(first);
                string changedFactor = BestChangedFactorLabel(g);
                return new MicroSpeakerReviewChangedFactor
                {
                    ChangeKey = changeKey,
                    ChangeDomain = first.ChangeDomain,
                    ChangedFactor = changedFactor,
                    BeforeCondition = first.BeforeCondition,
                    AfterCondition = first.AfterCondition,
                    ReviewedProcess = first.ReviewedProcess,
                    Evidence = string.Join("; ", g.Select(r => r.Evidence).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Take(4)),
                    Outcomes = g.Select(r => ToOutcome(r, changeKey)).ToList(),
                };
            })
            .ToList();
    }

    private static MicroSpeakerReviewOutcome ToOutcome(MicroSpeakerReviewCase row, string changedFactorKey)
        => new()
        {
            OutcomeId = row.CaseId,
            ChangedFactorKey = changedFactorKey,
            ExtractionSource = row.ExtractionSource,
            OutcomeDomain = row.OutcomeDomain,
            OutcomeMetric = row.OutcomeMetric,
            SheetName = row.SheetName,
            SourceRows = row.SourceRows,
            NormalInput = row.NormalInput,
            NormalNg = row.NormalNg,
            NormalRate = row.NormalRate,
            NormalRatePercent = row.NormalRatePercent,
            TestInput = row.TestInput,
            TestNg = row.TestNg,
            TestRate = row.TestRate,
            TestRatePercent = row.TestRatePercent,
            DeltaRate = row.DeltaRate,
            DeltaRatePercentPoint = row.DeltaRatePercentPoint,
            RelativeChangePercent = row.RelativeChangePercent,
            NormalMeasurement = row.NormalMeasurement,
            TestMeasurement = row.TestMeasurement,
            MeasurementDeltaAvg = row.MeasurementDeltaAvg,
            MeasurementRelativeChangePercent = row.MeasurementRelativeChangePercent,
            Judgement = row.Judgement,
            AggregationMethod = row.AggregationMethod,
            Confidence = row.Confidence,
            ConfidenceScore = row.ConfidenceScore,
            LimitReason = row.LimitReason,
            Evidence = row.Evidence,
        };

    private static List<MicroSpeakerReviewCase> BuildMeasurementCases(IReadOnlyList<ReviewMeasurementRow> rows)
    {
        var cases = new List<MicroSpeakerReviewCase>();
        var groups = rows
            .Where(r => IsMeasurementOutcomeCandidate(r.ItemLabel + " | " + r.ConditionLabel + " | " + r.RawRow))
            .GroupBy(r => new { r.FileId, Item = NormalizeAggregateKey(r.ItemLabel) });

        foreach (var group in groups)
        {
            List<ReviewMeasurementRow> ordered = group.OrderBy(r => r.RowNumber).ToList();
            List<ReviewMeasurementRow> normals = ordered.Where(r => IsNormalCondition(r.ConditionLabel)).ToList();
            List<ReviewMeasurementRow> tests = ordered.Where(r => IsTestCondition(r.ConditionLabel)).ToList();
            if (normals.Count == 0 || tests.Count == 0) continue;

            foreach (ReviewMeasurementRow test in tests)
            {
                ReviewMeasurementRow? normal = FindBestNormalMeasurement(test, normals);
                if (normal is null) continue;

                double? delta = test.AvgValue.HasValue && normal.AvgValue.HasValue
                    ? test.AvgValue.Value - normal.AvgValue.Value
                    : null;
                string sourceText = string.Join(" | ", test.FileName, test.ItemLabel, normal.ConditionLabel, test.ConditionLabel, test.RawRow);
                string confidence = InferConfidence([normal.ParseConfidence, test.ParseConfidence], "measurement_pair");

                cases.Add(new MicroSpeakerReviewCase
                {
                    CaseId = StableId("ms-measure", test.FileId.ToString(CultureInfo.InvariantCulture), test.ItemLabel, normal.ConditionLabel, test.ConditionLabel),
                    SourceSystem = "MicroSpeaker",
                    ExtractionSource = "measurement_stats",
                    FileId = test.FileId,
                    Dataset = test.Dataset,
                    SourceFile = test.FileName,
                    OriginalFileUrl = SourceFileUrl(test.FileId),
                    SheetName = test.SheetName,
                    SourceRows = [normal.RowEvidence, test.RowEvidence],
                    ReviewTitle = Brief(test.ItemLabel, 220),
                    ReviewPurpose = InferReviewPurpose(test.FileName, test.ItemLabel),
                    ChangeDomain = InferChangeDomain(sourceText),
                    ChangedFactor = InferChangedFactor(test.FileName, test.ItemLabel, test.ItemLabel, normal.ConditionLabel, test.ConditionLabel),
                    BeforeCondition = normal.ConditionLabel,
                    AfterCondition = test.ConditionLabel,
                    ReviewedProcess = InferReviewedProcess(test.FileName, test.ItemLabel),
                    OutcomeDomain = InferOutcomeDomain(test.ItemLabel + " | " + test.ConditionLabel),
                    OutcomeMetric = InferOutcomeMetric(test.ItemLabel, test.ItemLabel, isMeasurement: true),
                    NormalMeasurement = ToMeasurement(normal),
                    TestMeasurement = ToMeasurement(test),
                    MeasurementDeltaAvg = delta,
                    MeasurementRelativeChangePercent = RelativeChangePercent(normal.AvgValue, test.AvgValue),
                    Judgement = MeasurementJudgement(delta),
                    AggregationMethod = "measurement_pair",
                    Confidence = confidence,
                    ConfidenceScore = ConfidenceScore(confidence, "measurement_pair"),
                    LimitReason = "Continuous measurement direction is not judged without a spec target; compare n/min/max/avg as evidence.",
                    Evidence = string.Join("; ", new[] { normal.RawRow, test.RawRow }.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => Brief(x, 180))),
                });
            }
        }

        return cases;
    }

    private static List<ReviewPairRow> ReadPairRows(SqliteConnection conn, string query)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        string where = BuildQueryWhere(cmd, query,
        [
            "f.dataset", "f.path", "f.file_name", "f.sheet_names", "f.models", "f.categories", "f.term_summary",
            "p.table_title", "p.compare_item", "p.control_condition", "p.test_condition", "p.effect_direction", "p.evidence", "p.pair_confidence"
        ]);

        cmd.CommandText = $"""
            SELECT p.pair_id, p.file_id, f.dataset, f.file_name, f.models, f.categories, f.dates_found, f.term_summary,
                   p.table_title, p.compare_item, p.control_condition, p.test_condition,
                   p.control_input, p.control_ng, p.control_rate,
                   p.test_input, p.test_ng, p.test_rate,
                   p.delta_rate, p.improvement_rate, p.effect_direction, p.evidence, p.pair_confidence
            FROM comparison_pairs p
            JOIN files f ON f.file_id = p.file_id
            {where}
            ORDER BY p.file_id, p.pair_id;
            """;

        var rows = new List<ReviewPairRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            rows.Add(new ReviewPairRow(
                L(r, 0), L(r, 1), S(r, 2), S(r, 3), S(r, 4), S(r, 5), S(r, 6), S(r, 7),
                S(r, 8), S(r, 9), S(r, 10), S(r, 11),
                D(r, 12), D(r, 13), D(r, 14), D(r, 15), D(r, 16), D(r, 17),
                D(r, 18), D(r, 19), S(r, 20), S(r, 21), S(r, 22)));
        }

        return rows;
    }

    private static List<ReviewMeasurementRow> ReadMeasurementRows(SqliteConnection conn, string query)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        string where = BuildQueryWhere(cmd, query,
        [
            "f.dataset", "f.path", "f.file_name", "f.sheet_names", "f.models", "f.categories", "f.term_summary",
            "s.sheet_name", "s.item_label", "s.condition_label", "s.spec", "s.raw_row", "s.parse_confidence"
        ]);

        cmd.CommandText = $"""
            SELECT s.stat_id, s.file_id, f.dataset, f.file_name, f.models, f.categories, f.dates_found, f.term_summary,
                   s.sheet_name, s.row_number, s.item_label, s.condition_label, s.spec,
                   s.min_value, s.max_value, s.avg_value, s.sample_count, s.violation_count,
                   s.raw_row, s.parse_confidence
            FROM measurement_stats s
            JOIN files f ON f.file_id = s.file_id
            {where}
            ORDER BY s.file_id, s.stat_id;
            """;

        var rows = new List<ReviewMeasurementRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            rows.Add(new ReviewMeasurementRow(
                L(r, 0), L(r, 1), S(r, 2), S(r, 3), S(r, 4), S(r, 5), S(r, 6), S(r, 7),
                S(r, 8), L(r, 9), S(r, 10), S(r, 11), S(r, 12),
                D(r, 13), D(r, 14), D(r, 15), L(r, 16), L(r, 17), S(r, 18), S(r, 19)));
        }

        return rows;
    }

    private static MicroSpeakerReviewCaseAiFile? ReadAiFile(SqliteConnection conn, long fileId)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT f.file_id, f.dataset, f.path, f.file_name, f.extension, f.size_bytes, f.status, f.error,
                   f.sheet_count, f.sheet_names, f.max_rows_any_sheet, f.max_cols_any_sheet, f.non_empty_cells,
                   f.models, f.categories, f.dates_found, f.structure_family, f.structure_confidence, f.term_summary,
                   f.metric_candidate_count, f.measurement_stat_count, f.comparison_pair_count,
                   (SELECT COUNT(*) FROM sheet_rows sr WHERE sr.file_id=f.file_id) AS sheet_row_count,
                   (SELECT COUNT(*) FROM sheet_cells sc WHERE sc.file_id=f.file_id) AS sheet_cell_count
            FROM files f
            WHERE f.file_id=@id
            LIMIT 1;
            """;
        cmd.Parameters.AddWithValue("@id", fileId);

        using SqliteDataReader r = cmd.ExecuteReader();
        if (!r.Read()) return null;

        return new MicroSpeakerReviewCaseAiFile
        {
            FileId = L(r, 0),
            Dataset = S(r, 1),
            Path = S(r, 2),
            FileName = S(r, 3),
            Extension = S(r, 4),
            SizeBytes = L(r, 5),
            Status = S(r, 6),
            Error = S(r, 7),
            SheetCount = L(r, 8),
            SheetNames = S(r, 9),
            MaxRowsAnySheet = L(r, 10),
            MaxColsAnySheet = L(r, 11),
            NonEmptyCells = L(r, 12),
            Models = S(r, 13),
            Categories = S(r, 14),
            DatesFound = S(r, 15),
            StructureFamily = S(r, 16),
            StructureConfidence = S(r, 17),
            TermSummary = S(r, 18),
            MetricCandidateCount = L(r, 19),
            MeasurementStatCount = L(r, 20),
            ComparisonPairCount = L(r, 21),
            SheetRowCount = L(r, 22),
            SheetCellCount = L(r, 23),
            OriginalFileUrl = SourceFileUrl(fileId),
        };
    }

    private static List<MicroSpeakerReviewCaseAiSheet> ReadAiSheets(SqliteConnection conn, long fileId)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT sheet_name, row_count, col_count, non_empty_count, sample_text
            FROM sheets
            WHERE file_id=@id
            ORDER BY sheet_name;
            """;
        cmd.Parameters.AddWithValue("@id", fileId);

        var rows = new List<MicroSpeakerReviewCaseAiSheet>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            rows.Add(new MicroSpeakerReviewCaseAiSheet
            {
                SheetName = S(r, 0),
                RowCount = L(r, 1),
                ColCount = L(r, 2),
                NonEmptyCount = L(r, 3),
                SampleText = S(r, 4),
            });
        }

        return rows;
    }

    private static List<MicroSpeakerReviewCaseAiSheetRow> ReadAiSheetRows(SqliteConnection conn, long fileId, int limit)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT sheet_name, row_number, non_empty_count, row_text, cells_json
            FROM sheet_rows
            WHERE file_id=@id
            ORDER BY sheet_name, row_number
            LIMIT @limit;
            """;
        cmd.Parameters.AddWithValue("@id", fileId);
        cmd.Parameters.AddWithValue("@limit", limit);

        var rows = new List<MicroSpeakerReviewCaseAiSheetRow>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string sheet = S(r, 0);
            long rowNumber = L(r, 1);
            rows.Add(new MicroSpeakerReviewCaseAiSheetRow
            {
                RowId = RowId(sheet, rowNumber),
                SheetName = sheet,
                RowNumber = rowNumber,
                NonEmptyCount = L(r, 2),
                RowText = S(r, 3),
                CellsJson = S(r, 4),
            });
        }

        return rows;
    }

    private static void AttachAiCells(
        SqliteConnection conn,
        long fileId,
        IReadOnlyList<MicroSpeakerReviewCaseAiSheetRow> rows,
        int maxCells)
    {
        Dictionary<string, MicroSpeakerReviewCaseAiSheetRow> rowMap = rows
            .ToDictionary(r => RowId(r.SheetName, r.RowNumber), StringComparer.OrdinalIgnoreCase);
        if (rowMap.Count == 0) return;

        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT sheet_name, row_number, col_number, col_label, cell_value
            FROM sheet_cells
            WHERE file_id=@id
            ORDER BY sheet_name, row_number, col_number
            LIMIT @limit;
            """;
        cmd.Parameters.AddWithValue("@id", fileId);
        cmd.Parameters.AddWithValue("@limit", maxCells);

        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string sheet = S(r, 0);
            long rowNumber = L(r, 1);
            if (!rowMap.TryGetValue(RowId(sheet, rowNumber), out MicroSpeakerReviewCaseAiSheetRow? row))
                continue;

            row.Cells.Add(new MicroSpeakerReviewCaseAiCell
            {
                CellId = $"{sheet}!{ColumnLabel(L(r, 2))}{rowNumber.ToString(CultureInfo.InvariantCulture)}",
                ColNumber = L(r, 2),
                ColLabel = S(r, 3),
                Value = S(r, 4),
            });
        }
    }

    private static List<MicroSpeakerReviewCaseAiRowRef> BuildAiContextRows(
        IReadOnlyList<MicroSpeakerReviewCaseAiSheetRow> rows)
    {
        return rows
            .Where(r => IsAiContextRow(r.RowText))
            .Take(120)
            .Select(r => new MicroSpeakerReviewCaseAiRowRef
            {
                RowId = r.RowId,
                SheetName = r.SheetName,
                RowNumber = r.RowNumber,
                RowText = Brief(r.RowText, 420),
            })
            .ToList();
    }

    private static List<MicroSpeakerReviewCaseAiPairCandidate> ReadAiPairCandidates(
        SqliteConnection conn,
        long fileId,
        int limit)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT pair_id, table_title, compare_item, control_condition, test_condition,
                   control_input, control_ng, control_rate, test_input, test_ng, test_rate,
                   delta_rate, improvement_rate, effect_direction, evidence, pair_confidence
            FROM comparison_pairs
            WHERE file_id=@id
            ORDER BY pair_id
            LIMIT @limit;
            """;
        cmd.Parameters.AddWithValue("@id", fileId);
        cmd.Parameters.AddWithValue("@limit", limit);

        var rows = new List<MicroSpeakerReviewCaseAiPairCandidate>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string evidence = S(r, 14);
            rows.Add(new MicroSpeakerReviewCaseAiPairCandidate
            {
                PairId = L(r, 0),
                TableTitle = S(r, 1),
                CompareItem = S(r, 2),
                ControlCondition = S(r, 3),
                TestCondition = S(r, 4),
                ControlInput = D(r, 5),
                ControlNg = D(r, 6),
                ControlRate = D(r, 7),
                ControlRatePercent = RatePercent(D(r, 7)),
                TestInput = D(r, 8),
                TestNg = D(r, 9),
                TestRate = D(r, 10),
                TestRatePercent = RatePercent(D(r, 10)),
                DeltaRate = D(r, 11),
                DeltaRatePercentPoint = RatePercent(D(r, 11)),
                ImprovementRate = D(r, 12),
                EffectDirection = S(r, 13),
                Evidence = evidence,
                EvidenceRows = ParseEvidenceRowRefs(evidence),
                PairConfidence = S(r, 15),
            });
        }

        return rows;
    }

    private static List<MicroSpeakerReviewCaseAiMetricCandidate> ReadAiMetricCandidates(
        SqliteConnection conn,
        long fileId,
        int limit)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT metric_id, sheet_name, row_number, table_title, condition_label,
                   input_qty, ok_qty, ng_qty, ng_rate, detail, raw_row, parse_confidence
            FROM metric_candidates
            WHERE file_id=@id
            ORDER BY metric_id
            LIMIT @limit;
            """;
        cmd.Parameters.AddWithValue("@id", fileId);
        cmd.Parameters.AddWithValue("@limit", limit);

        var rows = new List<MicroSpeakerReviewCaseAiMetricCandidate>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string sheet = S(r, 1);
            long rowNumber = L(r, 2);
            rows.Add(new MicroSpeakerReviewCaseAiMetricCandidate
            {
                MetricId = L(r, 0),
                RowId = RowId(sheet, rowNumber),
                SheetName = sheet,
                RowNumber = rowNumber,
                TableTitle = S(r, 3),
                ConditionLabel = S(r, 4),
                InputQty = D(r, 5),
                OkQty = D(r, 6),
                NgQty = D(r, 7),
                NgRate = D(r, 8),
                NgRatePercent = RatePercent(D(r, 8)),
                Detail = S(r, 9),
                RawRow = S(r, 10),
                ParseConfidence = S(r, 11),
            });
        }

        return rows;
    }

    private static List<MicroSpeakerReviewCaseAiMeasurementCandidate> ReadAiMeasurementCandidates(
        SqliteConnection conn,
        long fileId,
        int limit)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT stat_id, sheet_name, row_number, item_label, condition_label, spec,
                   min_value, max_value, avg_value, sample_count, violation_count,
                   raw_row, parse_confidence
            FROM measurement_stats
            WHERE file_id=@id
            ORDER BY stat_id
            LIMIT @limit;
            """;
        cmd.Parameters.AddWithValue("@id", fileId);
        cmd.Parameters.AddWithValue("@limit", limit);

        var rows = new List<MicroSpeakerReviewCaseAiMeasurementCandidate>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string sheet = S(r, 1);
            long rowNumber = L(r, 2);
            rows.Add(new MicroSpeakerReviewCaseAiMeasurementCandidate
            {
                StatId = L(r, 0),
                RowId = RowId(sheet, rowNumber),
                SheetName = sheet,
                RowNumber = rowNumber,
                ItemLabel = S(r, 3),
                ConditionLabel = S(r, 4),
                Spec = S(r, 5),
                Min = D(r, 6),
                Max = D(r, 7),
                Avg = D(r, 8),
                SampleCount = L(r, 9),
                ViolationCount = L(r, 10),
                RawRow = S(r, 11),
                ParseConfidence = S(r, 12),
            });
        }

        return rows;
    }

    private static List<MicroSpeakerReviewCaseAiTermHint> ReadAiTermHints(
        SqliteConnection conn,
        long fileId,
        int limit)
    {
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT term_raw, term_type, normalized_name, korean_desc, hit_count, example_context
            FROM term_hits
            WHERE file_id=@id
            ORDER BY hit_count DESC, term_raw
            LIMIT @limit;
            """;
        cmd.Parameters.AddWithValue("@id", fileId);
        cmd.Parameters.AddWithValue("@limit", limit);

        var rows = new List<MicroSpeakerReviewCaseAiTermHint>();
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            rows.Add(new MicroSpeakerReviewCaseAiTermHint
            {
                TermRaw = S(r, 0),
                TermType = S(r, 1),
                NormalizedName = S(r, 2),
                KoreanDescription = S(r, 3),
                HitCount = L(r, 4),
                ExampleContext = S(r, 5),
            });
        }

        return rows;
    }

    private static MicroSpeakerReviewCaseAuditDecision? ReadAuditDecision(long fileId, string path)
    {
        if (!File.Exists(path)) return null;

        foreach (string line in File.ReadLines(path))
        {
            Match m = Regex.Match(line, @"^\|\s*(?<id>\d+)\s*\|\s*(?<decision>[^|]+)\|\s*(?<reason>[^|]+)\|\s*`?(?<file>[^|`]+)`?\s*\|");
            if (!m.Success) continue;
            if (!long.TryParse(m.Groups["id"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long id) || id != fileId)
                continue;

            return new MicroSpeakerReviewCaseAuditDecision
            {
                FileId = id,
                Decision = m.Groups["decision"].Value.Trim(),
                Reason = m.Groups["reason"].Value.Trim(),
                FileName = m.Groups["file"].Value.Trim(),
            };
        }

        return null;
    }

    private static List<MicroSpeakerReviewCaseAiRowRef> ParseEvidenceRowRefs(string evidence)
    {
        var rows = new List<MicroSpeakerReviewCaseAiRowRef>();
        foreach (Match m in Regex.Matches(evidence ?? "", @"(?<sheet>[^!;]+)!(?<row>\d+)"))
        {
            string sheet = m.Groups["sheet"].Value.Trim();
            if (!long.TryParse(m.Groups["row"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long rowNumber))
                continue;

            rows.Add(new MicroSpeakerReviewCaseAiRowRef
            {
                RowId = RowId(sheet, rowNumber),
                SheetName = sheet,
                RowNumber = rowNumber,
            });
        }

        return rows
            .GroupBy(r => r.RowId, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToList();
    }

    private static string BuildQueryWhere(SqliteCommand cmd, string query, string[] columns)
    {
        if (string.IsNullOrWhiteSpace(query)) return "";
        cmd.Parameters.AddWithValue("@q", "%" + query.Trim() + "%");
        return "WHERE " + string.Join(" OR ", columns.Select(c => $"COALESCE({c}, '') LIKE @q"));
    }

    private static ReviewPairRow? FindTotalEquivalentRow(IReadOnlyList<ReviewPairRow> rows)
    {
        if (rows.Count < 2) return null;

        foreach (ReviewPairRow candidate in rows)
        {
            if (!IsTotalLike(candidate.TestCondition) && !IsTotalLike(candidate.Evidence)) continue;
            if (!candidate.TestInput.HasValue || !candidate.TestNg.HasValue) continue;

            List<ReviewPairRow> others = rows.Where(r => !ReferenceEquals(r, candidate)).ToList();
            double? otherInput = SumNullable(others.Select(r => r.TestInput));
            double? otherNg = SumNullable(others.Select(r => r.TestNg));
            if (NearlyEqual(candidate.TestInput.Value, otherInput) && NearlyEqual(candidate.TestNg.Value, otherNg))
                return candidate;
        }

        return null;
    }

    private static ReviewMeasurementRow? FindBestNormalMeasurement(ReviewMeasurementRow test, IReadOnlyList<ReviewMeasurementRow> normals)
    {
        string testCore = ConditionCore(test.ConditionLabel);
        string testLast = LastConditionToken(test.ConditionLabel);

        return normals
            .Select(n =>
            {
                string normalCore = ConditionCore(n.ConditionLabel);
                string normalLast = LastConditionToken(n.ConditionLabel);
                int score =
                    string.Equals(testCore, normalCore, StringComparison.OrdinalIgnoreCase) ? 10 :
                    !string.IsNullOrWhiteSpace(testLast) && string.Equals(testLast, normalLast, StringComparison.OrdinalIgnoreCase) ? 8 :
                    ContainsEither(testCore, normalCore) ? 6 :
                    normals.Count == 1 ? 2 :
                    0;
                long distance = Math.Abs(test.RowNumber - n.RowNumber);
                return new { Row = n, Score = score, Distance = distance };
            })
            .Where(x => x.Score > 0)
            .OrderByDescending(x => x.Score)
            .ThenBy(x => x.Distance)
            .Select(x => x.Row)
            .FirstOrDefault();
    }

    private static MicroSpeakerReviewCaseMeasurement ToMeasurement(ReviewMeasurementRow row)
        => new()
        {
            Condition = row.ConditionLabel,
            SampleCount = row.SampleCount,
            ViolationCount = row.ViolationCount,
            Min = row.MinValue,
            Max = row.MaxValue,
            Avg = row.AvgValue,
            Spec = row.Spec,
        };

    private static string CanonicalChangedFactor(MicroSpeakerReviewCase row)
    {
        string sourceTitle = CleanFileTitle(row.SourceFile);
        string factor = (row.ChangedFactor ?? "").Trim();
        string after = StripTotalPrefix(row.AfterCondition);

        if (LooksLikeSourceTitle(factor, sourceTitle))
            return Brief(factor, 180);

        if (IsSpecificChangeStatement(after))
            return Brief(after, 180);

        if (!string.IsNullOrWhiteSpace(factor)
            && !LooksLikeGenericResultTitle(factor)
            && !LooksLikeSourceTitle(sourceTitle, factor))
            return Brief(factor, 180);

        if (!string.IsNullOrWhiteSpace(sourceTitle))
            return Brief(sourceTitle, 180);

        return Brief(FirstUseful(factor, row.ReviewTitle, row.OutcomeMetric), 180);
    }

    private static string ReviewCaseGroupKey(MicroSpeakerReviewCase row)
    {
        string sourceTitle = CleanFileTitle(row.SourceFile);
        string key = !string.IsNullOrWhiteSpace(sourceTitle)
            ? sourceTitle
            : CanonicalChangedFactor(row);
        return NormalizeAggregateKey(key);
    }

    private static string ChangeSignatureKey(MicroSpeakerReviewCase row)
        => StableId(
            "ms-change",
            row.FileId.ToString(CultureInfo.InvariantCulture),
            NormalizeAggregateKey(CanonicalChangedFactor(row)),
            NormalizeAggregateKey(row.ChangeDomain),
            NormalizeAggregateKey(row.BeforeCondition),
            NormalizeAggregateKey(row.AfterCondition),
            NormalizeAggregateKey(row.ReviewedProcess));

    private static string BestChangedFactorLabel(IEnumerable<MicroSpeakerReviewCase> rows)
    {
        List<MicroSpeakerReviewCase> list = rows.ToList();
        string sourceTitle = list.Count == 0 ? "" : CleanFileTitle(list[0].SourceFile);
        return list
            .Select(CanonicalChangedFactor)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .OrderBy(x => LooksLikeGenericResultTitle(x) ? 1 : 0)
            .ThenBy(x => LooksLikeSourceTitle(x, sourceTitle) ? 1 : 0)
            .ThenBy(x => x.Length)
            .FirstOrDefault() ?? "";
    }

    private static string BestReviewTitle(IReadOnlyList<MicroSpeakerReviewCase> rows)
    {
        string sourceTitle = CleanFileTitle(rows[0].SourceFile);
        if (!string.IsNullOrWhiteSpace(sourceTitle)) return Brief(sourceTitle, 220);

        return rows
            .Select(r => r.ReviewTitle)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .OrderByDescending(x => x.Length)
            .FirstOrDefault() ?? "";
    }

    private static string BestReviewPurpose(IReadOnlyList<MicroSpeakerReviewCase> rows)
        => rows
            .Select(r => r.ReviewPurpose)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .OrderByDescending(x => x.Length)
            .FirstOrDefault() ?? "";

    private static int OutcomeOrder(string outcomeDomain)
        => (outcomeDomain ?? "").ToLowerInvariant() switch
        {
            "process defect" => 0,
            "function defect" => 1,
            "measurement" => 2,
            "reliability" => 3,
            _ => 9,
        };

    private static string MergeConfidence(IEnumerable<string> values)
    {
        string[] confidences = values.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
        if (confidences.Any(x => x.Equals("high", StringComparison.OrdinalIgnoreCase))) return "high";
        if (confidences.Any(x => x.Equals("medium", StringComparison.OrdinalIgnoreCase))) return "medium";
        if (confidences.Any(x => x.Equals("low", StringComparison.OrdinalIgnoreCase))) return "low";
        return "unknown";
    }

    private static bool LooksLikeSourceTitle(string value, string sourceTitle)
    {
        string a = NormalizeTextForSimilarity(value);
        string b = NormalizeTextForSimilarity(sourceTitle);
        if (a.Length < 12 || b.Length < 12) return false;
        return a.Contains(b, StringComparison.OrdinalIgnoreCase)
               || b.Contains(a, StringComparison.OrdinalIgnoreCase)
               || TokenOverlapRatio(a, b) >= 0.62;
    }

    private static bool LooksLikeGenericResultTitle(string value)
    {
        string lower = (value ?? "").ToLowerInvariant();
        return ContainsAny(lower, "result checking", "result check", "checking function", "checking vision")
               || Regex.IsMatch(lower, @"^\d+\.\s*result", RegexOptions.IgnoreCase);
    }

    private static string NormalizeTextForSimilarity(string value)
    {
        string text = Regex.Replace((value ?? "").ToLowerInvariant(), @"_clean$|_\d{8,}(?:_\d{8,})?$", " ");
        text = Regex.Replace(text, @"[^a-z0-9]+", " ");
        return Regex.Replace(text, @"\s+", " ").Trim();
    }

    private static double TokenOverlapRatio(string a, string b)
    {
        HashSet<string> left = NormalizeTextForSimilarity(a)
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(t => t.Length > 2)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        HashSet<string> right = NormalizeTextForSimilarity(b)
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(t => t.Length > 2)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (left.Count == 0 || right.Count == 0) return 0;
        int overlap = left.Count(right.Contains);
        return overlap / (double)Math.Min(left.Count, right.Count);
    }

    private static string InferChangeDomain(string text)
    {
        string lower = (text ?? "").ToLowerInvariant();
        var domains = new List<string>();

        if (ContainsAny(lower, "supplier", "vendor")) domains.Add("supplier");
        if (Regex.IsMatch(lower, @"\blot\b|lot date|material lot")) domains.Add("lot");
        if (ContainsAny(lower, "material", "raw material", " mtr ", " mtr-", "mtr ")) domains.Add("material");
        if (ContainsAny(lower, "plating", "lating", "polish", "non polish", " tin ")) domains.Add("plating method");
        if (ContainsAny(lower, "bonding amount", "bond amount", "glue amount", "bonding quantity")) domains.Add("bonding amount");
        if (ContainsAny(lower, "machine", "equipment", "new mc", " mc ", " awf ")) domains.Add("equipment");
        if (ContainsAny(lower, "jig", "fixture")) domains.Add("jig");
        if (ContainsAny(lower, "mold", "mould")) domains.Add("mold");
        if (ContainsAny(lower, "method", "condition", "press", "dry", "uv", "plasma", "bonding", "assembly", "ass'y", "assy")) domains.Add("process method");
        if (ContainsAny(lower, "thickness", "dimension", "spec", "cutting", "%")) domains.Add("spec");
        if (ContainsAny(lower, "inspection", "retest", "checking")) domains.Add("inspection");

        return domains.Count == 0
            ? "unknown"
            : string.Join(" + ", domains.Distinct(StringComparer.OrdinalIgnoreCase).Take(3));
    }

    private static string InferChangedFactor(
        string fileName,
        string tableTitle,
        string compareItem,
        string beforeCondition,
        string afterCondition)
    {
        string after = StripTotalPrefix(afterCondition);
        if (IsSpecificChangeStatement(after)) return Brief(after, 180);

        string axis = ExtractChangedFactorAxis(tableTitle, compareItem);
        if (!string.IsNullOrWhiteSpace(axis)) return Brief(axis, 180);

        string before = StripTotalPrefix(beforeCondition);
        if (IsSpecificChangeStatement(before) && IsSpecificChangeStatement(after))
            return Brief(before + " -> " + after, 180);

        string title = CleanFileTitle(fileName);
        if (!string.IsNullOrWhiteSpace(title)) return Brief(title, 180);

        string fallback = FirstUseful(tableTitle, compareItem, after, before);
        return Brief(fallback, 180);
    }

    private static string ExtractChangedFactorAxis(params string[] values)
    {
        var candidates = values
            .SelectMany(v => (v ?? "").Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            .Select(v => Brief(v, 100))
            .Where(v => v.Length is >= 3 and <= 100)
            .Where(v => !IsOutcomeColumn(v))
            .Where(v => ContainsAny(v.ToLowerInvariant(),
                "mold", "mould", "supplier", "vendor", "lot", "material", "mtr",
                "plating", "polish", "machine", "equipment", "jig", "fixture",
                "bonding amount", "bond amount", "glue amount", "method", "condition",
                "press", "dry", "uv", "plasma", "assembly", "ass'y", "assy"))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(v => v.Length)
            .ToList();

        return candidates.FirstOrDefault() ?? "";
    }

    private static string InferReviewPurpose(string fileName, string title)
    {
        string text = FirstUseful(title, CleanFileTitle(fileName));
        return Brief(text, 220);
    }

    private static string InferReviewedProcess(string fileName, string title)
    {
        string text = (title ?? "") + " | " + CleanFileTitle(fileName);
        string lower = text.ToLowerInvariant();
        if (ContainsAny(lower, "spot")) return "spot welding";
        if (ContainsAny(lower, "bond", "bonding", "glue")) return "bonding";
        if (ContainsAny(lower, "assembly", "ass'y", "assy")) return "assembly";
        if (ContainsAny(lower, "press")) return "press";
        if (ContainsAny(lower, "plasma")) return "plasma";
        if (ContainsAny(lower, "dry")) return "dry";
        if (ContainsAny(lower, "uv")) return "UV";
        if (ContainsAny(lower, "vision")) return "vision inspection";
        if (ContainsAny(lower, "function", "hearing")) return "function inspection";
        return Brief(FirstUseful(title ?? "", CleanFileTitle(fileName)), 140);
    }

    private static string InferOutcomeDomain(string text)
    {
        string lower = (text ?? "").ToLowerInvariant();
        if (ContainsAny(lower, "function", "hearing", "sound", "audio", "rub", "buzz", "noise")) return "function defect";
        if (ContainsAny(lower, "tension", "tensile", "strength", "pull", "spl", "thd", "dcr", "imp", "f0")) return "measurement";
        if (ContainsAny(lower, "vision", "process", "spot", "ng rate", "defect", "fail")) return "process defect";
        if (ContainsAny(lower, "drop", "reliability")) return "reliability";
        return "unknown";
    }

    private static string InferOutcomeMetric(string tableTitle, string compareItem, bool isMeasurement)
    {
        string text = FirstUseful(tableTitle, compareItem);
        string lower = text.ToLowerInvariant();
        if (isMeasurement) return Brief(text, 180);
        if (ContainsAny(lower, "function", "hearing")) return "Function NG rate";
        if (ContainsAny(lower, "spot") && ContainsAny(lower, "vision")) return "Spot welding process defect rate";
        if (ContainsAny(lower, "vision", "process")) return "Process NG rate";
        if (ContainsAny(lower, "air leak")) return "Air leak NG rate";
        return Brief(text, 180);
    }

    private static bool IsSpecificChangeStatement(string value)
    {
        string text = StripTotalPrefix(value).Trim();
        if (text.Length < 4) return false;
        string lower = text.ToLowerInvariant();
        if (lower is "normal" or "test" or "total") return false;
        if (Regex.IsMatch(lower, @"^(normal|test)\s*$")) return false;
        return ContainsAny(lower,
            "new", "change", "improve", "machine", "jig", "mold", "lot", "supplier", "material",
            "bond", "amount", "plating", "polish", "method", "press", "dry", "uv", "plasma");
    }

    private static bool IsOutcomeColumn(string value)
    {
        string lower = (value ?? "").ToLowerInvariant();
        return ContainsAny(lower,
            "input", "ok", "ng", "rate", "total", "function", "hearing",
            "air leak", "sigma", "pass", "fail", "result", "date", "line");
    }

    private static bool IsMeasurementOutcomeCandidate(string text)
        => ContainsAny((text ?? "").ToLowerInvariant(),
            "tension", "tensile", "strength", "pull", "spl", "thd", "dcr", "imp", "f0", "hearing", "function");

    private static bool IsNormalCondition(string value)
        => Regex.IsMatch(value ?? "", @"(^|[^a-z])normal([^a-z]|$)|before", RegexOptions.IgnoreCase);

    private static bool IsTestCondition(string value)
        => Regex.IsMatch(value ?? "", @"(^|[^a-z])test([^a-z]|$)|after|new|change", RegexOptions.IgnoreCase);

    private static string ConditionCore(string value)
    {
        string text = (value ?? "").ToLowerInvariant();
        text = Regex.Replace(text, @"\b(total|normal|test|before|after|new|change|changed)\b", " ");
        text = Regex.Replace(text, @"[^a-z0-9]+", " ");
        return Regex.Replace(text, @"\s+", " ").Trim();
    }

    private static string LastConditionToken(string value)
    {
        string[] parts = (value ?? "").Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return parts.Length == 0 ? ConditionCore(value ?? "") : ConditionCore(parts[^1] ?? "");
    }

    private static string JudgementFromEffect(string effect)
        => effect.ToUpperInvariant() switch
        {
            "IMPROVED" => "improved",
            "WORSENED" => "worse",
            "NO_CHANGE" => "no_change",
            "MIXED" => "mixed",
            _ => "not_comparable",
        };

    private static string MeasurementJudgement(double? delta)
    {
        if (!delta.HasValue) return "not_comparable";
        return Math.Abs(delta.Value) < 0.0000001 ? "no_change" : "not_comparable";
    }

    private static string AggregateEffectDirection(double? delta, IReadOnlyList<ReviewPairRow> rows)
    {
        if (delta.HasValue)
        {
            if (Math.Abs(delta.Value) < 0.0000001) return "NO_CHANGE";
            return delta.Value > 0 ? "WORSENED" : "IMPROVED";
        }

        string[] effects = rows.Select(r => r.EffectDirection).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        return effects.Length == 1 ? effects[0] : "MIXED";
    }

    private static string PairLimitReason(string aggregationMethod)
        => aggregationMethod switch
        {
            "single_row" => "Single paired Normal/Test row from comparison_pairs.",
            "total_row" => "Total-equivalent row selected; daily/detail rows were not double-counted.",
            "summed_rows" => "No total-equivalent row found; repeated same-condition rows were summed.",
            _ => "",
        };

    private static string InferConfidence(IEnumerable<string> confidences, string aggregationMethod)
    {
        string[] values = confidences.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
        if (values.Any(x => x.Equals("HIGH", StringComparison.OrdinalIgnoreCase))) return "high";
        if (string.Equals(aggregationMethod, "total_row", StringComparison.OrdinalIgnoreCase)) return "high";
        if (values.Any(x => x.Equals("MEDIUM", StringComparison.OrdinalIgnoreCase))) return "medium";
        return "low";
    }

    private static double ConfidenceScore(string confidence, string aggregationMethod)
    {
        double score = confidence.ToLowerInvariant() switch
        {
            "high" => 0.9,
            "medium" => 0.7,
            "low" => 0.45,
            _ => 0.5,
        };
        if (string.Equals(aggregationMethod, "total_row", StringComparison.OrdinalIgnoreCase)) score += 0.05;
        if (string.Equals(aggregationMethod, "summed_rows", StringComparison.OrdinalIgnoreCase)) score -= 0.05;
        return Math.Clamp(score, 0, 1);
    }

    private static string CleanFileTitle(string fileName)
    {
        string title = Path.GetFileNameWithoutExtension(fileName ?? "");
        title = Regex.Replace(title, @"_clean$", "", RegexOptions.IgnoreCase);
        title = Regex.Replace(title, @"_\d{8,}(?:_\d{8,})?$", "");
        title = Regex.Replace(title, @"[_]+", " ");
        title = Regex.Replace(title, @"\s+", " ").Trim();
        return title;
    }

    private static string NormalizeAggregateKey(string value)
    {
        string normalized = Regex.Replace((value ?? "").Trim().ToLowerInvariant(), @"\s+", " ");
        normalized = Regex.Replace(normalized, @"^(total|subtotal)\s*\|\s*", "", RegexOptions.IgnoreCase);
        normalized = Regex.Replace(normalized, @"^(total|subtotal)\s+", "", RegexOptions.IgnoreCase);
        return normalized;
    }

    private static string StripTotalPrefix(string value)
    {
        string text = (value ?? "").Trim();
        text = Regex.Replace(text, @"^(total|subtotal)\s*\|\s*", "", RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"^(total|subtotal)\s+", "", RegexOptions.IgnoreCase);
        return text.Trim();
    }

    private static string FirstUseful(params string[] values)
        => values.Select(v => (v ?? "").Trim()).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";

    private static string Brief(string text, int max)
    {
        string value = Regex.Replace((text ?? "").Trim(), @"\s+", " ");
        return value.Length <= max ? value : value[..max].TrimEnd() + "...";
    }

    private static string ExtractSheetName(string evidence)
    {
        Match m = Regex.Match(evidence ?? "", @"(?<sheet>[^!;]+)!\d+");
        return m.Success ? m.Groups["sheet"].Value.Trim() : "";
    }

    private static string StableId(string prefix, params string[] parts)
    {
        byte[] bytes = SHA256.HashData(Encoding.UTF8.GetBytes(string.Join('\u001f', parts)));
        return prefix + "-" + Convert.ToHexString(bytes)[..12].ToLowerInvariant();
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
        => input.HasValue && ng.HasValue && Math.Abs(input.Value) > 0.0000001
            ? ng.Value / input.Value
            : null;

    private static double? RatePercent(double? value)
    {
        if (!value.HasValue) return null;
        double v = value.Value;
        return Math.Abs(v) <= 1.5 ? v * 100.0 : v;
    }

    private static double? RelativeChangePercent(double? normalValue, double? testValue)
    {
        if (!normalValue.HasValue || !testValue.HasValue) return null;
        if (Math.Abs(normalValue.Value) < 0.0000001) return null;
        return (testValue.Value / normalValue.Value - 1.0) * 100.0;
    }

    private static bool NearlyEqual(double expected, double? actual)
    {
        if (!actual.HasValue) return false;
        double tolerance = Math.Max(1.0, Math.Abs(expected) * 0.002);
        return Math.Abs(expected - actual.Value) <= tolerance;
    }

    private static bool IsTotalLike(string value)
        => Regex.IsMatch(value ?? "", @"(^|[^a-z])total([^a-z]|$)|subtotal", RegexOptions.IgnoreCase);

    private static bool ContainsAny(string text, params string[] terms)
        => terms.Any(term => !string.IsNullOrWhiteSpace(term)
                             && (text ?? "").Contains(term, StringComparison.OrdinalIgnoreCase));

    private static bool ContainsEither(string a, string b)
        => !string.IsNullOrWhiteSpace(a)
           && !string.IsNullOrWhiteSpace(b)
           && (a.Contains(b, StringComparison.OrdinalIgnoreCase) || b.Contains(a, StringComparison.OrdinalIgnoreCase));

    private static string RowId(string sheetName, long rowNumber)
        => string.IsNullOrWhiteSpace(sheetName)
            ? rowNumber.ToString(CultureInfo.InvariantCulture)
            : $"{sheetName}!{rowNumber.ToString(CultureInfo.InvariantCulture)}";

    private static string ColumnLabel(long number)
    {
        if (number <= 0) return "";

        var label = new StringBuilder();
        while (number > 0)
        {
            number--;
            label.Insert(0, (char)('A' + number % 26));
            number /= 26;
        }

        return label.ToString();
    }

    private static bool IsAiContextRow(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;

        string lower = text.ToLowerInvariant();
        return ContainsAny(lower,
            "purpose", "objective", "summary", "content", "condition", "standard", "spec",
            "result", "judgement", "judgment", "decision", "conclusion", "remark", "note",
            "problem", "reason", "cause", "countermeasure", "before", "after", "normal",
            "test", "sample", "lot", "supplier", "material", "machine", "m/c", "jig",
            "base left", "base right", "laser delay", "coating cream", "gauss", "tension",
            "function", "vision", "repair")
            || Regex.IsMatch(lower, @"^\s*(i{1,3}|iv|v|\d+)[.)]\s+", RegexOptions.IgnoreCase);
    }

    private static string SourceFileUrl(long fileId) => $"/microspeaker/source-file/{fileId}";

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

    private static string S(SqliteDataReader r, int i)
        => r.IsDBNull(i) ? "" : Convert.ToString(r.GetValue(i), CultureInfo.InvariantCulture) ?? "";

    private static long L(SqliteDataReader r, int i)
        => r.IsDBNull(i) ? 0 : Convert.ToInt64(r.GetValue(i), CultureInfo.InvariantCulture);

    private static double? D(SqliteDataReader r, int i)
        => r.IsDBNull(i) ? null : Convert.ToDouble(r.GetValue(i), CultureInfo.InvariantCulture);

    private sealed record ReviewPairRow(
        long PairId,
        long FileId,
        string Dataset,
        string FileName,
        string Models,
        string Categories,
        string DatesFound,
        string TermSummary,
        string TableTitle,
        string CompareItem,
        string ControlCondition,
        string TestCondition,
        double? ControlInput,
        double? ControlNg,
        double? ControlRate,
        double? TestInput,
        double? TestNg,
        double? TestRate,
        double? DeltaRate,
        double? ImprovementRate,
        string EffectDirection,
        string Evidence,
        string PairConfidence);

    private sealed record ReviewMeasurementRow(
        long StatId,
        long FileId,
        string Dataset,
        string FileName,
        string Models,
        string Categories,
        string DatesFound,
        string TermSummary,
        string SheetName,
        long RowNumber,
        string ItemLabel,
        string ConditionLabel,
        string Spec,
        double? MinValue,
        double? MaxValue,
        double? AvgValue,
        long SampleCount,
        long ViolationCount,
        string RawRow,
        string ParseConfidence)
    {
        public string RowEvidence => string.IsNullOrWhiteSpace(SheetName)
            ? RowNumber.ToString(CultureInfo.InvariantCulture)
            : $"{SheetName}!{RowNumber.ToString(CultureInfo.InvariantCulture)}";
    }
}

public sealed class MicroSpeakerReviewCaseSample
{
    [JsonPropertyName("createdAt")] public string CreatedAt { get; set; } = "";
    [JsonPropertyName("microSpeakerDatabasePath")] public string MicroSpeakerDatabasePath { get; set; } = "";
    [JsonPropertyName("databaseExists")] public bool DatabaseExists { get; set; }
    [JsonPropertyName("query")] public string Query { get; set; } = "";
    [JsonPropertyName("limit")] public int Limit { get; set; }
    [JsonPropertyName("sourcePairRows")] public int SourcePairRows { get; set; }
    [JsonPropertyName("sourceMeasurementRows")] public int SourceMeasurementRows { get; set; }
    [JsonPropertyName("reviewCaseCount")] public int ReviewCaseCount { get; set; }
    [JsonPropertyName("outcomeCount")] public int OutcomeCount { get; set; }
    [JsonPropertyName("flatCandidateCount")] public int FlatCandidateCount { get; set; }
    [JsonPropertyName("pairCaseCount")] public int PairCaseCount { get; set; }
    [JsonPropertyName("measurementCaseCount")] public int MeasurementCaseCount { get; set; }
    [JsonPropertyName("reviewCases")] public List<MicroSpeakerReviewCaseGroup> ReviewCases { get; set; } = [];
    [JsonPropertyName("flatCandidates")] public List<MicroSpeakerReviewCase> FlatCandidates { get; set; } = [];
    [JsonPropertyName("notes")] public List<string> Notes { get; set; } = [];
}

public sealed class MicroSpeakerReviewCaseAiPacket
{
    [JsonPropertyName("createdAt")] public string CreatedAt { get; set; } = "";
    [JsonPropertyName("microSpeakerDatabasePath")] public string MicroSpeakerDatabasePath { get; set; } = "";
    [JsonPropertyName("databaseExists")] public bool DatabaseExists { get; set; }
    [JsonPropertyName("sourceFileId")] public long SourceFileId { get; set; }
    [JsonPropertyName("fileFound")] public bool FileFound { get; set; }
    [JsonPropertyName("rowLimit")] public int RowLimit { get; set; }
    [JsonPropertyName("candidateLimit")] public int CandidateLimit { get; set; }
    [JsonPropertyName("reviewCasePromptPath")] public string ReviewCasePromptPath { get; set; } = "";
    [JsonPropertyName("calibrationReferencePath")] public string CalibrationReferencePath { get; set; } = "";
    [JsonPropertyName("auditDecisionPath")] public string AuditDecisionPath { get; set; } = "";
    [JsonPropertyName("file")] public MicroSpeakerReviewCaseAiFile? File { get; set; }
    [JsonPropertyName("auditDecision")] public MicroSpeakerReviewCaseAuditDecision? AuditDecision { get; set; }
    [JsonPropertyName("sheets")] public List<MicroSpeakerReviewCaseAiSheet> Sheets { get; set; } = [];
    [JsonPropertyName("contextRows")] public List<MicroSpeakerReviewCaseAiRowRef> ContextRows { get; set; } = [];
    [JsonPropertyName("sheetRows")] public List<MicroSpeakerReviewCaseAiSheetRow> SheetRows { get; set; } = [];
    [JsonPropertyName("pairCandidates")] public List<MicroSpeakerReviewCaseAiPairCandidate> PairCandidates { get; set; } = [];
    [JsonPropertyName("metricCandidates")] public List<MicroSpeakerReviewCaseAiMetricCandidate> MetricCandidates { get; set; } = [];
    [JsonPropertyName("measurementCandidates")] public List<MicroSpeakerReviewCaseAiMeasurementCandidate> MeasurementCandidates { get; set; } = [];
    [JsonPropertyName("termHints")] public List<MicroSpeakerReviewCaseAiTermHint> TermHints { get; set; } = [];
    [JsonPropertyName("notes")] public List<string> Notes { get; set; } = [];
}

public sealed class MicroSpeakerReviewCaseAiFile
{
    [JsonPropertyName("fileId")] public long FileId { get; set; }
    [JsonPropertyName("dataset")] public string Dataset { get; set; } = "";
    [JsonPropertyName("path")] public string Path { get; set; } = "";
    [JsonPropertyName("fileName")] public string FileName { get; set; } = "";
    [JsonPropertyName("extension")] public string Extension { get; set; } = "";
    [JsonPropertyName("sizeBytes")] public long SizeBytes { get; set; }
    [JsonPropertyName("status")] public string Status { get; set; } = "";
    [JsonPropertyName("error")] public string Error { get; set; } = "";
    [JsonPropertyName("sheetCount")] public long SheetCount { get; set; }
    [JsonPropertyName("sheetNames")] public string SheetNames { get; set; } = "";
    [JsonPropertyName("maxRowsAnySheet")] public long MaxRowsAnySheet { get; set; }
    [JsonPropertyName("maxColsAnySheet")] public long MaxColsAnySheet { get; set; }
    [JsonPropertyName("nonEmptyCells")] public long NonEmptyCells { get; set; }
    [JsonPropertyName("models")] public string Models { get; set; } = "";
    [JsonPropertyName("categories")] public string Categories { get; set; } = "";
    [JsonPropertyName("datesFound")] public string DatesFound { get; set; } = "";
    [JsonPropertyName("structureFamily")] public string StructureFamily { get; set; } = "";
    [JsonPropertyName("structureConfidence")] public string StructureConfidence { get; set; } = "";
    [JsonPropertyName("termSummary")] public string TermSummary { get; set; } = "";
    [JsonPropertyName("metricCandidateCount")] public long MetricCandidateCount { get; set; }
    [JsonPropertyName("measurementStatCount")] public long MeasurementStatCount { get; set; }
    [JsonPropertyName("comparisonPairCount")] public long ComparisonPairCount { get; set; }
    [JsonPropertyName("sheetRowCount")] public long SheetRowCount { get; set; }
    [JsonPropertyName("sheetCellCount")] public long SheetCellCount { get; set; }
    [JsonPropertyName("originalFileUrl")] public string OriginalFileUrl { get; set; } = "";
}

public sealed class MicroSpeakerReviewCaseAiSheet
{
    [JsonPropertyName("sheetName")] public string SheetName { get; set; } = "";
    [JsonPropertyName("rowCount")] public long RowCount { get; set; }
    [JsonPropertyName("colCount")] public long ColCount { get; set; }
    [JsonPropertyName("nonEmptyCount")] public long NonEmptyCount { get; set; }
    [JsonPropertyName("sampleText")] public string SampleText { get; set; } = "";
}

public sealed class MicroSpeakerReviewCaseAiSheetRow
{
    [JsonPropertyName("rowId")] public string RowId { get; set; } = "";
    [JsonPropertyName("sheetName")] public string SheetName { get; set; } = "";
    [JsonPropertyName("rowNumber")] public long RowNumber { get; set; }
    [JsonPropertyName("nonEmptyCount")] public long NonEmptyCount { get; set; }
    [JsonPropertyName("rowText")] public string RowText { get; set; } = "";
    [JsonPropertyName("cellsJson")] public string CellsJson { get; set; } = "";
    [JsonPropertyName("cells")] public List<MicroSpeakerReviewCaseAiCell> Cells { get; set; } = [];
}

public sealed class MicroSpeakerReviewCaseAiCell
{
    [JsonPropertyName("cellId")] public string CellId { get; set; } = "";
    [JsonPropertyName("colNumber")] public long ColNumber { get; set; }
    [JsonPropertyName("colLabel")] public string ColLabel { get; set; } = "";
    [JsonPropertyName("value")] public string Value { get; set; } = "";
}

public sealed class MicroSpeakerReviewCaseAiRowRef
{
    [JsonPropertyName("rowId")] public string RowId { get; set; } = "";
    [JsonPropertyName("sheetName")] public string SheetName { get; set; } = "";
    [JsonPropertyName("rowNumber")] public long RowNumber { get; set; }
    [JsonPropertyName("rowText")] public string RowText { get; set; } = "";
}

public sealed class MicroSpeakerReviewCaseAiPairCandidate
{
    [JsonPropertyName("pairId")] public long PairId { get; set; }
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
    [JsonPropertyName("improvementRate")] public double? ImprovementRate { get; set; }
    [JsonPropertyName("effectDirection")] public string EffectDirection { get; set; } = "";
    [JsonPropertyName("evidence")] public string Evidence { get; set; } = "";
    [JsonPropertyName("evidenceRows")] public List<MicroSpeakerReviewCaseAiRowRef> EvidenceRows { get; set; } = [];
    [JsonPropertyName("pairConfidence")] public string PairConfidence { get; set; } = "";
}

public sealed class MicroSpeakerReviewCaseAiMetricCandidate
{
    [JsonPropertyName("metricId")] public long MetricId { get; set; }
    [JsonPropertyName("rowId")] public string RowId { get; set; } = "";
    [JsonPropertyName("sheetName")] public string SheetName { get; set; } = "";
    [JsonPropertyName("rowNumber")] public long RowNumber { get; set; }
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
}

public sealed class MicroSpeakerReviewCaseAiMeasurementCandidate
{
    [JsonPropertyName("statId")] public long StatId { get; set; }
    [JsonPropertyName("rowId")] public string RowId { get; set; } = "";
    [JsonPropertyName("sheetName")] public string SheetName { get; set; } = "";
    [JsonPropertyName("rowNumber")] public long RowNumber { get; set; }
    [JsonPropertyName("itemLabel")] public string ItemLabel { get; set; } = "";
    [JsonPropertyName("conditionLabel")] public string ConditionLabel { get; set; } = "";
    [JsonPropertyName("spec")] public string Spec { get; set; } = "";
    [JsonPropertyName("min")] public double? Min { get; set; }
    [JsonPropertyName("max")] public double? Max { get; set; }
    [JsonPropertyName("avg")] public double? Avg { get; set; }
    [JsonPropertyName("sampleCount")] public long SampleCount { get; set; }
    [JsonPropertyName("violationCount")] public long ViolationCount { get; set; }
    [JsonPropertyName("rawRow")] public string RawRow { get; set; } = "";
    [JsonPropertyName("parseConfidence")] public string ParseConfidence { get; set; } = "";
}

public sealed class MicroSpeakerReviewCaseAiTermHint
{
    [JsonPropertyName("termRaw")] public string TermRaw { get; set; } = "";
    [JsonPropertyName("termType")] public string TermType { get; set; } = "";
    [JsonPropertyName("normalizedName")] public string NormalizedName { get; set; } = "";
    [JsonPropertyName("koreanDescription")] public string KoreanDescription { get; set; } = "";
    [JsonPropertyName("hitCount")] public long HitCount { get; set; }
    [JsonPropertyName("exampleContext")] public string ExampleContext { get; set; } = "";
}

public sealed class MicroSpeakerReviewCaseAuditDecision
{
    [JsonPropertyName("fileId")] public long FileId { get; set; }
    [JsonPropertyName("decision")] public string Decision { get; set; } = "";
    [JsonPropertyName("reason")] public string Reason { get; set; } = "";
    [JsonPropertyName("fileName")] public string FileName { get; set; } = "";
}

public sealed class MicroSpeakerReviewCaseGroup
{
    [JsonPropertyName("reviewCaseId")] public string ReviewCaseId { get; set; } = "";
    [JsonPropertyName("sourceSystem")] public string SourceSystem { get; set; } = "";
    [JsonPropertyName("fileId")] public long FileId { get; set; }
    [JsonPropertyName("dataset")] public string Dataset { get; set; } = "";
    [JsonPropertyName("sourceFile")] public string SourceFile { get; set; } = "";
    [JsonPropertyName("originalFileUrl")] public string OriginalFileUrl { get; set; } = "";
    [JsonPropertyName("reviewTitle")] public string ReviewTitle { get; set; } = "";
    [JsonPropertyName("reviewPurpose")] public string ReviewPurpose { get; set; } = "";
    [JsonPropertyName("changedFactors")] public List<MicroSpeakerReviewChangedFactor> ChangedFactors { get; set; } = [];
    [JsonPropertyName("outcomes")] public List<MicroSpeakerReviewOutcome> Outcomes { get; set; } = [];
    [JsonPropertyName("evidenceRows")] public List<string> EvidenceRows { get; set; } = [];
    [JsonPropertyName("confidence")] public string Confidence { get; set; } = "";
    [JsonPropertyName("confidenceScore")] public double ConfidenceScore { get; set; }
    [JsonPropertyName("limitReason")] public string LimitReason { get; set; } = "";
}

public sealed class MicroSpeakerReviewChangedFactor
{
    [JsonPropertyName("changeKey")] public string ChangeKey { get; set; } = "";
    [JsonPropertyName("changeDomain")] public string ChangeDomain { get; set; } = "";
    [JsonPropertyName("changedFactor")] public string ChangedFactor { get; set; } = "";
    [JsonPropertyName("beforeCondition")] public string BeforeCondition { get; set; } = "";
    [JsonPropertyName("afterCondition")] public string AfterCondition { get; set; } = "";
    [JsonPropertyName("reviewedProcess")] public string ReviewedProcess { get; set; } = "";
    [JsonPropertyName("evidence")] public string Evidence { get; set; } = "";
    [JsonPropertyName("outcomes")] public List<MicroSpeakerReviewOutcome> Outcomes { get; set; } = [];
}

public sealed class MicroSpeakerReviewOutcome
{
    [JsonPropertyName("outcomeId")] public string OutcomeId { get; set; } = "";
    [JsonPropertyName("changedFactorKey")] public string ChangedFactorKey { get; set; } = "";
    [JsonPropertyName("extractionSource")] public string ExtractionSource { get; set; } = "";
    [JsonPropertyName("outcomeDomain")] public string OutcomeDomain { get; set; } = "";
    [JsonPropertyName("outcomeMetric")] public string OutcomeMetric { get; set; } = "";
    [JsonPropertyName("sheetName")] public string SheetName { get; set; } = "";
    [JsonPropertyName("sourceRows")] public List<string> SourceRows { get; set; } = [];
    [JsonPropertyName("normalInput")] public double? NormalInput { get; set; }
    [JsonPropertyName("normalNg")] public double? NormalNg { get; set; }
    [JsonPropertyName("normalRate")] public double? NormalRate { get; set; }
    [JsonPropertyName("normalRatePercent")] public double? NormalRatePercent { get; set; }
    [JsonPropertyName("testInput")] public double? TestInput { get; set; }
    [JsonPropertyName("testNg")] public double? TestNg { get; set; }
    [JsonPropertyName("testRate")] public double? TestRate { get; set; }
    [JsonPropertyName("testRatePercent")] public double? TestRatePercent { get; set; }
    [JsonPropertyName("deltaRate")] public double? DeltaRate { get; set; }
    [JsonPropertyName("deltaRatePercentPoint")] public double? DeltaRatePercentPoint { get; set; }
    [JsonPropertyName("relativeChangePercent")] public double? RelativeChangePercent { get; set; }
    [JsonPropertyName("normalMeasurement")] public MicroSpeakerReviewCaseMeasurement? NormalMeasurement { get; set; }
    [JsonPropertyName("testMeasurement")] public MicroSpeakerReviewCaseMeasurement? TestMeasurement { get; set; }
    [JsonPropertyName("measurementDeltaAvg")] public double? MeasurementDeltaAvg { get; set; }
    [JsonPropertyName("measurementRelativeChangePercent")] public double? MeasurementRelativeChangePercent { get; set; }
    [JsonPropertyName("judgement")] public string Judgement { get; set; } = "";
    [JsonPropertyName("aggregationMethod")] public string AggregationMethod { get; set; } = "";
    [JsonPropertyName("confidence")] public string Confidence { get; set; } = "";
    [JsonPropertyName("confidenceScore")] public double ConfidenceScore { get; set; }
    [JsonPropertyName("limitReason")] public string LimitReason { get; set; } = "";
    [JsonPropertyName("evidence")] public string Evidence { get; set; } = "";
}

public sealed class MicroSpeakerReviewCase
{
    [JsonPropertyName("caseId")] public string CaseId { get; set; } = "";
    [JsonPropertyName("sourceSystem")] public string SourceSystem { get; set; } = "";
    [JsonPropertyName("extractionSource")] public string ExtractionSource { get; set; } = "";
    [JsonPropertyName("fileId")] public long FileId { get; set; }
    [JsonPropertyName("dataset")] public string Dataset { get; set; } = "";
    [JsonPropertyName("sourceFile")] public string SourceFile { get; set; } = "";
    [JsonPropertyName("originalFileUrl")] public string OriginalFileUrl { get; set; } = "";
    [JsonPropertyName("sheetName")] public string SheetName { get; set; } = "";
    [JsonPropertyName("sourceRows")] public List<string> SourceRows { get; set; } = [];
    [JsonPropertyName("reviewTitle")] public string ReviewTitle { get; set; } = "";
    [JsonPropertyName("reviewPurpose")] public string ReviewPurpose { get; set; } = "";
    [JsonPropertyName("changeDomain")] public string ChangeDomain { get; set; } = "";
    [JsonPropertyName("changedFactor")] public string ChangedFactor { get; set; } = "";
    [JsonPropertyName("beforeCondition")] public string BeforeCondition { get; set; } = "";
    [JsonPropertyName("afterCondition")] public string AfterCondition { get; set; } = "";
    [JsonPropertyName("reviewedProcess")] public string ReviewedProcess { get; set; } = "";
    [JsonPropertyName("outcomeDomain")] public string OutcomeDomain { get; set; } = "";
    [JsonPropertyName("outcomeMetric")] public string OutcomeMetric { get; set; } = "";
    [JsonPropertyName("normalInput")] public double? NormalInput { get; set; }
    [JsonPropertyName("normalNg")] public double? NormalNg { get; set; }
    [JsonPropertyName("normalRate")] public double? NormalRate { get; set; }
    [JsonPropertyName("normalRatePercent")] public double? NormalRatePercent { get; set; }
    [JsonPropertyName("testInput")] public double? TestInput { get; set; }
    [JsonPropertyName("testNg")] public double? TestNg { get; set; }
    [JsonPropertyName("testRate")] public double? TestRate { get; set; }
    [JsonPropertyName("testRatePercent")] public double? TestRatePercent { get; set; }
    [JsonPropertyName("deltaRate")] public double? DeltaRate { get; set; }
    [JsonPropertyName("deltaRatePercentPoint")] public double? DeltaRatePercentPoint { get; set; }
    [JsonPropertyName("relativeChangePercent")] public double? RelativeChangePercent { get; set; }
    [JsonPropertyName("normalMeasurement")] public MicroSpeakerReviewCaseMeasurement? NormalMeasurement { get; set; }
    [JsonPropertyName("testMeasurement")] public MicroSpeakerReviewCaseMeasurement? TestMeasurement { get; set; }
    [JsonPropertyName("measurementDeltaAvg")] public double? MeasurementDeltaAvg { get; set; }
    [JsonPropertyName("measurementRelativeChangePercent")] public double? MeasurementRelativeChangePercent { get; set; }
    [JsonPropertyName("judgement")] public string Judgement { get; set; } = "";
    [JsonPropertyName("aggregationMethod")] public string AggregationMethod { get; set; } = "";
    [JsonPropertyName("confidence")] public string Confidence { get; set; } = "";
    [JsonPropertyName("confidenceScore")] public double ConfidenceScore { get; set; }
    [JsonPropertyName("limitReason")] public string LimitReason { get; set; } = "";
    [JsonPropertyName("evidence")] public string Evidence { get; set; } = "";
}

public sealed class MicroSpeakerReviewCaseMeasurement
{
    [JsonPropertyName("condition")] public string Condition { get; set; } = "";
    [JsonPropertyName("sampleCount")] public long SampleCount { get; set; }
    [JsonPropertyName("violationCount")] public long ViolationCount { get; set; }
    [JsonPropertyName("min")] public double? Min { get; set; }
    [JsonPropertyName("max")] public double? Max { get; set; }
    [JsonPropertyName("avg")] public double? Avg { get; set; }
    [JsonPropertyName("spec")] public string Spec { get; set; } = "";
}
