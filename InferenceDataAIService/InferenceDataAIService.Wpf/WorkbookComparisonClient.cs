using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace InferenceDataAIService.Wpf;

internal sealed class WorkbookComparisonClient(AppPathSettings paths)
{
    private readonly string _benchmarkResultPath =
        paths.BenchmarkResultPath;

    internal Task<IReadOnlyList<WorkbookComparisonSummary>> ListAsync(
        string databasePath) =>
        Task.Run(() => List(databasePath));

    internal Task<IReadOnlyList<WorkbookComparisonSummary>> ListSourcesAsync(
        string databasePath,
        IReadOnlyList<string> sourcePaths) =>
        Task.Run(() => ListSources(
            databasePath,
            sourcePaths));

    internal Task<WorkbookComparisonDocument> LoadAsync(
        string databasePath,
        string publicAnalysisId) =>
        Task.Run(() => Load(databasePath, publicAnalysisId));

    private IReadOnlyList<WorkbookComparisonSummary> List(
        string databasePath)
    {
        using var connection = OpenReadOnly(databasePath);
        var benchmarkSources = ReadBenchmarkSources();
        if (benchmarkSources.Count == 0)
            return ListLatestAnalyses(connection);

        return ReadSummaries(
            connection,
            benchmarkSources);
    }

    private static IReadOnlyList<WorkbookComparisonSummary> ListSources(
        string databasePath,
        IReadOnlyList<string> sourcePaths)
    {
        using var connection = OpenReadOnly(databasePath);
        return ReadSummaries(
            connection,
            sourcePaths);
    }

    private static IReadOnlyList<WorkbookComparisonSummary> ReadSummaries(
        SqliteConnection connection,
        IReadOnlyList<string> sourcePaths)
    {
        var result = new List<WorkbookComparisonSummary>(
            sourcePaths.Count);
        for (var index = 0; index < sourcePaths.Count; index++)
        {
            var sourcePath = sourcePaths[index];
            result.Add(ReadSummary(
                connection,
                index + 1,
                sourcePath));
        }
        return result;
    }

    private WorkbookComparisonDocument Load(
        string databasePath,
        string publicAnalysisId)
    {
        using var connection = OpenReadOnly(databasePath);
        var sourcePath = ReadSourcePath(connection, publicAnalysisId);
        var studies = ReadStudies(connection, publicAnalysisId);
        var rows = ReadObservationRows(
            connection,
            publicAnalysisId,
            sourcePath);
        return new WorkbookComparisonDocument(
            publicAnalysisId,
            sourcePath,
            studies,
            rows);
    }

    private List<string> ReadBenchmarkSources()
    {
        if (!File.Exists(_benchmarkResultPath))
            return [];

        using var document = JsonDocument.Parse(
            File.ReadAllText(_benchmarkResultPath));
        if (!document.RootElement.TryGetProperty(
                "items",
                out var items)
            || items.ValueKind != JsonValueKind.Array)
            return [];

        var result = new List<string>();
        foreach (var item in items.EnumerateArray())
        {
            if (!item.TryGetProperty(
                    "sourcePath",
                    out var sourcePathElement))
                continue;
            var sourcePath = sourcePathElement.GetString();
            if (!string.IsNullOrWhiteSpace(sourcePath))
                result.Add(Path.GetFullPath(sourcePath));
        }
        return result;
    }

    private static SqliteConnection OpenReadOnly(string databasePath)
    {
        if (!File.Exists(databasePath))
            throw new FileNotFoundException(
                "검토 DB를 찾을 수 없습니다.",
                databasePath);

        var connection = new SqliteConnection(
            new SqliteConnectionStringBuilder
            {
                DataSource = Path.GetFullPath(databasePath),
                Mode = SqliteOpenMode.ReadOnly,
            }.ToString());
        connection.Open();
        return connection;
    }

    private static WorkbookComparisonSummary ReadSummary(
        SqliteConnection connection,
        int benchmarkNumber,
        string sourcePath)
    {
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                wa.public_analysis_id,
                wa.title,
                wa.verification_status,
                sd.source_path,
                (
                    SELECT COUNT(*)
                    FROM knowledge_studies s
                    WHERE s.workbook_analysis_id=wa.workbook_analysis_id
                ) AS study_count,
                (
                    SELECT COUNT(*)
                    FROM knowledge_observations v
                    JOIN knowledge_outcomes o
                      ON o.outcome_id=v.outcome_id
                    JOIN knowledge_studies s
                      ON s.study_id=o.study_id
                    WHERE s.workbook_analysis_id=wa.workbook_analysis_id
                ) AS observation_count
            FROM workbook_analyses wa
            JOIN source_documents sd
              ON sd.document_id=wa.document_id
            WHERE sd.source_path=$source
               OR sd.original_file_name=$file_name
            ORDER BY
                CASE WHEN sd.source_path=$source THEN 0 ELSE 1 END,
                wa.updated_at DESC
            LIMIT 1;
            """;
        command.Parameters.AddWithValue("$source", sourcePath);
        command.Parameters.AddWithValue(
            "$file_name",
            Path.GetFileName(sourcePath));
        using var reader = command.ExecuteReader();
        if (!reader.Read())
        {
            return new WorkbookComparisonSummary(
                benchmarkNumber,
                string.Empty,
                Path.GetFileName(sourcePath),
                string.Empty,
                "DB 없음",
                sourcePath,
                0,
                0,
                false);
        }

        return new WorkbookComparisonSummary(
            benchmarkNumber,
            reader.GetString(0),
            Path.GetFileName(reader.GetString(3)),
            reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
            reader.GetString(2),
            reader.GetString(3),
            reader.GetInt32(4),
            reader.GetInt32(5),
            true);
    }

    private static IReadOnlyList<WorkbookComparisonSummary>
        ListLatestAnalyses(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                wa.public_analysis_id,
                wa.title,
                wa.verification_status,
                sd.source_path,
                (
                    SELECT COUNT(*)
                    FROM knowledge_studies s
                    WHERE s.workbook_analysis_id=wa.workbook_analysis_id
                ) AS study_count,
                (
                    SELECT COUNT(*)
                    FROM knowledge_observations v
                    JOIN knowledge_outcomes o
                      ON o.outcome_id=v.outcome_id
                    JOIN knowledge_studies s
                      ON s.study_id=o.study_id
                    WHERE s.workbook_analysis_id=wa.workbook_analysis_id
                ) AS observation_count
            FROM workbook_analyses wa
            JOIN source_documents sd
              ON sd.document_id=wa.document_id
            ORDER BY wa.updated_at DESC
            LIMIT 30;
            """;
        using var reader = command.ExecuteReader();
        var result = new List<WorkbookComparisonSummary>();
        while (reader.Read())
        {
            result.Add(new WorkbookComparisonSummary(
                result.Count + 1,
                reader.GetString(0),
                Path.GetFileName(reader.GetString(3)),
                reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                reader.GetString(2),
                reader.GetString(3),
                reader.GetInt32(4),
                reader.GetInt32(5),
                true));
        }
        return result;
    }

    private static string ReadSourcePath(
        SqliteConnection connection,
        string publicAnalysisId)
    {
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT sd.source_path
            FROM workbook_analyses wa
            JOIN source_documents sd
              ON sd.document_id=wa.document_id
            WHERE wa.public_analysis_id=$analysis_id
            LIMIT 1;
            """;
        command.Parameters.AddWithValue(
            "$analysis_id",
            publicAnalysisId);
        return command.ExecuteScalar() as string
            ?? throw new InvalidOperationException(
                $"분석 ID를 찾을 수 없습니다: {publicAnalysisId}");
    }

    private static IReadOnlyList<StudyComparisonRow> ReadStudies(
        SqliteConnection connection,
        string publicAnalysisId)
    {
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                s.public_data_id,
                s.title,
                s.design_type,
                s.verification_status,
                s.comparability_status,
                s.confounding_status,
                (
                    SELECT COUNT(*)
                    FROM knowledge_arms a
                    WHERE a.study_id=s.study_id
                ) AS arm_count,
                (
                    SELECT COUNT(*)
                    FROM knowledge_outcomes o
                    WHERE o.study_id=s.study_id
                ) AS outcome_count,
                (
                    SELECT COUNT(*)
                    FROM knowledge_observations v
                    JOIN knowledge_outcomes o
                      ON o.outcome_id=v.outcome_id
                    WHERE o.study_id=s.study_id
                ) AS observation_count,
                COALESCE(e.sheet_name, ''),
                COALESCE(e.range_address, ''),
                COALESCE(e.source_text, '')
            FROM knowledge_studies s
            JOIN workbook_analyses wa
              ON wa.workbook_analysis_id=s.workbook_analysis_id
            LEFT JOIN entity_evidence_links l
              ON l.entity_evidence_link_id=(
                  SELECT l2.entity_evidence_link_id
                  FROM entity_evidence_links l2
                  WHERE l2.entity_type='STUDY'
                    AND l2.entity_uid=s.study_uid
                  ORDER BY
                      CASE WHEN l2.evidence_role='SOURCE'
                           THEN 0 ELSE 1 END,
                      l2.entity_evidence_link_id
                  LIMIT 1
              )
            LEFT JOIN evidence_items e
              ON e.evidence_id=l.evidence_id
            WHERE wa.public_analysis_id=$analysis_id
            ORDER BY s.study_id;
            """;
        command.Parameters.AddWithValue(
            "$analysis_id",
            publicAnalysisId);
        using var reader = command.ExecuteReader();
        var result = new List<StudyComparisonRow>();
        while (reader.Read())
        {
            result.Add(new StudyComparisonRow(
                result.Count + 1,
                reader.GetString(0),
                reader.GetString(1),
                reader.GetString(2),
                reader.GetString(3),
                reader.GetString(4),
                reader.GetString(5),
                reader.GetInt32(6),
                reader.GetInt32(7),
                reader.GetInt32(8),
                BuildLocation(
                    reader.GetString(9),
                    reader.GetString(10)),
                reader.GetString(11)));
        }
        return result;
    }

    private static IReadOnlyList<ObservationComparisonRow>
        ReadObservationRows(
            SqliteConnection connection,
            string publicAnalysisId,
            string sourcePath)
    {
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                s.public_data_id,
                s.title,
                a.label,
                a.arm_role,
                o.original_label,
                o.metric_type,
                o.original_unit,
                COALESCE(u.canonical_symbol, ''),
                v.value_number,
                v.value_text,
                v.numerator,
                v.denominator,
                v.rate_ppm,
                v.min_value,
                v.max_value,
                v.average_value,
                v.result_status,
                v.verification_status,
                COALESCE(e.public_evidence_id, ''),
                COALESCE(e.sheet_name, ''),
                COALESCE(e.range_address, ''),
                COALESCE(e.source_text, ''),
                v.observation_uid
            FROM knowledge_observations v
            JOIN knowledge_outcomes o
              ON o.outcome_id=v.outcome_id
            JOIN knowledge_studies s
              ON s.study_id=o.study_id
            JOIN workbook_analyses wa
              ON wa.workbook_analysis_id=s.workbook_analysis_id
            JOIN knowledge_arms a
              ON a.arm_id=v.arm_id
            LEFT JOIN knowledge_units u
              ON u.unit_id=o.unit_id
            LEFT JOIN entity_evidence_links l
              ON l.entity_type='OBSERVATION'
             AND l.entity_uid=v.observation_uid
            LEFT JOIN evidence_items e
              ON e.evidence_id=l.evidence_id
            WHERE wa.public_analysis_id=$analysis_id
            ORDER BY
                s.study_id,
                v.observation_id,
                CASE WHEN l.evidence_role='SOURCE'
                     THEN 0 ELSE 1 END,
                l.entity_evidence_link_id;
            """;
        command.Parameters.AddWithValue(
            "$analysis_id",
            publicAnalysisId);
        using var reader = command.ExecuteReader();
        var candidatesByObservation =
            new Dictionary<string, List<ObservationComparisonRow>>();
        while (reader.Read())
        {
            var sourceValue = reader.GetString(21);
            var excelActualValue = ExtractExcelActualValue(sourceValue);
            var databaseValue = BuildDatabaseValue(
                reader,
                sourceValue);
            var unit = PickUnit(
                reader.GetString(6),
                reader.GetString(7));
            var sheet = reader.GetString(19);
            var range = reader.GetString(20);
            var sourceLocation = BuildLocation(sheet, range);
            var comparison = ClassifyComparison(
                sourceValue,
                excelActualValue,
                databaseValue,
                reader.GetString(5),
                reader.GetString(4),
                unit,
                sourceLocation,
                HasCompositeSummary(reader));
            var candidate = new ObservationComparisonRow(
                0,
                sourcePath,
                reader.GetString(0),
                reader.GetString(1),
                reader.GetString(2),
                reader.GetString(3),
                reader.GetString(4),
                reader.GetString(5),
                unit,
                databaseValue,
                reader.GetString(16),
                reader.GetString(17),
                reader.GetString(18),
                sheet,
                range,
                sourceValue,
                excelActualValue,
                comparison.Status,
                comparison.ProblemDescription);
            var observationUid = reader.GetString(22);
            if (!candidatesByObservation.TryGetValue(
                    observationUid,
                    out var candidates))
            {
                candidates = [];
                candidatesByObservation.Add(
                    observationUid,
                    candidates);
            }
            candidates.Add(candidate);
        }

        var result = new List<ObservationComparisonRow>(
            candidatesByObservation.Count);
        foreach (var candidates in candidatesByObservation.Values)
        {
            var selected = SelectBestEvidenceCandidate(candidates);
            result.Add(selected with { Number = result.Count + 1 });
        }
        return result;
    }

    private static ObservationComparisonRow SelectBestEvidenceCandidate(
        IReadOnlyList<ObservationComparisonRow> candidates)
    {
        var matched = candidates.FirstOrDefault(
            candidate => string.Equals(
                candidate.ComparisonStatus,
                "일치",
                StringComparison.Ordinal));
        if (matched is not null)
            return matched;

        return candidates.FirstOrDefault(
                   candidate => !string.IsNullOrWhiteSpace(
                       candidate.EvidenceId))
               ?? candidates[0];
    }

    private static string BuildDatabaseValue(
        SqliteDataReader reader,
        string sourceValue)
    {
        if (string.Equals(
                reader.GetString(5),
                "success_count",
                StringComparison.OrdinalIgnoreCase)
            && !reader.IsDBNull(10)
            && !reader.IsDBNull(11))
        {
            if (ShouldDisplaySuccessPair(sourceValue))
            {
                return $"{FormatNumber(reader.GetDouble(11))}; "
                    + $"{FormatNumber(reader.GetDouble(10))}";
            }
        }

        if (!reader.IsDBNull(13)
            || !reader.IsDBNull(14)
            || !reader.IsDBNull(15))
        {
            var values = new List<string>();
            if (!reader.IsDBNull(13))
                values.Add(
                    $"Min {FormatNumber(reader.GetDouble(13))}");
            if (!reader.IsDBNull(14))
                values.Add(
                    $"Max {FormatNumber(reader.GetDouble(14))}");
            if (!reader.IsDBNull(15))
                values.Add(
                    $"Average {FormatNumber(reader.GetDouble(15))}");
            return string.Join("; ", values);
        }
        var valueText = reader.GetString(9).Trim();
        if (!reader.IsDBNull(8))
        {
            var valueNumber = reader.GetDouble(8);
            if (valueText.Length > 0
                && TextRepresentsNumericValue(
                    valueText,
                    valueNumber))
                return valueText;
            return FormatNumber(valueNumber);
        }
        if (valueText.Length > 0)
            return valueText;
        if (!reader.IsDBNull(10) && !reader.IsDBNull(11))
            return $"{FormatNumber(reader.GetDouble(10))}"
                + $"/{FormatNumber(reader.GetDouble(11))}";
        if (!reader.IsDBNull(12))
            return $"{FormatNumber(reader.GetDouble(12))} ppm";
        return string.Empty;
    }

    private static bool TextRepresentsNumericValue(
        string valueText,
        double valueNumber)
    {
        var numbers = ExtractNumbers(valueText);
        return numbers.Any(
            number => NearlyEqual(number, valueNumber)
                || DisplayRoundedPercentMatches(
                    valueText,
                    number,
                    valueNumber));
    }

    private static bool HasCompositeSummary(
        SqliteDataReader reader)
    {
        var componentCount = 0;
        if (!reader.IsDBNull(13)) componentCount++;
        if (!reader.IsDBNull(14)) componentCount++;
        if (!reader.IsDBNull(15)) componentCount++;
        return componentCount >= 2;
    }

    private static ComparisonResult ClassifyComparison(
        string sourceValue,
        string excelActualValue,
        string databaseValue,
        string metricType,
        string outcomeLabel,
        string unit,
        string sourceLocation,
        bool databaseHasSummary)
    {
        if (string.IsNullOrWhiteSpace(sourceValue))
        {
            return new ComparisonResult(
                "근거 없음",
                $"Excel 원본 근거 없음 · {sourceLocation}");
        }

        if (databaseHasSummary)
        {
            if (!TryParseSummary(
                    databaseValue,
                    out var databaseSummary))
            {
                return new ComparisonResult(
                    "확인 필요",
                    $"DB 통계값 구조 오류 · {sourceLocation}");
            }

            if (TryParseSummary(
                    excelActualValue,
                    out var labeledSourceSummary))
            {
                return SummaryMatches(
                    labeledSourceSummary,
                    databaseSummary)
                        ? new ComparisonResult(
                            "일치",
                            $"문제 없음 · {sourceLocation}")
                        : new ComparisonResult(
                            "확인 필요",
                            $"Min/Max/Average 값 불일치 · {sourceLocation}");
            }

            var hasRawValues = TryParseNumericSequence(
                excelActualValue,
                out var rawValues);
            if (!hasRawValues)
            {
                rawValues = ExtractNumbers(excelActualValue);
                hasRawValues = rawValues.Count >= 3;
            }
            if (hasRawValues)
            {
                var databaseValuesInSourceOrder =
                    ExtractNumbers(databaseValue);
                if (rawValues.Count
                        == databaseValuesInSourceOrder.Count
                    && rawValues
                        .Zip(databaseValuesInSourceOrder)
                        .All(pair => NearlyEqual(
                            pair.First,
                            pair.Second)))
                {
                    return new ComparisonResult(
                        "일치",
                        $"문제 없음 · {sourceLocation}");
                }

                if (SummaryComponentsAppearIn(
                        rawValues,
                        databaseSummary))
                {
                    return new ComparisonResult(
                        "일치",
                        $"문제 없음 · {sourceLocation}");
                }

                if (rawValues.Count == 3
                    && SummaryMatches(
                        new NumericSummary(
                            rawValues[0],
                            rawValues[1],
                            rawValues[2]),
                        databaseSummary))
                {
                    return new ComparisonResult(
                        "일치",
                        $"문제 없음 · {sourceLocation}");
                }

                var calculatedSummary = new NumericSummary(
                    rawValues.Min(),
                    rawValues.Max(),
                    rawValues.Average());
                return SummaryMatches(
                    calculatedSummary,
                    databaseSummary)
                        ? new ComparisonResult(
                            "일치",
                            $"문제 없음 · {sourceLocation}")
                        : new ComparisonResult(
                            "확인 필요",
                            $"Raw 데이터 자동 계산값 불일치 · {sourceLocation}");
            }

            return new ComparisonResult(
                "확인 필요",
                $"원본 근거 위치 오류: 값 셀이 아닌 헤더 연결 · {sourceLocation}");
        }

        if (TryMatchOutcomeValue(
                excelActualValue,
                databaseValue,
                outcomeLabel))
        {
            return new ComparisonResult(
                "일치",
                $"문제 없음 · {sourceLocation}");
        }

        if (TryParseNumericSequence(
                excelActualValue,
                out var sourceValues)
            && sourceValues.Count > 1
            && TryParseNumericSequence(
                databaseValue,
                out var databaseValues)
            && sourceValues.Count == databaseValues.Count)
        {
            return sourceValues
                .Zip(databaseValues)
                .All(pair => NumericValuesMatch(
                    pair.First,
                    pair.Second,
                    unit))
                    ? new ComparisonResult(
                        "일치",
                        $"문제 없음 · {sourceLocation}")
                    : new ComparisonResult(
                        "확인 필요",
                        $"다중 셀 값 불일치 · {sourceLocation}");
        }

        if (TryParseNumericSequence(
                excelActualValue,
                out sourceValues)
            && sourceValues.Count > 1)
        {
            var databaseNumbers = ExtractNumbers(
                databaseValue);
            if (sourceValues.Count == databaseNumbers.Count
                && sourceValues
                    .Order()
                    .Zip(databaseNumbers.Order())
                    .All(pair => NumericValuesMatch(
                        pair.First,
                        pair.Second,
                        unit)))
            {
                return new ComparisonResult(
                    "일치",
                    $"문제 없음 · {sourceLocation}");
            }
        }

        if (TryParseNumericSequence(
                excelActualValue,
                out sourceValues)
            && sourceValues.Count > 1
            && TryParseDatabaseNumber(
                databaseValue,
                out var componentDatabaseValue)
            && sourceValues.Any(
                sourceComponent => NumericValuesMatch(
                    sourceComponent,
                    componentDatabaseValue,
                    unit)))
        {
            return new ComparisonResult(
                "일치",
                $"문제 없음 · {sourceLocation}");
        }

        if (TryMatchFractionEvidence(
                excelActualValue,
                databaseValue,
                unit))
        {
            return new ComparisonResult(
                "일치",
                $"문제 없음 · {sourceLocation}");
        }

        var embeddedSourceNumbers =
            ExtractNumbers(excelActualValue);
        if (embeddedSourceNumbers.Count > 0
            && TryParseDatabaseNumber(
                databaseValue,
                out var embeddedDatabaseValue)
            && embeddedSourceNumbers.Any(
                embeddedSourceNumber => NumericValuesMatch(
                    embeddedSourceNumber,
                    embeddedDatabaseValue,
                    unit,
                    excelActualValue)))
        {
            return new ComparisonResult(
                "일치",
                $"문제 없음 · {sourceLocation}");
        }

        if (TryParseSourceNumber(
                sourceValue,
                excelActualValue,
                out var sourceNumber)
            && TryParseDatabaseNumber(
                databaseValue,
                out var databaseNumber))
        {
            return NumericValuesMatch(
                sourceNumber,
                databaseNumber,
                unit,
                excelActualValue)
                ? new ComparisonResult(
                    "일치",
                    $"문제 없음 · {sourceLocation}")
                : new ComparisonResult(
                    "확인 필요",
                    $"값 불일치 · {sourceLocation}");
        }

        var normalizedExcelValue =
            NormalizeText(excelActualValue);
        var normalizedDatabaseValue =
            NormalizeText(databaseValue);
        return string.Equals(
            normalizedExcelValue,
            normalizedDatabaseValue,
            StringComparison.OrdinalIgnoreCase)
            || normalizedExcelValue.StartsWith(
                normalizedDatabaseValue
                + " in each supplied cell",
                StringComparison.OrdinalIgnoreCase)
                ? new ComparisonResult(
                    "일치",
                    $"문제 없음 · {sourceLocation}")
                : new ComparisonResult(
                    "확인 필요",
                    $"값 불일치 · {sourceLocation}");
    }

    private static bool ShouldDisplaySuccessPair(
        string sourceValue)
    {
        if (sourceValue.Contains(
                "Input",
                StringComparison.OrdinalIgnoreCase)
            && sourceValue.Contains(
                "OK",
                StringComparison.OrdinalIgnoreCase))
            return true;
        return TryParseNumericSequence(
                sourceValue,
                out var values)
            && values.Count == 2;
    }

    private static bool TryMatchOutcomeValue(
        string excelActualValue,
        string databaseValue,
        string outcomeLabel)
    {
        if (excelActualValue.Contains(
                "Input",
                StringComparison.OrdinalIgnoreCase)
            && excelActualValue.Contains(
                "OK",
                StringComparison.OrdinalIgnoreCase)
            && TryExtractNamedNumber(
                excelActualValue,
                "Input",
                out var input)
            && TryExtractNamedNumber(
                excelActualValue,
                "OK",
                out var ok)
            && TryParseNumericSequence(
                databaseValue,
                out var databasePair)
            && databasePair.Count == 2)
        {
            return NearlyEqual(input, databasePair[0])
                && NearlyEqual(ok, databasePair[1]);
        }

        var parts = excelActualValue.Split(
            ';',
            StringSplitOptions.RemoveEmptyEntries
            | StringSplitOptions.TrimEntries);
        if (parts.Length != 2
            || !TryParseNumber(parts[1], out var sourceNumber)
            || !TryParseDatabaseNumber(
                databaseValue,
                out var databaseNumber))
            return false;

        var normalizedLabel = NormalizeText(outcomeLabel);
        var normalizedSourceLabel = NormalizeText(parts[0]);
        return normalizedLabel.Length > 0
            && (string.Equals(
                    normalizedSourceLabel,
                    normalizedLabel,
                    StringComparison.OrdinalIgnoreCase)
                || normalizedSourceLabel.Contains(
                    normalizedLabel,
                    StringComparison.OrdinalIgnoreCase))
            && NearlyEqual(sourceNumber, databaseNumber);
    }

    private static bool TryExtractNamedNumber(
        string value,
        string label,
        out double number)
    {
        var match = Regex.Match(
            value,
            $@"\b{Regex.Escape(label)}\s*[:=]?\s*"
            + @"([-+]?(?:\d+(?:,\d{3})*|\d*)"
            + @"(?:\.\d+)?(?:[Ee][-+]?\d+)?)",
            RegexOptions.IgnoreCase
            | RegexOptions.CultureInvariant);
        if (match.Success)
            return TryParseNumber(
                match.Groups[1].Value,
                out number);
        number = default;
        return false;
    }

    private static IReadOnlyList<double> ExtractNumbers(
        string value)
    {
        var matches = Regex.Matches(
            value,
            @"[-+]?(?:\d+(?:,\d{3})*|\d*)"
            + @"(?:\.\d+)?(?:[Ee][-+]?\d+)?",
            RegexOptions.CultureInvariant);
        var result = new List<double>(matches.Count);
        foreach (Match match in matches)
        {
            if (match.Value.Length > 0
                && TryParseNumber(
                    match.Value,
                    out var number))
                result.Add(number);
        }
        return result;
    }

    private static bool TryParseSummary(
        string value,
        out NumericSummary summary)
    {
        if (TryParseLabeledNumber(
                value,
                "Min",
                out var minimum)
            && TryParseLabeledNumber(
                value,
                "Max",
                out var maximum)
            && TryParseLabeledNumber(
                value,
                "Average",
                out var average))
        {
            summary = new NumericSummary(
                minimum,
                maximum,
                average);
            return true;
        }

        summary = default;
        return false;
    }

    private static bool TryParseLabeledNumber(
        string value,
        string label,
        out double number)
    {
        var match = Regex.Match(
            value,
            $@"\b{Regex.Escape(label)}\s*[:=]?\s*"
            + @"([-+]?(?:\d+(?:,\d{3})*|\d*)"
            + @"(?:\.\d+)?(?:[Ee][-+]?\d+)?)",
            RegexOptions.IgnoreCase
            | RegexOptions.CultureInvariant);
        if (match.Success)
            return TryParseNumber(
                match.Groups[1].Value,
                out number);
        number = default;
        return false;
    }

    private static bool TryParseNumericSequence(
        string value,
        out IReadOnlyList<double> numbers)
    {
        var parts = value.Split(
            ';',
            StringSplitOptions.RemoveEmptyEntries
            | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
        {
            numbers = [];
            return false;
        }

        var result = new List<double>(parts.Length);
        foreach (var part in parts)
        {
            if (!TryParseNumber(part, out var number))
            {
                numbers = [];
                return false;
            }
            result.Add(number);
        }
        numbers = result;
        return true;
    }

    private static bool TryParseSourceNumber(
        string sourceValue,
        string excelActualValue,
        out double number)
    {
        const string numberFormatMarker = "; number format ";
        var markerIndex = sourceValue.IndexOf(
            numberFormatMarker,
            StringComparison.OrdinalIgnoreCase);
        if (markerIndex > 0
            && TryParseNumber(
                sourceValue[..markerIndex],
                out number))
        {
            var numberFormat = sourceValue[
                (markerIndex + numberFormatMarker.Length)..];
            if (numberFormat.Contains(
                    '%',
                    StringComparison.Ordinal))
                number *= 100;
            return true;
        }
        return TryParseNumber(excelActualValue, out number);
    }

    private static bool TryParseDatabaseNumber(
        string databaseValue,
        out double number)
    {
        if (TryParseNumber(databaseValue, out number))
            return true;
        var embeddedNumbers = ExtractNumbers(databaseValue);
        if (embeddedNumbers.Count == 1)
        {
            number = embeddedNumbers[0];
            return true;
        }
        number = default;
        return false;
    }

    private static bool NearlyEqual(double left, double right)
    {
        var tolerance = Math.Max(
            1e-9,
            Math.Max(
                Math.Abs(left),
                Math.Abs(right)) * 1e-9);
        return Math.Abs(left - right) <= tolerance;
    }

    private static bool NumericValuesMatch(
        double sourceValue,
        double databaseValue,
        string unit,
        string sourceDisplay = "")
    {
        if (NearlyEqual(sourceValue, databaseValue))
            return true;
        if (DisplayRoundedNumberMatches(
                sourceDisplay,
                sourceValue,
                databaseValue))
            return true;
        if (!unit.Contains(
                '%',
                StringComparison.Ordinal))
            return false;
        return NearlyEqual(sourceValue * 100, databaseValue)
            || NearlyEqual(sourceValue, databaseValue * 100)
            || DisplayRoundedPercentMatches(
                sourceDisplay,
                sourceValue,
                databaseValue);
    }

    private static bool DisplayRoundedNumberMatches(
        string sourceDisplay,
        double sourceValue,
        double databaseValue)
    {
        var matches = Regex.Matches(
            sourceDisplay,
            @"[-+]?(?:\d+(?:,\d{3})*|\d*)"
            + @"(?:\.(\d+))?(?:[Ee][-+]?\d+)?",
            RegexOptions.CultureInvariant);
        foreach (Match match in matches)
        {
            if (match.Value.Length == 0
                || !TryParseNumber(
                    match.Value,
                    out var displayedValue)
                || !NearlyEqual(
                    displayedValue,
                    sourceValue))
                continue;
            var decimalPlaces =
                match.Groups[1].Success
                    ? match.Groups[1].Value.Length
                    : 0;
            var tolerance =
                0.5 * Math.Pow(10, -decimalPlaces)
                + 1e-12;
            if (Math.Abs(
                    databaseValue - displayedValue)
                <= tolerance)
                return true;
        }
        return false;
    }

    private static bool DisplayRoundedPercentMatches(
        string sourceDisplay,
        double sourceValue,
        double databaseValue)
    {
        var matches = Regex.Matches(
            sourceDisplay,
            @"([-+]?\d+(?:\.(\d+))?)\s*%",
            RegexOptions.CultureInvariant);
        foreach (Match match in matches)
        {
            if (!TryParseNumber(
                    match.Groups[1].Value,
                    out var displayedValue)
                || !NearlyEqual(
                    displayedValue,
                    sourceValue))
                continue;
            var decimalPlaces =
                match.Groups[2].Success
                    ? match.Groups[2].Value.Length
                    : 0;
            var tolerance =
                0.5 * Math.Pow(10, -decimalPlaces)
                + 1e-12;
            if (Math.Abs(
                    databaseValue - displayedValue)
                <= tolerance)
                return true;
        }
        return false;
    }

    private static bool TryMatchFractionEvidence(
        string excelActualValue,
        string databaseValue,
        string unit)
    {
        var databaseParts = databaseValue.Split(
            '/',
            StringSplitOptions.TrimEntries);
        if (databaseParts.Length != 2
            || !TryParseNumber(
                databaseParts[0],
                out var numerator))
            return false;

        var sourceFirstPart = excelActualValue.Split(
            ';',
            2,
            StringSplitOptions.TrimEntries)[0];
        return TryParseNumber(
                sourceFirstPart,
                out var sourceNumber)
            && NumericValuesMatch(
                sourceNumber,
                numerator,
                unit,
                sourceFirstPart);
    }

    private static bool SummaryMatches(
        NumericSummary left,
        NumericSummary right) =>
        NearlyEqual(left.Minimum, right.Minimum)
        && NearlyEqual(left.Maximum, right.Maximum)
        && NearlyEqual(left.Average, right.Average);

    private static bool SummaryComponentsAppearIn(
        IReadOnlyList<double> sourceValues,
        NumericSummary summary)
    {
        var expected = new[]
        {
            summary.Minimum,
            summary.Maximum,
            summary.Average,
        };
        return expected.All(
            expectedValue => sourceValues.Any(
                sourceValue => NearlyEqual(
                    sourceValue,
                    expectedValue)));
    }

    private static string ExtractExcelActualValue(string sourceValue)
    {
        var value = sourceValue.Trim();
        const string displayedWithMarker = "; displayed with ";
        var displayedWithIndex = value.IndexOf(
            displayedWithMarker,
            StringComparison.OrdinalIgnoreCase);
        if (displayedWithIndex > 0
            && value.EndsWith(
                " format",
                StringComparison.OrdinalIgnoreCase))
            return value[..displayedWithIndex].Trim();

        const string numberFormatMarker = "; number format ";
        var numberFormatIndex = value.IndexOf(
            numberFormatMarker,
            StringComparison.OrdinalIgnoreCase);
        if (numberFormatIndex > 0)
        {
            var rawValue = value[..numberFormatIndex].Trim();
            var numberFormat = value[
                (numberFormatIndex + numberFormatMarker.Length)..];
            if (numberFormat.Contains(
                    '%',
                    StringComparison.Ordinal)
                && TryParseNumber(rawValue, out var rawNumber))
            {
                return $"{rawValue} (= "
                    + $"{FormatNumber(rawNumber * 100)}%)";
            }
            return rawValue;
        }
        return value;
    }

    private static string CompareValues(
        string sourceValue,
        string databaseValue)
    {
        if (string.IsNullOrWhiteSpace(sourceValue))
            return "근거 없음";
        if (string.IsNullOrWhiteSpace(databaseValue))
            return "확인 필요";

        if (TryParseNumber(sourceValue, out var sourceNumber)
            && TryParseNumber(databaseValue, out var databaseNumber))
        {
            return NearlyEqual(sourceNumber, databaseNumber)
                ? "일치"
                : "확인 필요";
        }

        return string.Equals(
            NormalizeText(sourceValue),
            NormalizeText(databaseValue),
            StringComparison.OrdinalIgnoreCase)
                ? "일치"
                : "확인 필요";
    }

    private static bool TryParseNumber(
        string value,
        out double number)
    {
        var normalized = value
            .Trim()
            .Replace(",", string.Empty, StringComparison.Ordinal)
            .Replace("%", string.Empty, StringComparison.Ordinal);
        return double.TryParse(
            normalized,
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out number);
    }

    private static string NormalizeText(string value) =>
        string.Join(
            " ",
            value.Split(
                (char[]?)null,
                StringSplitOptions.RemoveEmptyEntries
                | StringSplitOptions.TrimEntries));

    private static string FormatNumber(double value) =>
        value.ToString("G15", CultureInfo.InvariantCulture);

    private static string PickUnit(
        string originalUnit,
        string canonicalUnit) =>
        !string.IsNullOrWhiteSpace(originalUnit)
            ? originalUnit
            : canonicalUnit;

    private static string BuildLocation(
        string sheet,
        string range) =>
        string.IsNullOrWhiteSpace(sheet)
            ? string.Empty
            : $"{sheet}!{range}";

    private readonly record struct NumericSummary(
        double Minimum,
        double Maximum,
        double Average);

    private sealed record ComparisonResult(
        string Status,
        string ProblemDescription);
}

internal sealed record WorkbookComparisonSummary(
    int BenchmarkNumber,
    string PublicAnalysisId,
    string FileName,
    string Title,
    string VerificationStatus,
    string SourcePath,
    int StudyCount,
    int ObservationCount,
    bool IsAvailable);

internal sealed record WorkbookComparisonDocument(
    string PublicAnalysisId,
    string SourcePath,
    IReadOnlyList<StudyComparisonRow> Studies,
    IReadOnlyList<ObservationComparisonRow> Observations);

internal sealed record StudyComparisonRow(
    int Number,
    string PublicDataId,
    string Title,
    string DesignType,
    string VerificationStatus,
    string ComparabilityStatus,
    string ConfoundingStatus,
    int ArmCount,
    int OutcomeCount,
    int ObservationCount,
    string SourceLocation,
    string SourceText);

internal sealed record ObservationComparisonRow(
    int Number,
    string SourcePath,
    string PublicDataId,
    string StudyTitle,
    string ArmLabel,
    string ArmRole,
    string OutcomeLabel,
    string MetricType,
    string Unit,
    string DatabaseValue,
    string ResultStatus,
    string VerificationStatus,
    string EvidenceId,
    string Sheet,
    string Range,
    string SourceValue,
    string ExcelActualValue,
    string ComparisonStatus,
    string ProblemDescription)
{
    public string SourceLocation =>
        string.IsNullOrWhiteSpace(Sheet)
            ? string.Empty
            : $"{Sheet}!{Range}";

    public string SearchText =>
        $"{PublicDataId} {StudyTitle} {ArmLabel} {ArmRole} "
        + $"{OutcomeLabel} {Unit} {DatabaseValue} {ResultStatus} "
        + $"{EvidenceId} {Sheet} {Range} {SourceValue} "
        + $"{ComparisonStatus}";
}
