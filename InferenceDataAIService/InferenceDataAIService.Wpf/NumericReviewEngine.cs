using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace InferenceDataAIService.Wpf;

// Batch-local numeric review. This consumes only numeric-capture-v1 facts and
// creates review-v1 rows. It does not reopen Excel, use COM, recalculate a
// formula, generate a narrative, or make a quality/release decision.
internal sealed class NumericReviewEngine
{
    private const string ReviewVersion = "numeric-review-v3";
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };
    private static readonly Regex NormalRole = new(@"(?<![a-z])normal(?![a-z])", RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex TestRole = new(@"(?<![a-z])test(?![a-z])", RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex EquipmentLabelPattern = new(@"\bline\s+[a-z0-9_-]+\b", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly Action<string>? _log;
    private NumericReviewEngine(Action<string>? log) => _log = log;

    internal static NumericReviewRunResult Run(NumericReviewRequest request, Action<string>? log = null, CancellationToken cancellationToken = default) =>
        new NumericReviewEngine(log).RunCore(request, cancellationToken);

    private NumericReviewRunResult RunCore(NumericReviewRequest request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.ServiceDirectory) || !Directory.Exists(request.ServiceDirectory))
            throw new ArgumentException($"서비스 폴더를 찾을 수 없습니다: {request.ServiceDirectory}");
        var batchDirectory = ResolveBatchDirectory(Path.GetFullPath(request.ServiceDirectory), request.StructureBatchId);
        var database = Path.Combine(batchDirectory, "numeric-capture.sqlite");
        if (!File.Exists(database)) throw new FileNotFoundException("숫자 표 원본 DB를 찾을 수 없습니다.", database);

        using var connection = new SqliteConnection($"Data Source={database}");
        connection.Open();
        Execute(connection, null, "PRAGMA foreign_keys=ON;");
        Execute(connection, null, "PRAGMA journal_mode=WAL;");
        EnsureSchema(connection);
        cancellationToken.ThrowIfCancellationRequested();
        _log?.Invoke($"숫자 검토 DB 변환: {request.StructureBatchId}");
        using (var transaction = connection.BeginTransaction())
        {
            RebuildReviews(connection, transaction, cancellationToken);
            transaction.Commit();
        }
        var summary = WriteOutputs(batchDirectory, connection);
        _log?.Invoke($"숫자 검토 완료: defect={summary.DefectRateFactCount}, repeatedBlock={summary.RepeatedDefectBlockFactCount}, measurement={summary.MeasurementSummaryFactCount}, comparisons={summary.TestNormalComparisonCount + summary.RepeatedBlockTestNormalComparisonCount}");
        return new NumericReviewRunResult(batchDirectory, summary);
    }

    private static string ResolveBatchDirectory(string serviceDirectory, string batchId)
    {
        if (string.IsNullOrWhiteSpace(batchId) || batchId.Length > 96 || batchId is "." or ".." ||
            batchId.Any(value => !(char.IsAsciiLetterOrDigit(value) || value is '.' or '_' or '-')))
            throw new ArgumentException("유효한 구조 배치 ID가 아닙니다.");
        var root = Path.GetFullPath(
            AppRuntimePaths.Current.BatchRootDirectory);
        var result = Path.GetFullPath(Path.Combine(root, batchId));
        var relative = Path.GetRelativePath(root, result);
        if (relative == ".." || relative.StartsWith($"..{Path.DirectorySeparatorChar}", StringComparison.Ordinal) || Path.IsPathRooted(relative))
            throw new ArgumentException("배치 경로가 출력 루트 밖을 가리킵니다.");
        return result;
    }

    private static void EnsureSchema(SqliteConnection connection)
    {
        Execute(connection, null, """
            CREATE TABLE IF NOT EXISTS numeric_review_metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS numeric_table_reviews (
                table_id INTEGER PRIMARY KEY REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
                profile_name TEXT NOT NULL, extraction_status TEXT NOT NULL, reason_code TEXT NOT NULL DEFAULT '',
                extracted_fact_count INTEGER NOT NULL DEFAULT 0, updated_at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_numeric_table_reviews_status ON numeric_table_reviews(profile_name, extraction_status);
            CREATE TABLE IF NOT EXISTS numeric_review_facts (
                fact_id INTEGER PRIMARY KEY,
                workbook_id INTEGER NOT NULL REFERENCES capture_workbooks(workbook_id) ON DELETE CASCADE,
                sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
                table_id INTEGER NOT NULL REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
                row_index INTEGER NOT NULL, measurement_date TEXT NOT NULL DEFAULT '', metric_type TEXT NOT NULL,
                condition_role TEXT NOT NULL, condition_label TEXT NOT NULL,
                input_value_text TEXT NOT NULL, input_value REAL, total_ng_value_text TEXT NOT NULL, total_ng_value REAL,
                reported_ng_rate_text TEXT NOT NULL, reported_ng_rate REAL, computed_ng_rate REAL,
                fact_status TEXT NOT NULL, validation_code TEXT NOT NULL DEFAULT '',
                date_row_index INTEGER, date_column_index INTEGER, condition_column_index INTEGER,
                input_column_index INTEGER, total_ng_column_index INTEGER, rate_column_index INTEGER,
                UNIQUE(table_id, row_index, metric_type)
            );
            CREATE INDEX IF NOT EXISTS idx_numeric_review_facts_pair ON numeric_review_facts(table_id, measurement_date, condition_role, metric_type);
            CREATE TABLE IF NOT EXISTS repeated_defect_block_facts (
                fact_id INTEGER PRIMARY KEY,
                workbook_id INTEGER NOT NULL REFERENCES capture_workbooks(workbook_id) ON DELETE CASCADE,
                sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
                table_id INTEGER NOT NULL REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
                block_key TEXT NOT NULL, block_label TEXT NOT NULL,
                row_index INTEGER NOT NULL, measurement_date TEXT NOT NULL DEFAULT '',
                condition_role TEXT NOT NULL, condition_label TEXT NOT NULL,
                input_value_text TEXT NOT NULL, input_value REAL,
                total_ng_value_text TEXT NOT NULL, total_ng_value REAL,
                reported_ng_rate_text TEXT NOT NULL, reported_ng_rate REAL, computed_ng_rate REAL,
                fact_status TEXT NOT NULL, validation_code TEXT NOT NULL DEFAULT '',
                date_row_index INTEGER, date_column_index INTEGER, condition_column_index INTEGER,
                input_column_index INTEGER, total_ng_column_index INTEGER, rate_column_index INTEGER,
                UNIQUE(table_id, block_key, row_index)
            );
            CREATE INDEX IF NOT EXISTS idx_repeated_defect_block_facts_pair ON repeated_defect_block_facts(table_id, block_key, measurement_date, condition_role);
            CREATE TABLE IF NOT EXISTS repeated_block_equipment_tables (
                table_id INTEGER PRIMARY KEY REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
                equipment_label TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS measurement_summary_facts (
                fact_id INTEGER PRIMARY KEY,
                workbook_id INTEGER NOT NULL REFERENCES capture_workbooks(workbook_id) ON DELETE CASCADE,
                sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
                table_id INTEGER NOT NULL REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
                row_index INTEGER NOT NULL, measurement_date TEXT NOT NULL DEFAULT '', condition_label TEXT NOT NULL,
                sample_value_text TEXT NOT NULL, sample_value REAL, average_value_text TEXT NOT NULL, average_value REAL,
                minimum_value_text TEXT NOT NULL, minimum_value REAL, maximum_value_text TEXT NOT NULL, maximum_value REAL,
                fact_status TEXT NOT NULL, validation_code TEXT NOT NULL DEFAULT '',
                date_row_index INTEGER, date_column_index INTEGER, sample_column_index INTEGER, average_column_index INTEGER,
                minimum_column_index INTEGER, maximum_column_index INTEGER, UNIQUE(table_id, row_index)
            );
            CREATE INDEX IF NOT EXISTS idx_measurement_summary_facts_workbook ON measurement_summary_facts(workbook_id, measurement_date);
            CREATE TABLE IF NOT EXISTS test_normal_comparisons (
                comparison_id INTEGER PRIMARY KEY,
                workbook_id INTEGER NOT NULL REFERENCES capture_workbooks(workbook_id) ON DELETE CASCADE,
                sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
                table_id INTEGER NOT NULL REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
                measurement_date TEXT NOT NULL, metric_type TEXT NOT NULL,
                test_fact_id INTEGER REFERENCES numeric_review_facts(fact_id) ON DELETE SET NULL,
                normal_fact_id INTEGER REFERENCES numeric_review_facts(fact_id) ON DELETE SET NULL,
                test_ng_rate REAL, normal_ng_rate REAL, absolute_delta REAL, relative_ratio REAL,
                comparison_status TEXT NOT NULL, validation_code TEXT NOT NULL DEFAULT '',
                UNIQUE(table_id, measurement_date, metric_type, test_fact_id)
            );
            CREATE INDEX IF NOT EXISTS idx_test_normal_comparisons_status ON test_normal_comparisons(comparison_status, measurement_date);
            CREATE TABLE IF NOT EXISTS repeated_block_test_normal_comparisons (
                comparison_id INTEGER PRIMARY KEY,
                workbook_id INTEGER NOT NULL REFERENCES capture_workbooks(workbook_id) ON DELETE CASCADE,
                sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
                table_id INTEGER NOT NULL REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
                block_key TEXT NOT NULL, block_label TEXT NOT NULL, measurement_date TEXT NOT NULL,
                test_fact_id INTEGER REFERENCES repeated_defect_block_facts(fact_id) ON DELETE SET NULL,
                normal_fact_id INTEGER REFERENCES repeated_defect_block_facts(fact_id) ON DELETE SET NULL,
                test_ng_rate REAL, normal_ng_rate REAL, absolute_delta REAL, relative_ratio REAL,
                comparison_status TEXT NOT NULL, validation_code TEXT NOT NULL DEFAULT '',
                UNIQUE(table_id, block_key, measurement_date, test_fact_id)
            );
            CREATE INDEX IF NOT EXISTS idx_repeated_block_comparisons_status ON repeated_block_test_normal_comparisons(comparison_status, measurement_date);
            """);
        Execute(connection, null, "INSERT OR REPLACE INTO numeric_review_metadata(key, value) VALUES ('schemaVersion', 'numeric-review-db-v3');");
        Execute(connection, null, "INSERT OR REPLACE INTO numeric_review_metadata(key, value) VALUES ('reviewVersion', $version);", ("$version", ReviewVersion));
    }

    private void RebuildReviews(SqliteConnection connection, SqliteTransaction transaction, CancellationToken cancellationToken)
    {
        Execute(connection, transaction, "DELETE FROM test_normal_comparisons;");
        Execute(connection, transaction, "DELETE FROM repeated_block_test_normal_comparisons;");
        Execute(connection, transaction, "DELETE FROM repeated_block_equipment_tables;");
        Execute(connection, transaction, "DELETE FROM numeric_review_facts;");
        Execute(connection, transaction, "DELETE FROM repeated_defect_block_facts;");
        Execute(connection, transaction, "DELETE FROM measurement_summary_facts;");
        Execute(connection, transaction, "DELETE FROM numeric_table_reviews;");
        var tables = ReadTables(connection, transaction);
        foreach (var table in tables)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var headers = ReadLabels(connection, transaction, table.TableId, "HEADER");
            if (RepeatedDefectBlocks(headers) is not null) ExtractRepeatedDefectBlocks(connection, transaction, table, headers);
            else if (DefectColumns(headers) is not null) ExtractDefectRateTable(connection, transaction, table, headers);
            else if (MeasurementColumns(headers) is not null) ExtractMeasurementSummaryTable(connection, transaction, table, headers);
            else InsertTableReview(connection, transaction, table.TableId, table.CandidateType, "NOT_IMPLEMENTED", "PROFILE_PENDING", 0);
        }
        CreateComparisons(connection, transaction);
        CreateRepeatedBlockComparisons(connection, transaction);
    }

    private static List<TableCandidate> ReadTables(SqliteConnection connection, SqliteTransaction transaction)
    {
        using var command = Command(connection, transaction, """
            SELECT candidate.table_id, candidate.sheet_id, sheet.workbook_id, candidate.start_row, candidate.end_row,
                   candidate.start_column, candidate.end_column, candidate.candidate_type
            FROM numeric_table_candidates AS candidate
            JOIN captured_sheets AS sheet ON sheet.sheet_id=candidate.sheet_id
            JOIN capture_workbooks AS workbook ON workbook.workbook_id=sheet.workbook_id
            WHERE workbook.capture_status='CAPTURED' ORDER BY candidate.table_id;
            """);
        using var reader = command.ExecuteReader();
        var values = new List<TableCandidate>();
        while (reader.Read()) values.Add(new TableCandidate(reader.GetInt64(0), reader.GetInt64(1), reader.GetInt64(2), reader.GetInt32(3), reader.GetInt32(4), reader.GetInt32(5), reader.GetInt32(6), reader.GetString(7)));
        return values;
    }

    private static List<Label> ReadLabels(SqliteConnection connection, SqliteTransaction transaction, long tableId, string role)
    {
        using var command = Command(connection, transaction, "SELECT row_index, column_index, label_text FROM numeric_table_labels WHERE table_id=$tableId AND label_role=$role ORDER BY row_index, column_index;", ("$tableId", tableId), ("$role", role));
        using var reader = command.ExecuteReader();
        var values = new List<Label>();
        while (reader.Read()) values.Add(new Label(reader.GetInt32(0), reader.GetInt32(1), reader.GetString(2)));
        return values;
    }

    // Repeated horizontal defect blocks cannot be safely represented by the
    // single-table defect fact key.  Each explicit Input → Total NG → NG Rate
    // block is retained as an independent observation and comparison scope.
    private static void ExtractRepeatedDefectBlocks(SqliteConnection connection, SqliteTransaction transaction, TableCandidate table, IReadOnlyList<Label> headers)
    {
        var layout = RepeatedDefectBlocks(headers);
        if (layout is null)
        {
            InsertTableReview(connection, transaction, table.TableId, "REPEATED_DEFECT_BLOCKS", "NEEDS_REVIEW", "REPEATED_BLOCK_HEADER_AMBIGUOUS_OR_MISSING", 0);
            return;
        }
        var values = ReadNumericCells(connection, transaction, table);
        var rowLabels = ReadLabels(connection, transaction, table.TableId, "ROW_LABEL");
        var equipmentLabel = EquipmentLabel(rowLabels);
        var labelsByRow = rowLabels
            .GroupBy(value => value.Row)
            .ToDictionary(group => group.Key, group => group.ToList());
        var dates = RowDates(connection, transaction, table);
        var factCount = 0;
        foreach (var block in layout.Blocks)
        {
            for (var row = table.StartRow; row <= table.EndRow; row++)
            {
                var input = ValueForColumn(values, row, block.Input);
                var total = ValueForColumn(values, row, block.TotalNg);
                var rate = ValueForColumn(values, row, block.Rate);
                if (input.Value is null && total.Value is null && rate.Value is null) continue;
                var labels = labelsByRow.GetValueOrDefault(row, []);
                var roleLabel = labels.Select(label => (Role: ExplicitRole(label.Text), Label: label)).FirstOrDefault(value => value.Role is not null);
                var role = roleLabel.Role ?? "UNSPECIFIED";
                var conditionLabel = roleLabel.Role is null ? RowLabelText(labels) : roleLabel.Label.Text;
                DateInfo? date = dates.TryGetValue(row, out var resolvedDate) ? resolvedDate : null;
                var validation = new List<string>();
                if (roleLabel.Role is null && equipmentLabel is null) validation.Add("CONDITION_ROLE_MISSING");
                if (date is null) validation.Add("DATE_MISSING");
                if (input.Value is null || total.Value is null) validation.Add("INPUT_OR_TOTAL_NG_MISSING");
                else if (input.Value <= 0) validation.Add("INPUT_NOT_POSITIVE");
                else if (total.Value < 0) validation.Add("TOTAL_NG_NEGATIVE");
                var usable = input.Value is > 0 && total.Value is >= 0;
                var computed = usable ? total.Value!.Value / input.Value!.Value : (double?)null;
                InsertRepeatedBlockFact(connection, transaction, table, block, row, date, role, conditionLabel, input, total, rate, computed,
                    validation.Count == 0 ? "OBSERVED" : "NEEDS_REVIEW", string.Join(';', validation), roleLabel.Label?.Column ?? 0);
                factCount++;
            }
        }
        var reviewNeeded = layout.ReviewBlockCount > 0;
        if (equipmentLabel is not null)
            Execute(connection, transaction, "INSERT INTO repeated_block_equipment_tables(table_id, equipment_label) VALUES ($tableId, $equipment);", ("$tableId", table.TableId), ("$equipment", equipmentLabel));
        InsertTableReview(connection, transaction, table.TableId, "REPEATED_DEFECT_BLOCKS", factCount > 0 && !reviewNeeded ? "EXTRACTED" : "NEEDS_REVIEW",
            reviewNeeded ? "BLOCK_HEADER_AMBIGUOUS_OR_MISSING" : factCount > 0 ? string.Empty : "REPEATED_BLOCK_NUMERIC_ROW_NOT_FOUND", factCount);
    }

    private static void ExtractDefectRateTable(SqliteConnection connection, SqliteTransaction transaction, TableCandidate table, IReadOnlyList<Label> headers)
    {
        var columns = DefectColumns(headers);
        if (columns is null)
        {
            InsertTableReview(connection, transaction, table.TableId, "DEFECT_RATE", "NEEDS_REVIEW", "REQUIRED_HEADER_AMBIGUOUS_OR_MISSING", 0);
            return;
        }
        var values = ReadNumericCells(connection, transaction, table);
        var labelsByRow = ReadLabels(connection, transaction, table.TableId, "ROW_LABEL").GroupBy(value => value.Row).ToDictionary(group => group.Key, group => group.ToList());
        var dates = RowDates(connection, transaction, table);
        var factCount = 0;
        for (var row = table.StartRow; row <= table.EndRow; row++)
        {
            labelsByRow.TryGetValue(row, out var labels);
            labels ??= [];
            var roleLabel = labels.Select(label => (Role: ExplicitRole(label.Text), Label: label)).FirstOrDefault(value => value.Role is not null);
            var conditionLabel = roleLabel.Role is null ? RowLabelText(labels) : roleLabel.Label.Text;
            var input = ValueForColumn(values, row, columns.Value.Input);
            var total = ValueForColumn(values, row, columns.Value.TotalNg);
            var rate = ValueForColumn(values, row, columns.Value.Rate);
            if (input.Value is null && total.Value is null && rate.Value is null) continue;
            DateInfo? date = dates.TryGetValue(row, out var resolvedDate) ? resolvedDate : null;
            var validation = new List<string>();
            if (columns.Value.RepeatedSection) validation.Add("REPEATED_DEFECT_HEADER_SECTION");
            if (roleLabel.Role is null) validation.Add("CONDITION_ROLE_UNSPECIFIED");
            if (date is null) validation.Add("DATE_MISSING");
            var usable = false;
            if (input.Value is null || total.Value is null) validation.Add("INPUT_OR_TOTAL_NG_MISSING");
            else if (input.Value <= 0) validation.Add("INPUT_NOT_POSITIVE");
            else if (total.Value < 0) validation.Add("TOTAL_NG_NEGATIVE");
            else usable = true;
            var computed = usable ? total.Value!.Value / input.Value!.Value : (double?)null;
            InsertReviewFact(connection, transaction, table, row, date, roleLabel.Role ?? "UNSPECIFIED", conditionLabel, input, total, rate, computed,
                validation.Count == 0 ? "OBSERVED" : "NEEDS_REVIEW", string.Join(';', validation), columns.Value with { ConditionColumn = roleLabel.Label?.Column ?? 0 });
            factCount++;
        }
        InsertTableReview(connection, transaction, table.TableId, "DEFECT_RATE", factCount > 0 ? "EXTRACTED" : "NEEDS_REVIEW", factCount > 0 ? string.Empty : "DEFECT_NUMERIC_ROW_NOT_FOUND", factCount);
    }

    private static void ExtractMeasurementSummaryTable(SqliteConnection connection, SqliteTransaction transaction, TableCandidate table, IReadOnlyList<Label> headers)
    {
        var columns = MeasurementColumns(headers);
        if (columns is null)
        {
            InsertTableReview(connection, transaction, table.TableId, "MEASUREMENT_SUMMARY", "NEEDS_REVIEW", "REQUIRED_HEADER_AMBIGUOUS_OR_MISSING", 0);
            return;
        }
        var values = ReadNumericCells(connection, transaction, table);
        var labelsByRow = ReadLabels(connection, transaction, table.TableId, "ROW_LABEL").GroupBy(value => value.Row).ToDictionary(group => group.Key, group => string.Join(" | ", group.Select(value => value.Text)));
        var dates = RowDates(connection, transaction, table);
        var factCount = 0;
        for (var row = table.StartRow; row <= table.EndRow; row++)
        {
            var average = ValueForColumn(values, row, columns.Value.Average);
            var minimum = ValueForColumn(values, row, columns.Value.Minimum);
            var maximum = ValueForColumn(values, row, columns.Value.Maximum);
            if (average.Value is null && minimum.Value is null && maximum.Value is null) continue;
            var sample = ValueForColumn(values, row, columns.Value.Sample);
            DateInfo? date = dates.TryGetValue(row, out var resolvedDate) ? resolvedDate : null;
            var validation = new List<string>();
            if (average.Value is null || minimum.Value is null || maximum.Value is null) validation.Add("AVERAGE_MIN_OR_MAX_MISSING");
            else if (minimum.Value > maximum.Value) validation.Add("MINIMUM_GREATER_THAN_MAXIMUM");
            else if (average.Value < minimum.Value || average.Value > maximum.Value) validation.Add("AVERAGE_OUTSIDE_MIN_MAX");
            if (columns.Value.Sample is not null && sample.Value is null) validation.Add("SAMPLE_MISSING");
            InsertMeasurementFact(connection, transaction, table, row, date, labelsByRow.GetValueOrDefault(row, string.Empty), sample, average, minimum, maximum,
                validation.Count == 0 ? "OBSERVED" : "NEEDS_REVIEW", string.Join(';', validation), columns.Value);
            factCount++;
        }
        InsertTableReview(connection, transaction, table.TableId, "MEASUREMENT_SUMMARY", factCount > 0 ? "EXTRACTED" : "NEEDS_REVIEW", factCount > 0 ? string.Empty : "NUMERIC_SUMMARY_ROW_NOT_FOUND", factCount);
    }

    private static Dictionary<(int Row, int Column), NumericCell> ReadNumericCells(SqliteConnection connection, SqliteTransaction transaction, TableCandidate table)
    {
        using var command = Command(connection, transaction, "SELECT row_index, column_index, value_text, numeric_value FROM numeric_cells WHERE sheet_id=$sheetId AND row_index BETWEEN $start AND $end;", ("$sheetId", table.SheetId), ("$start", table.StartRow), ("$end", table.EndRow));
        using var reader = command.ExecuteReader();
        var values = new Dictionary<(int, int), NumericCell>();
        while (reader.Read()) values[(reader.GetInt32(0), reader.GetInt32(1))] = new NumericCell(reader.GetString(2), reader.GetDouble(3));
        return values;
    }

    private static Dictionary<int, DateInfo> RowDates(SqliteConnection connection, SqliteTransaction transaction, TableCandidate table)
    {
        var values = new Dictionary<int, DateInfo>();
        var allDates = new Dictionary<(int, int), string>();
        using (var command = Command(connection, transaction, "SELECT row_index, column_index, date_value FROM date_cells WHERE sheet_id=$sheetId;", ("$sheetId", table.SheetId)))
        using (var reader = command.ExecuteReader())
            while (reader.Read())
            {
                var row = reader.GetInt32(0); var column = reader.GetInt32(1); var date = reader.GetString(2);
                allDates[(row, column)] = date;
                if (row >= table.StartRow && row <= table.EndRow) values[row] = new DateInfo(date, row, column);
            }
        using var mergeCommand = Command(connection, transaction, "SELECT range_ref FROM captured_merge_ranges WHERE sheet_id=$sheetId;", ("$sheetId", table.SheetId));
        using var merges = mergeCommand.ExecuteReader();
        while (merges.Read())
        {
            if (!TryRangeBounds(merges.GetString(0), out var minimumRow, out var minimumColumn, out var maximumRow, out _)) continue;
            if (!allDates.TryGetValue((minimumRow, minimumColumn), out var sourceDate)) continue;
            for (var row = Math.Max(minimumRow, table.StartRow); row <= Math.Min(maximumRow, table.EndRow); row++)
                values.TryAdd(row, new DateInfo(sourceDate, minimumRow, minimumColumn));
        }
        return values;
    }

    private static void CreateComparisons(SqliteConnection connection, SqliteTransaction transaction)
    {
        using var command = Command(connection, transaction, """
            SELECT fact_id, workbook_id, sheet_id, table_id, measurement_date, condition_role, computed_ng_rate, fact_status
            FROM numeric_review_facts WHERE metric_type='DEFECT_RATE' AND measurement_date <> ''
            ORDER BY table_id, measurement_date, condition_role, fact_id;
            """);
        using var reader = command.ExecuteReader();
        var facts = new List<ReviewFact>();
        while (reader.Read()) facts.Add(new ReviewFact(reader.GetInt64(0), reader.GetInt64(1), reader.GetInt64(2), reader.GetInt64(3), reader.GetString(4), reader.GetString(5), NullableDouble(reader, 6), reader.GetString(7)));
        foreach (var group in facts.GroupBy(value => (value.TableId, value.MeasurementDate)))
        {
            var tests = group.Where(value => value.Role == "TEST").ToList();
            var normals = group.Where(value => value.Role == "NORMAL").ToList();
            foreach (var test in tests)
            {
                var status = "VALID"; var validation = "SAME_TABLE_SAME_DATE_EXPLICIT_TEST_NORMAL"; ReviewFact? normal = null;
                if (test.Status != "OBSERVED") { status = "NO_COMPARISON_NEEDS_REVIEW"; validation = "TEST_FACT_INVALID"; }
                else if (normals.Count == 0) { status = "NO_SAME_DAY_NORMAL"; validation = "NORMAL_NOT_FOUND_FOR_SAME_DATE"; }
                else if (normals.Count > 1) { status = "NORMAL_AMBIGUOUS"; validation = "MULTIPLE_NORMAL_ROWS_FOR_SAME_DATE"; }
                else if (tests.Count > 1) { status = "TEST_AMBIGUOUS"; validation = "MULTIPLE_TEST_ROWS_FOR_SAME_DATE"; }
                else if (normals[0].Status != "OBSERVED") { status = "NO_COMPARISON_NEEDS_REVIEW"; validation = "NORMAL_FACT_INVALID"; }
                else normal = normals[0];
                var normalRate = normal?.ComputedRate;
                var delta = normalRate is null || test.ComputedRate is null ? (double?)null : test.ComputedRate.Value - normalRate.Value;
                var ratio = normalRate is null || normalRate == 0 || test.ComputedRate is null ? (double?)null : test.ComputedRate.Value / normalRate.Value;
                Execute(connection, transaction, """
                    INSERT INTO test_normal_comparisons(workbook_id, sheet_id, table_id, measurement_date, metric_type, test_fact_id, normal_fact_id, test_ng_rate, normal_ng_rate, absolute_delta, relative_ratio, comparison_status, validation_code)
                    VALUES ($workbookId, $sheetId, $tableId, $date, 'DEFECT_RATE', $testId, $normalId, $testRate, $normalRate, $delta, $ratio, $status, $validation);
                    """,
                    ("$workbookId", test.WorkbookId), ("$sheetId", test.SheetId), ("$tableId", test.TableId), ("$date", test.MeasurementDate), ("$testId", test.FactId),
                    ("$normalId", normal?.FactId ?? (object)DBNull.Value), ("$testRate", test.ComputedRate ?? (object)DBNull.Value), ("$normalRate", normalRate ?? (object)DBNull.Value),
                    ("$delta", delta ?? (object)DBNull.Value), ("$ratio", ratio ?? (object)DBNull.Value), ("$status", status), ("$validation", validation));
            }
        }
    }

    private static void CreateRepeatedBlockComparisons(SqliteConnection connection, SqliteTransaction transaction)
    {
        using var command = Command(connection, transaction, """
            SELECT fact_id, workbook_id, sheet_id, table_id, block_key, block_label, measurement_date, condition_role, computed_ng_rate, fact_status
            FROM repeated_defect_block_facts
            WHERE measurement_date <> '' AND condition_role IN ('TEST', 'NORMAL')
            ORDER BY table_id, block_key, measurement_date, condition_role, fact_id;
            """);
        using var reader = command.ExecuteReader();
        var facts = new List<RepeatedBlockFact>();
        while (reader.Read()) facts.Add(new RepeatedBlockFact(reader.GetInt64(0), reader.GetInt64(1), reader.GetInt64(2), reader.GetInt64(3), reader.GetString(4), reader.GetString(5), reader.GetString(6), reader.GetString(7), NullableDouble(reader, 8), reader.GetString(9)));
        foreach (var group in facts.GroupBy(value => (value.TableId, value.BlockKey, value.MeasurementDate)))
        {
            var tests = group.Where(value => value.Role == "TEST").ToList();
            var normals = group.Where(value => value.Role == "NORMAL").ToList();
            foreach (var test in tests)
            {
                var status = "VALID"; var validation = "SAME_TABLE_SAME_BLOCK_SAME_DATE_EXPLICIT_TEST_NORMAL"; RepeatedBlockFact? normal = null;
                if (test.Status != "OBSERVED") { status = "NO_COMPARISON_NEEDS_REVIEW"; validation = "TEST_FACT_INVALID"; }
                else if (normals.Count == 0) { status = "NO_SAME_DAY_NORMAL"; validation = "NORMAL_NOT_FOUND_FOR_SAME_TABLE_BLOCK_DATE"; }
                else if (normals.Count > 1) { status = "NORMAL_AMBIGUOUS"; validation = "MULTIPLE_NORMAL_ROWS_FOR_SAME_TABLE_BLOCK_DATE"; }
                else if (tests.Count > 1) { status = "TEST_AMBIGUOUS"; validation = "MULTIPLE_TEST_ROWS_FOR_SAME_TABLE_BLOCK_DATE"; }
                else if (normals[0].Status != "OBSERVED") { status = "NO_COMPARISON_NEEDS_REVIEW"; validation = "NORMAL_FACT_INVALID"; }
                else normal = normals[0];
                var normalRate = normal?.ComputedRate;
                var delta = normalRate is null || test.ComputedRate is null ? (double?)null : test.ComputedRate.Value - normalRate.Value;
                var ratio = normalRate is null || normalRate == 0 || test.ComputedRate is null ? (double?)null : test.ComputedRate.Value / normalRate.Value;
                Execute(connection, transaction, """
                    INSERT INTO repeated_block_test_normal_comparisons(workbook_id, sheet_id, table_id, block_key, block_label, measurement_date, test_fact_id, normal_fact_id, test_ng_rate, normal_ng_rate, absolute_delta, relative_ratio, comparison_status, validation_code)
                    VALUES ($workbookId, $sheetId, $tableId, $blockKey, $blockLabel, $date, $testId, $normalId, $testRate, $normalRate, $delta, $ratio, $status, $validation);
                    """,
                    ("$workbookId", test.WorkbookId), ("$sheetId", test.SheetId), ("$tableId", test.TableId), ("$blockKey", test.BlockKey), ("$blockLabel", test.BlockLabel), ("$date", test.MeasurementDate), ("$testId", test.FactId),
                    ("$normalId", normal?.FactId ?? (object)DBNull.Value), ("$testRate", test.ComputedRate ?? (object)DBNull.Value), ("$normalRate", normalRate ?? (object)DBNull.Value),
                    ("$delta", delta ?? (object)DBNull.Value), ("$ratio", ratio ?? (object)DBNull.Value), ("$status", status), ("$validation", validation));
            }
        }
    }

    private static NumericReviewSummary WriteOutputs(string batchDirectory, SqliteConnection connection)
    {
        var tableStatuses = CountRows(connection, "SELECT profile_name || ':' || extraction_status, COUNT(*) FROM numeric_table_reviews GROUP BY profile_name, extraction_status;");
        var comparisonStatuses = CountRows(connection, """
            SELECT comparison_status, COUNT(*)
            FROM (
                SELECT comparison_status FROM test_normal_comparisons
                UNION ALL
                SELECT comparison_status FROM repeated_block_test_normal_comparisons
            )
            GROUP BY comparison_status;
            """);
        var defects = Count(connection, "SELECT COUNT(*) FROM numeric_review_facts;");
        var repeatedBlockFacts = Count(connection, "SELECT COUNT(*) FROM repeated_defect_block_facts;");
        var measurements = Count(connection, "SELECT COUNT(*) FROM measurement_summary_facts;");
        var comparisons = Count(connection, "SELECT COUNT(*) FROM test_normal_comparisons;");
        var repeatedBlockComparisons = Count(connection, "SELECT COUNT(*) FROM repeated_block_test_normal_comparisons;");
        var summary = new NumericReviewSummary("numeric-review-summary-v3", ReviewVersion, UtcNow(), false, defects + repeatedBlockFacts + measurements, defects, repeatedBlockFacts, measurements, comparisons, repeatedBlockComparisons,
            new ReadOnlyDictionary<string, int>(tableStatuses), new ReadOnlyDictionary<string, int>(comparisonStatuses), "numeric-capture.sqlite",
            ["Only explicit Test and Normal labels inside a defect-rate numeric table are compared.", "A comparison requires the same workbook table and normalized calendar date.", "Repeated defect blocks require the same table, block, date, and explicit Test/Normal roles.", "No release, quality, improvement, causality, or other narrative decision is generated."]);
        AtomicWriteText(Path.Combine(batchDirectory, "numeric-review-summary.json"), JsonSerializer.Serialize(summary, JsonOptions) + Environment.NewLine, new UTF8Encoding(false));
        WriteStructureGroupSummary(Path.Combine(batchDirectory, "numeric-structure-groups.json"), connection, repeatedBlockFacts);
        WriteCsv(Path.Combine(batchDirectory, "numeric-review.csv"), connection);
        return summary;
    }

    private static void WriteStructureGroupSummary(string path, SqliteConnection connection, int repeatedBlockFactCount)
    {
        var tableGroups = CountRows(connection, "SELECT profile_name, COUNT(*) FROM numeric_table_reviews GROUP BY profile_name;");
        var factGroups = new SortedDictionary<string, int>(StringComparer.Ordinal)
        {
            ["DEFECT_RATE"] = Count(connection, "SELECT COUNT(*) FROM numeric_review_facts;"),
            ["MEASUREMENT_SUMMARY"] = Count(connection, "SELECT COUNT(*) FROM measurement_summary_facts;"),
            ["REPEATED_DEFECT_BLOCKS"] = repeatedBlockFactCount,
        };
        var summary = new NumericStructureGroupSummary(
            "numeric-structure-groups-v1", UtcNow(),
            new ReadOnlyDictionary<string, int>(tableGroups), new ReadOnlyDictionary<string, int>(factGroups),
            ["Groups are derived only from deterministic numeric header signatures.", "A group identifies an extraction/display profile, not a quality, release, or causal conclusion."]);
        AtomicWriteText(path, JsonSerializer.Serialize(summary, JsonOptions) + Environment.NewLine, new UTF8Encoding(false));
    }

    private static void WriteCsv(string path, SqliteConnection connection)
    {
        var rows = new List<string[]>();
        using var command = Command(connection, null, """
            SELECT workbook.relative_path, comparison.measurement_date, comparison.test_ng_rate, comparison.normal_ng_rate, comparison.absolute_delta, comparison.comparison_status
            FROM test_normal_comparisons AS comparison JOIN capture_workbooks AS workbook ON workbook.workbook_id=comparison.workbook_id
            ORDER BY workbook.relative_path, comparison.measurement_date;
            """);
        using var reader = command.ExecuteReader();
        while (reader.Read()) rows.Add([reader.GetString(0), reader.GetString(1), NullableDouble(reader, 2)?.ToString("G10", CultureInfo.InvariantCulture) ?? string.Empty,
            NullableDouble(reader, 3)?.ToString("G10", CultureInfo.InvariantCulture) ?? string.Empty, NullableDouble(reader, 4)?.ToString("G10", CultureInfo.InvariantCulture) ?? string.Empty, reader.GetString(5)]);
        string[] header = rows.Count == 0 ? ["relativePath", "measurementDate", "status"] : ["relativePath", "measurementDate", "testRate", "normalRate", "difference", "status"];
        var builder = new StringBuilder();
        builder.AppendLine(string.Join(',', header.Select(Csv)));
        foreach (var row in rows) builder.AppendLine(string.Join(',', row.Select(Csv)));
        AtomicWriteText(path, builder.ToString(), new UTF8Encoding(true));
    }

    private static void InsertTableReview(SqliteConnection connection, SqliteTransaction transaction, long tableId, string profile, string status, string reason, int count) =>
        Execute(connection, transaction, "INSERT INTO numeric_table_reviews(table_id, profile_name, extraction_status, reason_code, extracted_fact_count, updated_at) VALUES ($tableId, $profile, $status, $reason, $count, $now);",
            ("$tableId", tableId), ("$profile", profile), ("$status", status), ("$reason", reason), ("$count", count), ("$now", UtcNow()));

    private static void InsertReviewFact(SqliteConnection connection, SqliteTransaction transaction, TableCandidate table, int row, DateInfo? date, string role, string label, CellValue input, CellValue total, CellValue rate, double? computed, string status, string validation, DefectColumnsValue columns) =>
        Execute(connection, transaction, """
            INSERT INTO numeric_review_facts(workbook_id, sheet_id, table_id, row_index, measurement_date, metric_type, condition_role, condition_label, input_value_text, input_value, total_ng_value_text, total_ng_value, reported_ng_rate_text, reported_ng_rate, computed_ng_rate, fact_status, validation_code, date_row_index, date_column_index, condition_column_index, input_column_index, total_ng_column_index, rate_column_index)
            VALUES ($workbookId, $sheetId, $tableId, $row, $date, 'DEFECT_RATE', $role, $label, $inputText, $input, $totalText, $total, $rateText, $rate, $computed, $status, $validation, $dateRow, $dateColumn, $conditionColumn, $inputColumn, $totalColumn, $rateColumn);
            """,
            ("$workbookId", table.WorkbookId), ("$sheetId", table.SheetId), ("$tableId", table.TableId), ("$row", row), ("$date", date?.Value ?? string.Empty), ("$role", role), ("$label", label),
            ("$inputText", input.Text), ("$input", input.Value ?? (object)DBNull.Value), ("$totalText", total.Text), ("$total", total.Value ?? (object)DBNull.Value),
            ("$rateText", rate.Text), ("$rate", rate.Value ?? (object)DBNull.Value), ("$computed", computed ?? (object)DBNull.Value), ("$status", status), ("$validation", validation),
             ("$dateRow", date?.Row ?? (object)DBNull.Value), ("$dateColumn", date?.Column ?? (object)DBNull.Value), ("$conditionColumn", columns.ConditionColumn), ("$inputColumn", columns.Input), ("$totalColumn", columns.TotalNg), ("$rateColumn", columns.Rate ?? (object)DBNull.Value));

    private static void InsertRepeatedBlockFact(SqliteConnection connection, SqliteTransaction transaction, TableCandidate table, RepeatedDefectBlock block, int row, DateInfo? date, string role, string label, CellValue input, CellValue total, CellValue rate, double? computed, string status, string validation, int conditionColumn) =>
        Execute(connection, transaction, """
            INSERT INTO repeated_defect_block_facts(workbook_id, sheet_id, table_id, block_key, block_label, row_index, measurement_date, condition_role, condition_label, input_value_text, input_value, total_ng_value_text, total_ng_value, reported_ng_rate_text, reported_ng_rate, computed_ng_rate, fact_status, validation_code, date_row_index, date_column_index, condition_column_index, input_column_index, total_ng_column_index, rate_column_index)
            VALUES ($workbookId, $sheetId, $tableId, $blockKey, $blockLabel, $row, $date, $role, $label, $inputText, $input, $totalText, $total, $rateText, $rate, $computed, $status, $validation, $dateRow, $dateColumn, $conditionColumn, $inputColumn, $totalColumn, $rateColumn);
            """,
            ("$workbookId", table.WorkbookId), ("$sheetId", table.SheetId), ("$tableId", table.TableId), ("$blockKey", block.Key), ("$blockLabel", block.Label), ("$row", row),
            ("$date", date?.Value ?? string.Empty), ("$role", role), ("$label", label), ("$inputText", input.Text), ("$input", input.Value ?? (object)DBNull.Value),
            ("$totalText", total.Text), ("$total", total.Value ?? (object)DBNull.Value), ("$rateText", rate.Text), ("$rate", rate.Value ?? (object)DBNull.Value),
            ("$computed", computed ?? (object)DBNull.Value), ("$status", status), ("$validation", validation), ("$dateRow", date?.Row ?? (object)DBNull.Value),
            ("$dateColumn", date?.Column ?? (object)DBNull.Value), ("$conditionColumn", conditionColumn == 0 ? (object)DBNull.Value : conditionColumn),
            ("$inputColumn", block.Input), ("$totalColumn", block.TotalNg), ("$rateColumn", block.Rate));

    private static void InsertMeasurementFact(SqliteConnection connection, SqliteTransaction transaction, TableCandidate table, int row, DateInfo? date, string label, CellValue sample, CellValue average, CellValue minimum, CellValue maximum, string status, string validation, MeasurementColumnsValue columns) =>
        Execute(connection, transaction, """
            INSERT INTO measurement_summary_facts(workbook_id, sheet_id, table_id, row_index, measurement_date, condition_label, sample_value_text, sample_value, average_value_text, average_value, minimum_value_text, minimum_value, maximum_value_text, maximum_value, fact_status, validation_code, date_row_index, date_column_index, sample_column_index, average_column_index, minimum_column_index, maximum_column_index)
            VALUES ($workbookId, $sheetId, $tableId, $row, $date, $label, $sampleText, $sample, $averageText, $average, $minimumText, $minimum, $maximumText, $maximum, $status, $validation, $dateRow, $dateColumn, $sampleColumn, $averageColumn, $minimumColumn, $maximumColumn);
            """,
            ("$workbookId", table.WorkbookId), ("$sheetId", table.SheetId), ("$tableId", table.TableId), ("$row", row), ("$date", date?.Value ?? string.Empty), ("$label", label),
            ("$sampleText", sample.Text), ("$sample", sample.Value ?? (object)DBNull.Value), ("$averageText", average.Text), ("$average", average.Value ?? (object)DBNull.Value),
            ("$minimumText", minimum.Text), ("$minimum", minimum.Value ?? (object)DBNull.Value), ("$maximumText", maximum.Text), ("$maximum", maximum.Value ?? (object)DBNull.Value),
            ("$status", status), ("$validation", validation), ("$dateRow", date?.Row ?? (object)DBNull.Value), ("$dateColumn", date?.Column ?? (object)DBNull.Value),
            ("$sampleColumn", columns.Sample ?? (object)DBNull.Value), ("$averageColumn", columns.Average), ("$minimumColumn", columns.Minimum), ("$maximumColumn", columns.Maximum));

    private static DefectColumnsValue? DefectColumns(IReadOnlyList<Label> headers)
    {
        var inputs = HeaderColumns(headers, "INPUT"); var totals = HeaderColumns(headers, "TOTAL_NG"); var rates = HeaderColumns(headers, "NG_RATE");
        if (inputs.Count == 1 && totals.Count > 0)
        {
            var total = totals.FirstOrDefault(value => value > inputs[0]);
            if (total == 0) total = totals[0];
            var rate = rates.FirstOrDefault(value => value > total);
            int? nullableRate = rate != 0 ? rate : rates.Count == 1 ? rates[0] : null;
            return new DefectColumnsValue(inputs[0], total, nullableRate, totals.Count > 1 || rates.Count > 1, 0);
        }
        var input = inputs.Count == 1 ? inputs[0] : (int?)null; var singleRate = rates.Count == 1 ? rates[0] : (int?)null;
        if (input is null || singleRate is null) return null;
        var ng = HeaderColumns(headers, "NG_CUE");
        return ng.Count == 1 ? new DefectColumnsValue(input.Value, ng[0], singleRate, false, 0) : null;
    }

    private static RepeatedDefectBlockLayout? RepeatedDefectBlocks(IReadOnlyList<Label> headers)
    {
        var inputs = HeaderColumns(headers, "INPUT");
        if (inputs.Count < 2) return null;
        var totals = HeaderColumns(headers, "TOTAL_NG");
        var rates = HeaderColumns(headers, "NG_RATE");
        var blocks = new List<RepeatedDefectBlock>(inputs.Count);
        var reviewBlockCount = 0;
        for (var index = 0; index < inputs.Count; index++)
        {
            var input = inputs[index];
            var boundary = index + 1 < inputs.Count ? inputs[index + 1] : int.MaxValue;
            var blockTotals = totals.Where(value => value > input && value < boundary).ToList();
            var blockRates = rates.Where(value => value > input && value < boundary).ToList();
            if (blockTotals.Count != 1 || blockRates.Count != 1 || blockTotals[0] >= blockRates[0])
            {
                reviewBlockCount++;
                continue;
            }
            blocks.Add(new RepeatedDefectBlock($"BLOCK_{index + 1}", BlockLabel(headers, input, boundary, index + 1), input, blockTotals[0], blockRates[0]));
        }
        return new RepeatedDefectBlockLayout(blocks, reviewBlockCount);
    }

    private static string BlockLabel(IReadOnlyList<Label> headers, int inputColumn, int boundary, int ordinal)
    {
        var label = headers
            .Where(value => value.Column >= inputColumn && value.Column < boundary && HeaderToken(value.Text) is null)
            .OrderBy(value => value.Row)
            .ThenBy(value => value.Column)
            .Select(value => value.Text.Trim())
            .FirstOrDefault(value => !string.IsNullOrEmpty(value));
        return string.IsNullOrEmpty(label) ? $"블록 {ordinal}" : label;
    }

    private static MeasurementColumnsValue? MeasurementColumns(IReadOnlyList<Label> headers)
    {
        var sample = SingleColumn(headers, "SAMPLE"); var average = SingleColumn(headers, "AVERAGE"); var minimum = SingleColumn(headers, "MIN"); var maximum = SingleColumn(headers, "MAX");
        return average is null || minimum is null || maximum is null ? null : new MeasurementColumnsValue(sample, average.Value, minimum.Value, maximum.Value);
    }

    private static List<int> HeaderColumns(IReadOnlyList<Label> headers, string token) => headers.Where(value => HeaderToken(value.Text) == token).Select(value => value.Column).Distinct().OrderBy(value => value).ToList();
    private static int? SingleColumn(IReadOnlyList<Label> headers, string token) { var values = HeaderColumns(headers, token); return values.Count == 1 ? values[0] : null; }
    private static CellValue ValueForColumn(IReadOnlyDictionary<(int Row, int Column), NumericCell> values, int row, int? column) => column is not null && values.TryGetValue((row, column.Value), out var value) ? new CellValue(value.Text, value.Value) : new CellValue(string.Empty, null);
    private static string RowLabelText(IReadOnlyList<Label> labels) => string.Join(" | ", labels.Select(value => value.Text).Where(static value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.Ordinal));
    private static string? EquipmentLabel(IReadOnlyList<Label> labels)
    {
        var match = labels.Select(value => EquipmentLabelPattern.Match(value.Text)).FirstOrDefault(value => value.Success);
        return match?.Value is { Length: > 0 } value ? value.Trim() : null;
    }
    private static string? ExplicitRole(string label) { var value = label.Trim().ToLowerInvariant(); return NormalRole.IsMatch(value) ? "NORMAL" : TestRole.IsMatch(value) ? "TEST" : null; }

    private static string? HeaderToken(string value)
    {
        var normalized = NormalizeLabel(value);
        return normalized switch
        {
            "input" => "INPUT", "ok" => "OK", "totalng" or "ngtotal" or "totaldefect" or "totaldefects" => "TOTAL_NG",
            "ngrate" or "defectrate" or "totalngrate" => "NG_RATE", "sample" or "samples" => "SAMPLE",
            "average" or "avg" or "mean" => "AVERAGE", "max" or "maximum" => "MAX", "min" or "minimum" => "MIN",
            "normal" or "baseline" or "before" => "NORMAL_CUE", "test" or "trial" or "after" => "TEST_CUE",
            _ when normalized == "ng" || normalized.StartsWith("ng", StringComparison.Ordinal) => "NG_CUE", _ => null,
        };
    }

    private static string NormalizeLabel(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var character in value.Trim().ToLowerInvariant()) if (char.IsAsciiLetterOrDigit(character) || (character >= '가' && character <= '힣')) builder.Append(character);
        return builder.ToString();
    }

    private static bool TryRangeBounds(string range, out int minimumRow, out int minimumColumn, out int maximumRow, out int maximumColumn)
    {
        minimumRow = minimumColumn = maximumRow = maximumColumn = 0;
        var values = range.Split(':', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (values.Length is < 1 or > 2) return false;
        try
        {
            var first = CellAddress(values[0]); var last = CellAddress(values[^1]);
            minimumRow = Math.Min(first.Row, last.Row); maximumRow = Math.Max(first.Row, last.Row);
            minimumColumn = Math.Min(first.Column, last.Column); maximumColumn = Math.Max(first.Column, last.Column);
            return minimumRow > 0 && minimumColumn > 0 && maximumRow > 0 && maximumColumn > 0;
        }
        catch (OverflowException)
        {
            minimumRow = minimumColumn = maximumRow = maximumColumn = 0;
            return false;
        }
    }

    private static (int Row, int Column) CellAddress(string address)
    {
        var index = 0; var column = 0;
        while (index < address.Length && char.IsLetter(address[index])) { column = checked(column * 26 + char.ToUpperInvariant(address[index]) - 'A' + 1); index++; }
        return int.TryParse(address[index..], NumberStyles.None, CultureInfo.InvariantCulture, out var row) ? (row, column) : (0, 0);
    }

    private static SortedDictionary<string, int> CountRows(SqliteConnection connection, string sql)
    {
        using var command = Command(connection, null, sql); using var reader = command.ExecuteReader(); var values = new SortedDictionary<string, int>(StringComparer.Ordinal);
        while (reader.Read()) values[reader.GetString(0)] = reader.GetInt32(1); return values;
    }
    private static int Count(SqliteConnection connection, string sql) { using var command = Command(connection, null, sql); return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture); }
    private static double? NullableDouble(SqliteDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : reader.GetDouble(ordinal);
    private static SqliteCommand Command(SqliteConnection connection, SqliteTransaction? transaction, string sql, params (string Name, object Value)[] values)
    {
        var command = connection.CreateCommand(); command.Transaction = transaction; command.CommandText = sql;
        foreach (var (name, value) in values) command.Parameters.AddWithValue(name, value); return command;
    }
    private static void Execute(SqliteConnection connection, SqliteTransaction? transaction, string sql, params (string Name, object Value)[] values) { using var command = Command(connection, transaction, sql, values); command.ExecuteNonQuery(); }
    private static string UtcNow() => DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ", CultureInfo.InvariantCulture);
    private static string Csv(string value) => $"\"{value.Replace("\"", "\"\"")}\"";
    private static void AtomicWriteText(string path, string content, Encoding encoding) { Directory.CreateDirectory(Path.GetDirectoryName(path)!); var temporary = path + ".tmp"; File.WriteAllText(temporary, content, encoding); File.Move(temporary, path, true); }

    private sealed record TableCandidate(long TableId, long SheetId, long WorkbookId, int StartRow, int EndRow, int StartColumn, int EndColumn, string CandidateType);
    private sealed record Label(int Row, int Column, string Text);
    private readonly record struct NumericCell(string Text, double Value);
    private readonly record struct CellValue(string Text, double? Value);
    private readonly record struct DateInfo(string Value, int Row, int Column);
    private readonly record struct DefectColumnsValue(int Input, int TotalNg, int? Rate, bool RepeatedSection, int ConditionColumn);
    private readonly record struct MeasurementColumnsValue(int? Sample, int Average, int Minimum, int Maximum);
    private sealed record ReviewFact(long FactId, long WorkbookId, long SheetId, long TableId, string MeasurementDate, string Role, double? ComputedRate, string Status);
    private sealed record RepeatedBlockFact(long FactId, long WorkbookId, long SheetId, long TableId, string BlockKey, string BlockLabel, string MeasurementDate, string Role, double? ComputedRate, string Status);
    private sealed record RepeatedDefectBlock(string Key, string Label, int Input, int TotalNg, int Rate);
    private sealed record RepeatedDefectBlockLayout(IReadOnlyList<RepeatedDefectBlock> Blocks, int ReviewBlockCount);
}

internal sealed record NumericReviewRequest(string ServiceDirectory, string StructureBatchId);
internal sealed record NumericReviewRunResult(string BatchDirectory, NumericReviewSummary Summary);
internal sealed record NumericReviewSummary(
    string SchemaVersion,
    string ReviewVersion,
    string GeneratedAt,
    bool UsesCom,
    int NumericFactCount,
    int DefectRateFactCount,
    int RepeatedDefectBlockFactCount,
    int MeasurementSummaryFactCount,
    int TestNormalComparisonCount,
    int RepeatedBlockTestNormalComparisonCount,
    IReadOnlyDictionary<string, int> TableProfileStatusCounts,
    IReadOnlyDictionary<string, int> ComparisonStatusCounts,
    string Database,
    IReadOnlyList<string> Limitations);

internal sealed record NumericStructureGroupSummary(
    string SchemaVersion,
    string GeneratedAt,
    IReadOnlyDictionary<string, int> TableGroupCounts,
    IReadOnlyDictionary<string, int> FactGroupCounts,
    IReadOnlyList<string> Limitations);
