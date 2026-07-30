using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace InferenceDataAIService.Wpf;

// Presentation-only stage. It reads the batch-local capture/review database
// and emits numeric-only HTML. It never reads an Excel file or applies review
// rules, so source facts and HTML rendering remain independently testable.
internal sealed class NumericRendererEngine
{
    private const string RendererVersion = "numeric-renderer-v3";
    private static readonly Regex ManagedReportName = new("^[0-9a-f]{20}\\.html$", RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    private readonly Action<string>? _log;
    private NumericRendererEngine(Action<string>? log) => _log = log;

    internal static NumericRendererRunResult Run(NumericRendererRequest request, Action<string>? log = null, CancellationToken cancellationToken = default) =>
        new NumericRendererEngine(log).RunCore(request, cancellationToken);

    private NumericRendererRunResult RunCore(NumericRendererRequest request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.ServiceDirectory) || !Directory.Exists(request.ServiceDirectory))
            throw new ArgumentException($"서비스 폴더를 찾을 수 없습니다: {request.ServiceDirectory}");
        var batchDirectory = ResolveBatchDirectory(Path.GetFullPath(request.ServiceDirectory), request.StructureBatchId);
        var databasePath = Path.Combine(batchDirectory, "numeric-capture.sqlite");
        if (!File.Exists(databasePath)) throw new FileNotFoundException("숫자 표 원본 DB를 찾을 수 없습니다.", databasePath);
        var reportDirectory = Path.Combine(batchDirectory, "numeric-reports");
        Directory.CreateDirectory(reportDirectory);

        List<IndexRow> indexRows;
        SortedDictionary<string, int> statuses;
        using (var connection = new SqliteConnection($"Data Source={databasePath}"))
        {
            connection.Open();
            Execute(connection, "PRAGMA foreign_keys=ON;");
            EnsureHtmlReportTable(connection);
            var workbooks = ReadWorkbooks(connection);
            if (workbooks.Count == 0) throw new InvalidOperationException("숫자 표 원본 DB에 workbook snapshot 행이 없습니다.");
            indexRows = new List<IndexRow>(workbooks.Count);
            var expected = new HashSet<string>(StringComparer.Ordinal);
            using var transaction = connection.BeginTransaction();
            foreach (var workbook in workbooks)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var reportName = $"{StableId(workbook.RelativePath)}.html";
                expected.Add(reportName);
                var reportPath = Path.Combine(reportDirectory, reportName);
                AtomicWriteText(reportPath, WorkbookReport(connection, transaction, workbook), new UTF8Encoding(false));
                var relativeReport = $"numeric-reports/{reportName}";
                Execute(connection, transaction, """
                    INSERT INTO numeric_html_reports(workbook_id, source_fingerprint, capture_status, renderer_version, report_path, rendered_at)
                    VALUES ($workbookId, $fingerprint, $status, $renderer, $path, $renderedAt)
                    ON CONFLICT(workbook_id) DO UPDATE SET
                        source_fingerprint=excluded.source_fingerprint, capture_status=excluded.capture_status,
                        renderer_version=excluded.renderer_version, report_path=excluded.report_path, rendered_at=excluded.rendered_at;
                    """,
                    ("$workbookId", workbook.WorkbookId), ("$fingerprint", workbook.Fingerprint), ("$status", workbook.CaptureStatus), ("$renderer", RendererVersion), ("$path", relativeReport), ("$renderedAt", UtcNow()));
                indexRows.Add(new IndexRow(workbook.RelativePath, relativeReport, workbook.CaptureStatus));
            }
            Execute(connection, transaction, "DELETE FROM numeric_html_reports WHERE workbook_id NOT IN (SELECT workbook_id FROM capture_workbooks);");
            transaction.Commit();
            statuses = CaptureStatusCounts(connection);
            CleanupManagedReports(reportDirectory, expected);
        }

        AtomicWriteText(Path.Combine(batchDirectory, "numeric-report-index.html"), IndexHtml(indexRows), new UTF8Encoding(false));
        var summary = new NumericRendererSummary("numeric-render-summary-v3", RendererVersion, UtcNow(), indexRows.Count,
            new ReadOnlyDictionary<string, int>(statuses), "numeric-report-index.html", "numeric-reports",
            ["Reports render only numeric facts and same-date Test–Normal comparisons stored in the batch database.", "Source coordinates, formula text, raw cell samples, and DB paths are intentionally omitted from report HTML."]);
        AtomicWriteText(Path.Combine(batchDirectory, "numeric-render-summary.json"), JsonSerializer.Serialize(summary, JsonOptions) + "\n", new UTF8Encoding(false));
        _log?.Invoke($"숫자 HTML 생성 완료: {summary.ReportCount}개 보고서");
        return new NumericRendererRunResult(batchDirectory, summary);
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

    private static void EnsureHtmlReportTable(SqliteConnection connection) => Execute(connection, """
        CREATE TABLE IF NOT EXISTS numeric_html_reports (
            workbook_id INTEGER PRIMARY KEY REFERENCES capture_workbooks(workbook_id) ON DELETE CASCADE,
            source_fingerprint TEXT NOT NULL, capture_status TEXT NOT NULL, renderer_version TEXT NOT NULL,
            report_path TEXT NOT NULL, rendered_at TEXT NOT NULL
        );
        """);

    private static List<Workbook> ReadWorkbooks(SqliteConnection connection)
    {
        using var command = Command(connection, null, "SELECT workbook_id, relative_path, snapshot_fingerprint, capture_status FROM capture_workbooks ORDER BY relative_path;");
        using var reader = command.ExecuteReader();
        var rows = new List<Workbook>();
        while (reader.Read()) rows.Add(new Workbook(reader.GetInt64(0), reader.GetString(1), reader.GetString(2), reader.GetString(3)));
        return rows;
    }

    private static string WorkbookReport(SqliteConnection connection, SqliteTransaction transaction, Workbook workbook)
    {
        var defects = ReadDefectRows(connection, transaction, workbook.WorkbookId);
        var repeatedBlockFacts = ReadRepeatedBlockFacts(connection, transaction, workbook.WorkbookId);
        var repeatedBlockComparisons = ReadRepeatedBlockComparisons(connection, transaction, workbook.WorkbookId);
        var equipmentBlockFacts = ReadEquipmentBlockFacts(connection, transaction, workbook.WorkbookId);
        var measurements = ReadMeasurementRows(connection, transaction, workbook.WorkbookId);
        var comparisons = ReadComparisonRows(connection, transaction, workbook.WorkbookId);
        var unclassified = Count(connection, transaction, """
            SELECT COUNT(*) FROM numeric_table_reviews AS review
            JOIN numeric_table_candidates AS candidate ON candidate.table_id=review.table_id
            JOIN captured_sheets AS sheet ON sheet.sheet_id=candidate.sheet_id
            WHERE sheet.workbook_id=$workbookId AND review.extraction_status IN ('NOT_IMPLEMENTED', 'NEEDS_REVIEW');
            """, ("$workbookId", workbook.WorkbookId));

        var defectTable = HtmlTable(
            ["측정일", "조건", "Input", "Total NG", "불량률 (ppm)", "상태"],
            defects.Select(row => new[] { Escape(row.MeasurementDate), Escape(string.IsNullOrEmpty(row.ConditionLabel) ? row.ConditionRole : row.ConditionLabel), Number(row.Input), Number(row.TotalNg), Ppm(row.ComputedRate), Escape(StatusText(row.Status)) }).ToList(),
            "추출된 불량률 숫자 데이터가 없습니다.");
        var measurementTable = HtmlTable(
            ["측정일", "조건", "N", "Average", "Min", "Max", "상태"],
            measurements.Select(row => new[] { Escape(row.MeasurementDate), Escape(row.ConditionLabel), Number(row.Sample, 0), Number(row.Average), Number(row.Minimum), Number(row.Maximum), Escape(StatusText(row.Status)) }).ToList(),
            "추출된 측정 통계 숫자 데이터가 없습니다.");
        var comparisonTable = HtmlTable(
            ["측정일", "Test 불량률 (ppm)", "Normal 불량률 (ppm)", "차이 (ppm)", "상태"],
            comparisons.Select(row => new[] { Escape(row.MeasurementDate), Ppm(row.TestRate), Ppm(row.NormalRate), Ppm(row.Delta), Escape(StatusText(row.Status)) }).ToList(),
            "동일 날짜의 Test–Normal 비교 결과가 없습니다.");
        var repeatedBlockTable = HtmlTable(
            ["블록", "측정일", "조건", "Input", "Total NG", "불량률 (ppm)", "상태"],
            repeatedBlockFacts.Select(row => new[] { Escape(row.BlockLabel), Escape(row.MeasurementDate), Escape(string.IsNullOrEmpty(row.ConditionLabel) ? row.ConditionRole : row.ConditionLabel), Number(row.Input), Number(row.TotalNg), Ppm(row.ComputedRate), Escape(StatusText(row.Status)) }).ToList(),
            "추출된 다중 블록 불량률 데이터가 없습니다.");
        var repeatedBlockComparisonTable = HtmlTable(
            ["블록", "측정일", "Test 불량률 (ppm)", "Normal 불량률 (ppm)", "차이 (ppm)", "상태"],
            repeatedBlockComparisons.Select(row => new[] { Escape(row.BlockLabel), Escape(row.MeasurementDate), Ppm(row.TestRate), Ppm(row.NormalRate), Ppm(row.Delta), Escape(StatusText(row.Status)) }).ToList(),
            "다중 블록 Test–Normal 비교 결과가 없습니다.");
        var equipmentBlockTable = HtmlTable(
            ["설비", "블록", "측정일", "조건", "Input", "Total NG", "불량률 (ppm)", "상태"],
            equipmentBlockFacts.Select(row => new[] { Escape(row.EquipmentLabel), Escape(row.BlockLabel), Escape(row.MeasurementDate), Escape(row.ConditionLabel), Number(row.Input), Number(row.TotalNg), Ppm(row.ComputedRate), Escape(StatusText(row.Status)) }).ToList(),
            "동일한 날짜와 조건으로 나란히 표시할 설비별 수치가 없습니다.");
        var notices = new StringBuilder();
        if (workbook.CaptureStatus != "CAPTURED") notices.Append($"<p class='notice'>원본 숫자 적재 상태: <strong>{Escape(StatusText(workbook.CaptureStatus))}</strong></p>");
        if (unclassified > 0) notices.Append($"<p class='notice'>추출 규칙 검토 필요 숫자 표: <strong>{unclassified}</strong>개</p>");
        var title = Path.GetFileName(workbook.RelativePath);
        var equipmentBlockSection = equipmentBlockFacts.Count == 0 ? string.Empty : $"<h2>설비별 다중 블록 비교</h2><p class='sub'>같은 날짜와 조건의 설비 수치를 나란히 표시합니다. 기준 설비나 품질 결론은 만들지 않습니다.</p>{equipmentBlockTable}";
        var repeatedBlockSection = repeatedBlockFacts.Count == 0 ? string.Empty : $"<h2>다중 블록 불량률</h2>{repeatedBlockTable}<h2>다중 블록 Test–Normal 비교</h2>{repeatedBlockComparisonTable}";
        return $@"<!doctype html><html lang='ko'><head><meta charset='utf-8'><title>{Escape(title)} — 숫자 검토</title>
<style>
body{{font-family:'Segoe UI',sans-serif;margin:22px;color:#162d50;background:#f7f9fc}}main{{max-width:1600px;margin:auto}}h1{{font-size:22px;margin:0 0 18px}}h2{{font-size:17px;margin:28px 0 8px}}table{{border-collapse:collapse;width:100%;background:white;font-size:13px}}th,td{{border:1px solid #d5dfeb;padding:8px;text-align:left;vertical-align:middle}}th{{background:#eaf0f7;color:#183b67}}td.empty{{color:#5d6d7e;text-align:center;padding:16px}}.notice{{padding:10px 12px;background:#fff8e8;border:1px solid #f1dfad;border-radius:6px}}.sub{{color:#53657a;font-size:13px;margin-top:-12px}}</style></head>
<body><main><h1>검토 결과 — {Escape(title)}</h1><p class='sub'>숫자 표 기반 관측값</p>{notices}
{repeatedBlockSection}
{equipmentBlockSection}
<h2>불량률</h2>{defectTable}
<h2>Test–Normal 비교</h2>{comparisonTable}
<h2>측정 통계</h2>{measurementTable}
</main></body></html>";
    }

    private static string IndexHtml(IReadOnlyList<IndexRow> rows)
    {
        var body = string.Concat(rows.Select(row => $"<tr><td>{Escape(row.RelativePath)}</td><td>{Escape(StatusText(row.CaptureStatus))}</td><td><a href='{Escape(row.ReportPath)}'>열기</a></td></tr>"));
        return $@"<!doctype html><html lang='ko'><head><meta charset='utf-8'><title>숫자 검토 보고서 목록</title>
<style>body{{font-family:'Segoe UI',sans-serif;margin:24px;color:#162d50}}table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #d5dfeb;padding:8px;text-align:left}}th{{background:#eaf0f7}}</style></head>
<body><h1>Excel별 숫자 검토 보고서</h1><table><thead><tr><th>Excel</th><th>적재 상태</th><th>보고서</th></tr></thead><tbody>{body}</tbody></table></body></html>";
    }

    private static List<DefectRow> ReadDefectRows(SqliteConnection connection, SqliteTransaction transaction, long workbookId)
    {
        using var command = Command(connection, transaction, "SELECT measurement_date, condition_role, condition_label, input_value, total_ng_value, computed_ng_rate, fact_status FROM numeric_review_facts WHERE workbook_id=$workbookId ORDER BY measurement_date, row_index, fact_id;", ("$workbookId", workbookId));
        using var reader = command.ExecuteReader(); var rows = new List<DefectRow>();
        while (reader.Read()) rows.Add(new DefectRow(reader.GetString(0), reader.GetString(1), reader.GetString(2), NullableDouble(reader, 3), NullableDouble(reader, 4), NullableDouble(reader, 5), reader.GetString(6)));
        return rows;
    }

    private static List<RepeatedBlockFactRow> ReadRepeatedBlockFacts(SqliteConnection connection, SqliteTransaction transaction, long workbookId)
    {
        using var command = Command(connection, transaction, "SELECT block_label, measurement_date, condition_role, condition_label, input_value, total_ng_value, computed_ng_rate, fact_status FROM repeated_defect_block_facts WHERE workbook_id=$workbookId ORDER BY block_key, measurement_date, row_index, fact_id;", ("$workbookId", workbookId));
        using var reader = command.ExecuteReader(); var rows = new List<RepeatedBlockFactRow>();
        while (reader.Read()) rows.Add(new RepeatedBlockFactRow(reader.GetString(0), reader.GetString(1), reader.GetString(2), reader.GetString(3), NullableDouble(reader, 4), NullableDouble(reader, 5), NullableDouble(reader, 6), reader.GetString(7)));
        return rows;
    }

    private static List<RepeatedBlockComparisonRow> ReadRepeatedBlockComparisons(SqliteConnection connection, SqliteTransaction transaction, long workbookId)
    {
        using var command = Command(connection, transaction, "SELECT block_label, measurement_date, test_ng_rate, normal_ng_rate, absolute_delta, comparison_status FROM repeated_block_test_normal_comparisons WHERE workbook_id=$workbookId ORDER BY block_key, measurement_date, comparison_id;", ("$workbookId", workbookId));
        using var reader = command.ExecuteReader(); var rows = new List<RepeatedBlockComparisonRow>();
        while (reader.Read()) rows.Add(new RepeatedBlockComparisonRow(reader.GetString(0), reader.GetString(1), NullableDouble(reader, 2), NullableDouble(reader, 3), NullableDouble(reader, 4), reader.GetString(5)));
        return rows;
    }

    private static List<EquipmentBlockRow> ReadEquipmentBlockFacts(SqliteConnection connection, SqliteTransaction transaction, long workbookId)
    {
        using var command = Command(connection, transaction, """
            SELECT equipment.equipment_label, fact.block_label, fact.measurement_date, fact.condition_label,
                   fact.input_value, fact.total_ng_value, fact.computed_ng_rate, fact.fact_status
            FROM repeated_defect_block_facts AS fact
            JOIN repeated_block_equipment_tables AS equipment ON equipment.table_id=fact.table_id
            WHERE fact.workbook_id=$workbookId AND fact.measurement_date <> ''
            ORDER BY fact.measurement_date, fact.condition_label, equipment.equipment_label, fact.block_key, fact.row_index, fact.fact_id;
            """, ("$workbookId", workbookId));
        using var reader = command.ExecuteReader(); var rows = new List<EquipmentBlockRow>();
        while (reader.Read()) rows.Add(new EquipmentBlockRow(reader.GetString(0), reader.GetString(1), reader.GetString(2), reader.GetString(3), NullableDouble(reader, 4), NullableDouble(reader, 5), NullableDouble(reader, 6), reader.GetString(7)));
        return rows;
    }

    private static List<MeasurementRow> ReadMeasurementRows(SqliteConnection connection, SqliteTransaction transaction, long workbookId)
    {
        using var command = Command(connection, transaction, "SELECT measurement_date, condition_label, sample_value, average_value, minimum_value, maximum_value, fact_status FROM measurement_summary_facts WHERE workbook_id=$workbookId ORDER BY measurement_date, row_index, fact_id;", ("$workbookId", workbookId));
        using var reader = command.ExecuteReader(); var rows = new List<MeasurementRow>();
        while (reader.Read()) rows.Add(new MeasurementRow(reader.GetString(0), reader.GetString(1), NullableDouble(reader, 2), NullableDouble(reader, 3), NullableDouble(reader, 4), NullableDouble(reader, 5), reader.GetString(6)));
        return rows;
    }

    private static List<ComparisonRow> ReadComparisonRows(SqliteConnection connection, SqliteTransaction transaction, long workbookId)
    {
        using var command = Command(connection, transaction, "SELECT measurement_date, test_ng_rate, normal_ng_rate, absolute_delta, comparison_status FROM test_normal_comparisons WHERE workbook_id=$workbookId ORDER BY measurement_date, comparison_id;", ("$workbookId", workbookId));
        using var reader = command.ExecuteReader(); var rows = new List<ComparisonRow>();
        while (reader.Read()) rows.Add(new ComparisonRow(reader.GetString(0), NullableDouble(reader, 1), NullableDouble(reader, 2), NullableDouble(reader, 3), reader.GetString(4)));
        return rows;
    }

    private static string HtmlTable(IReadOnlyList<string> headers, IReadOnlyList<string[]> rows, string empty)
    {
        var head = string.Concat(headers.Select(header => $"<th>{Escape(header)}</th>"));
        var body = rows.Count == 0
            ? $"<tr><td class='empty' colspan='{headers.Count}'>{Escape(empty)}</td></tr>"
            : string.Concat(rows.Select(row => "<tr>" + string.Concat(row.Select(value => $"<td>{value}</td>")) + "</tr>"));
        return $"<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>";
    }

    private static string Number(double? value, int digits = 2)
    {
        if (value is null) return "—";
        var result = value.Value.ToString($"N{digits}", CultureInfo.InvariantCulture);
        return result.Contains('.', StringComparison.Ordinal) ? result.TrimEnd('0').TrimEnd('.') : result;
    }

    private static string Ppm(double? value) => value is null ? "—" : Number(value.Value * 1_000_000d);
    private static string Escape(string? value)
    {
        var source = string.IsNullOrEmpty(value) ? "—" : value;
        return source.Replace("&", "&amp;", StringComparison.Ordinal).Replace("<", "&lt;", StringComparison.Ordinal).Replace(">", "&gt;", StringComparison.Ordinal).Replace("\"", "&quot;", StringComparison.Ordinal).Replace("'", "&#x27;", StringComparison.Ordinal);
    }
    private static string StatusText(string value) => value switch
    {
        "OBSERVED" => "관측", "NEEDS_REVIEW" => "검토 필요", "VALID" => "유효", "NO_SAME_DAY_NORMAL" => "동일 날짜 Normal 없음",
        "NORMAL_AMBIGUOUS" => "Normal 중복", "TEST_AMBIGUOUS" => "Test 중복", "NO_COMPARISON_NEEDS_REVIEW" => "비교 검토 필요",
        "PENDING" => "원본 적재 대기", "TRUNCATED" => "원본 적재 보류", "CHANGED" => "원본 변경됨", "QUARANTINED" => "원본 확인 필요", "FAILED_RETRYABLE" => "원본 적재 실패", _ => value,
    };

    // The current archive contains Korean/ASCII paths, for which invariant
    // lower casing equals Python casefold. These additional folds cover common
    // non-ASCII filename cases without normalizing or changing path separators.
    private static string PythonCaseFold(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var character in value)
        {
            builder.Append(character switch
            {
                '\u00df' or '\u1e9e' => "ss", '\u03c2' => "σ", '\u0345' => "ι", '\u03d0' => "β", '\u03d1' => "θ", '\u03d5' => "φ", '\u03d6' => "π", '\u03f0' => "κ", '\u03f1' => "ρ", '\u03f5' => "ε", '\u212a' => "k", '\u212b' => "å", '\u017f' => "s", '\u00b5' => "μ",
                _ => char.ToLowerInvariant(character).ToString(),
            });
        }
        return builder.ToString();
    }

    private static string StableId(string relativePath) => Convert.ToHexString(SHA256.HashData(new UTF8Encoding(false).GetBytes(PythonCaseFold(relativePath)))).ToLowerInvariant()[..20];
    private static SortedDictionary<string, int> CaptureStatusCounts(SqliteConnection connection)
    {
        using var command = Command(connection, null, "SELECT capture_status, COUNT(*) FROM numeric_html_reports GROUP BY capture_status;");
        using var reader = command.ExecuteReader(); var values = new SortedDictionary<string, int>(StringComparer.Ordinal);
        while (reader.Read()) values[reader.GetString(0)] = reader.GetInt32(1); return values;
    }
    private static void CleanupManagedReports(string directory, IReadOnlySet<string> expected)
    {
        var root = Path.GetFullPath(directory);
        foreach (var path in Directory.EnumerateFiles(root, "*.html", SearchOption.TopDirectoryOnly))
        {
            var name = Path.GetFileName(path);
            if (!ManagedReportName.IsMatch(name) || expected.Contains(name)) continue;
            var full = Path.GetFullPath(path);
            if (!full.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) continue;
            File.Delete(full);
        }
    }
    private static int Count(SqliteConnection connection, SqliteTransaction transaction, string sql, params (string Name, object Value)[] values) { using var command = Command(connection, transaction, sql, values); return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture); }
    private static double? NullableDouble(SqliteDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : reader.GetDouble(ordinal);
    private static SqliteCommand Command(SqliteConnection connection, SqliteTransaction? transaction, string sql, params (string Name, object Value)[] values) { var command = connection.CreateCommand(); command.Transaction = transaction; command.CommandText = sql; foreach (var (name, value) in values) command.Parameters.AddWithValue(name, value); return command; }
    private static void Execute(SqliteConnection connection, string sql) { using var command = Command(connection, null, sql); command.ExecuteNonQuery(); }
    private static void Execute(SqliteConnection connection, SqliteTransaction transaction, string sql, params (string Name, object Value)[] values) { using var command = Command(connection, transaction, sql, values); command.ExecuteNonQuery(); }
    private static string UtcNow() => DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.ffffffZ", CultureInfo.InvariantCulture);
    private static void AtomicWriteText(string path, string content, Encoding encoding)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var temporary = path + ".tmp";
        // Python Path.write_text() uses the Windows text newline convention.
        // Normalizing once keeps the legacy renderer's raw HTML bytes stable.
        var platformText = content.Replace("\r\n", "\n", StringComparison.Ordinal).Replace("\n", Environment.NewLine, StringComparison.Ordinal);
        File.WriteAllText(temporary, platformText, encoding);
        File.Move(temporary, path, true);
    }

    private sealed record Workbook(long WorkbookId, string RelativePath, string Fingerprint, string CaptureStatus);
    private sealed record IndexRow(string RelativePath, string ReportPath, string CaptureStatus);
    private sealed record DefectRow(string MeasurementDate, string ConditionRole, string ConditionLabel, double? Input, double? TotalNg, double? ComputedRate, string Status);
    private sealed record RepeatedBlockFactRow(string BlockLabel, string MeasurementDate, string ConditionRole, string ConditionLabel, double? Input, double? TotalNg, double? ComputedRate, string Status);
    private sealed record RepeatedBlockComparisonRow(string BlockLabel, string MeasurementDate, double? TestRate, double? NormalRate, double? Delta, string Status);
    private sealed record EquipmentBlockRow(string EquipmentLabel, string BlockLabel, string MeasurementDate, string ConditionLabel, double? Input, double? TotalNg, double? ComputedRate, string Status);
    private sealed record MeasurementRow(string MeasurementDate, string ConditionLabel, double? Sample, double? Average, double? Minimum, double? Maximum, string Status);
    private sealed record ComparisonRow(string MeasurementDate, double? TestRate, double? NormalRate, double? Delta, string Status);
}

internal sealed record NumericRendererRequest(string ServiceDirectory, string StructureBatchId);
internal sealed record NumericRendererRunResult(string BatchDirectory, NumericRendererSummary Summary);
internal sealed record NumericRendererSummary(
    string SchemaVersion,
    string RendererVersion,
    string GeneratedAt,
    int ReportCount,
    IReadOnlyDictionary<string, int> CaptureStatusCounts,
    string Index,
    string ReportDirectory,
    IReadOnlyList<string> Limitations);
