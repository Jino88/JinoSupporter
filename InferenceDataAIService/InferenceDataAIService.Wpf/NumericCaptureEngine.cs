using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Numerics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Xml;
using Microsoft.Data.Sqlite;

namespace InferenceDataAIService.Wpf;

// Read-only OpenXML numeric capture.  The table and SQLite names intentionally
// match numeric-capture-v1 because the review/renderer engines still consume
// this batch-local database during the staged C# migration.
internal sealed class NumericCaptureEngine
{
    private const string CaptureVersion = "numeric-capture-v3-full-document-grid";
    private const int MaxZipEntries = 20_000;
    private const long MaxPackageUncompressedBytes = 512L * 1024 * 1024;
    private const long MaxWorksheetUncompressedBytes = 96L * 1024 * 1024;
    private const int MaxCompressionRatio = 250;
    private const int MaxCapturedCellsPerSheet = 1_000_000;
    private const int MaxCapturedMergesPerSheet = 100_000;
    private const int MaxHeaderRowsPerRegion = 2;
    private const int MaxHeaderLabelLength = 500;
    private const int MaxTableColumnGap = 3;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    private readonly Action<string>? _log;

    private NumericCaptureEngine(Action<string>? log) => _log = log;

    internal static NumericCaptureRunResult Run(NumericCaptureRequest request, Action<string>? log = null, CancellationToken cancellationToken = default) =>
        new NumericCaptureEngine(log).RunCore(request, cancellationToken);

    private NumericCaptureRunResult RunCore(NumericCaptureRequest request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.ServiceDirectory) || !Directory.Exists(request.ServiceDirectory))
            throw new ArgumentException($"서비스 폴더를 찾을 수 없습니다: {request.ServiceDirectory}");
        if (request.Limit < 0 || request.ProgressEvery < 0) throw new ArgumentException("limit과 progress 간격은 0 이상이어야 합니다.");
        var batchDirectory = ResolveBatchDirectory(Path.GetFullPath(request.ServiceDirectory), request.StructureBatchId);
        if (!File.Exists(Path.Combine(batchDirectory, "batch.json"))) throw new ArgumentException($"구조 스캔 배치를 찾을 수 없습니다: {request.StructureBatchId}");
        var records = SourceRows(batchDirectory);
        if (records.Count == 0) throw new InvalidOperationException("선택한 구조 배치에 .xlsx/.xlsm 원본이 없습니다.");

        using var connection = OpenCaptureDb(Path.Combine(batchDirectory, "numeric-capture.sqlite"));
        foreach (var item in records) EnsureWorkbook(connection, item);
        var selected = request.Limit > 0 ? records.Take(request.Limit).ToList() : records;
        var runCounts = new SortedDictionary<string, int>(StringComparer.Ordinal);
        for (var index = 0; index < selected.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var item = selected[index];
            _log?.Invoke($"숫자 표 적재: {item.RelativePath}");
            var status = CaptureWorkbook(connection, item, request.Force, cancellationToken);
            Increment(runCounts, status);
            if (status is not "CAPTURED" and not "SKIPPED" || (request.ProgressEvery > 0 && (index == 0 || index == selected.Count - 1 || (index + 1) % request.ProgressEvery == 0)))
                _log?.Invoke($"숫자 표 적재 {index + 1}/{selected.Count}: {status} — {item.RelativePath}");
        }
        var summary = WriteOutputs(batchDirectory, connection);
        return new NumericCaptureRunResult(batchDirectory, new ReadOnlyDictionary<string, int>(runCounts), summary);
    }

    private static string ResolveBatchDirectory(string serviceDirectory, string batchId)
    {
        if (string.IsNullOrWhiteSpace(batchId) || batchId.Length > 96 || batchId is "." or ".." ||
            batchId.Any(character => !(char.IsAsciiLetterOrDigit(character) || character is '.' or '_' or '-')))
            throw new ArgumentException("유효한 구조 배치 ID가 아닙니다.");
        var outputRoot = Path.GetFullPath(
            AppRuntimePaths.Current.BatchRootDirectory);
        var target = Path.GetFullPath(Path.Combine(outputRoot, batchId));
        if (!IsWithin(target, outputRoot)) throw new ArgumentException("배치 경로가 출력 루트 밖을 가리킵니다.");
        return target;
    }

    private static SqliteConnection OpenCaptureDb(string path)
    {
        var connection = new SqliteConnection($"Data Source={path}");
        connection.Open();
        Execute(connection, null, "PRAGMA foreign_keys=ON;");
        Execute(connection, null, "PRAGMA journal_mode=WAL;");
        Execute(connection, null, """
            CREATE TABLE IF NOT EXISTS capture_metadata (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS capture_workbooks (
                workbook_id INTEGER PRIMARY KEY,
                relative_path TEXT NOT NULL UNIQUE,
                source_path TEXT NOT NULL,
                extension TEXT NOT NULL,
                snapshot_fingerprint TEXT NOT NULL,
                current_fingerprint TEXT NOT NULL DEFAULT '',
                structure_scan_status TEXT NOT NULL,
                capture_status TEXT NOT NULL,
                attempts INTEGER NOT NULL DEFAULT 0,
                sheet_count_expected INTEGER NOT NULL DEFAULT 0,
                sheet_count_captured INTEGER NOT NULL DEFAULT 0,
                numeric_cell_count INTEGER NOT NULL DEFAULT 0,
                formula_count INTEGER NOT NULL DEFAULT 0,
                date_cell_count INTEGER NOT NULL DEFAULT 0,
                table_candidate_count INTEGER NOT NULL DEFAULT 0,
                merge_count INTEGER NOT NULL DEFAULT 0,
                error_text TEXT NOT NULL DEFAULT '',
                started_at TEXT NOT NULL DEFAULT '',
                finished_at TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_capture_workbooks_status ON capture_workbooks(capture_status, relative_path);
            CREATE TABLE IF NOT EXISTS captured_sheets (
                sheet_id INTEGER PRIMARY KEY,
                workbook_id INTEGER NOT NULL REFERENCES capture_workbooks(workbook_id) ON DELETE CASCADE,
                sheet_index INTEGER NOT NULL,
                sheet_name TEXT NOT NULL,
                sheet_state TEXT NOT NULL,
                declared_dimension TEXT NOT NULL DEFAULT '',
                max_row INTEGER NOT NULL,
                max_column INTEGER NOT NULL,
                merge_count INTEGER NOT NULL DEFAULT 0,
                numeric_cell_count INTEGER NOT NULL DEFAULT 0,
                formula_count INTEGER NOT NULL DEFAULT 0,
                date_cell_count INTEGER NOT NULL DEFAULT 0,
                capture_status TEXT NOT NULL,
                warning_text TEXT NOT NULL DEFAULT '',
                UNIQUE(workbook_id, sheet_index)
            );
            CREATE INDEX IF NOT EXISTS idx_captured_sheets_workbook ON captured_sheets(workbook_id, sheet_index);
            CREATE TABLE IF NOT EXISTS captured_merge_ranges (
                sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
                range_ref TEXT NOT NULL,
                PRIMARY KEY(sheet_id, range_ref)
            );
            CREATE TABLE IF NOT EXISTS numeric_cells (
                sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
                row_index INTEGER NOT NULL,
                column_index INTEGER NOT NULL,
                source_kind TEXT NOT NULL,
                value_text TEXT NOT NULL,
                numeric_value REAL NOT NULL,
                formula_text TEXT NOT NULL DEFAULT '',
                PRIMARY KEY(sheet_id, row_index, column_index)
            );
            CREATE INDEX IF NOT EXISTS idx_numeric_cells_sheet_position ON numeric_cells(sheet_id, row_index, column_index);
            CREATE TABLE IF NOT EXISTS captured_text_cells (
                sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
                row_index INTEGER NOT NULL, column_index INTEGER NOT NULL, text_value TEXT NOT NULL,
                PRIMARY KEY(sheet_id, row_index, column_index)
            );
            CREATE INDEX IF NOT EXISTS idx_captured_text_cells_sheet_position ON captured_text_cells(sheet_id, row_index, column_index);
            CREATE TABLE IF NOT EXISTS formula_cells (
                sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
                row_index INTEGER NOT NULL,
                column_index INTEGER NOT NULL,
                formula_text TEXT NOT NULL,
                cached_value_text TEXT NOT NULL DEFAULT '',
                cached_numeric_value REAL,
                cached_is_date INTEGER NOT NULL DEFAULT 0,
                PRIMARY KEY(sheet_id, row_index, column_index)
            );
            CREATE TABLE IF NOT EXISTS date_cells (
                sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
                row_index INTEGER NOT NULL,
                column_index INTEGER NOT NULL,
                date_value TEXT NOT NULL,
                source_kind TEXT NOT NULL,
                PRIMARY KEY(sheet_id, row_index, column_index)
            );
            CREATE INDEX IF NOT EXISTS idx_date_cells_sheet_date ON date_cells(sheet_id, date_value);
            CREATE TABLE IF NOT EXISTS numeric_table_candidates (
                table_id INTEGER PRIMARY KEY,
                sheet_id INTEGER NOT NULL REFERENCES captured_sheets(sheet_id) ON DELETE CASCADE,
                start_row INTEGER NOT NULL,
                end_row INTEGER NOT NULL,
                start_column INTEGER NOT NULL,
                end_column INTEGER NOT NULL,
                header_start_row INTEGER NOT NULL,
                header_end_row INTEGER NOT NULL,
                numeric_cell_count INTEGER NOT NULL,
                candidate_type TEXT NOT NULL,
                confidence TEXT NOT NULL,
                UNIQUE(sheet_id, start_row, end_row, start_column, end_column)
            );
            CREATE INDEX IF NOT EXISTS idx_numeric_table_candidates_sheet ON numeric_table_candidates(sheet_id, start_row);
            CREATE TABLE IF NOT EXISTS numeric_table_labels (
                table_id INTEGER NOT NULL REFERENCES numeric_table_candidates(table_id) ON DELETE CASCADE,
                row_index INTEGER NOT NULL,
                column_index INTEGER NOT NULL,
                label_text TEXT NOT NULL,
                label_role TEXT NOT NULL,
                PRIMARY KEY(table_id, row_index, column_index)
            );
            """);
        Execute(connection, null, "INSERT OR REPLACE INTO capture_metadata(key, value) VALUES ('schemaVersion', 'numeric-capture-db-v1');");
        Execute(connection, null, "INSERT OR REPLACE INTO capture_metadata(key, value) VALUES ('captureVersion', $version);", ("$version", CaptureVersion));
        return connection;
    }

    private static List<SourceItem> SourceRows(string batchDirectory)
    {
        var statePath = Path.Combine(batchDirectory, "state.sqlite");
        if (!File.Exists(statePath)) throw new InvalidOperationException($"구조 스캔 상태 DB가 없습니다: {statePath}");
        using var connection = new SqliteConnection($"Data Source={statePath};Mode=ReadOnly");
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT relative_path, source_path, extension, kind, fingerprint, status FROM items WHERE kind='openxml' ORDER BY relative_path;";
        using var reader = command.ExecuteReader();
        var rows = new List<SourceItem>();
        while (reader.Read()) rows.Add(new SourceItem(reader.GetString(0), reader.GetString(1), reader.GetString(2), reader.GetString(3), reader.GetString(4), reader.GetString(5)));
        return rows;
    }

    private static long EnsureWorkbook(SqliteConnection connection, SourceItem item)
    {
        Execute(connection, null, """
            INSERT INTO capture_workbooks(relative_path, source_path, extension, snapshot_fingerprint, structure_scan_status, capture_status)
            VALUES ($relativePath, $sourcePath, $extension, $fingerprint, $status, 'PENDING')
            ON CONFLICT(relative_path) DO UPDATE SET
                source_path=excluded.source_path,
                extension=excluded.extension,
                snapshot_fingerprint=excluded.snapshot_fingerprint,
                structure_scan_status=excluded.structure_scan_status;
            """,
            ("$relativePath", item.RelativePath), ("$sourcePath", item.SourcePath), ("$extension", item.Extension), ("$fingerprint", item.Fingerprint), ("$status", item.Status));
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT workbook_id FROM capture_workbooks WHERE relative_path=$relativePath;";
        command.Parameters.AddWithValue("$relativePath", item.RelativePath);
        return Convert.ToInt64(command.ExecuteScalar(), CultureInfo.InvariantCulture);
    }

    private string CaptureWorkbook(SqliteConnection connection, SourceItem item, bool force, CancellationToken cancellationToken)
    {
        var workbookId = EnsureWorkbook(connection, item);
        if (!File.Exists(item.SourcePath) || !string.Equals(SourceFingerprint(item.SourcePath), item.Fingerprint, StringComparison.Ordinal))
        {
            Execute(connection, null, "UPDATE capture_workbooks SET capture_status='CHANGED', error_text=$error, finished_at=$finishedAt WHERE workbook_id=$workbookId;",
                ("$error", "Source path, size, or modified time differs from the structure-scan snapshot."), ("$finishedAt", UtcNow()), ("$workbookId", workbookId));
            return "CHANGED";
        }
        var current = WorkbookStatus(connection, workbookId);
        if (!force && current.CaptureStatus == "CAPTURED" && string.Equals(current.CurrentFingerprint, item.Fingerprint, StringComparison.Ordinal)) return "SKIPPED";

        DeleteExistingCapture(connection, workbookId);
        Execute(connection, null, """
            UPDATE capture_workbooks
            SET capture_status='CAPTURING', attempts=attempts+1, current_fingerprint=$fingerprint,
                started_at=$startedAt, finished_at='', error_text=''
            WHERE workbook_id=$workbookId;
            """,
            ("$fingerprint", item.Fingerprint), ("$startedAt", UtcNow()), ("$workbookId", workbookId));

        try
        {
            using var archive = ZipFile.OpenRead(item.SourcePath);
            VerifyPackage(archive);
            var workbook = ReadWorkbook(archive);
            var styles = ReadStyles(archive);
            var sharedStrings = ReadSharedStrings(archive);
            var totals = new CaptureTotals();
            using var transaction = connection.BeginTransaction();
            foreach (var sheet in workbook.Sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var entry = FindEntry(archive, sheet.Part) ?? throw new QuarantinedPackageException($"Worksheet relationship target is missing from the package: {sheet.SheetName}");
                var captured = CaptureSheet(connection, transaction, workbookId, sheet, entry, sharedStrings, styles, cancellationToken);
                totals.Add(captured);
            }
            if (!string.Equals(SourceFingerprint(item.SourcePath), item.Fingerprint, StringComparison.Ordinal))
                throw new InvalidOperationException("Source path, size, or modified time changed while numeric capture was running.");
            Execute(connection, transaction, """
                UPDATE capture_workbooks
                SET capture_status='CAPTURED', sheet_count_expected=$expected, sheet_count_captured=$captured,
                    numeric_cell_count=$numeric, formula_count=$formulas, date_cell_count=$dates,
                    table_candidate_count=$tables, merge_count=$merges, finished_at=$finishedAt, error_text=''
                WHERE workbook_id=$workbookId;
                """,
                ("$expected", workbook.Sheets.Count), ("$captured", workbook.Sheets.Count), ("$numeric", totals.Numeric),
                ("$formulas", totals.Formulas), ("$dates", totals.Dates), ("$tables", totals.Tables), ("$merges", totals.Merges),
                ("$finishedAt", UtcNow()), ("$workbookId", workbookId));
            transaction.Commit();
            return "CAPTURED";
        }
        catch (CaptureLimitException exception)
        {
            ClearFailedCapture(connection, workbookId, "TRUNCATED", exception.Message);
            return "TRUNCATED";
        }
        catch (QuarantinedPackageException exception)
        {
            ClearFailedCapture(connection, workbookId, "QUARANTINED", exception.Message);
            return "QUARANTINED";
        }
        catch (InvalidDataException exception)
        {
            ClearFailedCapture(connection, workbookId, "QUARANTINED", exception.Message);
            return "QUARANTINED";
        }
        catch (XmlException exception)
        {
            ClearFailedCapture(connection, workbookId, "QUARANTINED", exception.Message);
            return "QUARANTINED";
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            ClearFailedCapture(connection, workbookId, "FAILED_RETRYABLE", $"{exception.GetType().Name}: {exception.Message}");
            return "FAILED_RETRYABLE";
        }
    }

    private static void DeleteExistingCapture(SqliteConnection connection, long workbookId)
    {
        Execute(connection, null, "DELETE FROM captured_sheets WHERE workbook_id=$workbookId;", ("$workbookId", workbookId));
        Execute(connection, null, """
            UPDATE capture_workbooks SET sheet_count_captured=0, numeric_cell_count=0, formula_count=0,
                date_cell_count=0, table_candidate_count=0, merge_count=0, error_text=''
            WHERE workbook_id=$workbookId;
            """, ("$workbookId", workbookId));
    }

    private static WorkbookCaptureStatus WorkbookStatus(SqliteConnection connection, long workbookId)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT capture_status, current_fingerprint FROM capture_workbooks WHERE workbook_id=$workbookId;";
        command.Parameters.AddWithValue("$workbookId", workbookId);
        using var reader = command.ExecuteReader();
        if (!reader.Read()) throw new InvalidOperationException("적재 워크북 레코드를 찾을 수 없습니다.");
        return new WorkbookCaptureStatus(reader.GetString(0), reader.GetString(1));
    }

    private static void ClearFailedCapture(SqliteConnection connection, long workbookId, string status, string error)
    {
        using var transaction = connection.BeginTransaction();
        Execute(connection, transaction, "DELETE FROM captured_sheets WHERE workbook_id=$workbookId;", ("$workbookId", workbookId));
        Execute(connection, transaction, "UPDATE capture_workbooks SET capture_status=$status, error_text=$error, finished_at=$finishedAt WHERE workbook_id=$workbookId;",
            ("$status", status), ("$error", Clip(error, 2000)), ("$finishedAt", UtcNow()), ("$workbookId", workbookId));
        transaction.Commit();
    }

    private static CaptureTotals CaptureSheet(SqliteConnection connection, SqliteTransaction transaction, long workbookId, WorkbookSheet sheet, ZipArchiveEntry entry, IReadOnlyList<string> sharedStrings, WorkbookStyles styles, CancellationToken cancellationToken)
    {
        var content = ReadSheet(entry, sharedStrings, styles, cancellationToken);
        if ((long)content.MaxRow * content.MaxColumn > MaxCapturedCellsPerSheet)
            throw new CaptureLimitException($"Declared cell count exceeds capture limit ({(long)content.MaxRow * content.MaxColumn} > {MaxCapturedCellsPerSheet}) on {sheet.SheetName}.");
        if (content.MergeRanges.Count > MaxCapturedMergesPerSheet)
            throw new CaptureLimitException($"Merge range count exceeds capture limit ({MaxCapturedMergesPerSheet}) on {sheet.SheetName}.");

        var sheetId = ExecuteInsert(connection, transaction, """
            INSERT INTO captured_sheets(workbook_id, sheet_index, sheet_name, sheet_state, declared_dimension, max_row, max_column, merge_count, capture_status)
            VALUES ($workbookId, $sheetIndex, $sheetName, $sheetState, $dimension, $maxRow, $maxColumn, $mergeCount, 'CAPTURING');
            """,
            ("$workbookId", workbookId), ("$sheetIndex", sheet.SheetIndex), ("$sheetName", sheet.SheetName), ("$sheetState", sheet.SheetState),
            ("$dimension", content.DeclaredDimension), ("$maxRow", content.MaxRow), ("$maxColumn", content.MaxColumn), ("$mergeCount", content.MergeRanges.Count));
        foreach (var merge in content.MergeRanges)
            Execute(connection, transaction, "INSERT INTO captured_merge_ranges(sheet_id, range_ref) VALUES ($sheetId, $range);", ("$sheetId", sheetId), ("$range", merge));

        var totals = new CaptureTotals { Merges = content.MergeRanges.Count };
        var numericPositions = new List<CellPosition>();
        var textByRow = new Dictionary<int, List<TextLabel>>();
        foreach (var cell in content.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (cell.IsFormula)
            {
                totals.Formulas++;
                var cached = CellObservation.From(cell.CachedValue, cell.StyleIndex, styles);
                Execute(connection, transaction, """
                    INSERT INTO formula_cells(sheet_id, row_index, column_index, formula_text, cached_value_text, cached_numeric_value, cached_is_date)
                    VALUES ($sheetId, $row, $column, $formula, $cachedText, $cachedNumber, $cachedIsDate);
                    """,
                    ("$sheetId", sheetId), ("$row", cell.Row), ("$column", cell.Column), ("$formula", cell.FormulaText),
                    ("$cachedText", cached.DisplayText), ("$cachedNumber", cached.Numeric?.Value ?? (object)DBNull.Value), ("$cachedIsDate", cached.Date is null ? 0 : 1));
                if (cached.Date is not null)
                {
                    totals.Dates++;
                    InsertDate(connection, transaction, sheetId, cell.Row, cell.Column, cached.Date.Value, "FORMULA_CACHED");
                }
                else if (cached.Numeric is not null)
                {
                    totals.Numeric++;
                    numericPositions.Add(new CellPosition(cell.Row, cell.Column));
                    InsertNumeric(connection, transaction, sheetId, cell.Row, cell.Column, "FORMULA_CACHED", cached.Numeric.Value, cell.FormulaText);
                }
                continue;
            }
            var observation = CellObservation.From(cell.Value, cell.StyleIndex, styles);
            if (observation.Date is not null)
            {
                totals.Dates++;
                InsertDate(connection, transaction, sheetId, cell.Row, cell.Column, observation.Date.Value, "DATE");
                continue;
            }
            if (observation.Numeric is not null)
            {
                totals.Numeric++;
                numericPositions.Add(new CellPosition(cell.Row, cell.Column));
                InsertNumeric(connection, transaction, sheetId, cell.Row, cell.Column, observation.Numeric.Value.Kind, observation.Numeric.Value, string.Empty);
                continue;
            }
            if (!string.IsNullOrWhiteSpace(observation.Text))
            {
                InsertText(connection, transaction, sheetId, cell.Row, cell.Column, observation.Text.Trim());
                if (observation.Text.Trim().Length > MaxHeaderLabelLength) continue;
                if (!textByRow.TryGetValue(cell.Row, out var labels)) textByRow[cell.Row] = labels = [];
                labels.Add(new TextLabel(cell.Row, cell.Column, observation.Text.Trim()));
            }
        }

        foreach (var region in NumericTableRegions(numericPositions))
        {
            var headerStart = Math.Max(1, region.StartRow - MaxHeaderRowsPerRegion);
            var headerEnd = Math.Max(headerStart, region.StartRow - 1);
            var headerEntries = Enumerable.Range(headerStart, headerEnd - headerStart + 1)
                .SelectMany(row => textByRow.TryGetValue(row, out var labels) ? labels : [])
                .Where(label => label.Column >= region.StartColumn && label.Column <= region.EndColumn).ToList();
            var rowLabels = Enumerable.Range(region.StartRow, region.EndRow - region.StartRow + 1)
                .SelectMany(row => textByRow.TryGetValue(row, out var labels) ? labels : [])
                .Where(label => label.Column >= region.StartColumn - 3 && label.Column <= region.EndColumn + 3).ToList();
            var (type, confidence) = CandidateType(headerEntries.Select(static item => item.Text));
            var tableId = ExecuteInsert(connection, transaction, """
                INSERT INTO numeric_table_candidates(sheet_id, start_row, end_row, start_column, end_column, header_start_row, header_end_row, numeric_cell_count, candidate_type, confidence)
                VALUES ($sheetId, $startRow, $endRow, $startColumn, $endColumn, $headerStart, $headerEnd, $numericCount, $type, $confidence);
                """,
                ("$sheetId", sheetId), ("$startRow", region.StartRow), ("$endRow", region.EndRow), ("$startColumn", region.StartColumn), ("$endColumn", region.EndColumn),
                ("$headerStart", headerStart), ("$headerEnd", headerEnd), ("$numericCount", region.NumericCellCount), ("$type", type), ("$confidence", confidence));
            foreach (var label in headerEntries)
                Execute(connection, transaction, "INSERT INTO numeric_table_labels(table_id, row_index, column_index, label_text, label_role) VALUES ($tableId, $row, $column, $text, 'HEADER');",
                    ("$tableId", tableId), ("$row", label.Row), ("$column", label.Column), ("$text", label.Text));
            foreach (var label in rowLabels)
                Execute(connection, transaction, "INSERT OR IGNORE INTO numeric_table_labels(table_id, row_index, column_index, label_text, label_role) VALUES ($tableId, $row, $column, $text, 'ROW_LABEL');",
                    ("$tableId", tableId), ("$row", label.Row), ("$column", label.Column), ("$text", label.Text));
            totals.Tables++;
        }
        Execute(connection, transaction, "UPDATE captured_sheets SET numeric_cell_count=$numeric, formula_count=$formulas, date_cell_count=$dates, capture_status='CAPTURED' WHERE sheet_id=$sheetId;",
            ("$numeric", totals.Numeric), ("$formulas", totals.Formulas), ("$dates", totals.Dates), ("$sheetId", sheetId));
        return totals;
    }

    private static void InsertNumeric(SqliteConnection connection, SqliteTransaction transaction, long sheetId, int row, int column, string kind, NumericValue value, string formula)
    {
        Execute(connection, transaction, """
            INSERT INTO numeric_cells(sheet_id, row_index, column_index, source_kind, value_text, numeric_value, formula_text)
            VALUES ($sheetId, $row, $column, $kind, $text, $value, $formula);
            """,
            ("$sheetId", sheetId), ("$row", row), ("$column", column), ("$kind", kind), ("$text", value.Text), ("$value", value.Value), ("$formula", formula));
    }

    private static void InsertText(SqliteConnection connection, SqliteTransaction transaction, long sheetId, int row, int column, string value) =>
        Execute(connection, transaction, "INSERT INTO captured_text_cells(sheet_id, row_index, column_index, text_value) VALUES ($sheetId, $row, $column, $value);",
            ("$sheetId", sheetId), ("$row", row), ("$column", column), ("$value", value));

    private static void InsertDate(SqliteConnection connection, SqliteTransaction transaction, long sheetId, int row, int column, DateOnly date, string kind) =>
        Execute(connection, transaction, "INSERT INTO date_cells(sheet_id, row_index, column_index, date_value, source_kind) VALUES ($sheetId, $row, $column, $date, $kind);",
            ("$sheetId", sheetId), ("$row", row), ("$column", column), ("$date", date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)), ("$kind", kind));

    private static SheetContent ReadSheet(ZipArchiveEntry entry, IReadOnlyList<string> sharedStrings, WorkbookStyles styles, CancellationToken cancellationToken)
    {
        var result = new SheetContent();
        using var reader = CreateXmlReader(entry, MaxWorksheetUncompressedBytes);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (reader.LocalName == "dimension" && string.IsNullOrEmpty(result.DeclaredDimension))
            {
                result.DeclaredDimension = reader.GetAttribute("ref") ?? string.Empty;
                var bounds = DimensionBounds(result.DeclaredDimension);
                if (bounds is not null)
                {
                    result.MaxRow = bounds.Value.Last.Row;
                    result.MaxColumn = bounds.Value.Last.Column;
                }
            }
            else if (reader.LocalName == "mergeCell")
            {
                var reference = reader.GetAttribute("ref");
                if (!string.IsNullOrWhiteSpace(reference)) result.MergeRanges.Add(reference);
                if (result.MergeRanges.Count > MaxCapturedMergesPerSheet) throw new CaptureLimitException($"Merge range count exceeds capture limit ({MaxCapturedMergesPerSheet}).");
            }
            else if (reader.LocalName == "c")
            {
                var cell = ReadCell(reader, sharedStrings);
                result.Cells.Add(cell);
                result.MaxRow = Math.Max(result.MaxRow, cell.Row);
                result.MaxColumn = Math.Max(result.MaxColumn, cell.Column);
            }
        }
        return result;
    }

    private static CellData ReadCell(XmlReader reader, IReadOnlyList<string> sharedStrings)
    {
        var reference = reader.GetAttribute("r") ?? string.Empty;
        var type = reader.GetAttribute("t") ?? string.Empty;
        var style = int.TryParse(reader.GetAttribute("s"), NumberStyles.None, CultureInfo.InvariantCulture, out var styleIndex) ? styleIndex : 0;
        var (row, column) = CellAddress(reference);
        var hasFormula = false;
        var formula = string.Empty;
        var rawValue = string.Empty;
        var inlineText = new StringBuilder();
        using var subtree = reader.ReadSubtree();
        while (subtree.Read())
        {
            if (subtree.NodeType != XmlNodeType.Element) continue;
            if (subtree.LocalName == "f")
            {
                hasFormula = true;
                formula = subtree.IsEmptyElement ? string.Empty : subtree.ReadElementContentAsString();
            }
            else if (subtree.LocalName == "v") rawValue = subtree.ReadElementContentAsString();
            else if (subtree.LocalName == "t") inlineText.Append(subtree.ReadElementContentAsString());
        }
        var value = ResolveValue(type, rawValue, inlineText.ToString(), sharedStrings);
        var formulaText = hasFormula ? "=" + formula : value.Text.StartsWith("=", StringComparison.Ordinal) ? value.Text : string.Empty;
        return new CellData(row, column, style, value, hasFormula || !string.IsNullOrEmpty(formulaText), formulaText, ResolveValue(type, rawValue, inlineText.ToString(), sharedStrings));
    }

    private static RawValue ResolveValue(string type, string rawValue, string inlineText, IReadOnlyList<string> sharedStrings)
    {
        if (type == "s" && int.TryParse(rawValue, NumberStyles.None, CultureInfo.InvariantCulture, out var sharedIndex) && sharedIndex >= 0 && sharedIndex < sharedStrings.Count)
            return new RawValue("TEXT", sharedStrings[sharedIndex]);
        if (type == "inlineStr") return new RawValue("TEXT", inlineText);
        if (type == "str") return new RawValue("TEXT", rawValue);
        if (type == "b") return new RawValue("BOOLEAN", rawValue);
        if (type == "e") return new RawValue("ERROR", rawValue);
        return new RawValue("NUMBER", rawValue);
    }

    private static WorkbookDocument ReadWorkbook(ZipArchive archive)
    {
        VerifyPackage(archive);
        var relationships = ReadRelationships(archive);
        var workbookEntry = FindEntry(archive, "xl/workbook.xml") ?? throw new QuarantinedPackageException("OpenXML workbook metadata is missing.");
        var sheets = new List<WorkbookSheet>();
        using var reader = CreateXmlReader(workbookEntry, MaxWorksheetUncompressedBytes);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "sheet") continue;
            var id = reader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") ?? string.Empty;
            if (!relationships.TryGetValue(id, out var part)) throw new QuarantinedPackageException("Worksheet relationship target is missing from the package.");
            sheets.Add(new WorkbookSheet(sheets.Count + 1, reader.GetAttribute("name") ?? $"Sheet{sheets.Count + 1}", reader.GetAttribute("state") ?? "visible", part));
        }
        return new WorkbookDocument(sheets);
    }

    private static Dictionary<string, string> ReadRelationships(ZipArchive archive)
    {
        var entry = FindEntry(archive, "xl/_rels/workbook.xml.rels") ?? throw new QuarantinedPackageException("OpenXML workbook relationship metadata is missing.");
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        using var reader = CreateXmlReader(entry, MaxWorksheetUncompressedBytes);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship") continue;
            var id = reader.GetAttribute("Id");
            if (!string.IsNullOrWhiteSpace(id)) values[id] = RelationshipTarget(reader.GetAttribute("Target") ?? string.Empty);
        }
        return values;
    }

    private static List<string> ReadSharedStrings(ZipArchive archive)
    {
        var entry = FindEntry(archive, "xl/sharedStrings.xml");
        if (entry is null) return [];
        var values = new List<string>();
        using var reader = CreateXmlReader(entry, MaxWorksheetUncompressedBytes);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "si") continue;
            var text = new StringBuilder();
            using var item = reader.ReadSubtree();
            while (item.Read())
                if (item.NodeType == XmlNodeType.Element && item.LocalName == "t") text.Append(item.ReadElementContentAsString());
            values.Add(text.ToString());
        }
        return values;
    }

    private static WorkbookStyles ReadStyles(ZipArchive archive)
    {
        var entry = FindEntry(archive, "xl/styles.xml");
        if (entry is null) return WorkbookStyles.Empty;
        var customFormats = new Dictionary<int, string>();
        var styleFormats = new List<int>();
        var inCellXfs = false;
        using var reader = CreateXmlReader(entry, MaxWorksheetUncompressedBytes);
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "numFmt" &&
                int.TryParse(reader.GetAttribute("numFmtId"), NumberStyles.None, CultureInfo.InvariantCulture, out var id))
                customFormats[id] = reader.GetAttribute("formatCode") ?? string.Empty;
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "cellXfs") inCellXfs = true;
            else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "cellXfs") inCellXfs = false;
            else if (inCellXfs && reader.NodeType == XmlNodeType.Element && reader.LocalName == "xf")
                styleFormats.Add(int.TryParse(reader.GetAttribute("numFmtId"), NumberStyles.None, CultureInfo.InvariantCulture, out var format) ? format : 0);
        }
        return new WorkbookStyles(customFormats, styleFormats);
    }

    private static void VerifyPackage(ZipArchive archive)
    {
        if (archive.Entries.Count > MaxZipEntries) throw new QuarantinedPackageException($"ZIP entry count exceeds limit ({archive.Entries.Count} > {MaxZipEntries}).");
        var names = new HashSet<string>(StringComparer.Ordinal);
        long compressed = 0;
        long uncompressed = 0;
        foreach (var entry in archive.Entries)
        {
            if (!names.Add(entry.FullName)) throw new QuarantinedPackageException($"ZIP package contains a duplicate entry: {entry.FullName}");
            try { compressed = checked(compressed + entry.CompressedLength); uncompressed = checked(uncompressed + entry.Length); }
            catch (OverflowException) { throw new QuarantinedPackageException("ZIP package size overflow."); }
        }
        if (uncompressed > MaxPackageUncompressedBytes) throw new QuarantinedPackageException("ZIP package uncompressed size exceeds scan limit.");
        if (compressed > 0 && uncompressed / compressed > MaxCompressionRatio) throw new QuarantinedPackageException("ZIP compression ratio exceeds scan limit.");
        if (!names.Contains("xl/workbook.xml") || !names.Contains("xl/_rels/workbook.xml.rels")) throw new QuarantinedPackageException("OpenXML workbook metadata is missing.");
    }

    private static List<NumericRegion> NumericTableRegions(IReadOnlyList<CellPosition> positions)
    {
        var byRow = positions.GroupBy(static value => value.Row).OrderBy(static group => group.Key).ToList();
        if (byRow.Count == 0) return [];
        var rowBands = new List<List<IGrouping<int, CellPosition>>>();
        var current = new List<IGrouping<int, CellPosition>>();
        foreach (var row in byRow)
        {
            if (current.Count > 0 && row.Key - current[^1].Key > 3) { rowBands.Add(current); current = []; }
            current.Add(row);
        }
        if (current.Count > 0) rowBands.Add(current);
        var regions = new List<NumericRegion>();
        foreach (var rows in rowBands)
        {
            var columns = rows.SelectMany(static group => group).Select(static value => value.Column).Distinct().OrderBy(static value => value).ToList();
            var columnBands = new List<List<int>>();
            var currentColumns = new List<int>();
            foreach (var column in columns)
            {
                if (currentColumns.Count > 0 && column - currentColumns[^1] > MaxTableColumnGap) { columnBands.Add(currentColumns); currentColumns = []; }
                currentColumns.Add(column);
            }
            if (currentColumns.Count > 0) columnBands.Add(currentColumns);
            foreach (var band in columnBands)
            {
                var start = band.Min();
                var end = band.Max();
                var count = rows.SelectMany(static group => group).Count(position => position.Column >= start && position.Column <= end);
                regions.Add(new NumericRegion(rows[0].Key, rows[^1].Key, start, end, count));
            }
        }
        return regions;
    }

    private static (string Type, string Confidence) CandidateType(IEnumerable<string> headers)
    {
        var values = headers.Select(HeaderToken).Where(static value => value is not null).Cast<string>().ToList();
        var tokens = values.ToHashSet(StringComparer.Ordinal);
        if (values.Count(static value => value == "INPUT") >= 2 && tokens.Contains("TOTAL_NG") && tokens.Contains("NG_RATE")) return ("REPEATED_DEFECT_BLOCK_NUMERIC_TABLE", "HIGH");
        if (new[] { "INPUT", "OK", "TOTAL_NG", "NG_RATE" }.All(tokens.Contains)) return ("DEFECT_RATE_NUMERIC_TABLE", "HIGH");
        if (new[] { "SAMPLE", "AVERAGE", "MAX", "MIN" }.All(tokens.Contains)) return ("MEASUREMENT_SUMMARY_NUMERIC_TABLE", "HIGH");
        if (tokens.Contains("NORMAL_CUE") && tokens.Contains("TEST_CUE")) return ("TEST_NORMAL_NUMERIC_TABLE", "MEDIUM");
        return ("NUMERIC_TABLE_UNCLASSIFIED", "LOW");
    }

    private static string? HeaderToken(string value)
    {
        var normalized = NormalizeLabel(value);
        return normalized switch
        {
            "input" => "INPUT", "ok" => "OK", "totalng" or "ngtotal" or "totaldefect" or "totaldefects" => "TOTAL_NG",
            "ngrate" or "defectrate" or "totalngrate" => "NG_RATE", "sample" or "samples" => "SAMPLE",
            "average" or "avg" or "mean" => "AVERAGE", "max" or "maximum" => "MAX", "min" or "minimum" => "MIN",
            "normal" or "baseline" or "before" => "NORMAL_CUE", "test" or "trial" or "after" => "TEST_CUE",
            _ when normalized == "ng" || normalized.StartsWith("ng", StringComparison.Ordinal) => "NG_CUE",
            _ => null,
        };
    }

    private static string NormalizeLabel(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var builder = new StringBuilder(value.Length);
        foreach (var character in value.Trim().ToLowerInvariant())
            if (char.IsAsciiLetterOrDigit(character) || (character >= '가' && character <= '힣')) builder.Append(character);
        return builder.ToString();
    }

    private static NumericCaptureSummary WriteOutputs(string batchDirectory, SqliteConnection connection)
    {
        var rows = CaptureWorkbookRows(connection);
        var statuses = rows.GroupBy(static row => row.CaptureStatus, StringComparer.Ordinal).OrderBy(static group => group.Key, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.Count(), StringComparer.Ordinal);
        var sourceStatuses = rows.GroupBy(static row => row.StructureScanStatus, StringComparer.Ordinal).OrderBy(static group => group.Key, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.Count(), StringComparer.Ordinal);
        var totals = rows.Where(static row => row.CaptureStatus == "CAPTURED").Aggregate(new CaptureTotals(), static (total, row) => new CaptureTotals
        {
            Numeric = total.Numeric + row.NumericCells,
            Formulas = total.Formulas + row.Formulas,
            Dates = total.Dates + row.Dates,
            Tables = total.Tables + row.Tables,
        });
        var summary = new NumericCaptureSummary(
            "numeric-capture-summary-v1", CaptureVersion, UtcNow(), false, rows.Count,
            new ReadOnlyDictionary<string, int>(statuses), new ReadOnlyDictionary<string, int>(sourceStatuses),
            totals.Numeric, totals.Formulas, totals.Dates, totals.Tables, "numeric-capture.sqlite", "numeric-capture.csv",
            [
                "All non-empty text cells, numeric cells, dates, formulas, merge ranges, and adjacent table labels are captured with stable sheet coordinates.",
                "Formulas are never recalculated; cached values are retained only when present in the workbook.",
                "No Test–Normal comparison or quality conclusion is produced by this capture stage.",
            ]);
        AtomicWriteJson(Path.Combine(batchDirectory, "numeric-capture-summary.json"), summary);
        var reports = rows.Select(row => new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["relativePath"] = row.RelativePath, ["status"] = row.CaptureStatus, ["sourceScanStatus"] = row.StructureScanStatus,
            ["sheetCount"] = row.SheetCount.ToString(CultureInfo.InvariantCulture), ["numericCells"] = row.NumericCells.ToString(CultureInfo.InvariantCulture),
            ["formulas"] = row.Formulas.ToString(CultureInfo.InvariantCulture), ["dates"] = row.Dates.ToString(CultureInfo.InvariantCulture),
            ["numericTableCandidates"] = row.Tables.ToString(CultureInfo.InvariantCulture), ["warningOrError"] = row.ErrorText,
        }).ToList();
        WriteCsv(Path.Combine(batchDirectory, "numeric-capture.csv"), reports);
        WriteHtml(Path.Combine(batchDirectory, "numeric-capture.html"), reports, statuses);
        return summary;
    }

    private static List<CaptureWorkbookRow> CaptureWorkbookRows(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT relative_path, capture_status, structure_scan_status, sheet_count_captured, numeric_cell_count, formula_count, date_cell_count, table_candidate_count, error_text FROM capture_workbooks ORDER BY relative_path;";
        using var reader = command.ExecuteReader();
        var rows = new List<CaptureWorkbookRow>();
        while (reader.Read()) rows.Add(new CaptureWorkbookRow(reader.GetString(0), reader.GetString(1), reader.GetString(2), reader.GetInt32(3), reader.GetInt32(4), reader.GetInt32(5), reader.GetInt32(6), reader.GetInt32(7), reader.GetString(8)));
        return rows;
    }

    private static void WriteCsv(string path, IReadOnlyList<Dictionary<string, string>> rows)
    {
        var columns = rows.Count > 0 ? rows[0].Keys.ToList() : new List<string> { "relativePath", "status" };
        var builder = new StringBuilder();
        builder.AppendLine(string.Join(',', columns.Select(Csv)));
        foreach (var row in rows) builder.AppendLine(string.Join(',', columns.Select(column => Csv(row.GetValueOrDefault(column, string.Empty)))));
        AtomicWriteText(path, builder.ToString(), new UTF8Encoding(true));
    }

    private static void WriteHtml(string path, IReadOnlyList<Dictionary<string, string>> rows, IReadOnlyDictionary<string, int> statuses)
    {
        var columns = rows.Count > 0 ? rows[0].Keys.ToList() : new List<string> { "relativePath", "status" };
        var statusRows = string.Concat(statuses.Select(pair => $"<li>{Html(pair.Key)}: {pair.Value}</li>"));
        var header = string.Concat(columns.Select(column => $"<th>{Html(column)}</th>"));
        var body = string.Concat(rows.Select(row => "<tr>" + string.Concat(columns.Select(column => $"<td>{Html(row.GetValueOrDefault(column, string.Empty))}</td>")) + "</tr>"));
        var document = "<!doctype html><html lang='ko'><head><meta charset='utf-8'><title>숫자 표 원본 적재</title>\n"
            + "<style>body{font-family:Segoe UI,sans-serif;margin:24px;color:#172b4d}table{border-collapse:collapse;width:100%;font-size:12px}th,td{border:1px solid #d7dde7;padding:6px;text-align:left;vertical-align:top}th{background:#edf2f7}</style></head>\n"
            + $"<body><h1>숫자 표 원본 적재</h1><p>Excel/COM을 사용하지 않았습니다. 이 결과는 숫자 사실 적재 현황이며, 비교·품질 판정을 포함하지 않습니다.</p><h2>상태</h2><ul>{statusRows}</ul><p><a href='numeric-capture.csv'>CSV</a> · <a href='numeric-capture-summary.json'>JSON</a></p><table><thead><tr>{header}</tr></thead><tbody>{body}</tbody></table></body></html>";
        AtomicWriteText(path, document, new UTF8Encoding(false));
    }

    private static string Csv(string value) => $"\"{value.Replace("\"", "\"\"")}\"";
    private static string Html(string value) => System.Net.WebUtility.HtmlEncode(value);

    private static void Execute(SqliteConnection connection, SqliteTransaction? transaction, string sql, params (string Name, object Value)[] parameters)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = sql;
        foreach (var (name, value) in parameters) command.Parameters.AddWithValue(name, value);
        command.ExecuteNonQuery();
    }

    private static long ExecuteInsert(SqliteConnection connection, SqliteTransaction transaction, string sql, params (string Name, object Value)[] parameters)
    {
        Execute(connection, transaction, sql, parameters);
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = "SELECT last_insert_rowid();";
        return Convert.ToInt64(command.ExecuteScalar(), CultureInfo.InvariantCulture);
    }

    private static ZipArchiveEntry? FindEntry(ZipArchive archive, string name) => archive.Entries.FirstOrDefault(entry => string.Equals(entry.FullName, name, StringComparison.Ordinal));

    private static XmlReader CreateXmlReader(ZipArchiveEntry entry, long maxCharacters) => XmlReader.Create(entry.Open(), new XmlReaderSettings
    {
        DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreWhitespace = true,
        MaxCharactersInDocument = maxCharacters, MaxCharactersFromEntities = 0, CloseInput = true,
    });

    private static string RelationshipTarget(string target)
    {
        var combined = target.Replace('\\', '/').Trim();
        combined = combined.StartsWith("/", StringComparison.Ordinal) ? combined.TrimStart('/') : $"xl/{combined}";
        var parts = new Stack<string>();
        foreach (var part in combined.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (part == ".") continue;
            if (part == "..") { if (parts.Count == 0) throw new QuarantinedPackageException("Worksheet relationship target escapes the workbook package."); parts.Pop(); continue; }
            parts.Push(part);
        }
        var result = string.Join('/', parts.Reverse());
        if (!result.StartsWith("xl/", StringComparison.Ordinal)) throw new QuarantinedPackageException("Worksheet relationship target escapes the workbook package.");
        return result;
    }

    private static (int Row, int Column) CellAddress(string address)
    {
        if (string.IsNullOrWhiteSpace(address)) return (0, 0);
        var index = 0;
        var column = 0;
        while (index < address.Length && char.IsLetter(address[index])) { column = checked(column * 26 + char.ToUpperInvariant(address[index]) - 'A' + 1); index++; }
        return int.TryParse(address[index..], NumberStyles.None, CultureInfo.InvariantCulture, out var row) ? (row, column) : (0, 0);
    }

    private static ((int Row, int Column) First, (int Row, int Column) Last)? DimensionBounds(string dimension)
    {
        if (string.IsNullOrWhiteSpace(dimension)) return null;
        var parts = dimension.ToUpperInvariant().Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length is < 1 or > 2) return null;
        return (CellAddress(parts[0]), CellAddress(parts[^1]));
    }

    private static string SourceFingerprint(string path)
    {
        var info = new FileInfo(Path.GetFullPath(path));
        var mtimeNs = checked((info.LastWriteTimeUtc.Ticks - DateTime.UnixEpoch.Ticks) * 100L);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes($"{info.FullName}|{info.Length}|{mtimeNs}"))).ToLowerInvariant();
    }

    private static string UtcNow() => DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ", CultureInfo.InvariantCulture);
    private static void Increment(IDictionary<string, int> values, string value) => values[value] = values.TryGetValue(value, out var existing) ? existing + 1 : 1;
    private static bool IsWithin(string child, string parent)
    {
        var relative = Path.GetRelativePath(Path.GetFullPath(parent), Path.GetFullPath(child));
        return !string.Equals(relative, "..", StringComparison.Ordinal) && !relative.StartsWith($"..{Path.DirectorySeparatorChar}", StringComparison.Ordinal) && !Path.IsPathRooted(relative);
    }
    private static void AtomicWriteJson(string path, object value) => AtomicWriteText(path, JsonSerializer.Serialize(value, JsonOptions) + Environment.NewLine, new UTF8Encoding(false));
    private static void AtomicWriteText(string path, string content, Encoding encoding)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var temporary = path + ".tmp";
        File.WriteAllText(temporary, content, encoding);
        File.Move(temporary, path, overwrite: true);
    }
    private static string Clip(string value, int maxLength) => value.Length <= maxLength ? value : value[..maxLength];

    private sealed class CaptureLimitException(string message) : Exception(message);
    private sealed class QuarantinedPackageException(string message) : Exception(message);

    private sealed record SourceItem(string RelativePath, string SourcePath, string Extension, string Kind, string Fingerprint, string Status);
    private sealed record WorkbookCaptureStatus(string CaptureStatus, string CurrentFingerprint);
    private sealed record WorkbookSheet(int SheetIndex, string SheetName, string SheetState, string Part);
    private sealed record WorkbookDocument(List<WorkbookSheet> Sheets);
    private sealed record RawValue(string Kind, string Text);
    private sealed record CellData(int Row, int Column, int StyleIndex, RawValue Value, bool IsFormula, string FormulaText, RawValue CachedValue);
    private sealed record CellPosition(int Row, int Column);
    private sealed record TextLabel(int Row, int Column, string Text);
    private sealed record NumericRegion(int StartRow, int EndRow, int StartColumn, int EndColumn, int NumericCellCount);
    private sealed record CaptureWorkbookRow(string RelativePath, string CaptureStatus, string StructureScanStatus, int SheetCount, int NumericCells, int Formulas, int Dates, int Tables, string ErrorText);

    private sealed class SheetContent
    {
        public string DeclaredDimension { get; set; } = string.Empty;
        public int MaxRow { get; set; }
        public int MaxColumn { get; set; }
        public List<string> MergeRanges { get; } = [];
        public List<CellData> Cells { get; } = [];
    }

    private sealed class CaptureTotals
    {
        public int Numeric { get; set; }
        public int Formulas { get; set; }
        public int Dates { get; set; }
        public int Tables { get; set; }
        public int Merges { get; set; }
        public void Add(CaptureTotals other) { Numeric += other.Numeric; Formulas += other.Formulas; Dates += other.Dates; Tables += other.Tables; Merges += other.Merges; }
    }

    private readonly record struct NumericValue(string Text, double Value, string Kind);

    private readonly record struct CellObservation(string? Text, NumericValue? Numeric, DateOnly? Date, string DisplayText)
    {
        public static CellObservation From(RawValue raw, int styleIndex, WorkbookStyles styles)
        {
            if (raw.Kind == "BOOLEAN") return new CellObservation(null, null, null, raw.Text);
            // openpyxl exposes cell errors (for example #DIV/0!) as strings.
            // Preserve them as short table-adjacent labels so the v1 DB matches
            // the existing Python capture contract, but never treat them as a number.
            if (raw.Kind == "ERROR") return new CellObservation(raw.Text, null, null, raw.Text);
            if (raw.Kind == "TEXT")
            {
                var numeric = NumericText(raw.Text);
                return new CellObservation(raw.Text, numeric, null, raw.Text);
            }
            if (string.IsNullOrWhiteSpace(raw.Text)) return new CellObservation(null, null, null, string.Empty);
            if (!decimal.TryParse(raw.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var decimalValue)) return new CellObservation(null, null, null, raw.Text);
            if (styles.IsDateStyle(styleIndex) && TryExcelDate(decimalValue, out var date))
                return new CellObservation(null, null, date, date.ToDateTime(TimeOnly.MinValue).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            if (!double.TryParse(raw.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number)) return new CellObservation(null, null, null, raw.Text);
            // Python's openpyxl path first creates a float, then uses
            // Decimal(str(float)).  Format the double first so XML artifacts
            // such as 0.07000000000000001 become the same 0.07 evidence text.
            var pythonCompatible = number.ToString(CultureInfo.InvariantCulture);
            var evidenceText = decimal.TryParse(pythonCompatible, NumberStyles.Float, CultureInfo.InvariantCulture, out var compatibilityDecimal)
                ? FixedDecimalText(compatibilityDecimal)
                : ExpandScientificNumber(pythonCompatible);
            return new CellObservation(null, new NumericValue(evidenceText, number, "NUMBER"), null, raw.Text);
        }

        private static NumericValue? NumericText(string text)
        {
            var value = text.Trim();
            if (value.Length == 0 || value.Length > 128) return null;
            var compact = value.Replace(",", string.Empty, StringComparison.Ordinal).Replace(" ", string.Empty, StringComparison.Ordinal);
            var sign = string.Empty;
            if (compact.StartsWith('(') && compact.EndsWith(')')) { sign = "-"; compact = compact[1..^1]; }
            var percent = compact.EndsWith('%');
            if (percent) compact = compact[..^1];
            if (!System.Text.RegularExpressions.Regex.IsMatch(sign + compact, @"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")) return null;
            if (!decimal.TryParse(sign + compact, NumberStyles.Float, CultureInfo.InvariantCulture, out var decimalValue)) return null;
            if (percent) decimalValue /= 100m;
            var number = (double)decimalValue;
            return double.IsFinite(number) ? new NumericValue(FixedDecimalText(decimalValue), number, percent ? "TEXT_PERCENT" : "TEXT_NUMBER") : null;
        }

        private static string FixedDecimalText(decimal value)
        {
            var bits = decimal.GetBits(value);
            var coefficient = (BigInteger)(uint)bits[0]
                | ((BigInteger)(uint)bits[1] << 32)
                | ((BigInteger)(uint)bits[2] << 64);
            var scale = (bits[3] >> 16) & 0x7f;
            var digits = coefficient.ToString(CultureInfo.InvariantCulture);
            var sign = (bits[3] & unchecked((int)0x80000000)) != 0 ? "-" : string.Empty;
            if (scale == 0) return sign + digits;
            if (digits.Length <= scale) return sign + "0." + new string('0', scale - digits.Length) + digits;
            return sign + digits[..^scale] + "." + digits[^scale..];
        }

        private static string ExpandScientificNumber(string value)
        {
            var exponentIndex = value.IndexOfAny(['E', 'e']);
            if (exponentIndex < 0) return value;
            var mantissa = value[..exponentIndex];
            if (!int.TryParse(value[(exponentIndex + 1)..], NumberStyles.Integer, CultureInfo.InvariantCulture, out var exponent)) return value;
            var sign = string.Empty;
            if (mantissa.StartsWith('-')) { sign = "-"; mantissa = mantissa[1..]; }
            else if (mantissa.StartsWith('+')) mantissa = mantissa[1..];
            var point = mantissa.IndexOf('.');
            var digits = mantissa.Replace(".", string.Empty, StringComparison.Ordinal);
            var position = (point < 0 ? mantissa.Length : point) + exponent;
            if (position <= 0) return sign + "0." + new string('0', -position) + digits;
            if (position >= digits.Length) return sign + digits + new string('0', position - digits.Length);
            return sign + digits[..position] + "." + digits[position..];
        }

        private static bool TryExcelDate(decimal value, out DateOnly date)
        {
            try
            {
                var dateTime = new DateTime(1899, 12, 30).AddDays((double)value);
                date = DateOnly.FromDateTime(dateTime);
                return true;
            }
            catch (ArgumentOutOfRangeException) { date = default; return false; }
        }
    }

    private sealed class WorkbookStyles(IReadOnlyDictionary<int, string> customFormats, IReadOnlyList<int> styleFormats)
    {
        private static readonly HashSet<int> BuiltInDateFormats = [14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58];
        public static WorkbookStyles Empty { get; } = new(new Dictionary<int, string>(), []);

        public bool IsDateStyle(int styleIndex)
        {
            if (styleIndex < 0 || styleIndex >= styleFormats.Count) return false;
            var formatId = styleFormats[styleIndex];
            if (BuiltInDateFormats.Contains(formatId)) return true;
            return customFormats.TryGetValue(formatId, out var format) && IsDateFormat(format);
        }

        private static bool IsDateFormat(string format)
        {
            var cleaned = new StringBuilder();
            var quoted = false;
            for (var index = 0; index < format.Length; index++)
            {
                var character = format[index];
                if (character == '"') { quoted = !quoted; continue; }
                if (quoted) continue;
                if (character == '\\' && index + 1 < format.Length) { index++; continue; }
                if (character == '[')
                {
                    while (index < format.Length && format[index] != ']') index++;
                    continue;
                }
                cleaned.Append(char.ToLowerInvariant(character));
            }
            var value = cleaned.ToString();
            return value.Contains('y') || value.Contains('d') || value.Contains('h') || value.Contains('s') || value.Contains("am/pm", StringComparison.Ordinal);
        }
    }
}

internal sealed record NumericCaptureRequest(string ServiceDirectory, string StructureBatchId, int Limit = 0, int ProgressEvery = 25, bool Force = false);
internal sealed record NumericCaptureRunResult(string BatchDirectory, IReadOnlyDictionary<string, int> RunStatusCounts, NumericCaptureSummary Summary);
internal sealed record NumericCaptureSummary(
    string SchemaVersion,
    string CaptureVersion,
    string GeneratedAt,
    bool UsesCom,
    int SourceWorkbookCount,
    IReadOnlyDictionary<string, int> CaptureStatusCounts,
    IReadOnlyDictionary<string, int> StructureScanStatusCounts,
    int CapturedNumericCellCount,
    int CapturedFormulaCount,
    int CapturedDateCellCount,
    int NumericTableCandidateCount,
    string Database,
    string ClassificationCsv,
    IReadOnlyList<string> Limitations);
