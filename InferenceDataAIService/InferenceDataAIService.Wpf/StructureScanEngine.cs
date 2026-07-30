using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Xml;
using Microsoft.Data.Sqlite;

namespace InferenceDataAIService.Wpf;

// This service deliberately has no WPF references.  It is the first C# engine
// replacement and preserves the structure-scan-v1 batch contract consumed by
// the remaining numeric Python compatibility helpers.
internal sealed class StructureScanEngine
{
    private const string ScannerVersion = "structure-scan-v1";
    private const int MaxZipEntries = 20_000;
    private const long MaxPackageUncompressedBytes = 512L * 1024 * 1024;
    private const long MaxWorksheetUncompressedBytes = 96L * 1024 * 1024;
    private const int MaxCompressionRatio = 250;
    private const long MaxDeclaredCells = 2_000_000;
    private const int MaxScannedCells = 250_000;
    private const int MaxMergeSample = 200;
    private const int MaxHeaderRows = 80;
    private const int MaxHeaderTokensPerRow = 12;

    private static readonly HashSet<string> OpenXmlExtensions = new(StringComparer.OrdinalIgnoreCase) { ".xlsx", ".xlsm" };
    private static readonly HashSet<string> BinaryExtensions = new(StringComparer.OrdinalIgnoreCase) { ".xls", ".xlsb" };
    private static readonly HashSet<string> InventoryExtensions = new(StringComparer.OrdinalIgnoreCase) { ".xlsx", ".xlsm", ".xls", ".xlsb", ".html", ".htm" };
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    private readonly Action<string>? _log;

    private StructureScanEngine(Action<string>? log) => _log = log;

    internal static StructureScanRunResult Run(StructureScanRequest request, Action<string>? log = null, CancellationToken cancellationToken = default) =>
        new StructureScanEngine(log).RunCore(request, cancellationToken);

    private StructureScanRunResult RunCore(StructureScanRequest request, CancellationToken cancellationToken)
    {
        var serviceDirectory = Path.GetFullPath(request.ServiceDirectory);
        if (!Directory.Exists(serviceDirectory)) throw new ArgumentException($"서비스 폴더를 찾을 수 없습니다: {serviceDirectory}");
        if (string.IsNullOrWhiteSpace(request.BatchFolder) == string.IsNullOrWhiteSpace(request.ResumeBatchId))
            throw new ArgumentException("새 배치 폴더 또는 재개할 배치 ID 중 하나만 지정해야 합니다.");
        if (request.Pilot > 0 && request.Limit > 0) throw new ArgumentException("pilot과 limit은 함께 지정할 수 없습니다.");
        if (request.Pilot < 0 || request.Limit < 0) throw new ArgumentException("pilot과 limit은 0 이상이어야 합니다.");

        var isNewBatch = !string.IsNullOrWhiteSpace(request.BatchFolder);
        var batchDirectory = isNewBatch
            ? CreateBatch(serviceDirectory, Path.GetFullPath(request.BatchFolder!), request.BatchId ?? $"structure-scan-{DateTime.UtcNow:yyyyMMddTHHmmssZ}", request.Pilot)
            : ResolveBatchDirectory(serviceDirectory, request.ResumeBatchId!);

        if (!isNewBatch && !File.Exists(Path.Combine(batchDirectory, "batch.json")))
            throw new ArgumentException($"구조 스캔 배치를 찾을 수 없습니다: {request.ResumeBatchId}");
        Directory.CreateDirectory(Path.Combine(batchDirectory, "logs"));

        using var connection = OpenState(Path.Combine(batchDirectory, "state.sqlite"));
        foreach (var item in PendingItems(connection, request.RetryFailed, request.Limit, includeDeferred: !isNewBatch))
        {
            cancellationToken.ThrowIfCancellationRequested();
            ProcessItem(connection, batchDirectory, item, cancellationToken);
        }
        var summary = BuildOutputs(batchDirectory, connection);
        return new StructureScanRunResult(Path.GetFileName(batchDirectory), batchDirectory, summary);
    }

    private static string BatchOutputRoot(string serviceDirectory)
    {
        var root = Path.GetFullPath(
            AppRuntimePaths.Current.BatchRootDirectory);
        Directory.CreateDirectory(root);
        return root;
    }

    private static string ResolveBatchDirectory(string serviceDirectory, string batchId)
    {
        var root = BatchOutputRoot(serviceDirectory);
        var safeId = ValidateBatchId(batchId);
        var target = Path.GetFullPath(Path.Combine(root, safeId));
        if (!IsWithin(target, root)) throw new InvalidOperationException("배치 경로가 출력 루트 밖을 가리킵니다.");
        return target;
    }

    private static string ValidateBatchId(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 96 || value is "." or ".." ||
            value.Any(character => !(char.IsAsciiLetterOrDigit(character) || character is '.' or '_' or '-')))
            throw new ArgumentException("배치 ID는 영문/숫자/.-_만 사용해 1~96자로 지정해야 합니다.");
        return value;
    }

    private string CreateBatch(string serviceDirectory, string sourceRoot, string batchId, int pilot)
    {
        if (!Directory.Exists(sourceRoot) || !Path.IsPathFullyQualified(sourceRoot))
            throw new ArgumentException($"배치 폴더는 존재하는 전체 경로여야 합니다: {sourceRoot}");
        var target = ResolveBatchDirectory(serviceDirectory, ValidateBatchId(batchId));
        if (Directory.Exists(target))
        {
            var entries = Directory.EnumerateFileSystemEntries(target).Select(Path.GetFileName).Where(static name => !string.IsNullOrEmpty(name)).ToList();
            var unexpected = entries.Where(name => !string.Equals(name, "logs", StringComparison.OrdinalIgnoreCase)).ToList();
            var logsPath = Path.Combine(target, "logs");
            if (unexpected.Count > 0 || (File.Exists(logsPath) && !Directory.Exists(logsPath)))
                throw new InvalidOperationException($"이미 초기화된 배치 산출물이 있습니다: {batchId}. 기존 배치를 재개하세요.");
        }
        else Directory.CreateDirectory(target);

        var records = InventoryFiles(sourceRoot);
        if (records.Count == 0) throw new InvalidOperationException($"Excel 또는 인벤토리 파일이 없습니다: {sourceRoot}");
        var selected = PilotSelection(records, pilot);
        var batch = new
        {
            schemaVersion = "structure-scan-batch-v1",
            scannerVersion = ScannerVersion,
            batchId,
            createdAt = UtcNow(),
            rootPath = sourceRoot,
            options = new { pilot, recursive = true, readOnly = true, usesCom = false },
            discovered = records.Count,
            selectedInitially = selected.Count,
        };
        AtomicWriteJson(Path.Combine(target, "batch.json"), batch);

        using var connection = OpenState(Path.Combine(target, "state.sqlite"));
        using var transaction = connection.BeginTransaction();
        foreach (var record in records)
        {
            var status = record.Kind switch
            {
                "non_workbook" => "NON_WORKBOOK",
                "unsupported_binary" => "UNSUPPORTED",
                _ when pilot > 0 && !selected.Contains(record.RelativePath) => "DEFERRED",
                _ => "PENDING",
            };
            Execute(connection, transaction, """
                INSERT INTO items(relative_path, source_path, extension, kind, size_bytes, mtime_ns, fingerprint, selected, status)
                VALUES ($relativePath, $sourcePath, $extension, $kind, $sizeBytes, $mtimeNs, $fingerprint, $selected, $status);
                """,
                ("$relativePath", record.RelativePath),
                ("$sourcePath", record.SourcePath),
                ("$extension", record.Extension),
                ("$kind", record.Kind),
                ("$sizeBytes", record.SizeBytes),
                ("$mtimeNs", record.MtimeNs),
                ("$fingerprint", record.Fingerprint),
                ("$selected", selected.Contains(record.RelativePath) ? 1 : 0),
                ("$status", status));
        }
        transaction.Commit();
        return target;
    }

    private static SqliteConnection OpenState(string path)
    {
        var connection = new SqliteConnection($"Data Source={path}");
        connection.Open();
        Execute(connection, null, "PRAGMA journal_mode=WAL;");
        Execute(connection, null, """
            CREATE TABLE IF NOT EXISTS items (
                relative_path TEXT PRIMARY KEY,
                source_path TEXT NOT NULL,
                extension TEXT NOT NULL,
                kind TEXT NOT NULL,
                size_bytes INTEGER NOT NULL,
                mtime_ns INTEGER NOT NULL,
                fingerprint TEXT NOT NULL,
                selected INTEGER NOT NULL,
                status TEXT NOT NULL,
                attempts INTEGER NOT NULL DEFAULT 0,
                result_path TEXT NOT NULL DEFAULT '',
                error_text TEXT NOT NULL DEFAULT '',
                started_at TEXT NOT NULL DEFAULT '',
                finished_at TEXT NOT NULL DEFAULT ''
            );
            """);
        Execute(connection, null, "CREATE INDEX IF NOT EXISTS idx_items_status ON items(status, relative_path);");
        return connection;
    }

    private static List<InventoryRecord> InventoryFiles(string root)
    {
        var records = new List<InventoryRecord>();
        foreach (var path in Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
                     .OrderBy(static value => value, StringComparer.OrdinalIgnoreCase))
        {
            var name = Path.GetFileName(path);
            if (name.StartsWith("~$", StringComparison.Ordinal)) continue;
            var extension = Path.GetExtension(path).ToLowerInvariant();
            if (!InventoryExtensions.Contains(extension)) continue;
            var fullPath = Path.GetFullPath(path);
            var info = new FileInfo(fullPath);
            var relativePath = Path.GetRelativePath(root, fullPath).Replace(Path.DirectorySeparatorChar, '/');
            var kind = OpenXmlExtensions.Contains(extension) ? "openxml" : BinaryExtensions.Contains(extension) ? "unsupported_binary" : "non_workbook";
            records.Add(new InventoryRecord(
                relativePath,
                fullPath,
                extension,
                kind,
                info.Length,
                MtimeNanoseconds(info),
                SourceFingerprint(fullPath),
                FamilyTag(relativePath)));
        }
        return records;
    }

    private static HashSet<string> PilotSelection(IReadOnlyList<InventoryRecord> records, int count)
    {
        var eligible = records.Where(static record => record.Kind == "openxml").ToList();
        if (count <= 0 || count >= eligible.Count) return eligible.Select(static record => record.RelativePath).ToHashSet(StringComparer.Ordinal);

        var buckets = eligible.GroupBy(static record => record.FamilyTag, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.OrderBy(static record => StableSourceId(record.RelativePath), StringComparer.Ordinal).ToList(), StringComparer.Ordinal);
        static List<InventoryRecord> Bucket(IReadOnlyDictionary<string, List<InventoryRecord>> values, string name) => values.TryGetValue(name, out var records) ? records : [];
        var targets = new Dictionary<string, int>(StringComparer.Ordinal)
        {
            ["BRS"] = (count * 40 + 99) / 100,
            ["TIU"] = (count * 25 + 99) / 100,
            ["MSU"] = (count * 20 + 99) / 100,
        };
        targets["Other"] = Math.Max(0, count - targets.Values.Sum());
        var selected = new List<InventoryRecord>();
        foreach (var name in new[] { "BRS", "TIU", "MSU", "Other" }) selected.AddRange(Bucket(buckets, name).Take(targets[name]));
        if (selected.Count < count)
        {
            var selectedPaths = selected.Select(static item => item.RelativePath).ToHashSet(StringComparer.Ordinal);
            selected.AddRange(eligible.Where(record => !selectedPaths.Contains(record.RelativePath))
                .OrderBy(static record => StableSourceId(record.RelativePath), StringComparer.Ordinal)
                .Take(count - selected.Count));
        }
        return selected.Take(count).Select(static record => record.RelativePath).ToHashSet(StringComparer.Ordinal);
    }

    private static List<ItemRecord> PendingItems(SqliteConnection connection, bool retryFailed, int limit, bool includeDeferred)
    {
        Execute(connection, null, "UPDATE items SET status='INTERRUPTED' WHERE status='SCANNING';");
        Execute(connection, null, "UPDATE items SET status='PENDING' WHERE status='INTERRUPTED';");
        var statuses = new List<string> { "PENDING" };
        if (includeDeferred) statuses.Add("DEFERRED");
        if (retryFailed) statuses.Add("FAILED_RETRYABLE");
        var placeholders = string.Join(",", statuses.Select((_, index) => $"$status{index}"));
        var sql = $"SELECT relative_path, source_path, extension, kind, size_bytes, mtime_ns, fingerprint, selected, status, attempts, result_path, error_text FROM items WHERE status IN ({placeholders}) ORDER BY relative_path";
        if (limit > 0) sql += " LIMIT $limit";
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        for (var index = 0; index < statuses.Count; index++) command.Parameters.AddWithValue($"$status{index}", statuses[index]);
        if (limit > 0) command.Parameters.AddWithValue("$limit", limit);
        using var reader = command.ExecuteReader();
        var result = new List<ItemRecord>();
        while (reader.Read()) result.Add(ReadItem(reader));
        return result;
    }

    private void ProcessItem(SqliteConnection connection, string batchDirectory, ItemRecord item, CancellationToken cancellationToken)
    {
        if (!CurrentFingerprintMatches(item))
        {
            MarkChanged(connection, batchDirectory, item, "Source path, size, or modified time changed from the batch snapshot.");
            return;
        }
        Execute(connection, null, "UPDATE items SET status='SCANNING', attempts=attempts+1, started_at=$startedAt, error_text='' WHERE relative_path=$relativePath;",
            ("$startedAt", UtcNow()), ("$relativePath", item.RelativePath));
        _log?.Invoke($"구조 스캔: {item.RelativePath}");
        var timer = Stopwatch.StartNew();
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var result = item.Kind == "unsupported_binary" ? UnsupportedResult(item) : ScanOpenXml(item, cancellationToken);
            result.ElapsedSeconds = Math.Round(timer.Elapsed.TotalSeconds, 3, MidpointRounding.AwayFromZero);
            if (!CurrentFingerprintMatches(item))
            {
                MarkChanged(connection, batchDirectory, item, "Source changed while it was being scanned.");
                return;
            }
            var resultPath = WriteResult(batchDirectory, item.RelativePath, result);
            var status = string.Equals(result.ScanStatus, "TRUNCATED", StringComparison.Ordinal) ? "TRUNCATED" : "SCANNED";
            Execute(connection, null, "UPDATE items SET status=$status, result_path=$resultPath, finished_at=$finishedAt WHERE relative_path=$relativePath;",
                ("$status", status), ("$resultPath", resultPath), ("$finishedAt", UtcNow()), ("$relativePath", item.RelativePath));
            AppendEvent(batchDirectory, new Dictionary<string, object?> { ["at"] = UtcNow(), ["relativePath"] = item.RelativePath, ["status"] = status, ["result"] = resultPath });
        }
        catch (QuarantinedPackageException exception)
        {
            var result = QuarantinedResult(item, exception.Message, timer.Elapsed);
            var resultPath = WriteResult(batchDirectory, item.RelativePath, result);
            Execute(connection, null, "UPDATE items SET status='QUARANTINED', result_path=$resultPath, error_text=$error, finished_at=$finishedAt WHERE relative_path=$relativePath;",
                ("$resultPath", resultPath), ("$error", Clip(exception.Message, 2000)), ("$finishedAt", UtcNow()), ("$relativePath", item.RelativePath));
            AppendEvent(batchDirectory, new Dictionary<string, object?> { ["at"] = UtcNow(), ["relativePath"] = item.RelativePath, ["status"] = "QUARANTINED", ["error"] = exception.Message });
        }
        catch (InvalidDataException exception)
        {
            var result = QuarantinedResult(item, exception.Message, timer.Elapsed);
            var resultPath = WriteResult(batchDirectory, item.RelativePath, result);
            Execute(connection, null, "UPDATE items SET status='QUARANTINED', result_path=$resultPath, error_text=$error, finished_at=$finishedAt WHERE relative_path=$relativePath;",
                ("$resultPath", resultPath), ("$error", Clip(exception.Message, 2000)), ("$finishedAt", UtcNow()), ("$relativePath", item.RelativePath));
            AppendEvent(batchDirectory, new Dictionary<string, object?> { ["at"] = UtcNow(), ["relativePath"] = item.RelativePath, ["status"] = "QUARANTINED", ["error"] = exception.Message });
        }
        catch (XmlException exception)
        {
            var result = QuarantinedResult(item, exception.Message, timer.Elapsed);
            var resultPath = WriteResult(batchDirectory, item.RelativePath, result);
            Execute(connection, null, "UPDATE items SET status='QUARANTINED', result_path=$resultPath, error_text=$error, finished_at=$finishedAt WHERE relative_path=$relativePath;",
                ("$resultPath", resultPath), ("$error", Clip(exception.Message, 2000)), ("$finishedAt", UtcNow()), ("$relativePath", item.RelativePath));
            AppendEvent(batchDirectory, new Dictionary<string, object?> { ["at"] = UtcNow(), ["relativePath"] = item.RelativePath, ["status"] = "QUARANTINED", ["error"] = exception.Message });
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            var error = $"{exception.GetType().Name}: {exception.Message}";
            Execute(connection, null, "UPDATE items SET status='FAILED_RETRYABLE', error_text=$error, finished_at=$finishedAt WHERE relative_path=$relativePath;",
                ("$error", Clip(error, 2000)), ("$finishedAt", UtcNow()), ("$relativePath", item.RelativePath));
            AppendEvent(batchDirectory, new Dictionary<string, object?> { ["at"] = UtcNow(), ["relativePath"] = item.RelativePath, ["status"] = "FAILED_RETRYABLE", ["error"] = error });
        }
    }

    private static void MarkChanged(SqliteConnection connection, string batchDirectory, ItemRecord item, string reason)
    {
        Execute(connection, null, "UPDATE items SET status='CHANGED', error_text=$error, finished_at=$finishedAt WHERE relative_path=$relativePath;",
            ("$error", reason), ("$finishedAt", UtcNow()), ("$relativePath", item.RelativePath));
        AppendEvent(batchDirectory, new Dictionary<string, object?> { ["at"] = UtcNow(), ["relativePath"] = item.RelativePath, ["status"] = "CHANGED" });
    }

    private static bool CurrentFingerprintMatches(ItemRecord item) =>
        File.Exists(item.SourcePath) && string.Equals(SourceFingerprint(item.SourcePath), item.Fingerprint, StringComparison.Ordinal);

    private static StructureResult ScanOpenXml(ItemRecord item, CancellationToken cancellationToken)
    {
        using var archive = ZipFile.OpenRead(item.SourcePath);
        var package = ZipPreflight(archive);
        var relationships = ReadRelationships(archive);
        var sheets = ReadWorkbookSheets(archive, relationships);
        var sharedStrings = ReadSharedStrings(archive);
        var anyTruncated = false;
        foreach (var sheet in sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ScanSheet(archive, sheet, sharedStrings, ref anyTruncated, cancellationToken);
        }
        package["hasVbaProject"] = ContentTypesContainVbaProject(archive);
        var sections = sheets.SelectMany(static sheet => sheet.Sections).ToList();
        return new StructureResult
        {
            SchemaVersion = "structure-scan-result-v1",
            ScannerVersion = ScannerVersion,
            Source = new SourceRecord(item.RelativePath, item.SourcePath, item.Extension, item.SizeBytes, item.MtimeNs, item.Fingerprint, FamilyTag(item.RelativePath)),
            ScanStatus = anyTruncated ? "TRUNCATED" : "OK",
            ReadOnly = true,
            UsesCom = false,
            Package = package,
            Sheets = sheets,
            PrimaryStructure = PrimaryStructure(sheets, anyTruncated),
            StructuralTypes = sections.Select(static section => section.Type).Distinct(StringComparer.Ordinal).OrderBy(static value => value, StringComparer.Ordinal).ToList(),
            Limitations =
            [
                "This is a header-and-merge layout scan only; it does not calculate formulas or assess data values.",
                "Candidate sections are not acceptance, quality, causality, or Test/Normal conclusions.",
            ],
        };
    }

    private static StructureResult UnsupportedResult(ItemRecord item) => new()
    {
        SchemaVersion = "structure-scan-result-v1",
        ScannerVersion = ScannerVersion,
        Source = new SourceRecord(item.RelativePath, item.SourcePath, item.Extension, item.SizeBytes, item.MtimeNs, item.Fingerprint, FamilyTag(item.RelativePath)),
        ScanStatus = "UNSUPPORTED",
        PrimaryStructure = "TRUNCATED_OR_UNREADABLE",
        StructuralTypes = [],
        Limitations = [item.Extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) ? "UNSUPPORTED_XLS: this read-only OpenXML scanner does not open .xls files." : "UNSUPPORTED_XLSB: this read-only OpenXML scanner does not open .xlsb files.", "No COM or automatic conversion fallback was used."],
    };

    private static StructureResult QuarantinedResult(ItemRecord item, string reason, TimeSpan elapsed) => new()
    {
        SchemaVersion = "structure-scan-result-v1",
        ScannerVersion = ScannerVersion,
        Source = new SourceRecord(item.RelativePath, item.SourcePath, item.Extension, item.SizeBytes, item.MtimeNs, item.Fingerprint, FamilyTag(item.RelativePath)),
        ScanStatus = "QUARANTINED",
        PrimaryStructure = "TRUNCATED_OR_UNREADABLE",
        StructuralTypes = [],
        Limitations = [reason],
        ElapsedSeconds = Math.Round(elapsed.TotalSeconds, 3, MidpointRounding.AwayFromZero),
    };

    private static Dictionary<string, object> ZipPreflight(ZipArchive archive)
    {
        if (archive.Entries.Count > MaxZipEntries)
            throw new QuarantinedPackageException($"ZIP entry count exceeds limit ({archive.Entries.Count} > {MaxZipEntries}).");
        var names = new HashSet<string>(StringComparer.Ordinal);
        long compressed = 0;
        long uncompressed = 0;
        foreach (var entry in archive.Entries)
        {
            if (!names.Add(entry.FullName)) throw new QuarantinedPackageException($"ZIP package contains a duplicate entry: {entry.FullName}");
            try
            {
                uncompressed = checked(uncompressed + entry.Length);
                compressed = checked(compressed + entry.CompressedLength);
            }
            catch (OverflowException) { throw new QuarantinedPackageException("ZIP package size overflow."); }
        }
        if (uncompressed > MaxPackageUncompressedBytes) throw new QuarantinedPackageException("ZIP package uncompressed size exceeds scan limit.");
        if (compressed > 0 && uncompressed / compressed > MaxCompressionRatio) throw new QuarantinedPackageException("ZIP compression ratio exceeds scan limit.");
        if (!names.Contains("xl/workbook.xml") || !names.Contains("xl/_rels/workbook.xml.rels"))
            throw new QuarantinedPackageException("OpenXML workbook metadata is missing.");
        return new Dictionary<string, object>(StringComparer.Ordinal)
        {
            ["entryCount"] = archive.Entries.Count,
            ["compressedBytes"] = compressed,
            ["uncompressedBytes"] = uncompressed,
        };
    }

    private static Dictionary<string, string> ReadRelationships(ZipArchive archive)
    {
        var entry = FindEntry(archive, "xl/_rels/workbook.xml.rels") ?? throw new QuarantinedPackageException("OpenXML workbook relationship metadata is missing.");
        var relationships = new Dictionary<string, string>(StringComparer.Ordinal);
        using var reader = CreateXmlReader(entry, MaxWorksheetUncompressedBytes);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || !string.Equals(reader.LocalName, "Relationship", StringComparison.Ordinal)) continue;
            var id = reader.GetAttribute("Id") ?? string.Empty;
            var target = reader.GetAttribute("Target") ?? string.Empty;
            if (!string.IsNullOrEmpty(id)) relationships[id] = RelationshipTarget(target);
        }
        return relationships;
    }

    private static List<SheetRecord> ReadWorkbookSheets(ZipArchive archive, IReadOnlyDictionary<string, string> relationships)
    {
        var entry = FindEntry(archive, "xl/workbook.xml") ?? throw new QuarantinedPackageException("OpenXML workbook metadata is missing.");
        var sheets = new List<SheetRecord>();
        using var reader = CreateXmlReader(entry, MaxWorksheetUncompressedBytes);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || !string.Equals(reader.LocalName, "sheet", StringComparison.Ordinal)) continue;
            var id = reader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") ?? string.Empty;
            var part = relationships.TryGetValue(id, out var target) ? target : string.Empty;
            sheets.Add(new SheetRecord
            {
                SheetIndex = sheets.Count + 1,
                SheetName = reader.GetAttribute("name") ?? $"Sheet{sheets.Count + 1}",
                SheetState = reader.GetAttribute("state") ?? "visible",
                Part = part,
            });
        }
        return sheets;
    }

    private static List<string> ReadSharedStrings(ZipArchive archive)
    {
        var entry = FindEntry(archive, "xl/sharedStrings.xml");
        if (entry is null) return [];
        var strings = new List<string>();
        using var reader = CreateXmlReader(entry, MaxWorksheetUncompressedBytes);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || !string.Equals(reader.LocalName, "si", StringComparison.Ordinal)) continue;
            var text = new StringBuilder();
            using var item = reader.ReadSubtree();
            while (item.Read())
            {
                if (item.NodeType == XmlNodeType.Element && string.Equals(item.LocalName, "t", StringComparison.Ordinal))
                    text.Append(item.ReadElementContentAsString());
            }
            strings.Add(text.ToString());
        }
        return strings;
    }

    private static bool ContentTypesContainVbaProject(ZipArchive archive)
    {
        var entry = FindEntry(archive, "[Content_Types].xml");
        if (entry is null) return false;
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: false);
        var content = reader.ReadToEnd();
        return content.Contains("vbaProject", StringComparison.OrdinalIgnoreCase);
    }

    private static void ScanSheet(ZipArchive archive, SheetRecord sheet, IReadOnlyList<string> sharedStrings, ref bool anyTruncated, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(sheet.Part))
        {
            sheet.Warnings.Add("Worksheet relationship target is missing from the package.");
            sheet.ScanStatus = "TRUNCATED";
            anyTruncated = true;
            return;
        }
        var entry = FindEntry(archive, sheet.Part);
        if (entry is null)
        {
            sheet.Warnings.Add("Worksheet relationship target is missing from the package.");
            sheet.ScanStatus = "TRUNCATED";
            anyTruncated = true;
            return;
        }
        if (entry.Length > MaxWorksheetUncompressedBytes)
        {
            sheet.Warnings.Add("Worksheet XML exceeds scan limit; header scan will be truncated.");
            sheet.ScanStatus = "TRUNCATED";
            anyTruncated = true;
            return;
        }

        var tokenRows = new Dictionary<int, List<CellToken>>();
        var visibleRows = new SortedDictionary<int, List<string>>();
        var truncated = false;
        var maximumObservedRow = 0;
        var maximumObservedColumn = 0;
        using var reader = CreateXmlReader(entry, MaxWorksheetUncompressedBytes);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (string.Equals(reader.LocalName, "dimension", StringComparison.Ordinal) && string.IsNullOrEmpty(sheet.DeclaredDimension))
                sheet.DeclaredDimension = reader.GetAttribute("ref") ?? string.Empty;
            else if (string.Equals(reader.LocalName, "mergeCell", StringComparison.Ordinal))
            {
                sheet.MergeCount++;
                var reference = reader.GetAttribute("ref");
                if (!string.IsNullOrWhiteSpace(reference) && sheet.MergeRangeSample.Count < MaxMergeSample) sheet.MergeRangeSample.Add(reference);
            }
            else if (string.Equals(reader.LocalName, "drawing", StringComparison.Ordinal)) sheet.DrawingRelationshipCount++;
            else if (string.Equals(reader.LocalName, "c", StringComparison.Ordinal))
            {
                if (DeclaredCells(sheet.DeclaredDimension) is long declared && declared > MaxDeclaredCells)
                {
                    sheet.Warnings.Add("Declared worksheet dimension exceeds cell budget; header scan truncated.");
                    truncated = true;
                    break;
                }
                var cell = ReadCell(reader, sharedStrings);
                maximumObservedRow = Math.Max(maximumObservedRow, cell.Row);
                maximumObservedColumn = Math.Max(maximumObservedColumn, cell.Column);
                if (GridCellIndex(cell.Row, cell.Column, sheet.DeclaredDimension) > MaxScannedCells)
                {
                    // openpyxl iter_rows visits blank grid positions as well.  Once
                    // their virtual position crosses this budget, it stops before
                    // reading this physical cell too.
                    truncated = true;
                    break;
                }
                // openpyxl exposes both formula elements and literal strings
                // beginning with '=' as a string that begins with '='.  Keep the
                // v1 scanner's formulaCount behaviour for either representation.
                if (cell.IsFormula || cell.Value?.StartsWith("=", StringComparison.Ordinal) == true)
                {
                    sheet.FormulaCount++;
                    continue;
                }
                var token = HeaderToken(cell.Value);
                if (token is null || cell.Row <= 0 || cell.Column <= 0) continue;
                if (!tokenRows.TryGetValue(cell.Row, out var tokens)) tokenRows[cell.Row] = tokens = [];
                tokens.Add(new CellToken(cell.Row, cell.Column, token));
                if (!visibleRows.TryGetValue(cell.Row, out var visible)) visibleRows[cell.Row] = visible = [];
                if (visible.Count < MaxHeaderTokensPerRow) visible.Add(token);
            }
        }
        sheet.ScannedCellCount = ScannedCellCount(sheet.DeclaredDimension, maximumObservedRow, maximumObservedColumn, truncated);
        sheet.HeaderRows = visibleRows.Take(MaxHeaderRows).Select(pair => new HeaderRow(pair.Key, pair.Value)).ToList();
        sheet.Sections = ClassifyHeaderRows(sheet.SheetName, sheet.SheetState, tokenRows, sheet.MergeCount);
        if (truncated)
        {
            if (!sheet.Warnings.Any(static warning => warning.StartsWith("Declared worksheet", StringComparison.Ordinal))) sheet.Warnings.Add("Header scan cell budget reached.");
            sheet.ScanStatus = "TRUNCATED";
            anyTruncated = true;
        }
        else sheet.ScanStatus = "OK";
    }

    private static ParsedCell ReadCell(XmlReader reader, IReadOnlyList<string> sharedStrings)
    {
        var reference = reader.GetAttribute("r") ?? string.Empty;
        var type = reader.GetAttribute("t") ?? string.Empty;
        var (row, column) = CellAddress(reference);
        var formula = false;
        var text = new StringBuilder();
        var rawValue = string.Empty;
        using var subtree = reader.ReadSubtree();
        while (subtree.Read())
        {
            if (subtree.NodeType != XmlNodeType.Element) continue;
            if (string.Equals(subtree.LocalName, "f", StringComparison.Ordinal)) formula = true;
            else if (string.Equals(subtree.LocalName, "v", StringComparison.Ordinal)) rawValue = subtree.ReadElementContentAsString();
            else if (string.Equals(subtree.LocalName, "t", StringComparison.Ordinal)) text.Append(subtree.ReadElementContentAsString());
        }
        string? value = type switch
        {
            "s" when int.TryParse(rawValue, NumberStyles.None, CultureInfo.InvariantCulture, out var sharedIndex) && sharedIndex >= 0 && sharedIndex < sharedStrings.Count => sharedStrings[sharedIndex],
            "inlineStr" => text.ToString(),
            "str" => rawValue,
            _ => rawValue,
        };
        return new ParsedCell(row, column, value, formula);
    }

    private static List<SectionRecord> ClassifyHeaderRows(string sheetName, string state, IReadOnlyDictionary<int, List<CellToken>> rows, int mergeCount)
    {
        var sections = new Dictionary<string, SectionRecord>(StringComparer.Ordinal);
        foreach (var rowNumber in rows.Keys.OrderBy(static value => value))
        {
            var window = new List<CellToken>();
            if (rows.TryGetValue(rowNumber, out var current)) window.AddRange(current);
            if (rows.TryGetValue(rowNumber + 1, out var next)) window.AddRange(next);
            var tokens = window.Select(static cell => cell.Token).ToHashSet(StringComparer.Ordinal);
            var inputCount = window.Count(static cell => cell.Token == "INPUT");
            var headerRange = HeaderRange(window);
            void Add(string type, string confidence)
            {
                var section = new SectionRecord(sheetName, state, headerRange, mergeCount, type, confidence);
                sections[$"{type}\u001f{headerRange}"] = section;
            }
            if (new[] { "INPUT", "OK", "TOTAL_NG", "NG_RATE" }.All(tokens.Contains)) Add("DEFECT_ACCOUNTING_LAYOUT_CANDIDATE", "HIGH");
            if (inputCount >= 2 && tokens.Contains("TOTAL_NG") && tokens.Contains("NG_RATE")) Add("REPEATED_DEFECT_BLOCK_LAYOUT_CANDIDATE", "HIGH");
            if (new[] { "SAMPLE", "AVERAGE", "MAX", "MIN" }.All(tokens.Contains)) Add("RAW_MEASUREMENT_SUMMARY_LAYOUT_CANDIDATE", "HIGH");
            if (tokens.Contains("NG_CUE") && tokens.Contains("TOTAL_NG")) Add("NG_BREAKDOWN_MATRIX_CANDIDATE", "MEDIUM");
            if (tokens.Contains("NORMAL_CUE") && tokens.Contains("TEST_CUE")) Add("EXPLICIT_COHORT_COMPARISON_CANDIDATE", "LOW");
        }
        return sections.Values.ToList();
    }

    private static string? HeaderToken(string? value)
    {
        var normalized = NormalizeLabel(value);
        return normalized switch
        {
            "input" => "INPUT",
            "ok" => "OK",
            "totalng" or "ngtotal" or "totaldefect" or "totaldefects" => "TOTAL_NG",
            "ngrate" or "defectrate" or "totalngrate" => "NG_RATE",
            "sample" or "samples" => "SAMPLE",
            "average" or "avg" or "mean" => "AVERAGE",
            "max" or "maximum" => "MAX",
            "min" or "minimum" => "MIN",
            "normal" or "baseline" or "before" => "NORMAL_CUE",
            "test" or "trial" or "after" => "TEST_CUE",
            _ when normalized == "ng" || normalized.StartsWith("ng", StringComparison.Ordinal) => "NG_CUE",
            _ => null,
        };
    }

    private static string NormalizeLabel(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var builder = new StringBuilder(value.Length);
        foreach (var character in value.Trim().ToLowerInvariant())
        {
            if (char.IsAsciiLetterOrDigit(character) || (character >= '가' && character <= '힣')) builder.Append(character);
        }
        return builder.ToString();
    }

    private static (int Row, int Column) CellAddress(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference)) return (0, 0);
        var column = 0;
        var index = 0;
        while (index < reference.Length && char.IsLetter(reference[index]))
        {
            column = checked(column * 26 + char.ToUpperInvariant(reference[index]) - 'A' + 1);
            index++;
        }
        return int.TryParse(reference[index..], NumberStyles.None, CultureInfo.InvariantCulture, out var row) ? (row, column) : (0, 0);
    }

    private static long? DeclaredCells(string? dimension)
    {
        var bounds = DimensionBounds(dimension);
        if (bounds is null) return null;
        var (first, last) = bounds.Value;
        if (first.Row <= 0 || first.Column <= 0 || last.Row <= 0 || last.Column <= 0) return null;
        return checked((long)(Math.Abs(last.Column - first.Column) + 1) * (Math.Abs(last.Row - first.Row) + 1));
    }

    private static ((int Row, int Column) First, (int Row, int Column) Last)? DimensionBounds(string? dimension)
    {
        if (string.IsNullOrWhiteSpace(dimension)) return null;
        var parts = dimension.ToUpperInvariant().Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length is < 1 or > 2) return null;
        return (CellAddress(parts[0]), CellAddress(parts[^1]));
    }

    private static long GridCellIndex(int row, int column, string dimension)
    {
        var bounds = DimensionBounds(dimension);
        var maximumColumn = bounds?.Last.Column ?? column;
        return row <= 0 || column <= 0 || maximumColumn <= 0 ? 0 : checked((long)(row - 1) * maximumColumn + column);
    }

    private static int ScannedCellCount(string dimension, int maximumObservedRow, int maximumObservedColumn, bool truncated)
    {
        if (truncated) return MaxScannedCells + 1;
        if (maximumObservedRow <= 0 || maximumObservedColumn <= 0) return 0;
        var bounds = DimensionBounds(dimension);
        var maximumRow = bounds?.Last.Row ?? maximumObservedRow;
        var maximumColumn = bounds?.Last.Column ?? maximumObservedColumn;
        if (maximumRow <= 0 || maximumColumn <= 0) return 0;
        var gridCells = checked((long)maximumRow * maximumColumn);
        return gridCells > int.MaxValue ? int.MaxValue : (int)gridCells;
    }

    private static string HeaderRange(IReadOnlyList<CellToken> cells)
    {
        if (cells.Count == 0) return string.Empty;
        return $"{ColumnLabel(cells.Min(static cell => cell.Column))}{cells.Min(static cell => cell.Row)}:{ColumnLabel(cells.Max(static cell => cell.Column))}{cells.Max(static cell => cell.Row)}";
    }

    private static string ColumnLabel(int column)
    {
        var letters = string.Empty;
        while (column > 0)
        {
            column--;
            letters = (char)('A' + column % 26) + letters;
            column /= 26;
        }
        return letters;
    }

    private static string PrimaryStructure(IReadOnlyCollection<SheetRecord> sheets, bool truncated)
    {
        if (truncated) return "TRUNCATED_OR_UNREADABLE";
        var count = sheets.Sum(static sheet => sheet.Sections.Count);
        return count switch { 0 => "NO_RECOGNIZED_TABLE", 1 => "TABULAR_SINGLE", _ => "TABULAR_MULTI_SECTION" };
    }

    private static ZipArchiveEntry? FindEntry(ZipArchive archive, string path) => archive.Entries.FirstOrDefault(entry => string.Equals(entry.FullName, path, StringComparison.Ordinal));

    private static XmlReader CreateXmlReader(ZipArchiveEntry entry, long maximumCharacters)
    {
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreWhitespace = true,
            MaxCharactersInDocument = maximumCharacters,
            MaxCharactersFromEntities = 0,
            CloseInput = true,
        };
        return XmlReader.Create(entry.Open(), settings);
    }

    private static string RelationshipTarget(string target)
    {
        var normalized = target.Replace('\\', '/').Trim();
        var combined = normalized.StartsWith("/", StringComparison.Ordinal) ? normalized.TrimStart('/') : $"xl/{normalized}";
        var parts = new Stack<string>();
        foreach (var part in combined.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (part == ".") continue;
            if (part == "..")
            {
                if (parts.Count == 0) throw new QuarantinedPackageException("Worksheet relationship target escapes the workbook package.");
                parts.Pop();
                continue;
            }
            parts.Push(part);
        }
        var resolved = string.Join('/', parts.Reverse());
        if (!resolved.StartsWith("xl/", StringComparison.Ordinal)) throw new QuarantinedPackageException("Worksheet relationship target escapes the workbook package.");
        return resolved;
    }

    private static string WriteResult(string batchDirectory, string relativePath, StructureResult result)
    {
        var relative = $"results/{StableSourceId(relativePath)}.json";
        AtomicWriteJson(Path.Combine(batchDirectory, relative.Replace('/', Path.DirectorySeparatorChar)), result);
        return relative;
    }

    private static void AppendEvent(string batchDirectory, IReadOnlyDictionary<string, object?> value)
    {
        var line = JsonSerializer.Serialize(value, JsonOptions);
        File.AppendAllText(Path.Combine(batchDirectory, "events.jsonl"), line + Environment.NewLine, new UTF8Encoding(false));
        var path = Path.Combine(batchDirectory, "logs", "batch.log");
        File.AppendAllText(path, $"[{UtcNow()}] {value["relativePath"]} {value["status"]}{Environment.NewLine}", new UTF8Encoding(false));
    }

    private static StructureScanSummary BuildOutputs(string batchDirectory, SqliteConnection connection)
    {
        var rows = AllItems(connection);
        var statusCounts = rows.GroupBy(static row => row.Status, StringComparer.Ordinal).OrderBy(static group => group.Key, StringComparer.Ordinal).ToDictionary(static group => group.Key, static group => group.Count(), StringComparer.Ordinal);
        var structureCounts = new SortedDictionary<string, int>(StringComparer.Ordinal);
        var sectionCounts = new SortedDictionary<string, int>(StringComparer.Ordinal);
        var reportRows = new List<Dictionary<string, string>>();
        var failures = new List<Dictionary<string, string>>();
        foreach (var row in rows)
        {
            using var result = ReadResult(batchDirectory, row.ResultPath);
            var structure = JsonString(result, "primaryStructure");
            var types = JsonStrings(result, "structuralTypes");
            if (!string.IsNullOrEmpty(structure)) Increment(structureCounts, structure);
            foreach (var type in types) Increment(sectionCounts, type);
            var sheets = result is not null && result.RootElement.TryGetProperty("sheets", out var sheetsElement) && sheetsElement.ValueKind == JsonValueKind.Array ? sheetsElement : default;
            var sheetCount = sheets.ValueKind == JsonValueKind.Array ? sheets.GetArrayLength() : 0;
            var mergeCount = 0;
            if (sheets.ValueKind == JsonValueKind.Array)
                foreach (var sheet in sheets.EnumerateArray())
                    if (sheet.TryGetProperty("mergeCount", out var countElement) && countElement.TryGetInt32(out var count)) mergeCount += count;
            var warnings = string.Join("; ", JsonStrings(result, "limitations"));
            var report = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["relativePath"] = row.RelativePath,
                ["extension"] = row.Extension,
                ["familyTag"] = FamilyTag(row.RelativePath),
                ["status"] = row.Status,
                ["primaryStructure"] = structure,
                ["structuralTypes"] = string.Join("; ", types),
                ["sheetCount"] = sheetCount.ToString(CultureInfo.InvariantCulture),
                ["mergeCount"] = mergeCount.ToString(CultureInfo.InvariantCulture),
                ["resultPath"] = row.ResultPath,
                ["warningOrError"] = string.IsNullOrEmpty(row.ErrorText) ? warnings : row.ErrorText,
            };
            reportRows.Add(report);
            if (row.Status is "FAILED_RETRYABLE" or "QUARANTINED" or "CHANGED") failures.Add(report);
        }
        var summary = new StructureScanSummary(
            "structure-scan-summary-v1",
            ScannerVersion,
            UtcNow(),
            rows.Count,
            new ReadOnlyDictionary<string, int>(statusCounts),
            new ReadOnlyDictionary<string, int>(structureCounts),
            new ReadOnlyDictionary<string, int>(sectionCounts),
            "classification.csv",
            "classification.html",
            "failures.csv");
        AtomicWriteJson(Path.Combine(batchDirectory, "summary.json"), summary);
        WriteCsv(Path.Combine(batchDirectory, "classification.csv"), reportRows);
        WriteCsv(Path.Combine(batchDirectory, "failures.csv"), failures);
        WriteClassificationHtml(Path.Combine(batchDirectory, "classification.html"), reportRows, statusCounts);
        return summary;
    }

    private static List<ItemRecord> AllItems(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT relative_path, source_path, extension, kind, size_bytes, mtime_ns, fingerprint, selected, status, attempts, result_path, error_text FROM items ORDER BY relative_path;";
        using var reader = command.ExecuteReader();
        var rows = new List<ItemRecord>();
        while (reader.Read()) rows.Add(ReadItem(reader));
        return rows;
    }

    private static JsonDocument? ReadResult(string batchDirectory, string relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath)) return null;
        try
        {
            var fullPath = Path.GetFullPath(Path.Combine(batchDirectory, relativePath.Replace('/', Path.DirectorySeparatorChar)));
            if (!IsWithin(fullPath, batchDirectory) || !File.Exists(fullPath)) return null;
            return JsonDocument.Parse(File.ReadAllText(fullPath, Encoding.UTF8));
        }
        catch (IOException) { return null; }
        catch (JsonException) { return null; }
    }

    private static string JsonString(JsonDocument? document, string property) =>
        document is not null && document.RootElement.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() ?? string.Empty : string.Empty;

    private static List<string> JsonStrings(JsonDocument? document, string property)
    {
        if (document is null || !document.RootElement.TryGetProperty(property, out var value) || value.ValueKind != JsonValueKind.Array) return [];
        return value.EnumerateArray().Where(static item => item.ValueKind == JsonValueKind.String).Select(static item => item.GetString() ?? string.Empty).ToList();
    }

    private static void Increment(IDictionary<string, int> values, string key) => values[key] = values.TryGetValue(key, out var count) ? count + 1 : 1;

    private static void WriteCsv(string path, IReadOnlyList<Dictionary<string, string>> rows)
    {
        var columns = rows.Count > 0 ? rows[0].Keys.ToList() : new List<string> { "relativePath", "status" };
        var builder = new StringBuilder();
        builder.AppendLine(string.Join(',', columns.Select(Csv)));
        foreach (var row in rows) builder.AppendLine(string.Join(',', columns.Select(column => Csv(row.GetValueOrDefault(column, string.Empty)))));
        AtomicWriteText(path, builder.ToString(), new UTF8Encoding(true));
    }

    private static string Csv(string text) => $"\"{text.Replace("\"", "\"\"")}\"";

    private static void WriteClassificationHtml(string path, IReadOnlyList<Dictionary<string, string>> rows, IReadOnlyDictionary<string, int> statusCounts)
    {
        var columns = rows.Count > 0 ? rows[0].Keys.ToList() : new List<string> { "relativePath", "status" };
        var countRows = string.Concat(statusCounts.Select(pair => $"<li>{Html(pair.Key)}: {pair.Value}</li>"));
        var bodyRows = string.Concat(rows.Select(row => "<tr>" + string.Concat(columns.Select(column => $"<td>{Html(row.GetValueOrDefault(column, string.Empty))}</td>")) + "</tr>"));
        var header = string.Concat(columns.Select(column => $"<th>{Html(column)}</th>"));
        var document = $"<!doctype html><html lang='ko'><head><meta charset='utf-8'><title>구조 사전검사</title>\n"
            + "<style>body{font-family:Segoe UI,sans-serif;margin:24px;color:#172b4d}table{border-collapse:collapse;width:100%;font-size:12px}th,td{border:1px solid #d7dde7;padding:6px;text-align:left;vertical-align:top}th{background:#edf2f7}code{font-family:Consolas,monospace}</style></head>\n"
            + $"<body><h1>Excel 구조 사전검사</h1><p>이 결과는 header/merge layout 후보만 다루며 품질·비교·승인 결론을 만들지 않습니다. Excel COM, Office, universal-grid DB는 사용하지 않았습니다.</p>\n"
            + $"<h2>상태</h2><ul>{countRows}</ul><p><a href='classification.csv'>CSV</a> · <a href='summary.json'>JSON</a> · <a href='failures.csv'>실패 CSV</a></p>\n"
            + $"<table><thead><tr>{header}</tr></thead><tbody>{bodyRows}</tbody></table></body></html>";
        AtomicWriteText(path, document, new UTF8Encoding(false));
    }

    private static string Html(string value) => System.Net.WebUtility.HtmlEncode(value);

    private static void AtomicWriteJson(string path, object value) => AtomicWriteText(path, JsonSerializer.Serialize(value, JsonOptions) + Environment.NewLine, new UTF8Encoding(false));

    private static void AtomicWriteText(string path, string content, Encoding encoding)
    {
        var directory = Path.GetDirectoryName(path) ?? throw new InvalidOperationException("출력 디렉터리를 확인할 수 없습니다.");
        Directory.CreateDirectory(directory);
        var temporary = path + ".tmp";
        File.WriteAllText(temporary, content, encoding);
        File.Move(temporary, path, overwrite: true);
    }

    private static void Execute(SqliteConnection connection, SqliteTransaction? transaction, string sql, params (string Name, object Value)[] parameters)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = sql;
        foreach (var (name, value) in parameters) command.Parameters.AddWithValue(name, value);
        command.ExecuteNonQuery();
    }

    private static ItemRecord ReadItem(SqliteDataReader reader) => new(
        reader.GetString(0), reader.GetString(1), reader.GetString(2), reader.GetString(3), reader.GetInt64(4), reader.GetInt64(5), reader.GetString(6), reader.GetInt32(7), reader.GetString(8), reader.GetInt32(9), reader.GetString(10), reader.GetString(11));

    private static string SourceFingerprint(string path)
    {
        var info = new FileInfo(Path.GetFullPath(path));
        var input = $"{info.FullName}|{info.Length}|{MtimeNanoseconds(info)}";
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(input))).ToLowerInvariant();
    }

    private static long MtimeNanoseconds(FileInfo info)
    {
        var unixTicks = info.LastWriteTimeUtc.Ticks - DateTime.UnixEpoch.Ticks;
        return checked(unixTicks * 100L);
    }

    private static string StableSourceId(string relativePath) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(relativePath.ToLowerInvariant()))).ToLowerInvariant()[..20];

    private static string FamilyTag(string relativePath)
    {
        var name = Path.GetFileName(relativePath).ToLowerInvariant();
        if (name.Contains("brs", StringComparison.Ordinal)) return "BRS";
        if (name.Contains("tiu", StringComparison.Ordinal)) return "TIU";
        if (name.Contains("msu", StringComparison.Ordinal)) return "MSU";
        return "Other";
    }

    private static string UtcNow() => DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ", CultureInfo.InvariantCulture);

    private static bool IsWithin(string child, string parent)
    {
        var relative = Path.GetRelativePath(Path.GetFullPath(parent), Path.GetFullPath(child));
        return !string.Equals(relative, "..", StringComparison.Ordinal) && !relative.StartsWith($"..{Path.DirectorySeparatorChar}", StringComparison.Ordinal) && !Path.IsPathRooted(relative);
    }

    private static string Clip(string value, int maximum) => value.Length <= maximum ? value : value[..maximum];

    private sealed class QuarantinedPackageException(string message) : Exception(message);

    private sealed record InventoryRecord(string RelativePath, string SourcePath, string Extension, string Kind, long SizeBytes, long MtimeNs, string Fingerprint, string FamilyTag);
    private sealed record ItemRecord(string RelativePath, string SourcePath, string Extension, string Kind, long SizeBytes, long MtimeNs, string Fingerprint, int Selected, string Status, int Attempts, string ResultPath, string ErrorText);
    private sealed record CellToken(int Row, int Column, string Token);
    private sealed record ParsedCell(int Row, int Column, string? Value, bool IsFormula);
    private sealed record HeaderRow(int Row, List<string> Tokens);
    private sealed record SectionRecord(string SheetName, string SheetState, string HeaderRange, int MergeCount, string Type, string Confidence);
    private sealed record SourceRecord(string RelativePath, string SourcePath, string Extension, long SizeBytes, long MtimeNs, string Fingerprint, string FamilyTag);

    private sealed class SheetRecord
    {
        public int SheetIndex { get; init; }
        public string SheetName { get; init; } = string.Empty;
        public string SheetState { get; init; } = "visible";
        public string Part { get; init; } = string.Empty;
        public string DeclaredDimension { get; set; } = string.Empty;
        public int MergeCount { get; set; }
        public List<string> MergeRangeSample { get; } = [];
        public int DrawingRelationshipCount { get; set; }
        public List<string> Warnings { get; } = [];
        public int FormulaCount { get; set; }
        public int ScannedCellCount { get; set; }
        public List<HeaderRow> HeaderRows { get; set; } = [];
        public List<SectionRecord> Sections { get; set; } = [];
        public string ScanStatus { get; set; } = string.Empty;
    }

    private sealed class StructureResult
    {
        public string SchemaVersion { get; init; } = string.Empty;
        public string ScannerVersion { get; init; } = string.Empty;
        public SourceRecord? Source { get; init; }
        public string ScanStatus { get; init; } = string.Empty;
        public bool ReadOnly { get; init; }
        public bool UsesCom { get; init; }
        public Dictionary<string, object> Package { get; init; } = new(StringComparer.Ordinal);
        public List<SheetRecord> Sheets { get; init; } = [];
        public string PrimaryStructure { get; init; } = string.Empty;
        public List<string> StructuralTypes { get; init; } = [];
        public List<string> Limitations { get; init; } = [];
        public double ElapsedSeconds { get; set; }
    }
}

internal sealed record StructureScanRequest(
    string ServiceDirectory,
    string? BatchFolder,
    string? ResumeBatchId,
    string? BatchId,
    int Pilot = 0,
    int Limit = 0,
    bool RetryFailed = false);

internal sealed record StructureScanRunResult(string BatchId, string BatchDirectory, StructureScanSummary Summary);

internal sealed record StructureScanSummary(
    string SchemaVersion,
    string ScannerVersion,
    string GeneratedAt,
    int TotalItems,
    IReadOnlyDictionary<string, int> StatusCounts,
    IReadOnlyDictionary<string, int> PrimaryStructureCounts,
    IReadOnlyDictionary<string, int> SectionCandidateCounts,
    string ClassificationCsv,
    string ClassificationHtml,
    string FailuresCsv);
