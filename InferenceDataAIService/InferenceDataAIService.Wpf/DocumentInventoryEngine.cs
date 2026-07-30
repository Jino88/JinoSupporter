using System.Security.Cryptography;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace InferenceDataAIService.Wpf;

// Bounded AI input. It summarizes the complete coordinate capture without
// exposing the full SQLite grid or relying on workbook filenames for grouping.
internal static class DocumentInventoryEngine
{
    internal static DocumentInventoryResult Write(string batchDirectory)
    {
        var database = Path.Combine(batchDirectory, "numeric-capture.sqlite");
        if (!File.Exists(database)) throw new FileNotFoundException("Document capture database is missing.", database);
        var rawRecords = new List<InventoryRecord>();
        using var connection = new SqliteConnection($"Data Source={database};Mode=ReadOnly"); connection.Open();
        using var workbooks = connection.CreateCommand(); workbooks.CommandText = "SELECT workbook_id, relative_path, capture_status FROM capture_workbooks ORDER BY relative_path;";
        using var reader = workbooks.ExecuteReader();
        while (reader.Read()) rawRecords.Add(ReadWorkbook(connection, reader.GetInt64(0), reader.GetString(1), reader.GetString(2)));
        var clusterSeeds = rawRecords.GroupBy(record => record.LayoutFamilyId, StringComparer.Ordinal)
            .Select(group => new LayoutClusterSeed(group.Key, group.ToList())).ToList();
        var clusterByFamily = clusterSeeds.ToDictionary(seed => seed.FamilyId, seed => CreateClusterId(seed.FamilyId), StringComparer.Ordinal);
        var records = rawRecords.Select(record => record with
        {
            LayoutClusterId = clusterByFamily[record.LayoutFamilyId],
            RoutingState = RoutingStateFor(record)
        }).ToList();
        var inventoryPath = Path.Combine(batchDirectory, "document-inventory.jsonl");
        File.WriteAllLines(inventoryPath, records.Select(record => JsonSerializer.Serialize(record)), new UTF8Encoding(false));
        var signatures = records.GroupBy(record => record.LayoutSignatureId).Select(group => new SignatureRecord(group.Key, group.Count(), group.Select(record => record.RelativePath).ToList(), group.First())).ToList();
        var signaturePath = Path.Combine(batchDirectory, "layout-signatures.json");
        File.WriteAllText(signaturePath, JsonSerializer.Serialize(new { schemaVersion = "layout-signatures-v1", signatures }, new JsonSerializerOptions { WriteIndented = true }) + "\n", new UTF8Encoding(false));
        var clusters = clusterSeeds.Select(seed =>
        {
            var members = records.Where(record => string.Equals(record.LayoutFamilyId, seed.FamilyId, StringComparison.Ordinal)).ToList();
            return new LayoutClusterRecord(
                clusterByFamily[seed.FamilyId],
                seed.FamilyId,
                members.Count,
                members.Select(member => member.RelativePath).ToList(),
                members.Select(member => member.LayoutSignatureId).Distinct(StringComparer.Ordinal).ToList(),
                members.Select(member => member.RoutingState).Distinct(StringComparer.Ordinal).ToList(),
                members[0]);
        }).OrderByDescending(cluster => cluster.FileCount).ToList();
        var clusterPath = Path.Combine(batchDirectory, "layout-clusters.json");
        File.WriteAllText(clusterPath, JsonSerializer.Serialize(new { schemaVersion = "layout-clusters-v1", clusters }, new JsonSerializerOptions { WriteIndented = true }) + "\n", new UTF8Encoding(false));
        var summaryPath = Path.Combine(batchDirectory, "layout-cluster-summary.json");
        var summary = clusters.Select(cluster =>
        {
            var sections = cluster.Representative.Sheets.SelectMany(sheet => sheet.Sections).ToList();
            return new
            {
                cluster.LayoutClusterId,
                cluster.FileCount,
                cluster.RoutingStates,
                // The refinement prompt receives a bounded, path-free structural
                // description plus at most two traceable examples. It never has to
                // discover or read the large local capture artifacts itself.
                representativePaths = cluster.MemberPaths.Take(2).ToList(),
                tableCount = sections.Count,
                tableShape = sections.Select(section => new
                {
                    candidateType = section.CandidateType,
                    rowSpan = SpanBucket(section.EndRow - section.StartRow + 1),
                    columnSpan = SpanBucket(section.EndColumn - section.StartColumn + 1)
                }).ToList(),
                sectionRecipe = sections.Select(section => new
                {
                    section.CandidateType,
                    title = StructuralTokens([section.Title]).Take(6),
                    headers = StructuralTokens(section.Headers).Take(12),
                    logicalRowFacets = StructuralTokens(section.LogicalRowFacets).Take(12)
                }).ToList()
            };
        });
        File.WriteAllText(summaryPath, JsonSerializer.Serialize(new { schemaVersion = "layout-cluster-summary-v1", clusters = summary }, new JsonSerializerOptions { WriteIndented = true }) + "\n", new UTF8Encoding(false));
        var semanticPath = Path.Combine(batchDirectory, "layout-semantic-summary.json");
        var semantic = clusters.GroupBy(SemanticCategory).Select(group => new
        {
            category = group.Key,
            topLevelCategory = TopLevelCategory(group.Key),
            fileCount = group.Sum(cluster => cluster.FileCount),
            layoutClusterIds = group.Select(cluster => cluster.LayoutClusterId).ToList()
        }).OrderByDescending(group => group.fileCount);
        File.WriteAllText(semanticPath, JsonSerializer.Serialize(new { schemaVersion = "layout-semantic-summary-v1", categories = semantic }, new JsonSerializerOptions { WriteIndented = false }) + "\n", new UTF8Encoding(false));
        return new DocumentInventoryResult(inventoryPath, signaturePath, clusterPath, summaryPath, semanticPath, records.Count, signatures.Count, clusters.Count);
    }

    private static InventoryRecord ReadWorkbook(SqliteConnection connection, long workbookId, string path, string status)
    {
        using var sheets = connection.CreateCommand(); sheets.CommandText = "SELECT sheet_id, sheet_name, declared_dimension, merge_count FROM captured_sheets WHERE workbook_id=$id ORDER BY sheet_index;"; sheets.Parameters.AddWithValue("$id", workbookId);
        using var reader = sheets.ExecuteReader(); var values = new List<InventorySheet>();
        while (reader.Read())
        {
            var sheetId = reader.GetInt64(0); values.Add(new InventorySheet(reader.GetString(1), reader.GetString(2), reader.GetInt32(3), ReadSections(connection, sheetId)));
        }
        var canonical = string.Join("|", values.Select(sheet => $"{sheet.Dimension}:{sheet.MergeCount}:{string.Join("/", sheet.Sections.Select(section => section.Canonical))}"));
        // Grouping must survive routine row/column movement and table growth.
        // Keep only the ordered section kinds and the meaningful data axes here;
        // strict coordinates remain available through LayoutSignatureId for evidence.
        var familyCanonical = string.Join("|", values.Select(sheet => $"S:{string.Join("/", sheet.Sections.Select(section => section.FamilyCanonical))}"));
        var signature = "sig-" + Hash(canonical);
        var family = "family-" + Hash(familyCanonical);
        return new InventoryRecord(path, status, signature, family, string.Empty, string.Empty, values);
    }

    private static List<InventorySection> ReadSections(SqliteConnection connection, long sheetId)
    {
        using var tables = connection.CreateCommand(); tables.CommandText = "SELECT table_id, start_row, end_row, start_column, end_column, header_start_row, header_end_row, candidate_type FROM numeric_table_candidates WHERE sheet_id=$sheet ORDER BY start_row;"; tables.Parameters.AddWithValue("$sheet", sheetId);
        using var reader = tables.ExecuteReader(); var values = new List<InventorySection>();
        while (reader.Read())
        {
            var tableId = reader.GetInt64(0);
            var context = CapturedTableContextResolver.Resolve(connection, new CapturedTableRegion(
                sheetId, string.Empty, reader.GetInt32(1), reader.GetInt32(2), reader.GetInt32(3), reader.GetInt32(4), reader.GetInt32(5), reader.GetInt32(6)));
            var headers = context.HeaderLabels.Count > 0 ? context.HeaderLabels : Labels(connection, tableId, "HEADER");
            var title = context.Title;
            values.Add(new InventorySection(title, reader.GetInt32(1), reader.GetInt32(2), reader.GetInt32(3), reader.GetInt32(4), reader.GetString(7), headers, context.LogicalRowFacets));
        }
        return values;
    }

    private static List<string> Labels(SqliteConnection connection, long tableId, string role)
    {
        using var command = connection.CreateCommand(); command.CommandText = "SELECT label_text FROM numeric_table_labels WHERE table_id=$id AND label_role=$role ORDER BY row_index, column_index;"; command.Parameters.AddWithValue("$id", tableId); command.Parameters.AddWithValue("$role", role); using var reader = command.ExecuteReader(); var values = new List<string>(); while (reader.Read()) values.Add(reader.GetString(0)); return values;
    }
    private static string TextAt(SqliteConnection connection, long sheetId, int row, int column)
    {
        using var command = connection.CreateCommand(); command.CommandText = "SELECT text_value FROM captured_text_cells WHERE sheet_id=$sheet AND row_index=$row AND column_index=$column;"; command.Parameters.AddWithValue("$sheet", sheetId); command.Parameters.AddWithValue("$row", row); command.Parameters.AddWithValue("$column", column); return command.ExecuteScalar() as string ?? string.Empty;
    }
    private static string TitleNearTable(SqliteConnection connection, long sheetId, int headerStart, int startColumn, int endColumn)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT text_value FROM captured_text_cells WHERE sheet_id=$sheet AND row_index BETWEEN $first AND $last AND column_index BETWEEN $start AND $end AND text_value <> '' ORDER BY row_index, column_index;";
        command.Parameters.AddWithValue("$sheet", sheetId); command.Parameters.AddWithValue("$first", Math.Max(1, headerStart - 2)); command.Parameters.AddWithValue("$last", Math.Max(1, headerStart - 1)); command.Parameters.AddWithValue("$start", startColumn); command.Parameters.AddWithValue("$end", endColumn);
        using var reader = command.ExecuteReader(); var values = new List<string>();
        while (reader.Read()) values.Add(reader.GetString(0));
        return string.Join(" ", values);
    }
    private static string RoutingStateFor(InventoryRecord record) =>
        !string.Equals(record.CaptureStatus, "CAPTURED", StringComparison.OrdinalIgnoreCase) ? "CAPTURE_INCOMPLETE" :
        record.Sheets.All(sheet => sheet.Sections.Count == 0) ? "EMPTY_LAYOUT" : "ASSIGNABLE";
    private static string CreateClusterId(string familyId) => "cluster-" + familyId["family-".Length..];
    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant()[..16];
    private static string Bucket(int value) => value switch { <= 0 => "0", <= 3 => "1-3", <= 10 => "4-10", _ => "11+" };
    private static string SpanBucket(int value) => value switch { <= 1 => "1", <= 3 => "2-3", <= 6 => "4-6", <= 12 => "7-12", <= 25 => "13-25", _ => "26+" };
    private static readonly HashSet<string> StructuralTerms = ["INPUT", "OK", "NG", "RATE", "SIGMA", "HEARING", "NOISE", "TOUCH", "SPL", "THD", "DIMENSION", "MIN", "MAX", "AVG", "TENSION", "BENDING", "FRAME", "MOLD", "DATE", "LINE", "TOTAL", "FUNCTION", "TEST", "NORMAL", "BAKO", "VISUAL", "DEFECT", "VISION", "COIL", "YOKE", "POSITION", "P1", "P2", "P3", "P4", "P5", "P6"];
    private static IReadOnlyList<string> StructuralTokens(IEnumerable<string> values) => values.SelectMany(value => Regex.Matches(value.ToUpperInvariant(), "[A-Z][A-Z0-9+/-]*").Select(match => match.Value))
        .Where(token => StructuralTerms.Contains(token)).Distinct(StringComparer.Ordinal).OrderBy(token => token, StringComparer.Ordinal).Take(12).ToList();
    private static string SemanticCategory(LayoutClusterRecord cluster)
    {
        var tokens = StructuralTokens(cluster.Representative.Sheets.SelectMany(sheet => sheet.Sections).SelectMany(section => new[] { section.Title }.Concat(section.Headers).Concat(section.LogicalRowFacets))).ToHashSet(StringComparer.Ordinal);
        var axis = tokens.Overlaps(["SIGMA", "HEARING", "SPL", "THD", "NOISE", "TOUCH"]) ? "acoustic-ng-dashboard" :
            tokens.Overlaps(["DIMENSION", "MIN", "MAX", "AVG", "POSITION"]) ? "measurement-dimension-dashboard" :
            tokens.Contains("TENSION") ? "tension-dashboard" :
            tokens.Overlaps(["FUNCTION", "PROCESS", "BAKO", "VISUAL"]) ? "function-process-dashboard" :
            tokens.Overlaps(["INPUT", "OK", "NG", "RATE", "DEFECT"]) ? "quality-ng-dashboard" : "general-table-dashboard";
        var sections = cluster.Representative.Sheets.SelectMany(sheet => sheet.Sections).ToList();
        var pattern = string.Join("-", sections.Select(section => section.CandidateType
            .Replace("_LAYOUT_CANDIDATE", string.Empty, StringComparison.Ordinal)
            .Replace("NUMERIC_TABLE_", "TABLE_", StringComparison.Ordinal))
            .Distinct(StringComparer.Ordinal).Take(4));
        return $"{axis}--s{Math.Min(sections.Count, 9)}--{(string.IsNullOrWhiteSpace(pattern) ? "EMPTY" : pattern)}";
    }
    private static string TopLevelCategory(string semanticCategory)
    {
        var separator = semanticCategory.IndexOf("--", StringComparison.Ordinal);
        return separator < 0 ? semanticCategory : semanticCategory[..separator];
    }
    private sealed record InventoryRecord(string RelativePath, string CaptureStatus, string LayoutSignatureId, string LayoutFamilyId, string LayoutClusterId, string RoutingState, IReadOnlyList<InventorySheet> Sheets);
    private sealed record InventorySheet(string Name, string Dimension, int MergeCount, IReadOnlyList<InventorySection> Sections);
    private sealed record InventorySection(string Title, int StartRow, int EndRow, int StartColumn, int EndColumn, string CandidateType, IReadOnlyList<string> Headers, IReadOnlyList<string> LogicalRowFacets)
    {
        public string Canonical => $"{StartRow}-{EndRow}:{StartColumn}-{EndColumn}:{CandidateType}:{string.Join(",", Headers)}";
        public string FamilyCanonical => $"{CandidateType}:{string.Join(",", StructuralTokens(new[] { Title }.Concat(Headers).Concat(LogicalRowFacets)))}";
    }
    private sealed record LayoutClusterSeed(string FamilyId, IReadOnlyList<InventoryRecord> Members);
    private sealed record LayoutClusterRecord(string LayoutClusterId, string LayoutFamilyId, int FileCount, IReadOnlyList<string> MemberPaths, IReadOnlyList<string> StrictSignatureIds, IReadOnlyList<string> RoutingStates, InventoryRecord Representative);
    private sealed record SignatureRecord(string SignatureId, int FileCount, IReadOnlyList<string> MemberPaths, InventoryRecord Representative);
}
internal sealed record DocumentInventoryResult(string InventoryPath, string SignaturePath, string ClusterPath, string ClusterSummaryPath, string SemanticSummaryPath, int WorkbookCount, int SignatureCount, int ClusterCount);
