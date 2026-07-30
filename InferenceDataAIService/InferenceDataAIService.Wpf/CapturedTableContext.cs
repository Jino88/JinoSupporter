using Microsoft.Data.Sqlite;

namespace InferenceDataAIService.Wpf;

// Shared persisted-capture contract used by both AI inventory evidence and HTML
// rendering.  Numeric candidates identify a data island, while this resolver
// restores the title band, label gutter, and fully-contained merged ranges that
// give those numbers their workbook meaning.
internal readonly record struct CapturedMergeRange(int Top, int Left, int Bottom, int Right)
{
    internal bool Intersects(int top, int left, int bottom, int right) => !(Bottom < top || Top > bottom || Right < left || Left > right);
    internal bool Contains(int row, int column) => row >= Top && row <= Bottom && column >= Left && column <= Right;
}

internal sealed record CapturedTableRegion(long SheetId, string SheetName, int StartRow, int EndRow, int StartColumn, int EndColumn, int HeaderStartRow, int HeaderEndRow);

internal sealed record CapturedTableContext(
    int TopRow,
    int BottomRow,
    int LeftColumn,
    int RightColumn,
    int DataStartRow,
    string Title,
    IReadOnlyList<string> HeaderLabels,
    IReadOnlyList<string> LogicalRowFacets,
    IReadOnlyList<CapturedMergeRange> MergeRanges);

internal static class CapturedTableContextResolver
{
    private const int TitleLookbackRows = 8;
    private const int LabelGutterColumns = 8;
    private const int MaxDataColumns = 32;
    private const int MaxRenderedColumns = 40;
    private const int MaxLogicalRows = 8;

    internal static CapturedTableContext Resolve(SqliteConnection connection, CapturedTableRegion region)
    {
        var dataStart = Math.Max(region.StartRow, region.HeaderEndRow + 1);
        var bottom = Math.Min(region.EndRow, dataStart + 249);
        var top = Math.Min(region.HeaderStartRow, dataStart);
        var left = Math.Min(region.StartColumn, FindLabelGutterLeft(connection, region.SheetId, top, bottom, region.StartColumn));
        var right = Math.Min(region.EndColumn, region.StartColumn + MaxDataColumns - 1);
        var merges = ReadMergeRanges(connection, region.SheetId);

        top = Math.Min(top, FindPreHeaderContextTop(connection, region.SheetId, top, left, right));
        ExpandByMerges(merges, ref top, ref bottom, ref left, ref right);
        left = Math.Max(1, left);
        right = Math.Min(right, left + MaxRenderedColumns - 1);
        var includedMerges = merges.Where(merge => merge.Top >= top && merge.Bottom <= bottom && merge.Left >= left && merge.Right <= right).ToList();
        var texts = ReadCells(connection, "SELECT row_index, column_index, text_value FROM captured_text_cells WHERE sheet_id=$sheet AND row_index BETWEEN $start AND $end;", region.SheetId, top, bottom);
        var values = ReadCells(connection, "SELECT row_index, column_index, value_text FROM numeric_cells WHERE sheet_id=$sheet AND row_index BETWEEN $start AND $end;", region.SheetId, top, bottom);
        var dates = ReadCells(connection, "SELECT row_index, column_index, date_value FROM date_cells WHERE sheet_id=$sheet AND row_index BETWEEN $start AND $end;", region.SheetId, top, bottom);
        var title = ResolveTitle(texts, includedMerges, region, top, left, right);
        var headers = HeaderLabels(texts, includedMerges, top, region.HeaderEndRow, left, right);
        var rows = LogicalRows(texts, values, dates, includedMerges, region, top, bottom, left, right, dataStart);
        return new CapturedTableContext(top, bottom, left, right, dataStart, title, headers, rows, includedMerges);
    }

    internal static string EffectiveValue(IReadOnlyDictionary<(int Row, int Column), string> texts, IReadOnlyDictionary<(int Row, int Column), string> dates, IReadOnlyDictionary<(int Row, int Column), string> values, IReadOnlyList<CapturedMergeRange> merges, int row, int column)
    {
        var merge = merges.FirstOrDefault(range => range.Contains(row, column));
        var key = merge == default ? (row, column) : (merge.Top, merge.Left);
        return texts.GetValueOrDefault(key) ?? dates.GetValueOrDefault(key) ?? values.GetValueOrDefault(key) ?? string.Empty;
    }

    internal static CapturedMergeRange? MergeAt(IReadOnlyList<CapturedMergeRange> merges, int row, int column)
    {
        foreach (var merge in merges)
            if (merge.Contains(row, column)) return merge;
        return null;
    }

    internal static IReadOnlyDictionary<(int Row, int Column), string> ReadText(SqliteConnection connection, long sheetId, int start, int end) =>
        ReadCells(connection, "SELECT row_index, column_index, text_value FROM captured_text_cells WHERE sheet_id=$sheet AND row_index BETWEEN $start AND $end;", sheetId, start, end);
    internal static IReadOnlyDictionary<(int Row, int Column), string> ReadNumeric(SqliteConnection connection, long sheetId, int start, int end) =>
        ReadCells(connection, "SELECT row_index, column_index, value_text FROM numeric_cells WHERE sheet_id=$sheet AND row_index BETWEEN $start AND $end;", sheetId, start, end);
    internal static IReadOnlyDictionary<(int Row, int Column), string> ReadDates(SqliteConnection connection, long sheetId, int start, int end) =>
        ReadCells(connection, "SELECT row_index, column_index, date_value FROM date_cells WHERE sheet_id=$sheet AND row_index BETWEEN $start AND $end;", sheetId, start, end);

    private static void ExpandByMerges(IReadOnlyList<CapturedMergeRange> merges, ref int top, ref int bottom, ref int left, ref int right)
    {
        for (var pass = 0; pass < 4; pass++)
        {
            var changed = false;
            foreach (var merge in merges)
            {
                if (!merge.Intersects(top, left, bottom, right)) continue;
                var nextTop = Math.Min(top, merge.Top); var nextBottom = Math.Max(bottom, merge.Bottom);
                var nextLeft = Math.Min(left, merge.Left); var nextRight = Math.Max(right, merge.Right);
                changed |= nextTop != top || nextBottom != bottom || nextLeft != left || nextRight != right;
                top = nextTop; bottom = nextBottom; left = nextLeft; right = nextRight;
            }
            if (!changed) break;
        }
    }

    private static int FindLabelGutterLeft(SqliteConnection connection, long sheetId, int top, int bottom, int firstColumn)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT MIN(column_index) FROM captured_text_cells WHERE sheet_id=$sheet AND row_index BETWEEN $top AND $bottom AND column_index BETWEEN $left AND $right;";
        command.Parameters.AddWithValue("$sheet", sheetId); command.Parameters.AddWithValue("$top", top); command.Parameters.AddWithValue("$bottom", bottom);
        command.Parameters.AddWithValue("$left", Math.Max(1, firstColumn - LabelGutterColumns)); command.Parameters.AddWithValue("$right", Math.Max(1, firstColumn - 1));
        return command.ExecuteScalar() is long value ? Convert.ToInt32(value) : firstColumn;
    }

    private static int FindPreHeaderContextTop(SqliteConnection connection, long sheetId, int headerStart, int left, int right)
    {
        var lower = Math.Max(1, headerStart - TitleLookbackRows);
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT row_index, text_value FROM captured_text_cells WHERE sheet_id=$sheet AND row_index BETWEEN $first AND $last AND column_index BETWEEN $left AND $right ORDER BY row_index DESC, column_index;";
        command.Parameters.AddWithValue("$sheet", sheetId); command.Parameters.AddWithValue("$first", lower); command.Parameters.AddWithValue("$last", headerStart - 1); command.Parameters.AddWithValue("$left", left); command.Parameters.AddWithValue("$right", right);
        using var reader = command.ExecuteReader();
        var candidateRows = new HashSet<int>();
        while (reader.Read()) if (reader.GetString(1).Trim().Length >= 6) candidateRows.Add(reader.GetInt32(0));
        return candidateRows.Count == 0 ? headerStart : candidateRows.Min();
    }

    private static string ResolveTitle(IReadOnlyDictionary<(int Row, int Column), string> texts, IReadOnlyList<CapturedMergeRange> merges, CapturedTableRegion region, int top, int left, int right)
    {
        var values = texts
            .Where(cell => cell.Key.Row >= top && cell.Key.Row <= region.HeaderEndRow && cell.Key.Column >= left && cell.Key.Column <= right && cell.Value.Trim().Length >= 4)
            .Select(cell => new
            {
                cell.Key.Row,
                cell.Key.Column,
                Text = cell.Value.Trim(),
                IsWideMerge = merges.Any(merge => merge.Top == cell.Key.Row && merge.Left == cell.Key.Column && merge.Right - merge.Left >= 2 && merge.Intersects(top, region.StartColumn, region.HeaderEndRow, right))
            })
            .OrderByDescending(value => value.IsWideMerge)
            .ThenByDescending(value => value.Row)
            .ThenByDescending(value => value.Text.Length)
            .ToList();
        return values.FirstOrDefault()?.Text ?? string.Empty;
    }

    private static IReadOnlyList<string> HeaderLabels(IReadOnlyDictionary<(int Row, int Column), string> texts, IReadOnlyList<CapturedMergeRange> merges, int top, int headerEnd, int left, int right)
    {
        var labels = new List<string>();
        for (var column = left; column <= right; column++)
        {
            var parts = new List<string>();
            for (var row = top; row <= headerEnd; row++)
            {
                var value = EffectiveValue(texts, EmptyCells, EmptyCells, merges, row, column).Trim();
                if (value.Length > 0 && !parts.Contains(value, StringComparer.OrdinalIgnoreCase)) parts.Add(value);
            }
            if (parts.Count > 0) labels.Add(string.Join(" / ", parts));
        }
        return labels;
    }

    private static IReadOnlyList<string> LogicalRows(IReadOnlyDictionary<(int Row, int Column), string> texts, IReadOnlyDictionary<(int Row, int Column), string> values, IReadOnlyDictionary<(int Row, int Column), string> dates, IReadOnlyList<CapturedMergeRange> merges, CapturedTableRegion region, int top, int bottom, int left, int right, int dataStart)
    {
        var result = new List<string>();
        for (var row = dataStart; row <= bottom && result.Count < MaxLogicalRows; row++)
        {
            var labels = new List<string>();
            for (var column = left; column < region.StartColumn; column++)
            {
                var value = EffectiveValue(texts, dates, values, merges, row, column).Trim();
                if (value.Length > 0 && !labels.Contains(value, StringComparer.OrdinalIgnoreCase)) labels.Add(value);
            }
            var metrics = new List<string>();
            var seenMetricAnchors = new HashSet<(int Row, int Column)>();
            for (var column = region.StartColumn; column <= right; column++)
            {
                var merge = MergeAt(merges, row, column);
                var anchor = merge is { } range ? (range.Top, range.Left) : (row, column);
                if (!seenMetricAnchors.Add(anchor) || !values.TryGetValue(anchor, out var value)) continue;
                var header = HeaderForColumn(texts, merges, top, region.HeaderEndRow, column);
                if (header.Length > 0 && header.Contains("RATE", StringComparison.OrdinalIgnoreCase) && TryReadNumeric(value, out var rate) && rate is >= 0 and <= 1)
                    metrics.Add($"{header}={rate * 100d:F2}%");
            }
            if (labels.Count > 0 || metrics.Count > 0)
                result.Add($"row {row}: {string.Join("; ", labels.Take(4).Concat(metrics.Take(4)))}");
        }
        return result;
    }

    private static string HeaderForColumn(IReadOnlyDictionary<(int Row, int Column), string> texts, IReadOnlyList<CapturedMergeRange> merges, int top, int headerEnd, int column)
    {
        for (var row = headerEnd; row >= top; row--)
        {
            var value = EffectiveValue(texts, EmptyCells, EmptyCells, merges, row, column).Trim();
            if (value.Length > 0) return value;
        }
        return string.Empty;
    }

    private static bool TryReadNumeric(string value, out double number)
    {
        var normalized = value.Trim();
        var isPercent = normalized.EndsWith('%');
        if (isPercent) normalized = normalized[..^1].Trim();
        if (!double.TryParse(normalized, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number)) return false;
        if (isPercent) number /= 100d;
        return true;
    }

    private static IReadOnlyList<CapturedMergeRange> ReadMergeRanges(SqliteConnection connection, long sheetId)
    {
        using var command = connection.CreateCommand(); command.CommandText = "SELECT range_ref FROM captured_merge_ranges WHERE sheet_id=$sheet;"; command.Parameters.AddWithValue("$sheet", sheetId);
        using var reader = command.ExecuteReader(); var ranges = new List<CapturedMergeRange>();
        while (reader.Read() && TryParseRange(reader.GetString(0), out var range)) ranges.Add(range);
        return ranges;
    }

    private static bool TryParseRange(string value, out CapturedMergeRange range)
    {
        range = default; var parts = value.Split(':', 2);
        if (parts.Length != 2 || !TryParseCell(parts[0], out var top, out var left) || !TryParseCell(parts[1], out var bottom, out var right)) return false;
        range = new CapturedMergeRange(Math.Min(top, bottom), Math.Min(left, right), Math.Max(top, bottom), Math.Max(left, right)); return true;
    }

    private static bool TryParseCell(string value, out int row, out int column)
    {
        row = 0; column = 0; var index = 0;
        while (index < value.Length && char.IsLetter(value[index])) { column = column * 26 + char.ToUpperInvariant(value[index]) - 'A' + 1; index++; }
        return index > 0 && index < value.Length && int.TryParse(value[index..], out row) && row > 0;
    }

    private static IReadOnlyDictionary<(int Row, int Column), string> ReadCells(SqliteConnection connection, string sql, long sheetId, int start, int end)
    {
        using var command = connection.CreateCommand(); command.CommandText = sql; command.Parameters.AddWithValue("$sheet", sheetId); command.Parameters.AddWithValue("$start", start); command.Parameters.AddWithValue("$end", end);
        using var reader = command.ExecuteReader(); var cells = new Dictionary<(int, int), string>();
        while (reader.Read()) cells[(reader.GetInt32(0), reader.GetInt32(1))] = reader.GetString(2);
        return cells;
    }

    private static readonly IReadOnlyDictionary<(int Row, int Column), string> EmptyCells = new Dictionary<(int, int), string>();
}
