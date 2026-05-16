using System.Globalization;

namespace BmesNgRateStandalone.Services.GraphMaker;

public enum GraphMakerInputMode
{
    HeaderTable,
    DateValue,
    NoXMultiY,
    ProcessTrend,
    HeatMap
}

public enum GraphMakerGraphType
{
    Line,
    Scatter,
    Average,
    Cpk,
    NormalDistribution,
    NoXMultiY,
    ProcessTrend,
    HeatMap
}

public enum GraphMakerScatterPointLayout
{
    Aligned,
    Spread
}

public enum GraphMakerScatterLabelPosition
{
    OnTick,
    BetweenTicks
}

public enum GraphMakerLineMode
{
    Line,
    Smoothing
}

public enum GraphMakerRateUnit
{
    Percent,
    Ppm
}

public sealed record GraphMakerParseOptions(string Delimiter, int HeaderRowNumber, bool HeaderNotUse = false);

public sealed class GraphMakerTable
{
    public List<string> Headers { get; } = [];
    public List<string[]> Rows { get; } = [];

    public bool HasData => Headers.Count > 0 && Rows.Count > 0;

    public int IndexOf(string column) =>
        Headers.FindIndex(h => string.Equals(h, column, StringComparison.Ordinal));

    public string GetCell(string[] row, int index) =>
        index >= 0 && index < row.Length ? row[index] : string.Empty;

    public double? TryGetDouble(string[] row, int index)
    {
        if (index < 0 || index >= row.Length)
        {
            return null;
        }

        return double.TryParse(row[index], NumberStyles.Any, CultureInfo.InvariantCulture, out double value)
            || double.TryParse(row[index], out value)
            ? value
            : null;
    }

    public double?[] BuildNumericValues(string column)
    {
        int index = IndexOf(column);
        return Rows.Select(row => TryGetDouble(row, index)).ToArray();
    }

    public List<double> GetNumericValues(string column)
    {
        int index = IndexOf(column);
        return Rows
            .Select(row => TryGetDouble(row, index))
            .Where(v => v.HasValue)
            .Select(v => v!.Value)
            .ToList();
    }
}

public sealed class GraphMakerRequest
{
    public GraphMakerInputMode InputMode { get; init; }
    public GraphMakerGraphType GraphType { get; init; }
    public string XColumn { get; init; } = string.Empty;
    public List<string> XColumns { get; init; } = [];
    public List<string> YColumns { get; init; } = [];
    public string HeatMapXColumn { get; init; } = string.Empty;
    public string HeatMapYColumn { get; init; } = string.Empty;
    public string HeatMapValueColumn { get; init; } = string.Empty;
    public bool HeatMapUseNgRate { get; init; }
    public string HeatMapNgColumn { get; init; } = string.Empty;
    public string HeatMapInputColumn { get; init; } = string.Empty;
    public GraphMakerRateUnit HeatMapRateUnit { get; init; }
    public string GraphSpecColumn { get; init; } = string.Empty;
    public string GraphUpperColumn { get; init; } = string.Empty;
    public string GraphLowerColumn { get; init; } = string.Empty;
    public Dictionary<string, string> GraphColumnGroups { get; init; } = new(StringComparer.Ordinal);
    public double? UpperLimit { get; init; }
    public double? SpecLimit { get; init; }
    public double? LowerLimit { get; init; }
    public int? ProcessSpecRowNumber { get; init; }
    public int? ProcessUpperRowNumber { get; init; }
    public int? ProcessLowerRowNumber { get; init; }
    public bool ShowProcessContour { get; init; }
    public bool ShowProcessTrendLine { get; init; }
    public bool ShowProcessMax { get; init; }
    public bool ShowProcessMin { get; init; }
    public bool ShowProcessAvg { get; init; }
    public bool ShowProcessStd { get; init; }
    public bool ShowProcessCpu { get; init; }
    public bool ShowProcessCpl { get; init; }
    public bool ShowProcessCpk { get; init; }
    public GraphMakerScatterPointLayout ScatterPointLayout { get; init; }
    public GraphMakerScatterLabelPosition ScatterLabelPosition { get; init; }
    public GraphMakerLineMode LineMode { get; init; }
    public bool ScatterHeaderNotUse { get; init; }
    public bool ShowScatterMax { get; init; }
    public bool ShowScatterMin { get; init; }
    public bool ShowScatterAvg { get; init; }
    public bool ShowScatterStd { get; init; }
    public bool ShowScatterCpu { get; init; }
    public bool ShowScatterCpl { get; init; }
    public bool ShowScatterCpk { get; init; }
}

public sealed record GraphMakerPoint(double X, double Y);
public sealed record GraphMakerStat(string Name, double Value);
public sealed record GraphMakerStatGroup(double X, string Label, List<GraphMakerStat> Stats);

public sealed class GraphMakerSeries
{
    public string Name { get; init; } = string.Empty;
    public double?[]? Data { get; init; }
    public List<GraphMakerPoint>? Points { get; init; }
    public bool IsLimit { get; init; }
    public string? Color { get; init; }
    public bool ShowLine { get; init; }
    public bool Dashed { get; init; }
    public double? LabelValue { get; init; }
    public int? PointRadius { get; init; }
}

public sealed record GraphMakerLimit(string Name, double? Value, string Color);

public sealed class GraphMakerChartPayload
{
    public string Kind { get; init; } = string.Empty;
    public string Title { get; init; } = string.Empty;
    public List<string> Labels { get; init; } = [];
    public List<GraphMakerSeries> Series { get; init; } = [];
    public List<GraphMakerLimit> Limits { get; init; } = [];
    public List<GraphMakerStat> Stats { get; init; } = [];
    public List<GraphMakerStatGroup> StatGroups { get; init; } = [];
    public bool UseCategoryXAxis { get; init; }
    public string ScatterLabelPosition { get; init; } = string.Empty;
    public string LineMode { get; init; } = string.Empty;
    public List<string> XLabels { get; init; } = [];
    public List<string> YLabels { get; init; } = [];
    public double?[][] Matrix { get; init; } = [];
    public string?[][] MatrixLabels { get; init; } = [];
}

public static class GraphMakerParser
{
    public static GraphMakerTable Parse(string text, GraphMakerParseOptions options)
    {
        var table = new GraphMakerTable();
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n')
            .Split('\n', StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length == 0)
        {
            return table;
        }

        string delimiter = ResolveDelimiter(lines, options.Delimiter);
        if (options.HeaderNotUse)
        {
            int columnCount = lines.Select(line => SplitLine(line, delimiter).Length).DefaultIfEmpty(0).Max();
            table.Headers.AddRange(Enumerable.Range(1, columnCount).Select(i => $"Column {i}"));
            foreach (var line in lines)
            {
                var cells = SplitLine(line, delimiter);
                if (cells.Any(c => !string.IsNullOrWhiteSpace(c)))
                {
                    table.Rows.Add(cells);
                }
            }

            return table;
        }

        int headerIndex = Math.Clamp(options.HeaderRowNumber - 1, 0, lines.Length - 1);
        table.Headers.AddRange(MakeUnique(SplitLine(lines[headerIndex], delimiter)));

        for (int i = headerIndex + 1; i < lines.Length; i++)
        {
            var cells = SplitLine(lines[i], delimiter);
            if (cells.Any(c => !string.IsNullOrWhiteSpace(c)))
            {
                table.Rows.Add(cells);
            }
        }

        return table;
    }

    private static string ResolveDelimiter(string[] lines, string delimiter)
    {
        if (!string.Equals(delimiter, "auto", StringComparison.OrdinalIgnoreCase))
        {
            return delimiter;
        }

        var candidates = new[] { "\t", ",", ";", "|" };
        return candidates
            .Select(d => new { Delimiter = d, Count = lines.Take(20).Sum(line => line.Count(ch => ch.ToString() == d)) })
            .OrderByDescending(x => x.Count)
            .First().Delimiter;
    }

    private static string[] SplitLine(string line, string delimiter) =>
        line.Split(delimiter).Select(cell => cell.Trim().Trim('"')).ToArray();

    private static IEnumerable<string> MakeUnique(IEnumerable<string> headers)
    {
        var used = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var raw in headers)
        {
            var name = string.IsNullOrWhiteSpace(raw) ? "Column" : raw.Trim();
            if (!used.TryGetValue(name, out int count))
            {
                used[name] = 1;
                yield return name;
            }
            else
            {
                used[name] = count + 1;
                yield return $"{name}_{count + 1}";
            }

        }
    }
}

public static class GraphMakerPayloadBuilder
{
    public static GraphMakerChartPayload Build(GraphMakerTable table, GraphMakerRequest request)
    {
        var graphType = NormalizeGraphType(request);
        return graphType == GraphMakerGraphType.HeatMap
            ? BuildHeatMapPayload(table, request)
            : BuildSeriesPayload(table, request, graphType);
    }

    public static double? ParseNullable(string text)
    {
        return double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double value)
            || double.TryParse(text, out value)
            ? value
            : null;
    }

    private static GraphMakerGraphType NormalizeGraphType(GraphMakerRequest request)
    {
        return request.InputMode switch
        {
            GraphMakerInputMode.HeatMap => GraphMakerGraphType.HeatMap,
            GraphMakerInputMode.ProcessTrend => GraphMakerGraphType.ProcessTrend,
            GraphMakerInputMode.NoXMultiY => GraphMakerGraphType.NoXMultiY,
            _ => request.GraphType
        };
    }

    private static GraphMakerChartPayload BuildSeriesPayload(
        GraphMakerTable table,
        GraphMakerRequest request,
        GraphMakerGraphType graphType)
    {
        var labels = BuildLabels(table, request.XColumn);
        var selected = request.YColumns.Count > 0
            ? request.YColumns
            : table.Headers.Where(h => h != request.XColumn).Take(1).ToList();

        var series = new List<GraphMakerSeries>();
        var statGroups = new List<GraphMakerStatGroup>();
        bool useCategoryXAxisForPayload = false;
        if (graphType == GraphMakerGraphType.NormalDistribution)
        {
            foreach (var column in selected)
            {
                var points = BuildDistributionPoints(table.GetNumericValues(column));
                if (points.Count > 0)
                {
                    series.Add(new GraphMakerSeries { Name = column, Points = points });
                }
            }
        }
        else if (graphType == GraphMakerGraphType.Average)
        {
            foreach (var column in selected)
            {
                series.Add(new GraphMakerSeries { Name = $"{column} avg", Data = BuildAverageValues(table, request.XColumn, column) });
            }
        }
        else if (graphType == GraphMakerGraphType.Cpk)
        {
            foreach (var column in selected)
            {
                series.Add(new GraphMakerSeries
                {
                    Name = $"{column} CPK",
                    Data = BuildRollingCpkValues(table, column, request.UpperLimit, request.LowerLimit)
                });
            }
        }
        else if (graphType == GraphMakerGraphType.ProcessTrend)
        {
            string xName = string.IsNullOrWhiteSpace(request.XColumn) ? table.Headers.FirstOrDefault() ?? string.Empty : request.XColumn;
            foreach (var column in selected.Where(c => c != xName))
            {
                var points = BuildScatterPoints(table, xName, column, request);
                series.Add(new GraphMakerSeries { Name = $"{xName} -> {column}", Points = points });
                AddProcessOptionSeries(request, series, xName, column, points);
                AddProcessLimitSeries(table, request, series, column, points);
            }
        }
        else if (graphType == GraphMakerGraphType.NoXMultiY)
        {
            labels = selected;
            foreach (var rowIndex in Enumerable.Range(0, Math.Min(table.Rows.Count, 200)))
            {
                var values = selected.Select(column => table.TryGetDouble(table.Rows[rowIndex], table.IndexOf(column))).ToArray();
                series.Add(new GraphMakerSeries { Name = $"Row {rowIndex + 1}", Data = values });
            }
        }
        else if (graphType == GraphMakerGraphType.Scatter)
        {
            var xColumns = request.XColumns.Count > 0
                ? request.XColumns
                : string.IsNullOrWhiteSpace(request.XColumn) ? [] : [request.XColumn];

            bool useHeaderlessScatter = request.ScatterHeaderNotUse && xColumns.Count > 0;
            bool useCategoryXAxis = useHeaderlessScatter || xColumns.Count != 1 || !IsNumericColumn(table, xColumns[0]);
            useCategoryXAxisForPayload = useCategoryXAxis;
            double categoryOffset = request.ScatterLabelPosition == GraphMakerScatterLabelPosition.BetweenTicks ? 0.5 : 0.0;
            if (xColumns.Count == 0)
            {
                labels = selected;
                foreach (var column in selected)
                {
                    series.Add(new GraphMakerSeries
                    {
                        Name = column,
                        Points = BuildScatterPointsByYCategory(table, selected, column, request.ScatterPointLayout, categoryOffset)
                    });
                }

                statGroups = BuildScatterStatsByYCategory(table, selected, request, categoryOffset);
            }
            else if (useHeaderlessScatter)
            {
                labels = BuildCombinedCategoryLabels(table, xColumns);
                var yColumns = selected.Where(c => !xColumns.Contains(c, StringComparer.Ordinal)).ToList();
                series.Add(new GraphMakerSeries
                {
                    Name = "Value",
                    Points = BuildHeaderlessScatterPointsByCategory(table, xColumns, labels, yColumns, request.ScatterPointLayout, categoryOffset)
                });
                statGroups = BuildHeaderlessScatterStatsByCategory(table, xColumns, labels, yColumns, request, categoryOffset);
            }
            else
            {
                if (useCategoryXAxis)
                {
                    labels = BuildCombinedCategoryLabels(table, xColumns);
                }

                foreach (var column in selected.Where(c => !xColumns.Contains(c, StringComparer.Ordinal)))
                {
                    series.Add(new GraphMakerSeries
                    {
                        Name = column,
                        Points = useCategoryXAxis
                            ? BuildScatterPointsByCategory(table, xColumns, labels, column, request.ScatterPointLayout, categoryOffset)
                            : BuildScatterPoints(table, xColumns[0], column)
                    });
                }

                statGroups = useCategoryXAxis
                    ? BuildScatterStatsByCategory(table, xColumns, labels, selected.Where(c => !xColumns.Contains(c, StringComparer.Ordinal)).ToList(), request, categoryOffset)
                    : [];
            }

            AddScatterLimitSeries(request, series, labels, useCategoryXAxis);
        }
        else if (graphType == GraphMakerGraphType.Line)
        {
            var groupColors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var column in selected)
            {
                string groupName = request.GraphColumnGroups.TryGetValue(column, out string? configuredGroup) && !string.IsNullOrWhiteSpace(configuredGroup)
                    ? configuredGroup.Trim()
                    : BuildGraphGroupName(column);
                if (!groupColors.TryGetValue(groupName, out string? color))
                {
                    color = GraphGroupColor(groupColors.Count);
                    groupColors[groupName] = color;
                }

                series.Add(new GraphMakerSeries
                {
                    Name = column,
                    Data = table.BuildNumericValues(column),
                    Color = color
                });
            }

            AddGraphColumnLimitSeries(table, request, series);
        }
        else
        {
            foreach (var column in selected)
            {
                series.Add(new GraphMakerSeries { Name = column, Data = table.BuildNumericValues(column) });
            }
        }

        return new GraphMakerChartPayload
        {
            Kind = graphType.ToString(),
            Title = graphType == GraphMakerGraphType.ProcessTrend && selected.Count == 1
                ? $"{request.XColumn} -> {selected[0]}"
                : BuildTitle(graphType),
            Labels = labels,
            Series = series,
            UseCategoryXAxis = graphType == GraphMakerGraphType.Scatter && useCategoryXAxisForPayload,
            ScatterLabelPosition = request.ScatterLabelPosition.ToString(),
            LineMode = request.LineMode.ToString(),
            Limits = graphType is GraphMakerGraphType.ProcessTrend or GraphMakerGraphType.Scatter or GraphMakerGraphType.Line ? [] : BuildLimits(request),
            StatGroups = statGroups,
            Stats = graphType == GraphMakerGraphType.ProcessTrend && selected.Count == 1
                ? BuildProcessStats(series.FirstOrDefault(s => !s.IsLimit)?.Points ?? [], request, table, selected[0])
                : graphType == GraphMakerGraphType.Scatter
                    ? (statGroups.Count == 0 ? BuildScatterStats(series.Where(s => !s.IsLimit), request) : [])
                : []
        };
    }

    private static GraphMakerChartPayload BuildHeatMapPayload(GraphMakerTable table, GraphMakerRequest request)
    {
        var xValues = GetDistinctValues(table, request.HeatMapXColumn);
        var yValues = GetDistinctValues(table, request.HeatMapYColumn);
        int xIndex = table.IndexOf(request.HeatMapXColumn);
        int yIndex = table.IndexOf(request.HeatMapYColumn);
        int valueIndex = table.IndexOf(request.HeatMapValueColumn);
        int ngIndex = table.IndexOf(request.HeatMapNgColumn);
        int inputIndex = table.IndexOf(request.HeatMapInputColumn);

        var matrixRows = new List<double?[]>();
        var labelRows = new List<string?[]>();
        foreach (var y in yValues)
        {
            var matrixRow = new List<double?>();
            var labelRow = new List<string?>();
            foreach (var x in xValues)
            {
                var rows = table.Rows.Where(row => table.GetCell(row, xIndex) == x && table.GetCell(row, yIndex) == y).ToList();
                if (request.HeatMapUseNgRate)
                {
                    // For duplicated X/Y heatmap cells, NG Rate is calculated from totals, not row averages.
                    double ng = rows.Select(row => table.TryGetDouble(row, ngIndex)).Where(v => v.HasValue).Sum(v => v!.Value);
                    double input = rows.Select(row => table.TryGetDouble(row, inputIndex)).Where(v => v.HasValue).Sum(v => v!.Value);
                    if (input <= 0)
                    {
                        matrixRow.Add(null);
                        labelRow.Add(null);
                        continue;
                    }

                    double scale = request.HeatMapRateUnit == GraphMakerRateUnit.Ppm ? 1_000_000 : 100;
                    double rate = ng / input * scale;
                    matrixRow.Add(rate);
                    string unit = request.HeatMapRateUnit == GraphMakerRateUnit.Ppm ? " PPM" : "%";
                    labelRow.Add($"NG: {FormatStat(ng)}\nInput: {FormatStat(input)}\nNG Rate: {FormatStat(rate)}{unit}");
                    continue;
                }

                // For duplicated X/Y heatmap cells in value mode, show the cell average.
                var values = rows.Select(row => table.TryGetDouble(row, valueIndex)).Where(v => v.HasValue).Select(v => v!.Value).ToList();
                if (values.Count == 0)
                {
                    matrixRow.Add(null);
                    labelRow.Add(null);
                    continue;
                }

                double avg = values.Average();
                matrixRow.Add(avg);
                labelRow.Add($"Max: {FormatStat(values.Max())}\nMin: {FormatStat(values.Min())}\nAvg: {FormatStat(avg)}\nStd: {FormatStat(SampleStd(values))}");
            }

            matrixRows.Add(matrixRow.ToArray());
            labelRows.Add(labelRow.ToArray());
        }

        var matrix = matrixRows.ToArray();

        var statValues = request.HeatMapUseNgRate
            ? matrix.SelectMany(row => row).Where(v => v.HasValue).Select(v => v!.Value).ToList()
            : table.GetNumericValues(request.HeatMapValueColumn);
        string title = request.HeatMapUseNgRate
            ? $"NG Rate ({(request.HeatMapRateUnit == GraphMakerRateUnit.Ppm ? "PPM" : "%")}) Heatmap"
            : $"{request.HeatMapValueColumn} Heatmap";

        return new GraphMakerChartPayload
        {
            Kind = GraphMakerGraphType.HeatMap.ToString(),
            Title = title,
            XLabels = xValues,
            YLabels = yValues,
            Matrix = matrix,
            MatrixLabels = labelRows.ToArray(),
            Stats = BuildBasicStats(statValues)
        };
    }

    private static string FormatStat(double value) =>
        Math.Abs(value) >= 1000
            ? value.ToString("N1", CultureInfo.InvariantCulture)
            : value.ToString("0.####", CultureInfo.InvariantCulture);

    private static List<GraphMakerStat> BuildBasicStats(IReadOnlyList<double> values)
    {
        if (values.Count == 0)
        {
            return [];
        }

        return
        [
            new("Max", values.Max()),
            new("Min", values.Min()),
            new("Avg", values.Average()),
            new("Std", SampleStd(values))
        ];
    }

    private static List<GraphMakerLimit> BuildLimits(GraphMakerRequest request)
    {
        return
        [
            new("USL", request.UpperLimit, "#dc2626"),
            new("Spec", request.SpecLimit, "#f59e0b"),
            new("LSL", request.LowerLimit, "#16a34a")
        ];
    }

    private static void AddGraphColumnLimitSeries(
        GraphMakerTable table,
        GraphMakerRequest request,
        List<GraphMakerSeries> series)
    {
        AddGraphColumnLimitSeries(table, series, "REF", request.GraphSpecColumn, "#111827");
        AddGraphColumnLimitSeries(table, series, "UP", request.GraphUpperColumn, "#111827");
        AddGraphColumnLimitSeries(table, series, "LOW", request.GraphLowerColumn, "#111827");
    }

    private static void AddGraphColumnLimitSeries(
        GraphMakerTable table,
        List<GraphMakerSeries> series,
        string name,
        string column,
        string color)
    {
        if (string.IsNullOrWhiteSpace(column) || table.IndexOf(column) < 0)
        {
            return;
        }

        series.Add(new GraphMakerSeries
        {
            Name = name,
            Data = table.BuildNumericValues(column),
            Color = color,
            Dashed = true,
            PointRadius = 0
        });
    }

    private static string BuildGraphGroupName(string column)
    {
        string trimmed = column.Trim();
        int marker = trimmed.LastIndexOf('#');
        if (marker > 0 && trimmed[(marker + 1)..].All(char.IsDigit))
        {
            return trimmed[..marker].TrimEnd();
        }

        return trimmed;
    }

    private static string GraphGroupColor(int index)
    {
        string[] colors =
        [
            "#2563eb",
            "#16a34a",
            "#dc2626",
            "#9333ea",
            "#ea580c",
            "#0891b2",
            "#be123c",
            "#4f46e5"
        ];
        return colors[index % colors.Length];
    }

    private static List<string> BuildLabels(GraphMakerTable table, string xColumn)
    {
        if (string.IsNullOrWhiteSpace(xColumn))
        {
            return Enumerable.Range(1, table.Rows.Count).Select(i => i.ToString(CultureInfo.InvariantCulture)).ToList();
        }

        int index = table.IndexOf(xColumn);
        return table.Rows.Select((row, i) =>
        {
            var text = table.GetCell(row, index);
            return string.IsNullOrWhiteSpace(text) ? (i + 1).ToString(CultureInfo.InvariantCulture) : text;
        }).ToList();
    }

    private static double?[] BuildAverageValues(GraphMakerTable table, string xColumn, string yColumn)
    {
        int valueIndex = table.IndexOf(yColumn);
        if (string.IsNullOrWhiteSpace(xColumn))
        {
            return table.BuildNumericValues(yColumn);
        }

        int xIndex = table.IndexOf(xColumn);
        return table.Rows
            .GroupBy(row => table.GetCell(row, xIndex))
            .Select(group =>
            {
                var values = group.Select(row => table.TryGetDouble(row, valueIndex)).Where(v => v.HasValue).Select(v => v!.Value).ToList();
                return values.Count == 0 ? (double?)null : values.Average();
            })
            .ToArray();
    }

    private static double?[] BuildRollingCpkValues(GraphMakerTable table, string column, double? upper, double? lower)
    {
        var raw = table.BuildNumericValues(column).Select(v => v ?? double.NaN).ToList();
        const int window = 10;
        return raw.Select((_, i) =>
        {
            var values = raw.Skip(Math.Max(0, i - window + 1)).Take(Math.Min(window, i + 1)).Where(v => !double.IsNaN(v)).ToList();
            if (values.Count < 2 || (!upper.HasValue && !lower.HasValue))
            {
                return null;
            }

            double avg = values.Average();
            double std = SampleStd(values);
            if (std <= 0)
            {
                return null;
            }

            double cpkUpper = upper.HasValue ? (upper.Value - avg) / (3 * std) : double.PositiveInfinity;
            double cpkLower = lower.HasValue ? (avg - lower.Value) / (3 * std) : double.PositiveInfinity;
            return (double?)Math.Min(cpkUpper, cpkLower);
        }).ToArray();
    }

    private static List<GraphMakerPoint> BuildScatterPoints(GraphMakerTable table, string xColumn, string yColumn)
    {
        return BuildScatterPoints(table, xColumn, yColumn, null);
    }

    private static List<string> BuildCombinedCategoryLabels(GraphMakerTable table, IReadOnlyList<string> xColumns)
    {
        var indexes = xColumns.Select(table.IndexOf).ToArray();
        return table.Rows
            .Select(row => BuildCombinedKey(table, row, indexes))
            .Distinct(StringComparer.Ordinal)
            .ToList();
    }

    private static bool IsNumericColumn(GraphMakerTable table, string column)
    {
        int index = table.IndexOf(column);
        if (index < 0)
        {
            return false;
        }

        var populated = table.Rows
            .Select(row => table.GetCell(row, index))
            .Where(cell => !string.IsNullOrWhiteSpace(cell))
            .Take(50)
            .ToList();
        return populated.Count > 0 && populated.All(cell => GraphMakerPayloadBuilder.ParseNullable(cell).HasValue);
    }

    private static List<GraphMakerStatGroup> BuildScatterStatsByYCategory(
        GraphMakerTable table,
        IReadOnlyList<string> yColumns,
        GraphMakerRequest request,
        double categoryOffset)
    {
        var groups = new List<GraphMakerStatGroup>();
        for (int i = 0; i < yColumns.Count; i++)
        {
            var values = table.GetNumericValues(yColumns[i]);
            var stats = BuildScatterStatList(values, request, yColumns[i]);
            if (stats.Count > 0)
            {
                groups.Add(new GraphMakerStatGroup(i + categoryOffset, yColumns[i], stats));
            }
        }

        return groups;
    }

    private static List<GraphMakerStatGroup> BuildScatterStatsByCategory(
        GraphMakerTable table,
        IReadOnlyList<string> xColumns,
        IReadOnlyList<string> labels,
        IReadOnlyList<string> yColumns,
        GraphMakerRequest request,
        double categoryOffset)
    {
        var xIndexes = xColumns.Select(table.IndexOf).ToArray();
        var groups = new List<GraphMakerStatGroup>();
        for (int i = 0; i < labels.Count; i++)
        {
            string label = labels[i];
            var rows = table.Rows.Where(row => string.Equals(BuildCombinedKey(table, row, xIndexes), label, StringComparison.Ordinal)).ToList();
            var stats = new List<GraphMakerStat>();
            foreach (var yColumn in yColumns)
            {
                int yIndex = table.IndexOf(yColumn);
                var values = rows
                    .Select(row => table.TryGetDouble(row, yIndex))
                    .Where(v => v.HasValue)
                    .Select(v => v!.Value)
                    .ToList();
                stats.AddRange(BuildScatterStatList(values, request, yColumn));
            }

            if (stats.Count > 0)
            {
                groups.Add(new GraphMakerStatGroup(i + categoryOffset, label, stats));
            }
        }

        return groups;
    }

    private static List<GraphMakerPoint> BuildHeaderlessScatterPointsByCategory(
        GraphMakerTable table,
        IReadOnlyList<string> xColumns,
        IReadOnlyList<string> labels,
        IReadOnlyList<string> yColumns,
        GraphMakerScatterPointLayout pointLayout,
        double categoryOffset)
    {
        var xIndexes = xColumns.Select(table.IndexOf).ToArray();
        var yIndexes = yColumns.Select(table.IndexOf).Where(index => index >= 0).ToArray();
        var labelMap = labels.Select((label, index) => new { label, index })
            .ToDictionary(x => x.label, x => x.index, StringComparer.Ordinal);
        var points = new List<GraphMakerPoint>();

        foreach (var (row, rowIndex) in table.Rows.Select((row, index) => (row, index)))
        {
            string key = BuildCombinedKey(table, row, xIndexes);
            if (!labelMap.TryGetValue(key, out int x))
            {
                continue;
            }

            foreach (int yIndex in yIndexes)
            {
                var y = table.TryGetDouble(row, yIndex);
                if (!y.HasValue)
                {
                    continue;
                }

                double plottedX = pointLayout == GraphMakerScatterPointLayout.Spread
                    ? x + categoryOffset + DeterministicJitter(rowIndex + yIndex * 17)
                    : x + categoryOffset;
                points.Add(new GraphMakerPoint(plottedX, y.Value));
            }
        }

        return points;
    }

    private static List<GraphMakerStatGroup> BuildHeaderlessScatterStatsByCategory(
        GraphMakerTable table,
        IReadOnlyList<string> xColumns,
        IReadOnlyList<string> labels,
        IReadOnlyList<string> yColumns,
        GraphMakerRequest request,
        double categoryOffset)
    {
        var xIndexes = xColumns.Select(table.IndexOf).ToArray();
        var yIndexes = yColumns.Select(table.IndexOf).Where(index => index >= 0).ToArray();
        var groups = new List<GraphMakerStatGroup>();
        var allValues = table.Rows
            .SelectMany(row => yIndexes.Select(index => table.TryGetDouble(row, index)))
            .Where(v => v.HasValue)
            .Select(v => v!.Value)
            .ToList();
        var allStats = BuildScatterStatList(allValues, request, string.Empty);
        if (allStats.Count > 0)
        {
            groups.Add(new GraphMakerStatGroup(-1, "All", allStats));
        }

        for (int i = 0; i < labels.Count; i++)
        {
            string label = labels[i];
            var values = table.Rows
                .Where(row => string.Equals(BuildCombinedKey(table, row, xIndexes), label, StringComparison.Ordinal))
                .SelectMany(row => yIndexes.Select(index => table.TryGetDouble(row, index)))
                .Where(v => v.HasValue)
                .Select(v => v!.Value)
                .ToList();
            var stats = BuildScatterStatList(values, request, string.Empty);
            if (stats.Count > 0)
            {
                groups.Add(new GraphMakerStatGroup(i + categoryOffset, label, stats));
            }
        }

        return groups;
    }

    private static List<GraphMakerPoint> BuildScatterPointsByCategory(
        GraphMakerTable table,
        IReadOnlyList<string> xColumns,
        IReadOnlyList<string> labels,
        string yColumn,
        GraphMakerScatterPointLayout pointLayout,
        double categoryOffset)
    {
        var xIndexes = xColumns.Select(table.IndexOf).ToArray();
        int yIndex = table.IndexOf(yColumn);
        var labelMap = labels.Select((label, index) => new { label, index })
            .ToDictionary(x => x.label, x => x.index, StringComparer.Ordinal);

        return table.Rows.Select((row, rowIndex) =>
        {
            string key = BuildCombinedKey(table, row, xIndexes);
            if (!labelMap.TryGetValue(key, out int x))
            {
                return null;
            }

            var y = table.TryGetDouble(row, yIndex);
            double plottedX = pointLayout == GraphMakerScatterPointLayout.Spread
                ? x + categoryOffset + DeterministicJitter(rowIndex)
                : x + categoryOffset;
            return y.HasValue ? new GraphMakerPoint(plottedX, y.Value) : null;
        }).Where(p => p is not null).Cast<GraphMakerPoint>().ToList();
    }

    private static List<GraphMakerPoint> BuildScatterPointsByYCategory(
        GraphMakerTable table,
        IReadOnlyList<string> selectedYColumns,
        string yColumn,
        GraphMakerScatterPointLayout pointLayout,
        double categoryOffset)
    {
        int yCategoryIndex = selectedYColumns
            .Select((column, index) => new { column, index })
            .FirstOrDefault(x => string.Equals(x.column, yColumn, StringComparison.Ordinal))?.index ?? 0;
        int yIndex = table.IndexOf(yColumn);
        return table.Rows.Select((row, rowIndex) =>
        {
            var y = table.TryGetDouble(row, yIndex);
            if (!y.HasValue)
            {
                return null;
            }

            double plottedX = pointLayout == GraphMakerScatterPointLayout.Spread
                ? yCategoryIndex + categoryOffset + DeterministicJitter(rowIndex)
                : yCategoryIndex + categoryOffset;
            return new GraphMakerPoint(plottedX, y.Value);
        }).Where(p => p is not null).Cast<GraphMakerPoint>().ToList();
    }

    private static void AddScatterLimitSeries(
        GraphMakerRequest request,
        List<GraphMakerSeries> series,
        IReadOnlyList<string> labels,
        bool useCategoryXAxis)
    {
        var points = series.Where(s => !s.IsLimit).SelectMany(s => s.Points ?? []).ToList();
        if (points.Count == 0)
        {
            return;
        }

        double minX = useCategoryXAxis ? -0.5 : points.Min(p => p.X);
        double maxX = useCategoryXAxis ? Math.Max(0.5, labels.Count - 0.5) : points.Max(p => p.X);
        if (Math.Abs(maxX - minX) < double.Epsilon)
        {
            minX -= 0.5;
            maxX += 0.5;
        }

        AddScatterLimitSeries(series, "USL", request.UpperLimit, minX, maxX);
        AddScatterLimitSeries(series, "SPEC", request.SpecLimit, minX, maxX);
        AddScatterLimitSeries(series, "LSL", request.LowerLimit, minX, maxX);
    }

    private static void AddScatterLimitSeries(
        List<GraphMakerSeries> series,
        string name,
        double? value,
        double minX,
        double maxX)
    {
        if (!value.HasValue)
        {
            return;
        }

        series.Add(new GraphMakerSeries
        {
            Name = name,
            Points = [new GraphMakerPoint(minX, value.Value), new GraphMakerPoint(maxX, value.Value)],
            IsLimit = true,
            Color = "#111827",
            ShowLine = true,
            Dashed = true,
            LabelValue = value.Value,
            PointRadius = 0
        });
    }

    private static double DeterministicJitter(int rowIndex)
    {
        int slot = rowIndex % 9;
        return (slot - 4) * 0.035;
    }

    private static string BuildCombinedKey(GraphMakerTable table, string[] row, IReadOnlyList<int> indexes)
    {
        return string.Join(" + ", indexes.Select(index =>
        {
            var value = table.GetCell(row, index);
            return string.IsNullOrWhiteSpace(value) ? "(blank)" : value;
        }));
    }

    private static List<GraphMakerPoint> BuildScatterPoints(
        GraphMakerTable table,
        string xColumn,
        string yColumn,
        GraphMakerRequest? request)
    {
        int xIndex = table.IndexOf(xColumn);
        int yIndex = table.IndexOf(yColumn);
        var excluded = request is null ? new HashSet<int>() : BuildExcludedDataRowIndexes(request);
        return table.Rows.Select((row, i) =>
        {
            if (excluded.Contains(i))
            {
                return null;
            }

            var x = table.TryGetDouble(row, xIndex) ?? i + 1;
            var y = table.TryGetDouble(row, yIndex);
            return y.HasValue ? new GraphMakerPoint(x, y.Value) : null;
        }).Where(p => p is not null).Cast<GraphMakerPoint>().ToList();
    }

    private static void AddProcessLimitSeries(
        GraphMakerTable table,
        GraphMakerRequest request,
        List<GraphMakerSeries> series,
        string yColumn,
        IReadOnlyList<GraphMakerPoint> points)
    {
        if (points.Count == 0)
        {
            return;
        }

        double minX = points.Min(p => p.X);
        double maxX = points.Max(p => p.X);
        int yIndex = table.IndexOf(yColumn);

        AddProcessLimitSeries(table, request.ProcessUpperRowNumber, request, yIndex, minX, maxX, "USL", series);
        AddProcessLimitSeries(table, request.ProcessSpecRowNumber, request, yIndex, minX, maxX, "SPEC", series);
        AddProcessLimitSeries(table, request.ProcessLowerRowNumber, request, yIndex, minX, maxX, "LSL", series);
    }

    private static void AddProcessOptionSeries(
        GraphMakerRequest request,
        List<GraphMakerSeries> series,
        string xName,
        string yName,
        IReadOnlyList<GraphMakerPoint> points)
    {
        if (points.Count < 2)
        {
            return;
        }

        if (request.ShowProcessContour)
        {
            var hull = BuildConvexHull(points);
            if (hull.Count >= 3)
            {
                hull.Add(hull[0]);
                series.Add(new GraphMakerSeries
                {
                    Name = "Contour",
                    Points = hull,
                    Color = "#475569",
                    ShowLine = true,
                    Dashed = false,
                    PointRadius = 0
                });
            }
        }

        if (request.ShowProcessTrendLine && TryCalculateTrendLine(points, out double slope, out double intercept))
        {
            double minX = points.Min(p => p.X);
            double maxX = points.Max(p => p.X);
            series.Add(new GraphMakerSeries
            {
                Name = "Trend",
                Points =
                [
                    new GraphMakerPoint(minX, slope * minX + intercept),
                    new GraphMakerPoint(maxX, slope * maxX + intercept)
                ],
                Color = "#0f172a",
                ShowLine = true,
                Dashed = true,
                PointRadius = 0
            });
        }
    }

    private static List<GraphMakerStat> BuildProcessStats(
        IReadOnlyList<GraphMakerPoint> points,
        GraphMakerRequest request,
        GraphMakerTable table,
        string yColumn)
    {
        if (points.Count == 0)
        {
            return [];
        }

        var values = points.Select(p => p.Y).ToList();
        double avg = values.Average();
        double std = SampleStd(values);
        var stats = new List<GraphMakerStat>();
        if (request.ShowProcessMax) stats.Add(new GraphMakerStat("Max", values.Max()));
        if (request.ShowProcessMin) stats.Add(new GraphMakerStat("Min", values.Min()));
        if (request.ShowProcessAvg) stats.Add(new GraphMakerStat("Avg", avg));
        if (request.ShowProcessStd) stats.Add(new GraphMakerStat("Std", std));

        if (std > 0)
        {
            int yIndex = table.IndexOf(yColumn);
            double? upper = TryGetProcessLimitValue(table, request.ProcessUpperRowNumber, yIndex);
            double? lower = TryGetProcessLimitValue(table, request.ProcessLowerRowNumber, yIndex);
            double? cpu = upper.HasValue ? (upper.Value - avg) / (3 * std) : null;
            double? cpl = lower.HasValue ? (avg - lower.Value) / (3 * std) : null;
            double? cpk = cpu.HasValue && cpl.HasValue ? Math.Min(cpu.Value, cpl.Value) : cpu ?? cpl;

            if (request.ShowProcessCpu && cpu.HasValue) stats.Add(new GraphMakerStat("CPU", cpu.Value));
            if (request.ShowProcessCpl && cpl.HasValue) stats.Add(new GraphMakerStat("CPL", cpl.Value));
            if (request.ShowProcessCpk && cpk.HasValue) stats.Add(new GraphMakerStat("CPK", cpk.Value));
        }

        return stats;
    }

    private static double? TryGetProcessLimitValue(GraphMakerTable table, int? rowNumber, int yIndex)
    {
        int rowIndex = ToDataRowIndex(rowNumber);
        return rowIndex >= 0 && rowIndex < table.Rows.Count
            ? table.TryGetDouble(table.Rows[rowIndex], yIndex)
            : null;
    }

    private static List<GraphMakerStat> BuildScatterStats(IEnumerable<GraphMakerSeries> series, GraphMakerRequest request)
    {
        var stats = new List<GraphMakerStat>();
        foreach (var item in series)
        {
            var values = item.Points?.Select(p => p.Y).ToList() ?? [];
            if (values.Count == 0)
            {
                continue;
            }

            double avg = values.Average();
            double std = SampleStd(values);
            stats.AddRange(BuildScatterStatList(values, request, item.Name));
        }

        return stats;
    }

    private static List<GraphMakerStat> BuildScatterStatList(
        IReadOnlyList<double> values,
        GraphMakerRequest request,
        string prefix)
    {
        if (values.Count == 0)
        {
            return [];
        }

        double avg = values.Average();
        double std = SampleStd(values);
        var stats = new List<GraphMakerStat>();
        string Name(string metric) => string.IsNullOrWhiteSpace(prefix) ? metric : $"{prefix} {metric}";
        if (request.ShowScatterMax) stats.Add(new GraphMakerStat(Name("Max"), values.Max()));
        if (request.ShowScatterMin) stats.Add(new GraphMakerStat(Name("Min"), values.Min()));
        if (request.ShowScatterAvg) stats.Add(new GraphMakerStat(Name("Avg"), avg));
        if (request.ShowScatterStd) stats.Add(new GraphMakerStat(Name("Std"), std));

        if (std > 0)
        {
            double? cpu = request.UpperLimit.HasValue ? (request.UpperLimit.Value - avg) / (3 * std) : null;
            double? cpl = request.LowerLimit.HasValue ? (avg - request.LowerLimit.Value) / (3 * std) : null;
            double? cpk = cpu.HasValue && cpl.HasValue
                ? Math.Min(cpu.Value, cpl.Value)
                : cpu ?? cpl;

            if (request.ShowScatterCpu && cpu.HasValue) stats.Add(new GraphMakerStat(Name("CPU"), cpu.Value));
            if (request.ShowScatterCpl && cpl.HasValue) stats.Add(new GraphMakerStat(Name("CPL"), cpl.Value));
            if (request.ShowScatterCpk && cpk.HasValue) stats.Add(new GraphMakerStat(Name("CPK"), cpk.Value));
        }

        return stats;
    }

    private static bool TryCalculateTrendLine(IReadOnlyList<GraphMakerPoint> points, out double slope, out double intercept)
    {
        slope = 0;
        intercept = 0;
        if (points.Count < 2)
        {
            return false;
        }

        double avgX = points.Average(p => p.X);
        double avgY = points.Average(p => p.Y);
        double denominator = points.Sum(p => Math.Pow(p.X - avgX, 2));
        if (Math.Abs(denominator) <= 1e-12)
        {
            return false;
        }

        slope = points.Sum(p => (p.X - avgX) * (p.Y - avgY)) / denominator;
        intercept = avgY - slope * avgX;
        return true;
    }

    private static List<GraphMakerPoint> BuildConvexHull(IReadOnlyList<GraphMakerPoint> points)
    {
        var sorted = points
            .Distinct()
            .OrderBy(p => p.X)
            .ThenBy(p => p.Y)
            .ToList();

        if (sorted.Count <= 2)
        {
            return sorted;
        }

        var lower = new List<GraphMakerPoint>();
        foreach (var point in sorted)
        {
            while (lower.Count >= 2 && Cross(lower[^2], lower[^1], point) <= 0)
            {
                lower.RemoveAt(lower.Count - 1);
            }
            lower.Add(point);
        }

        var upper = new List<GraphMakerPoint>();
        for (int i = sorted.Count - 1; i >= 0; i--)
        {
            var point = sorted[i];
            while (upper.Count >= 2 && Cross(upper[^2], upper[^1], point) <= 0)
            {
                upper.RemoveAt(upper.Count - 1);
            }
            upper.Add(point);
        }

        lower.RemoveAt(lower.Count - 1);
        upper.RemoveAt(upper.Count - 1);
        lower.AddRange(upper);
        return lower;
    }

    private static double Cross(GraphMakerPoint origin, GraphMakerPoint a, GraphMakerPoint b) =>
        (a.X - origin.X) * (b.Y - origin.Y) - (a.Y - origin.Y) * (b.X - origin.X);

    private static void AddProcessLimitSeries(
        GraphMakerTable table,
        int? sheetRowNumber,
        GraphMakerRequest request,
        int yIndex,
        double minX,
        double maxX,
        string name,
        List<GraphMakerSeries> series)
    {
        int dataIndex = ToDataRowIndex(sheetRowNumber);
        if (dataIndex < 0 || dataIndex >= table.Rows.Count)
        {
            return;
        }

        double? value = table.TryGetDouble(table.Rows[dataIndex], yIndex);
        if (!value.HasValue)
        {
            return;
        }

        series.Add(new GraphMakerSeries
        {
            Name = name,
            Points = [new GraphMakerPoint(minX, value.Value), new GraphMakerPoint(maxX, value.Value)],
            IsLimit = true,
            Color = "#111827",
            ShowLine = true,
            Dashed = true,
            LabelValue = value.Value
        });
    }

    private static HashSet<int> BuildExcludedDataRowIndexes(GraphMakerRequest request)
    {
        return new[]
            {
                request.ProcessSpecRowNumber,
                request.ProcessUpperRowNumber,
                request.ProcessLowerRowNumber
            }
            .Select(ToDataRowIndex)
            .Where(index => index >= 0)
            .ToHashSet();
    }

    private static int ToDataRowIndex(int? sheetRowNumber)
    {
        return sheetRowNumber.HasValue ? sheetRowNumber.Value - 1 : -1;
    }

    private static List<GraphMakerPoint> BuildDistributionPoints(List<double> values)
    {
        if (values.Count < 2)
        {
            return [];
        }

        double avg = values.Average();
        double std = SampleStd(values);
        if (std <= 0)
        {
            return [];
        }

        double min = values.Min() - std;
        double max = values.Max() + std;
        return Enumerable.Range(0, 80).Select(i =>
        {
            double x = min + (max - min) * i / 79.0;
            double y = 1.0 / (std * Math.Sqrt(2 * Math.PI)) * Math.Exp(-0.5 * Math.Pow((x - avg) / std, 2));
            return new GraphMakerPoint(x, y);
        }).ToList();
    }

    private static List<string> GetDistinctValues(GraphMakerTable table, string column)
    {
        int index = table.IndexOf(column);
        return table.Rows.Select(row => table.GetCell(row, index))
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.Ordinal)
            .Take(80)
            .ToList();
    }

    private static string BuildTitle(GraphMakerGraphType graphType) => graphType switch
    {
        GraphMakerGraphType.NormalDistribution => "Normal Distribution",
        GraphMakerGraphType.ProcessTrend => "Process Trend",
        GraphMakerGraphType.NoXMultiY => "No X Multi Y",
        GraphMakerGraphType.Cpk => "CPK",
        GraphMakerGraphType.Line => "Graph",
        _ => $"{graphType} Graph"
    };

    private static double SampleStd(IReadOnlyList<double> values)
    {
        if (values.Count < 2)
        {
            return 0;
        }

        double avg = values.Average();
        double variance = values.Sum(v => Math.Pow(v - avg, 2)) / (values.Count - 1);
        return Math.Sqrt(variance);
    }
}
