using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

public sealed class ProcessMaterialNgService(
    NgRateSettingsService settings,
    ProcessMaterialMappingService mappingService,
    WebRepository repo)
{
    private readonly NgRateSettingsService _settings = settings;
    private readonly ProcessMaterialMappingService _mappingService = mappingService;
    private readonly WebRepository _repo = repo;

    public string? FindMostRecentDb()
    {
        string dir = _settings.DbSaveDirectory;
        if (!Directory.Exists(dir)) return null;

        return Directory.GetFiles(dir, "*.db")
            .Where(f => !Path.GetFileName(f).Equals("ngrate_settings.db", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(File.GetLastWriteTime)
            .FirstOrDefault();
    }

    public Task<ProcessMaterialNgReport> GenerateAsync(
        string dbPath,
        string modelQuery,
        IProgress<string>? progress = null)
        => Task.Run(() => Generate(dbPath, modelQuery, null, null, progress));

    public Task<ProcessMaterialNgReport> GenerateAsync(
        string dbPath,
        string modelQuery,
        DateTime startDate,
        DateTime endDate,
        IProgress<string>? progress = null)
        => Task.Run(() => Generate(dbPath, modelQuery, startDate.Date, endDate.Date, progress));

    private ProcessMaterialNgReport Generate(
        string dbPath,
        string modelQuery,
        DateTime? startDate,
        DateTime? endDate,
        IProgress<string>? progress)
    {
        if (startDate.HasValue && endDate.HasValue && endDate.Value < startDate.Value)
            throw new InvalidOperationException("End date must be on or after start date.");
        if (string.IsNullOrWhiteSpace(dbPath) || !File.Exists(dbPath))
            throw new FileNotFoundException("NG Rate DB was not found.", dbPath);

        string query = (modelQuery ?? string.Empty).Trim();
        var mappings = _mappingService.GetAll()
            .Where(IsSelectableMapping)
            .ToList();
        var processSettings = _mappingService.GetProcessRows();
        var routingRows = _settings.GetRoutingRows();
        var modelFilter = BuildModelFilter(query, mappings, processSettings);

        progress?.Report("Loading NG rows...");
        var rawRows = LoadNgRows(dbPath, startDate, endDate)
            .Where(r => MatchesModelFilter(r, modelFilter))
            .ToList();
        DateTime reportStart = startDate ?? rawRows.Select(r => r.ProductDate).DefaultIfEmpty(DateTime.Today).Min();
        DateTime reportEnd = endDate ?? rawRows.Select(r => r.ProductDate).DefaultIfEmpty(DateTime.Today).Max();

        progress?.Report($"NG rows: {rawRows.Count:N0}");
        var processCatalog = BuildProcessCatalog(rawRows, routingRows, mappings, processSettings, query);
        var effectiveCache = new Dictionary<string, List<ProcessMaterialMappingRow>>(StringComparer.OrdinalIgnoreCase);
        var processRows = new Dictionary<string, ProcessMaterialNgProcessSummaryRow>(StringComparer.OrdinalIgnoreCase);
        var processMaterialRows = new Dictionary<string, ProcessMaterialNgProcessMaterialRow>(StringComparer.OrdinalIgnoreCase);

        var processWeekRows = rawRows
            .Select(r => r with { WeekKey = BuildWeekKey(r.ProductDate) })
            .Where(r => !string.IsNullOrWhiteSpace(r.WeekKey))
            .GroupBy(r => (
                Model: NormalizeKey(r.MaterialName),
                r.MaterialName,
                Code: NormalizeKey(r.ProcessCode),
                r.ProcessCode,
                Name: NormalizeKey(r.ProcessName),
                r.ProcessName,
                Type: NormalizeKey(r.ProcessType),
                r.ProcessType,
                r.WeekKey))
            .Select(g => new ProcessWeekNg(
                g.Key.MaterialName.Trim(),
                g.Key.ProcessCode.Trim(),
                g.Key.ProcessName.Trim(),
                g.Key.ProcessType.Trim(),
                g.Key.WeekKey,
                g.Sum(r => r.QtyInput),
                g.Sum(r => r.QtyNg)))
            .Where(r => r.NgQty > 0)
            .ToList();

        foreach (var row in processWeekRows)
        {
            var process = new ProcessMaterialNgProcess(
                row.ModelName,
                row.ProcessCode,
                row.ProcessName,
                ResolveProcessType(row, routingRows));

            string processKey = BuildExactProcessKey(process);
            if (!processRows.TryGetValue(processKey, out var processSummary))
            {
                var effectiveMappings = GetEffectiveMappingsForProcess(
                    process,
                    processCatalog,
                    mappings,
                    processSettings,
                    effectiveCache);

                var materialRefs = effectiveMappings
                    .GroupBy(m => BuildMaterialKey(m.RawMaterialCode, m.RawMaterialName), StringComparer.OrdinalIgnoreCase)
                    .Select(g => g.First())
                    .OrderBy(m => m.RawMaterialName, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(m => m.RawMaterialCode, StringComparer.OrdinalIgnoreCase)
                    .Select(m => new ProcessMaterialNgMaterialRef(
                        m.RawMaterialCode,
                        m.RawMaterialName,
                        m.UsageQty,
                        m.UsageUnit))
                    .ToList();

                processSummary = new ProcessMaterialNgProcessSummaryRow
                {
                    ModelName = process.ModelName,
                    ProcessNo = GetProcessNo(process, processSettings),
                    ProcessCode = process.ProcessCode,
                    ProcessName = process.ProcessName,
                    ProcessType = process.ProcessType,
                    Materials = materialRefs,
                };
                processRows[processKey] = processSummary;
            }

            AddQty(processSummary.NgByWeek, row.WeekKey, row.NgQty);
            AddQty(processSummary.InputByWeek, row.WeekKey, row.InputQty);
            processSummary.TotalInputQty += row.InputQty;
            processSummary.TotalNgQty += row.NgQty;

            foreach (var material in processSummary.Materials)
            {
                string detailKey = processKey + "\t" + BuildMaterialKey(material.RawMaterialCode, material.RawMaterialName);
                if (!processMaterialRows.TryGetValue(detailKey, out var detail))
                {
                    detail = new ProcessMaterialNgProcessMaterialRow
                    {
                        ModelName = process.ModelName,
                        ProcessNo = processSummary.ProcessNo,
                        ProcessCode = process.ProcessCode,
                        ProcessName = process.ProcessName,
                        ProcessType = process.ProcessType,
                        RawMaterialCode = material.RawMaterialCode,
                        RawMaterialName = material.RawMaterialName,
                        UsageQty = material.UsageQty,
                        UsageUnit = material.UsageUnit,
                    };
                    processMaterialRows[detailKey] = detail;
                }

                AddQty(detail.NgByWeek, row.WeekKey, row.NgQty);
                AddQty(detail.InputByWeek, row.WeekKey, row.InputQty);
                detail.TotalInputQty += row.InputQty;
                detail.TotalNgQty += row.NgQty;
            }
        }

        var weekColumns = BuildWeekColumns(startDate, endDate, processWeekRows.Select(r => r.WeekKey));
        var processList = processRows.Values
            .OrderBy(r => string.IsNullOrWhiteSpace(r.ProcessNo) ? 1 : 0)
            .ThenBy(r => BuildProcessNoSortKey(r.ProcessNo), StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.ModelName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => NormalizeKey(r.ProcessCode), StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => NormalizeKey(r.ProcessName), StringComparer.OrdinalIgnoreCase)
            .ToList();
        var detailList = processMaterialRows.Values
            .OrderBy(r => string.IsNullOrWhiteSpace(r.ProcessNo) ? 1 : 0)
            .ThenBy(r => BuildProcessNoSortKey(r.ProcessNo), StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.ModelName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => NormalizeKey(r.ProcessCode), StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => NormalizeKey(r.ProcessName), StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.RawMaterialName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.RawMaterialCode, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var materialSummary = BuildMaterialSummary(detailList);

        return new ProcessMaterialNgReport
        {
            DbPath = dbPath,
            ModelFilter = query,
            StartDate = reportStart,
            EndDate = reportEnd,
            GeneratedAt = DateTime.Now,
            SourceRowCount = rawRows.Count,
            LineShiftCount = rawRows.Select(r => r.LineShift).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
            WeekColumns = weekColumns,
            ProcessRows = processList,
            ProcessMaterialRows = detailList,
            MaterialSummaryRows = materialSummary,
            TotalProcessNgQty = processList.Sum(r => r.TotalNgQty),
            TotalProcessInputQty = processList.Sum(r => r.TotalInputQty),
            TotalMaterialNgQty = materialSummary.Sum(r => r.TotalNgQty),
        };
    }

    private ModelFilter BuildModelFilter(
        string query,
        IReadOnlyList<ProcessMaterialMappingRow> mappings,
        IReadOnlyList<ProcessMaterialProcessRow> processSettings)
    {
        var filter = new ModelFilter(query);
        if (string.IsNullOrWhiteSpace(query))
            return filter;

        foreach (var group in _repo.GetModelGroups())
        {
            bool groupMatch = Contains(group.Name, query);
            foreach (var mid in group.MidGroups)
            {
                bool midMatch = groupMatch ||
                                Contains(mid.Material, query) ||
                                mid.LineShifts.Any(ls => Contains(ls, query));
                if (!midMatch) continue;

                AddIfNotBlank(filter.ModelNames, mid.Material);
                foreach (string lineShift in mid.LineShifts)
                    AddIfNotBlank(filter.LineShifts, lineShift);
            }
        }

        foreach (var modelName in mappings.Select(r => r.ModelName).Concat(processSettings.Select(r => r.ModelName)))
        {
            if (Contains(modelName, query))
                AddIfNotBlank(filter.ModelNames, modelName);
        }

        return filter;
    }

    private static List<RawNgRow> LoadNgRows(string dbPath, DateTime? startDate, DateTime? endDate)
    {
        var rows = new List<RawNgRow>();
        var csb = new SqliteConnectionStringBuilder
        {
            DataSource = dbPath,
            Mode = SqliteOpenMode.ReadOnly,
        };
        using var conn = new SqliteConnection(csb.ToString());
        conn.Open();
        if (!TableExists(conn, "OrginalTable")) return rows;

        if (!ColumnExists(conn, "OrginalTable", "MATERIALNAME") ||
            !ColumnExists(conn, "OrginalTable", "PROCESSCODE") ||
            !ColumnExists(conn, "OrginalTable", "PROCESSNAME") ||
            !ColumnExists(conn, "OrginalTable", "QTYINPUT") ||
            !ColumnExists(conn, "OrginalTable", "QTYNG") ||
            !ColumnExists(conn, "OrginalTable", "PRODUCT_DATE"))
        {
            return rows;
        }

        bool hasLineShift = ColumnExists(conn, "OrginalTable", "LineShift");
        string lineShiftExpr = hasLineShift
            ? "[LineShift]"
            : "([MATERIALNAME] || '_' || [PRODUCTION_LINE])";

        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"""
            SELECT [MATERIALNAME], {lineShiftExpr}, [PROCESSCODE], [PROCESSNAME],
                   [QTYINPUT], [QTYNG], [PRODUCT_DATE]
            FROM [OrginalTable]
            """;

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            string productDateText = ReadStr(reader, 6);
            if (!DateTime.TryParse(productDateText, out var productDate))
                continue;

            productDate = productDate.Date;
            if (startDate.HasValue && productDate < startDate.Value.Date)
                continue;
            if (endDate.HasValue && productDate > endDate.Value.Date)
                continue;

            rows.Add(new RawNgRow(
                ReadStr(reader, 0),
                ReadStr(reader, 1),
                ReadStr(reader, 2),
                NormalizeText(ReadStr(reader, 3)),
                string.Empty,
                productDate,
                ReadDouble(reader, 4),
                ReadDouble(reader, 5),
                string.Empty));
        }

        return rows;
    }

    private static List<ProcessMaterialNgMaterialSummaryRow> BuildMaterialSummary(
        IReadOnlyList<ProcessMaterialNgProcessMaterialRow> detailRows)
        => detailRows
            .GroupBy(r => BuildMaterialKey(r.RawMaterialCode, r.RawMaterialName), StringComparer.OrdinalIgnoreCase)
            .Select(g =>
            {
                var first = g.First();
                var summary = new ProcessMaterialNgMaterialSummaryRow
                {
                    RawMaterialCode = first.RawMaterialCode,
                    RawMaterialName = first.RawMaterialName,
                    ProcessCount = g.Select(ProcessIdentityOf).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                };

                foreach (var row in g)
                {
                    foreach (var kv in row.NgByWeek)
                        AddQty(summary.NgByWeek, kv.Key, kv.Value);
                    summary.TotalNgQty += row.TotalNgQty;
                }

                return summary;
            })
            .OrderByDescending(r => r.TotalNgQty)
            .ThenBy(r => r.RawMaterialName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.RawMaterialCode, StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static List<ProcessMaterialNgProcess> BuildProcessCatalog(
        IReadOnlyList<RawNgRow> rawRows,
        IReadOnlyList<RoutingRow> routingRows,
        IReadOnlyList<ProcessMaterialMappingRow> mappings,
        IReadOnlyList<ProcessMaterialProcessRow> processSettings,
        string query)
    {
        var modelFamilies = rawRows
            .Select(r => NormalizeSideAgnosticModelName(r.MaterialName))
            .Where(s => s.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var catalog = new Dictionary<string, ProcessMaterialNgProcess>(StringComparer.OrdinalIgnoreCase);

        void Add(ProcessMaterialNgProcess process)
        {
            if (string.IsNullOrWhiteSpace(process.ModelName) ||
                (string.IsNullOrWhiteSpace(process.ProcessCode) && string.IsNullOrWhiteSpace(process.ProcessName)))
            {
                return;
            }

            catalog.TryAdd(BuildExactProcessKey(process), process);
        }

        foreach (var row in rawRows)
            Add(new ProcessMaterialNgProcess(row.MaterialName.Trim(), row.ProcessCode.Trim(), row.ProcessName.Trim(), row.ProcessType.Trim()));

        foreach (var row in routingRows)
        {
            string family = NormalizeSideAgnosticModelName(row.ModelName);
            if (modelFamilies.Contains(family) || Contains(row.ModelName, query))
                Add(new ProcessMaterialNgProcess(row.ModelName.Trim(), row.ProcessCode.Trim(), row.ProcessName.Trim(), row.ProcessType.Trim()));
        }

        foreach (var row in mappings)
        {
            string family = NormalizeSideAgnosticModelName(row.ModelName);
            if (modelFamilies.Contains(family) || Contains(row.ModelName, query))
                Add(new ProcessMaterialNgProcess(row.ModelName.Trim(), row.ProcessCode.Trim(), row.ProcessName.Trim(), string.Empty));
        }

        foreach (var row in processSettings)
        {
            string family = NormalizeSideAgnosticModelName(row.ModelName);
            if (modelFamilies.Contains(family) || Contains(row.ModelName, query))
                Add(new ProcessMaterialNgProcess(row.ModelName.Trim(), row.ProcessCode.Trim(), row.ProcessName.Trim(), string.Empty));
        }

        return catalog.Values.ToList();
    }

    private static List<ProcessMaterialMappingRow> GetEffectiveMappingsForProcess(
        ProcessMaterialNgProcess process,
        IReadOnlyList<ProcessMaterialNgProcess> processCatalog,
        IReadOnlyList<ProcessMaterialMappingRow> mappings,
        IReadOnlyList<ProcessMaterialProcessRow> processSettings,
        Dictionary<string, List<ProcessMaterialMappingRow>> cache,
        int depth = 0)
    {
        string cacheKey = BuildExactProcessKey(process);
        if (cache.TryGetValue(cacheKey, out var cached))
            return cached;

        var direct = GetMappingsForProcess(process, mappings);
        if (depth >= 8)
        {
            cache[cacheKey] = direct;
            return direct;
        }

        string referenceNo = GetReferenceProcessNo(process, processSettings);
        if (string.IsNullOrWhiteSpace(referenceNo))
        {
            cache[cacheKey] = direct;
            return direct;
        }

        var referenceProcess = FindProcessByNo(process.ModelName, referenceNo, processCatalog, processSettings);
        if (referenceProcess is null || IsSameProcess(referenceProcess, process))
        {
            cache[cacheKey] = direct;
            return direct;
        }

        var referenceMappings = GetEffectiveMappingsForProcess(
            referenceProcess,
            processCatalog,
            mappings,
            processSettings,
            cache,
            depth + 1);

        if (referenceMappings.Count == 0)
        {
            cache[cacheKey] = direct;
            return direct;
        }

        var directKeys = direct
            .Select(r => NormalizeSideAgnosticMaterialKey(r.RawMaterialCode, r.RawMaterialName))
            .Where(s => s.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var result = referenceMappings
            .Where(r => !directKeys.Contains(NormalizeSideAgnosticMaterialKey(r.RawMaterialCode, r.RawMaterialName)))
            .Concat(direct)
            .Where(IsSelectableMapping)
            .GroupBy(r => BuildMaterialKey(r.RawMaterialCode, r.RawMaterialName), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToList();

        cache[cacheKey] = result;
        return result;
    }

    private static List<ProcessMaterialMappingRow> GetMappingsForProcess(
        ProcessMaterialNgProcess process,
        IReadOnlyList<ProcessMaterialMappingRow> mappings)
        => mappings
            .Where(r => string.Equals(r.ModelName.Trim(), process.ModelName, StringComparison.OrdinalIgnoreCase))
            .Where(r => ProcessMatches(r.ProcessCode, r.ProcessName, process.ProcessCode, process.ProcessName))
            .Where(IsSelectableMapping)
            .OrderBy(r => r.RawMaterialName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.RawMaterialCode, StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static ProcessMaterialNgProcess? FindProcessByNo(
        string modelName,
        string processNo,
        IReadOnlyList<ProcessMaterialNgProcess> processCatalog,
        IReadOnlyList<ProcessMaterialProcessRow> processSettings)
    {
        string normalizedNo = NormalizeProcessNo(processNo);
        if (normalizedNo.Length == 0)
            return null;

        var exactModel = processCatalog
            .Where(p => string.Equals(p.ModelName, modelName, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var candidates = exactModel.Count > 0
            ? exactModel
            : processCatalog
                .Where(p => string.Equals(
                    NormalizeSideAgnosticModelName(p.ModelName),
                    NormalizeSideAgnosticModelName(modelName),
                    StringComparison.OrdinalIgnoreCase))
                .ToList();

        return candidates.FirstOrDefault(p =>
            string.Equals(NormalizeProcessNo(GetProcessNo(p, processSettings)), normalizedNo, StringComparison.OrdinalIgnoreCase));
    }

    private static string ResolveProcessType(ProcessWeekNg row, IReadOnlyList<RoutingRow> routingRows)
    {
        var exact = routingRows.FirstOrDefault(r =>
            string.Equals(r.ModelName.Trim(), row.ModelName, StringComparison.OrdinalIgnoreCase) &&
            ProcessMatches(r.ProcessCode, r.ProcessName, row.ProcessCode, row.ProcessName));
        if (exact is not null)
            return exact.ProcessType.Trim();

        var linked = routingRows.FirstOrDefault(r =>
            string.Equals(NormalizeSideAgnosticModelName(r.ModelName), NormalizeSideAgnosticModelName(row.ModelName), StringComparison.OrdinalIgnoreCase) &&
            ProcessMatches(r.ProcessCode, r.ProcessName, row.ProcessCode, row.ProcessName));
        return linked?.ProcessType.Trim() ?? string.Empty;
    }

    private static string GetProcessNo(
        ProcessMaterialNgProcess process,
        IReadOnlyList<ProcessMaterialProcessRow> processSettings)
        => GetProcessSetting(process, processSettings)?.ProcessNo.Trim() ?? string.Empty;

    private static string GetReferenceProcessNo(
        ProcessMaterialNgProcess process,
        IReadOnlyList<ProcessMaterialProcessRow> processSettings)
        => GetProcessSetting(process, processSettings)?.ReferenceProcessNo.Trim() ?? string.Empty;

    private static ProcessMaterialProcessRow? GetProcessSetting(
        ProcessMaterialNgProcess process,
        IReadOnlyList<ProcessMaterialProcessRow> processSettings)
        => processSettings.FirstOrDefault(s => ProcessSettingMatches(s, process)) ??
           processSettings.FirstOrDefault(s => ProcessSettingMatchesLinked(s, process));

    private static bool ProcessSettingMatches(ProcessMaterialProcessRow row, ProcessMaterialNgProcess process)
        => string.Equals(row.ModelName.Trim(), process.ModelName, StringComparison.OrdinalIgnoreCase) &&
           string.Equals(NormalizeKey(row.ProcessCode), NormalizeKey(process.ProcessCode), StringComparison.OrdinalIgnoreCase) &&
           string.Equals(NormalizeKey(row.ProcessName), NormalizeKey(process.ProcessName), StringComparison.OrdinalIgnoreCase);

    private static bool ProcessSettingMatchesLinked(ProcessMaterialProcessRow row, ProcessMaterialNgProcess process)
        => string.Equals(NormalizeSideAgnosticModelName(row.ModelName), NormalizeSideAgnosticModelName(process.ModelName), StringComparison.OrdinalIgnoreCase) &&
           (!string.IsNullOrWhiteSpace(process.ProcessName)
                ? string.Equals(NormalizeKey(row.ProcessName), NormalizeKey(process.ProcessName), StringComparison.OrdinalIgnoreCase)
                : string.Equals(NormalizeKey(row.ProcessCode), NormalizeKey(process.ProcessCode), StringComparison.OrdinalIgnoreCase));

    private static bool ProcessMatches(string rowCode, string rowName, string processCode, string processName)
    {
        bool codeMatches =
            !string.IsNullOrWhiteSpace(rowCode) &&
            !string.IsNullOrWhiteSpace(processCode) &&
            string.Equals(rowCode.Trim(), processCode, StringComparison.OrdinalIgnoreCase);

        bool nameMatches =
            !string.IsNullOrWhiteSpace(rowName) &&
            !string.IsNullOrWhiteSpace(processName) &&
            string.Equals(rowName.Trim(), processName, StringComparison.OrdinalIgnoreCase);

        return codeMatches || nameMatches;
    }

    private static bool IsSameProcess(ProcessMaterialNgProcess left, ProcessMaterialNgProcess right)
        => string.Equals(left.ModelName, right.ModelName, StringComparison.OrdinalIgnoreCase) &&
           string.Equals(NormalizeKey(left.ProcessCode), NormalizeKey(right.ProcessCode), StringComparison.OrdinalIgnoreCase) &&
           string.Equals(NormalizeKey(left.ProcessName), NormalizeKey(right.ProcessName), StringComparison.OrdinalIgnoreCase);

    private static bool MatchesModelFilter(RawNgRow row, ModelFilter filter)
    {
        if (string.IsNullOrWhiteSpace(filter.Query))
            return true;

        return Contains(row.MaterialName, filter.Query) ||
               Contains(row.LineShift, filter.Query) ||
               filter.ModelNames.Contains(row.MaterialName.Trim()) ||
               filter.LineShifts.Contains(row.LineShift.Trim());
    }

    private static List<ProcessMaterialNgWeekColumn> BuildWeekColumns(
        DateTime? startDate,
        DateTime? endDate,
        IEnumerable<string> extraWeekKeys)
    {
        var keys = new HashSet<string>(StringComparer.Ordinal);
        if (startDate.HasValue && endDate.HasValue)
        {
            for (var date = startDate.Value.Date; date <= endDate.Value.Date; date = date.AddDays(1))
            {
                string key = BuildWeekKey(date);
                if (key.Length > 0)
                    keys.Add(key);
            }
        }

        foreach (string key in extraWeekKeys)
            if (!string.IsNullOrWhiteSpace(key))
                keys.Add(key);

        return keys
            .OrderByDescending(k => k, StringComparer.Ordinal)
            .Select(k => new ProcessMaterialNgWeekColumn(k, FormatWeekHeader(k)))
            .ToList();
    }

    private static string BuildWeekKey(DateTime date)
    {
        int year = ISOWeek.GetYear(date);
        int week = ISOWeek.GetWeekOfYear(date);
        return $"W:{year:0000}{week:00}";
    }

    private static string FormatWeekHeader(string key)
    {
        string raw = key.StartsWith("W:", StringComparison.Ordinal) ? key[2..] : key;
        return raw.Length >= 6 &&
               int.TryParse(raw[..4], NumberStyles.Integer, CultureInfo.InvariantCulture, out int year) &&
               int.TryParse(raw[4..], NumberStyles.Integer, CultureInfo.InvariantCulture, out int week)
            ? $"{year % 100:00}-W{week:00}"
            : key;
    }

    private static string ProcessIdentityOf(ProcessMaterialNgProcessMaterialRow row)
        => string.Join('\t', NormalizeKey(row.ModelName), NormalizeKey(row.ProcessCode), NormalizeKey(row.ProcessName));

    private static string BuildExactProcessKey(ProcessMaterialNgProcess process)
        => string.Join('\t', NormalizeKey(process.ModelName), NormalizeKey(process.ProcessCode), NormalizeKey(process.ProcessName));

    private static string BuildMaterialKey(string rawMaterialCode, string rawMaterialName)
    {
        string code = NormalizeKey(rawMaterialCode);
        if (code.Length > 0)
            return "C:" + code;
        return "N:" + NormalizeKey(rawMaterialName);
    }

    private static void AddQty(Dictionary<string, double> dict, string key, double qty)
        => dict[key] = dict.GetValueOrDefault(key) + qty;

    private static bool IsSelectableMapping(ProcessMaterialMappingRow row)
        => IsSelectableMaterial(row.RawMaterialCode, row.RawMaterialName);

    private static bool IsSelectableMaterial(string code, string name)
    {
        bool isCs = code.StartsWith("C-S-", StringComparison.OrdinalIgnoreCase);
        bool isRs = code.StartsWith("R-S-", StringComparison.OrdinalIgnoreCase);
        return (isCs || isRs) && !(isCs && IsAssyFrameMaterial(name));
    }

    private static bool IsAssyFrameMaterial(string name)
        => NormalizeKey(name).Contains("ASSY FRAME", StringComparison.OrdinalIgnoreCase);

    private static string NormalizeSideAgnosticMaterialKey(string rawMaterialCode, string rawMaterialName)
    {
        string nameKey = NormalizeSideAgnosticText(rawMaterialName);
        if (!string.IsNullOrWhiteSpace(nameKey))
            return nameKey;

        return NormalizeSideAgnosticText(StripMaterialCodePrefix(rawMaterialCode));
    }

    private static string NormalizeSideAgnosticText(string? value)
    {
        string text = NormalizeKey(value ?? string.Empty);
        if (text.Length == 0)
            return string.Empty;

        var sb = new StringBuilder(text.Length);
        int index = 0;
        while (index < text.Length)
        {
            if (!char.IsLetterOrDigit(text[index]))
            {
                sb.Append(text[index++]);
                continue;
            }

            int start = index;
            while (index < text.Length && char.IsLetterOrDigit(text[index]))
                index++;

            string token = text[start..index];
            sb.Append(IsSideToken(token) ? "*" : token);
        }

        return sb.ToString();
    }

    private static bool IsSideToken(string token)
        => string.Equals(token, "L", StringComparison.OrdinalIgnoreCase) ||
           string.Equals(token, "R", StringComparison.OrdinalIgnoreCase);

    private static string StripMaterialCodePrefix(string value)
    {
        string code = (value ?? string.Empty).Trim();
        return code.StartsWith("C-S-", StringComparison.OrdinalIgnoreCase) ||
               code.StartsWith("R-S-", StringComparison.OrdinalIgnoreCase)
            ? code[4..]
            : code;
    }

    private static string NormalizeSideAgnosticModelName(string? value)
    {
        string modelName = NormalizeKey(value ?? string.Empty);
        if (modelName.Length == 0)
            return string.Empty;

        string[] parts = modelName.Split('-', StringSplitOptions.None);
        for (int i = 0; i < parts.Length; i++)
        {
            if (string.Equals(parts[i], "L", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(parts[i], "R", StringComparison.OrdinalIgnoreCase))
            {
                parts[i] = "*";
            }
        }

        return string.Join('-', parts);
    }

    private static string BuildProcessNoSortKey(string? value)
    {
        string processNo = NormalizeProcessNo(value);
        if (processNo.Length == 0)
            return string.Empty;

        var parts = new List<string>();
        int index = 0;
        while (index < processNo.Length)
        {
            int start = index;
            bool isDigit = char.IsDigit(processNo[index]);
            while (index < processNo.Length && char.IsDigit(processNo[index]) == isDigit)
                index++;

            string part = processNo[start..index];
            parts.Add(isDigit ? part.PadLeft(10, '0') : part);
        }

        return string.Concat(parts);
    }

    private static string NormalizeProcessNo(string? value)
        => (value ?? string.Empty).Trim().TrimStart('#').Trim().ToUpperInvariant();

    private static string NormalizeKey(string? value)
        => (value ?? string.Empty).Trim().ToUpperInvariant();

    private static string NormalizeText(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        input = input.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        input = input.Replace("\u2018", "'").Replace("\u2019", "'")
            .Replace("\u201c", "\"").Replace("\u201d", "\"");
        input = input.Replace("'", " ").Replace("\"", " ").Replace("~", " ");
        input = input.Replace("[", "").Replace("]", "_").Replace("+", " ");
        input = Regex.Replace(input, @"\s{2,}", " ");
        return input.Trim();
    }

    private static bool Contains(string? value, string query)
        => string.IsNullOrWhiteSpace(query) ||
           (value ?? string.Empty).Contains(query.Trim(), StringComparison.OrdinalIgnoreCase);

    private static void AddIfNotBlank(HashSet<string> set, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
            set.Add(value.Trim());
    }

    private static bool TableExists(SqliteConnection conn, string tableName)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@t";
        cmd.Parameters.AddWithValue("@t", tableName);
        return (long)cmd.ExecuteScalar()! > 0;
    }

    private static bool ColumnExists(SqliteConnection conn, string table, string column)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"PRAGMA table_info([{table}])";
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            if (reader.GetString(1).Equals(column, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static string ReadStr(SqliteDataReader reader, int index)
        => reader.IsDBNull(index) ? string.Empty : reader.GetValue(index).ToString() ?? string.Empty;

    private static double ReadDouble(SqliteDataReader reader, int index)
    {
        if (reader.IsDBNull(index)) return 0;
        object value = reader.GetValue(index);
        if (value is double d) return d;
        if (value is long l) return l;
        return double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double result)
            ? result
            : 0;
    }

    private sealed record ModelFilter(string Query)
    {
        public HashSet<string> ModelNames { get; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> LineShifts { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private sealed record RawNgRow(
        string MaterialName,
        string LineShift,
        string ProcessCode,
        string ProcessName,
        string ProcessType,
        DateTime ProductDate,
        double QtyInput,
        double QtyNg,
        string WeekKey);

    private sealed record ProcessWeekNg(
        string ModelName,
        string ProcessCode,
        string ProcessName,
        string ProcessType,
        string WeekKey,
        double InputQty,
        double NgQty);

    private sealed record ProcessMaterialNgProcess(
        string ModelName,
        string ProcessCode,
        string ProcessName,
        string ProcessType);
}

public sealed class ProcessMaterialNgReport
{
    public string DbPath { get; init; } = string.Empty;
    public string ModelFilter { get; init; } = string.Empty;
    public DateTime StartDate { get; init; }
    public DateTime EndDate { get; init; }
    public DateTime GeneratedAt { get; init; } = DateTime.Now;
    public int SourceRowCount { get; init; }
    public int LineShiftCount { get; init; }
    public double TotalProcessNgQty { get; init; }
    public double TotalProcessInputQty { get; init; }
    public double TotalMaterialNgQty { get; init; }
    public List<ProcessMaterialNgWeekColumn> WeekColumns { get; init; } = [];
    public List<ProcessMaterialNgProcessSummaryRow> ProcessRows { get; init; } = [];
    public List<ProcessMaterialNgProcessMaterialRow> ProcessMaterialRows { get; init; } = [];
    public List<ProcessMaterialNgMaterialSummaryRow> MaterialSummaryRows { get; init; } = [];
    public int MappedProcessCount => ProcessRows.Count(r => r.Materials.Count > 0);
    public int UnmappedProcessCount => ProcessRows.Count(r => r.Materials.Count == 0);
}

public sealed record ProcessMaterialNgWeekColumn(string Key, string Header);

public sealed record ProcessMaterialNgMaterialRef(
    string RawMaterialCode,
    string RawMaterialName,
    decimal UsageQty,
    string UsageUnit);

public sealed class ProcessMaterialNgProcessSummaryRow
{
    public string ModelName { get; init; } = string.Empty;
    public string ProcessNo { get; init; } = string.Empty;
    public string ProcessCode { get; init; } = string.Empty;
    public string ProcessName { get; init; } = string.Empty;
    public string ProcessType { get; init; } = string.Empty;
    public List<ProcessMaterialNgMaterialRef> Materials { get; init; } = [];
    public Dictionary<string, double> NgByWeek { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, double> InputByWeek { get; } = new(StringComparer.Ordinal);
    public double TotalNgQty { get; set; }
    public double TotalInputQty { get; set; }
}

public sealed class ProcessMaterialNgProcessMaterialRow
{
    public string ModelName { get; init; } = string.Empty;
    public string ProcessNo { get; init; } = string.Empty;
    public string ProcessCode { get; init; } = string.Empty;
    public string ProcessName { get; init; } = string.Empty;
    public string ProcessType { get; init; } = string.Empty;
    public string RawMaterialCode { get; init; } = string.Empty;
    public string RawMaterialName { get; init; } = string.Empty;
    public decimal UsageQty { get; init; }
    public string UsageUnit { get; init; } = string.Empty;
    public Dictionary<string, double> NgByWeek { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, double> InputByWeek { get; } = new(StringComparer.Ordinal);
    public double TotalNgQty { get; set; }
    public double TotalInputQty { get; set; }
}

public sealed class ProcessMaterialNgMaterialSummaryRow
{
    public string RawMaterialCode { get; init; } = string.Empty;
    public string RawMaterialName { get; init; } = string.Empty;
    public int ProcessCount { get; init; }
    public Dictionary<string, double> NgByWeek { get; } = new(StringComparer.Ordinal);
    public double TotalNgQty { get; set; }
}
