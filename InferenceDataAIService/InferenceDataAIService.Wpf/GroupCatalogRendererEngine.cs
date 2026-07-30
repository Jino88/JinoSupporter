using System.Collections.ObjectModel;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace InferenceDataAIService.Wpf;

// Catalog-aware rendering boundary.  It consumes only the immutable grouping
// catalog and the existing batch-local numeric database: it never reopens Excel.
internal sealed class GroupCatalogRendererEngine
{
    private const string Version = "group-catalog-renderer-v5-support-index-suppression";
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
    private static readonly IReadOnlyDictionary<(int Row, int Column), string> EmptyCells = new Dictionary<(int, int), string>();
    private static readonly IReadOnlyDictionary<string, GroupProfile> Profiles = new Dictionary<string, GroupProfile>(StringComparer.Ordinal)
    {
        ["renderSingleRawMeasurementSummary"] = new("renderSingleRawMeasurementSummary", "measurement-summary-v1", ["MeasurementSummary"], true),
        ["renderMultiRepeatedDefectBlocks"] = new("renderMultiRepeatedDefectBlocks", "repeated-defect-blocks-v1", ["RepeatedDefectBlocks"], true),
        ["renderNeedsReview"] = new("renderNeedsReview", "structure-section-panels-v1", [], true),
        ["renderSingleDefectAccounting"] = new("renderSingleDefectAccounting", "defect-accounting-v1", ["DefectAccounting"], true),
        ["renderMultiDefectAccounting"] = new("renderMultiDefectAccounting", "defect-accounting-v1", ["DefectAccounting"], true),
        ["renderMultiDefectNgBreakdown"] = Pending("renderMultiDefectNgBreakdown", "DefectAccounting", "NgBreakdown"),
        ["renderMultiDefectNgRepeated"] = Pending("renderMultiDefectNgRepeated", "DefectAccounting", "NgBreakdown", "RepeatedDefectBlocks"),
        ["renderMultiDefectCohortNg"] = Pending("renderMultiDefectCohortNg", "DefectAccounting", "CohortComparison", "NgBreakdown"),
        ["renderMultiDefectRepeated"] = new("renderMultiDefectRepeated", "defect-repeated-v1", ["DefectAccounting", "RepeatedDefectBlocks"], true),
        ["renderMultiDefectCohortNgRepeated"] = Pending("renderMultiDefectCohortNgRepeated", "DefectAccounting", "CohortComparison", "NgBreakdown", "RepeatedDefectBlocks"),
        ["renderMultiDefectCohort"] = Pending("renderMultiDefectCohort", "DefectAccounting", "CohortComparison"),
        ["renderSingleCohortComparison"] = Pending("renderSingleCohortComparison", "CohortComparison"),
        ["renderMultiNgRepeated"] = Pending("renderMultiNgRepeated", "NgBreakdown", "RepeatedDefectBlocks"),
        ["renderSingleNgBreakdown"] = Pending("renderSingleNgBreakdown", "NgBreakdown"),
        ["renderMultiCohortComparison"] = Pending("renderMultiCohortComparison", "CohortComparison"),
        ["renderMultiDefectRawMeasurement"] = new("renderMultiDefectRawMeasurement", "defect-measurement-v1", ["DefectAccounting", "MeasurementSummary"], true),
        ["renderMultiDefectNgRawMeasurement"] = Pending("renderMultiDefectNgRawMeasurement", "DefectAccounting", "NgBreakdown", "MeasurementSummary"),
        ["renderMultiNgBreakdown"] = Pending("renderMultiNgBreakdown", "NgBreakdown"),
        ["renderMultiDefectCohortRepeated"] = Pending("renderMultiDefectCohortRepeated", "DefectAccounting", "CohortComparison", "RepeatedDefectBlocks"),
        ["renderMultiCohortNgRepeated"] = Pending("renderMultiCohortNgRepeated", "CohortComparison", "NgBreakdown", "RepeatedDefectBlocks"),
        ["defect-cohort-repeated-result-v1"] = new("defect-cohort-repeated-result-v1", "defect-cohort-repeated-result-v1", [], true),
        ["cohort-ng-repeated-block-v1"] = new("cohort-ng-repeated-block-v1", "cohort-ng-repeated-block-v1", [], true),
        ["dimension-summary-comparison-v1"] = new("dimension-summary-comparison-v1", "dimension-summary-comparison-v1", [], true),
        // AI group-plan renderer keys.  These profiles render the full captured
        // source sections while dedicated semantic component renderers are refined.
        ["dimension-comparison-panels-v1"] = new("dimension-comparison-panels-v1", "dimension-comparison-panels-v1", [], true),
        ["process-function-cohort-report-v1"] = new("process-function-cohort-report-v1", "process-function-cohort-report-v1", [], true),
        ["doe-function-defect-block-v1"] = new("doe-function-defect-block-v1", "doe-function-defect-block-v1", [], true),
        ["acoustic-dashboard-v1"] = new("acoustic-dashboard-v1", "acoustic-dashboard-v1", [], true),
        ["quality-dashboard-v1"] = new("quality-dashboard-v1", "quality-dashboard-v1", [], true),
        ["measurement-dashboard-v1"] = new("measurement-dashboard-v1", "measurement-dashboard-v1", [], true),
        ["function-process-dashboard-v1"] = new("function-process-dashboard-v1", "function-process-dashboard-v1", [], true),
        ["tension-dashboard-v1"] = new("tension-dashboard-v1", "tension-dashboard-v1", [], true),
        ["general-table-dashboard-v1"] = new("general-table-dashboard-v1", "general-table-dashboard-v1", [], true),
    };

    // Candidate types are persisted by NumericCaptureEngine.  Keeping this
    // registry here makes the group-plan componentRecipe an executable and
    // validated rendering contract instead of a free-text hint.
    private static readonly IReadOnlyDictionary<string, ComponentDescriptor> CapturedComponents = new Dictionary<string, ComponentDescriptor>(StringComparer.Ordinal)
    {
        ["NUMERIC_TABLE_UNCLASSIFIED"] = new("NUMERIC_TABLE_UNCLASSIFIED", "Captured numeric tables"),
        ["DEFECT_RATE_NUMERIC_TABLE"] = new("DEFECT_RATE_NUMERIC_TABLE", "Defect-rate tables"),
        ["REPEATED_DEFECT_BLOCK_NUMERIC_TABLE"] = new("REPEATED_DEFECT_BLOCK_NUMERIC_TABLE", "Repeated defect blocks"),
        ["MEASUREMENT_SUMMARY_NUMERIC_TABLE"] = new("MEASUREMENT_SUMMARY_NUMERIC_TABLE", "Measurement-summary tables"),
        ["TEST_NORMAL_NUMERIC_TABLE"] = new("TEST_NORMAL_NUMERIC_TABLE", "Test/normal comparison tables"),
        ["EMPTY_LAYOUT"] = new("EMPTY_LAYOUT", "Scanner-proven empty layout"),
        ["CAPTURE_INCOMPLETE"] = new("CAPTURE_INCOMPLETE", "Scanner-proven incomplete capture"),
        ["EMPTY_LAYOUT_OR_CAPTURE_INCOMPLETE"] = new("EMPTY_LAYOUT_OR_CAPTURE_INCOMPLETE", "Scanner-proven capture exception"),
    };
    private static readonly IReadOnlyDictionary<string, string> ComponentAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["TABLE_UNCLASSIFIED"] = "NUMERIC_TABLE_UNCLASSIFIED",
        ["NUMERIC_TABLE_UNCLASSIFIED"] = "NUMERIC_TABLE_UNCLASSIFIED",
        ["DEFECT_RATE_NUMERIC_TABLE"] = "DEFECT_RATE_NUMERIC_TABLE",
        ["REPEATED_DEFECT_BLOCK_NUMERIC_TABLE"] = "REPEATED_DEFECT_BLOCK_NUMERIC_TABLE",
        ["MEASUREMENT_SUMMARY_NUMERIC_TABLE"] = "MEASUREMENT_SUMMARY_NUMERIC_TABLE",
        ["TEST_NORMAL_NUMERIC_TABLE"] = "TEST_NORMAL_NUMERIC_TABLE",
        ["EMPTY_LAYOUT"] = "EMPTY_LAYOUT",
        ["CAPTURE_INCOMPLETE"] = "CAPTURE_INCOMPLETE",
        ["EMPTY_LAYOUT_OR_CAPTURE_INCOMPLETE"] = "EMPTY_LAYOUT_OR_CAPTURE_INCOMPLETE",
    };

    private readonly Action<string>? _log;
    private GroupCatalogRendererEngine(Action<string>? log) => _log = log;
    internal static GroupCatalogRendererRunResult Run(GroupCatalogRendererRequest request, Action<string>? log = null, CancellationToken cancellationToken = default) => new GroupCatalogRendererEngine(log).RunCore(request, cancellationToken);
    internal static bool IsRegisteredRendererKey(string key) => Profiles.ContainsKey(key);
    internal static bool IsSupportedComponentRecipe(IReadOnlyList<string> recipe) => recipe.Count > 0 && recipe.All(component => CapturedComponents.ContainsKey(NormalizeComponentKey(component)));
    private static string NormalizeComponentKey(string component) => ComponentAliases.GetValueOrDefault(component.Trim(), component.Trim());

    private GroupCatalogRendererRunResult RunCore(GroupCatalogRendererRequest request, CancellationToken cancellationToken)
    {
        var batch = Path.Combine(
            AppRuntimePaths.Current.BatchRootDirectory,
            request.StructureBatchId);
        _log?.Invoke("Catalog renderer: loading catalog.");
        var catalogPath = Path.Combine(batch, "group-catalog.json");
        if (!File.Exists(catalogPath)) return new GroupCatalogRendererRunResult(batch, false, null);
        EnsureValidatedTwoLevelCatalog(batch, catalogPath);
        var catalog = Catalog.Load(catalogPath);
        var unknownKeys = catalog.Groups.Values.Select(group => group.RendererKey).Where(key => !Profiles.ContainsKey(key)).Distinct().ToList();
        if (unknownKeys.Count > 0) throw new InvalidOperationException($"Catalog uses unregistered renderer keys: {string.Join(", ", unknownKeys)}.");
        var fallbackGroup = catalog.Groups.Values.FirstOrDefault(group => string.Equals(group.Id, "needs-review-fallback", StringComparison.Ordinal))
            ?? catalog.Groups.Values.FirstOrDefault(group => string.Equals(group.RendererKey, "renderNeedsReview", StringComparison.Ordinal));
        if (fallbackGroup is null) throw new InvalidOperationException("Catalog requires a renderNeedsReview fallback group.");
        var database = Path.Combine(batch, "numeric-capture.sqlite");
        if (!File.Exists(database)) throw new FileNotFoundException("Numeric capture database is required for group rendering.", database);
        var reportDirectory = Path.Combine(batch, "group-reports");
        var rows = new List<GroupRenderRow>();
        _log?.Invoke("Catalog renderer: opening numeric capture database.");
        using var connection = new SqliteConnection($"Data Source={database}");
        connection.Open();
        _log?.Invoke("Catalog renderer: reading workbook list.");
        var workbooks = new List<(long Id, string Path)>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText = "SELECT workbook_id, relative_path FROM capture_workbooks ORDER BY relative_path;";
            using var reader = command.ExecuteReader();
            while (reader.Read()) workbooks.Add((reader.GetInt64(0), reader.GetString(1)));
        }
        ValidateCatalogComponentContracts(connection, catalog, workbooks);
        // Do not delete a previous successful render until every assigned Excel
        // has proved that its persisted candidate sequence matches the immutable
        // component recipe.  This makes contract failure non-destructive.
        if (Directory.Exists(reportDirectory)) Directory.Delete(reportDirectory, recursive: true);
        Directory.CreateDirectory(reportDirectory);
        _log?.Invoke($"Catalog renderer: rendering {workbooks.Count} workbooks.");
        for (var workbookIndex = 0; workbookIndex < workbooks.Count; workbookIndex++)
        {
            var (id, path) = workbooks[workbookIndex];
            cancellationToken.ThrowIfCancellationRequested();
            if (workbookIndex == 0 || (workbookIndex + 1) % 10 == 0 || workbookIndex + 1 == workbooks.Count)
                _log?.Invoke($"Catalog renderer: {workbookIndex + 1}/{workbooks.Count} rendering {path}");
            // Captured but non-SCANNED files (for example truncated layouts) are
            // intentionally outside the immutable catalog assignment coverage.
            // They remain visible as needs-review, never silently rendered as data.
            if (!catalog.Assignments.TryGetValue(path, out var group)) group = fallbackGroup;
            var profile = Profiles[group.RendererKey];
            // The two-level catalog is a rendering contract, not a label.  Its
            // component recipe selects which captured numeric-table components are
            // materialized for this particular workbook.  The previous renderer
            // discarded that recipe and dumped every candidate table as an arbitrary
            // list, so an AI form group had no effect on the dashboard itself.
            var componentRender = string.Equals(profile.Key, "renderNeedsReview", StringComparison.Ordinal)
                ? null
                : RenderCatalogComponents(connection, id, group);
            // Batch capture contains structural sections for every workbook, but
            // legacy numeric-review tables exist only in the old single-file path.
            // Do not query those optional tables for structure-driven profiles.
            var facts = componentRender?.ComponentCounts
                ?? (profile.RequiredComponents.Count == 0
                    ? new Dictionary<string, int>(StringComparer.Ordinal)
                    : ComponentCounts(connection, id));
            var status = componentRender?.Status
                ?? (string.Equals(profile.Key, "renderNeedsReview", StringComparison.Ordinal)
                    ? "STRUCTURE_READY"
                    : profile.Key.EndsWith("-v1", StringComparison.Ordinal) && profile.RequiredComponents.Count == 0
                        ? "REVIEW_RENDERED"
                    : profile.IsImplemented ? (profile.RequiredComponents.All(component => facts.TryGetValue(component, out var count) && count > 0) ? "READY" : "NEEDS_REVIEW") : "PENDING_PROFILE");
            var reportPath = $"group-reports/{StableId(path)}.html";
            var sourceSections = componentRender?.SourceSections ?? (string.Equals(profile.Key, "renderNeedsReview", StringComparison.Ordinal) ||
                                 profile.Key is "defect-cohort-repeated-result-v1" or "cohort-ng-repeated-block-v1" or "dimension-summary-comparison-v1"
                                     or "dimension-comparison-panels-v1" or "process-function-cohort-report-v1" or "doe-function-defect-block-v1"
                                     or "acoustic-dashboard-v1" or "quality-dashboard-v1" or "measurement-dashboard-v1" or "function-process-dashboard-v1" or "tension-dashboard-v1" or "general-table-dashboard-v1"
                ? RenderCapturedSections(connection, id, resultTitlesOnly: false)
                : string.Equals(group.Id, "multi-repeated-defect-blocks", StringComparison.Ordinal) ? RenderCapturedSections(connection, id, resultTitlesOnly: true) : []);
            AtomicWrite(Path.Combine(batch, reportPath.Replace('/', Path.DirectorySeparatorChar)), GroupReport(connection, id, path, request.StructureBatchId, group, profile, status, componentRender, sourceSections));
            for (var sectionIndex = 0; sectionIndex < sourceSections.Count; sectionIndex++)
            {
                var sectionPath = $"group-reports/{StableId(path)}-review-{sectionIndex + 1}.html";
                AtomicWrite(Path.Combine(batch, sectionPath.Replace('/', Path.DirectorySeparatorChar)), Page(path, group, profile, "PARTIAL", $"<section><h2>{Escape(sourceSections[sectionIndex].Title)}</h2><div class='table-wrap'>{sourceSections[sectionIndex].Table}</div></section>"));
            }
            rows.Add(new GroupRenderRow(path, group.Id, profile.Key, profile.Version, status, reportPath, facts));
        }
        var summary = new GroupCatalogRendererSummary("group-catalog-render-summary-v1", Version, DateTimeOffset.UtcNow.ToString("O"), rows.Count,
            new ReadOnlyDictionary<string, int>(rows.GroupBy(row => row.Status).ToDictionary(group => group.Key, group => group.Count())), "group-render-index.html", "group-render-index.json");
        AtomicWrite(Path.Combine(batch, summary.Json), JsonSerializer.Serialize(new { summary, rows }, JsonOptions) + "\n");
        AtomicWrite(Path.Combine(batch, summary.Html), IndexHtml(rows, summary));
        _log?.Invoke($"Catalog renderer completed: {rows.Count} workbooks, {summary.StatusCounts.GetValueOrDefault("COMPONENTS_RENDERED", 0)} component-contract renders.");
        return new GroupCatalogRendererRunResult(batch, true, summary);
    }

    private static void EnsureValidatedTwoLevelCatalog(string batch, string catalogPath)
    {
        var planPath = Path.Combine(batch, "group-plan.json");
        var validationPath = Path.Combine(batch, "group-validation-report.json");
        if (!File.Exists(planPath) || !File.Exists(validationPath))
            throw new InvalidOperationException("Group rendering requires a validated two-level plan and validation report.");
        try
        {
            using var plan = JsonDocument.Parse(File.ReadAllText(planPath, Encoding.UTF8));
            using var validation = JsonDocument.Parse(File.ReadAllText(validationPath, Encoding.UTF8));
            var schema = plan.RootElement.GetProperty("schemaVersion").GetString();
            var valid = validation.RootElement.GetProperty("isValid").GetBoolean();
            if (!string.Equals(schema, TwoLevelGroupAnalysisEngine.PlanSchemaVersion, StringComparison.Ordinal) || !valid)
                throw new InvalidOperationException("Group rendering is blocked because the catalog did not pass two-level grouping validation.");
        }
        catch (JsonException exception)
        {
            throw new InvalidOperationException($"Group rendering is blocked because validation artifacts are invalid JSON: {catalogPath}.", exception);
        }
    }

    private static GroupProfile Pending(string key, params string[] components) => new(key, "pending-v1", components, false);
    private static Dictionary<string, int> ComponentCounts(SqliteConnection connection, long workbookId) => new(StringComparer.Ordinal)
    {
        ["DefectAccounting"] = Count(connection, "SELECT COUNT(*) FROM numeric_review_facts WHERE workbook_id=$id;", workbookId),
        ["RepeatedDefectBlocks"] = Count(connection, "SELECT COUNT(*) FROM repeated_defect_block_facts WHERE workbook_id=$id;", workbookId),
        ["MeasurementSummary"] = Count(connection, "SELECT COUNT(*) FROM measurement_summary_facts WHERE workbook_id=$id;", workbookId),
        ["NgBreakdown"] = 0,
        ["CohortComparison"] = 0,
    };
    private static int Count(SqliteConnection connection, string sql, long workbookId) { using var command = connection.CreateCommand(); command.CommandText = sql; command.Parameters.AddWithValue("$id", workbookId); return Convert.ToInt32(command.ExecuteScalar()); }

    private void ValidateCatalogComponentContracts(SqliteConnection connection, Catalog catalog, IReadOnlyList<(long Id, string Path)> workbooks)
    {
        var errors = new List<string>();
        var validated = 0;
        foreach (var workbook in workbooks)
        {
            if (!catalog.Assignments.TryGetValue(workbook.Path, out var group) || IsCaptureExceptionRecipe(group.ComponentRecipe))
                continue;
            var actual = CapturedComponentRecipe(connection, workbook.Id);
            if (!actual.SequenceEqual(group.ComponentRecipe, StringComparer.Ordinal))
                errors.Add($"{workbook.Path}: catalog=[{string.Join(" | ", group.ComponentRecipe)}], captured=[{string.Join(" | ", actual)}]");
            validated++;
        }
        if (errors.Count > 0)
            throw new InvalidOperationException($"Component-contract preflight failed for {errors.Count} workbook(s). No existing group report was deleted. {string.Join("; ", errors.Take(5))}");
        _log?.Invoke($"Catalog renderer: component-contract preflight passed for {validated} assigned workbooks.");
    }

    private static IReadOnlyList<string> CapturedComponentRecipe(SqliteConnection connection, long workbookId)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT t.candidate_type FROM numeric_table_candidates t JOIN captured_sheets s ON s.sheet_id=t.sheet_id WHERE s.workbook_id=$id ORDER BY s.sheet_index, t.start_row, t.start_column;";
        command.Parameters.AddWithValue("$id", workbookId);
        using var reader = command.ExecuteReader();
        var recipe = new List<string>();
        while (reader.Read())
        {
            var component = NormalizeComponentKey(reader.GetString(0));
            if (!recipe.Contains(component, StringComparer.Ordinal)) recipe.Add(component);
        }
        return recipe;
    }

    private static bool IsCaptureExceptionRecipe(IReadOnlyList<string> recipe) => recipe.Any(component => component is "EMPTY_LAYOUT" or "CAPTURE_INCOMPLETE" or "EMPTY_LAYOUT_OR_CAPTURE_INCOMPLETE");
    private static void AtomicWrite(string path, string text) { var temp = path + ".tmp"; File.WriteAllText(temp, text, new UTF8Encoding(false)); File.Move(temp, path, true); }
    private static string IndexHtml(IReadOnlyList<GroupRenderRow> rows, GroupCatalogRendererSummary summary)
    {
        var body = string.Concat(rows.Select(row =>
        {
            var components = string.Join("; ", row.ComponentCounts.Where(component => component.Value > 0).OrderBy(component => component.Key, StringComparer.Ordinal).Select(component => $"{component.Key}={component.Value}"));
            return $"<tr><td>{Escape(row.RelativePath)}</td><td>{Escape(row.GroupId)}</td><td>{Escape(row.RendererKey)}</td><td>{Escape(row.ProfileVersion)}</td><td>{Escape(components)}</td><td>{Escape(row.Status)}</td><td><a href='{Escape(row.ReportPath)}'>Open</a></td></tr>";
        }));
        return $"<!doctype html><html><head><meta charset='utf-8'><title>Group render index</title><style>body{{font-family:Segoe UI,sans-serif;margin:24px;color:#162d50}}table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #d5dfeb;padding:8px;text-align:left}}th{{background:#eaf0f7}}</style></head><body><h1>AI group render index</h1><p>Workbooks: {summary.WorkbookCount}; component-contract rendered: {summary.StatusCounts.GetValueOrDefault("COMPONENTS_RENDERED", 0)}; contract mismatch: {summary.StatusCounts.GetValueOrDefault("CONTRACT_MISMATCH", 0)}.</p><table><thead><tr><th>Excel</th><th>Group</th><th>Renderer</th><th>Profile</th><th>Captured components</th><th>Status</th><th>Report</th></tr></thead><tbody>{body}</tbody></table></body></html>";
    }
    private static string GroupReport(SqliteConnection connection, long workbookId, string path, string batchId, CatalogGroup group, GroupProfile profile, string status, ComponentRender? componentRender, IReadOnlyList<SourceSection>? sourceSections = null)
    {
        if (status == "PENDING_PROFILE") return Page(path, group, profile, status, $"<p class='notice'>This renderer requires components not implemented yet: {Escape(string.Join(", ", profile.RequiredComponents))}.</p>");
        if (componentRender is not null) return ComponentDashboard(path, batchId, group, profile, componentRender);
        if (sourceSections is { Count: > 0 })
        {
            var document = new DashboardDocument(
                Path.GetFileName(path),
                new DashboardSourceTrace(path, batchId, group.Id, "numeric-capture-v3-full-document-grid", "group-catalog.json", profile.Version),
                [new DashboardSummaryItem("양식 그룹", group.Id, "REVIEW_REQUIRED", "AI catalog"), new DashboardSummaryItem("추출 범위", $"{sourceSections.Count}개 검토 섹션", "REVIEW_REQUIRED", "captured section inventory"), new DashboardSummaryItem("렌더러", profile.Version, "REVIEW_REQUIRED", "renderer registry")],
                sourceSections.Select(section => new DashboardSection(section.Title, "원본 표 섹션", section.Evidence, $"<div class='table-wrap'>{section.Table}</div>")).ToList(),
                [],
                new DashboardAssessment("REVIEW_REQUIRED", "표 구조와 수치는 추출됐지만, 이 그룹의 업무 판정·비교 규칙은 별도 프로필 검증이 필요합니다."),
                ["원본 조건·lot·모델·표본 선택 규칙을 확인하세요.", "명시적으로 같은 날짜·조건인 행만 비교하세요.", "허용 한계와 승인 기준은 원본 규격으로 확인하세요."]);
            return DashboardHtmlRenderer.Render(document);
        }
        var sections = new StringBuilder();
        if (profile.RequiredComponents.Contains("DefectAccounting")) sections.Append(Section("Defect accounting", Table(connection, workbookId, "SELECT condition_label, input_value, total_ng_value, computed_ng_rate, fact_status FROM numeric_review_facts WHERE workbook_id=$id ORDER BY sheet_id, table_id, row_index;", ["Condition", "Input", "Total NG", "NG rate", "Status"])));
        if (profile.RequiredComponents.Contains("RepeatedDefectBlocks")) sections.Append(Section("Repeated defect blocks", RepeatedBlocksTable(connection, workbookId)));
        if (profile.RequiredComponents.Contains("MeasurementSummary")) sections.Append(Section("Measurement summary", Table(connection, workbookId, "SELECT condition_label, sample_value, minimum_value, average_value, maximum_value, fact_status FROM measurement_summary_facts WHERE workbook_id=$id ORDER BY sheet_id, table_id, row_index;", ["Condition", "Sample", "Min", "Average", "Max", "Status"])));
        if (status == "NEEDS_REVIEW") sections.Insert(0, "<p class='notice'>The required component has no validated facts for this workbook.</p>");
        return Page(path, group, profile, status, sections.ToString());
    }

    // A component is the persisted candidate type from numeric-capture.sqlite.
    // The recipe order comes from the validated two-level catalog and therefore
    // drives dashboard order; source-table order is retained only within a
    // component.  This is the point where group/extraction contracts become
    // executable rather than display-only prose.
    private static ComponentRender RenderCatalogComponents(SqliteConnection connection, long workbookId, CatalogGroup group)
    {
        var allSections = ReadCapturedSections(connection, workbookId);
        var expected = group.ComponentRecipe.Select(NormalizeComponentKey).Distinct(StringComparer.Ordinal).ToList();
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        var dashboardSections = new List<DashboardSection>();
        var selected = new List<SourceSection>();
        var missing = new List<string>();

        foreach (var componentKey in expected)
        {
            var descriptor = CapturedComponents[componentKey];
            var componentCandidates = allSections.Where(section => string.Equals(section.ComponentKey, componentKey, StringComparison.Ordinal)).ToList();
            var componentTables = componentCandidates.Where(section => !section.IsSupportOnly).ToList();
            var suppressedSupportCount = componentCandidates.Count - componentTables.Count;
            counts[componentKey] = componentTables.Count;
            if (componentTables.Count == 0 && suppressedSupportCount == 0) missing.Add(componentKey);
            selected.AddRange(componentTables.Select(section => new SourceSection(section.ComponentKey, section.Title, section.Table, section.Evidence, section.NumericCellCount)));
            if (componentTables.Count > 0 || suppressedSupportCount == 0)
                dashboardSections.Add(new DashboardSection(
                    descriptor.Title,
                    componentKey,
                    $"catalog component recipe; {componentTables.Count} captured table(s)",
                    ComponentTablesHtml(descriptor, componentTables)));
        }

        // A strict component recipe must not hide persisted data.  Unexpected
        // candidate types are rendered in an explicit reconciliation panel and
        // make the output a contract mismatch, rather than being mixed silently
        // into an unrelated dashboard component.
        var unexpected = allSections
            .Where(section => !section.IsSupportOnly && !expected.Contains(section.ComponentKey, StringComparer.Ordinal))
            .GroupBy(section => section.ComponentKey, StringComparer.Ordinal)
            .OrderBy(grouping => grouping.Key, StringComparer.Ordinal)
            .ToList();
        foreach (var unexpectedComponent in unexpected)
        {
            var descriptor = CapturedComponents.GetValueOrDefault(unexpectedComponent.Key, new ComponentDescriptor(unexpectedComponent.Key, "Uncontracted captured component"));
            var tables = unexpectedComponent.ToList();
            counts[unexpectedComponent.Key] = tables.Count;
            selected.AddRange(tables.Select(section => new SourceSection(section.ComponentKey, section.Title, section.Table, section.Evidence, section.NumericCellCount)));
            dashboardSections.Add(new DashboardSection(
                $"Uncontracted: {descriptor.Title}",
                unexpectedComponent.Key,
                "Captured data is present but not named by the validated component recipe.",
                ComponentTablesHtml(descriptor, tables)));
        }

        return new ComponentRender(
            dashboardSections,
            selected,
            new ReadOnlyDictionary<string, int>(counts),
            missing,
            unexpected.Select(grouping => grouping.Key).ToList());
    }

    private static string ComponentDashboard(string path, string batchId, CatalogGroup group, GroupProfile profile, ComponentRender render)
    {
        var componentText = string.Join(" → ", group.ComponentRecipe.Select(NormalizeComponentKey));
        var tableCount = render.ComponentCounts.Values.Sum();
        var renderedComponentTypeCount = render.ComponentCounts.Count(component => component.Value > 0);
        var findings = new List<DashboardFinding>();
        foreach (var component in render.MissingComponents)
            findings.Add(new DashboardFinding("Missing captured component", component, "REVIEW_REQUIRED", "The catalog component recipe names this component, but numeric-capture.sqlite has no matching candidate table."));
        foreach (var component in render.UnexpectedComponents)
            findings.Add(new DashboardFinding("Uncontracted captured component", component, "REVIEW_REQUIRED", "The persisted candidate is visible below and requires catalog reconciliation before a fully contract-matched render."));

        var requirements = group.OpenQuestions.Concat([
            "Compare only rows that share the source workbook's stated condition, lot, model, and sample basis.",
            "Business acceptance, causality, and pass/fail conclusions remain outside this capture-only renderer."])
            .Distinct(StringComparer.Ordinal)
            .ToList();
        var assessment = render.IsContractSatisfied
            ? new DashboardAssessment("CAPTURED_COMPONENTS", "The dashboard is materialized from the validated component recipe and captured numeric tables. It does not create a business decision.")
            : new DashboardAssessment("CONTRACT_MISMATCH", "Captured data was retained, but the persisted tables do not exactly match the validated component recipe. Review the explicit reconciliation panels.");
        var document = new DashboardDocument(
            Path.GetFileName(path),
            new DashboardSourceTrace(path, batchId, group.Id, "numeric-capture-v3-component-materialization", "group-catalog.json", profile.Version),
            [
                new DashboardSummaryItem("Form group", group.Name, "CAPTURED_COMPONENTS", group.Id),
                new DashboardSummaryItem("Component recipe", componentText, render.IsContractSatisfied ? "CAPTURED_COMPONENTS" : "CONTRACT_MISMATCH", "validated group-catalog.json"),
                new DashboardSummaryItem("Captured data", $"{tableCount} table(s) across {renderedComponentTypeCount} component type(s)", "CAPTURED_COMPONENTS", "numeric-capture.sqlite / numeric_table_candidates"),
                new DashboardSummaryItem("Important data", string.Join("; ", group.ImportantData), "REVIEW_REQUIRED", "validated group-catalog.json"),
                new DashboardSummaryItem("Extraction contract", group.ExtractionRule, "REVIEW_REQUIRED", "validated group-catalog.json"),
                new DashboardSummaryItem("HTML contract", group.HtmlRule, "REVIEW_REQUIRED", "validated group-catalog.json")
            ],
            render.Sections,
            findings,
            assessment,
            requirements);
        return DashboardHtmlRenderer.Render(document);
    }

    private static string ComponentTablesHtml(ComponentDescriptor descriptor, IReadOnlyList<CapturedSection> tables)
    {
        if (tables.Count == 0)
            return $"<p class='empty'>No captured {Escape(descriptor.Title)} table matches this component contract.</p>";
        var numericCells = tables.Sum(table => table.NumericCellCount);
        var body = new StringBuilder($"<p class='component'>Captured {tables.Count} table(s), {numericCells:N0} numeric cell(s). Each table is materialized directly from numeric-capture.sqlite.</p>");
        foreach (var table in tables)
        {
            body.Append($"<article class='source-card'><h3>{Escape(table.Title)}</h3><p class='evidence'>{Escape(table.Evidence)}</p><div class='table-wrap'>{table.Table}</div></article>");
        }
        return body.ToString();
    }

    // Preserves the original multi-section grid for the Frame-bending layout.
    // This is deliberately separate from generic repeated-block facts: the latter
    // cannot retain the Date/Line/Mold Frame axes or the THD/SPL/Noise/Touch columns.
    private static List<SourceSection> RenderCapturedSections(SqliteConnection connection, long workbookId, bool resultTitlesOnly)
    {
        return ReadCapturedSections(connection, workbookId)
            .Where(section => !section.IsSupportOnly && (!resultTitlesOnly || section.Title.Contains("RESULT", StringComparison.OrdinalIgnoreCase)))
            .Select(section => new SourceSection(section.ComponentKey, section.Title, section.Table, section.Evidence, section.NumericCellCount))
            .ToList();
    }

    private static List<CapturedSection> ReadCapturedSections(SqliteConnection connection, long workbookId)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT t.sheet_id, s.sheet_name, t.start_row, t.end_row, t.start_column, t.end_column, t.header_start_row, t.header_end_row, t.candidate_type, t.numeric_cell_count FROM numeric_table_candidates t JOIN captured_sheets s ON s.sheet_id=t.sheet_id WHERE s.workbook_id=$id ORDER BY s.sheet_index, t.start_row, t.start_column;";
        command.Parameters.AddWithValue("$id", workbookId); using var reader = command.ExecuteReader();
        var candidates = new List<CapturedCandidate>();
        while (reader.Read()) candidates.Add(new CapturedCandidate(
            reader.GetInt64(0), reader.GetString(1), reader.GetInt32(2), reader.GetInt32(3), reader.GetInt32(4), reader.GetInt32(5), reader.GetInt32(6), reader.GetInt32(7),
            NormalizeComponentKey(reader.GetString(8)), reader.GetInt32(9)));
        var supportCandidates = DetectSupportIndexCandidates(connection, candidates);
        var result = new List<CapturedSection>();
        foreach (var candidate in candidates)
        {
            var context = CapturedTableContextResolver.Resolve(connection, new CapturedTableRegion(
                candidate.Sheet, candidate.SheetName, candidate.Start, candidate.End, candidate.FirstColumn,
                candidate.LastColumn, candidate.HeaderStart, candidate.HeaderEnd));
            var title = string.IsNullOrWhiteSpace(context.Title)
                ? CapturedComponents.GetValueOrDefault(candidate.ComponentKey, new ComponentDescriptor(candidate.ComponentKey, "Captured numeric table")).Title
                : context.Title;
            var candidateRange = $"{ColumnName(candidate.FirstColumn)}{candidate.Start}:{ColumnName(candidate.LastColumn)}{candidate.End}";
            var contextRange = $"{ColumnName(context.LeftColumn)}{context.TopRow}:{ColumnName(context.RightColumn)}{context.BottomRow}";
            var evidence = candidateRange == contextRange
                ? $"{candidate.SheetName}!{contextRange}"
                : $"{candidate.SheetName}!{contextRange} (numeric candidate {candidateRange})";
            var isSupportOnly = supportCandidates.Contains(candidate);
            result.Add(new CapturedSection(candidate.ComponentKey, title, isSupportOnly ? string.Empty : SourceGridTable(connection, candidate, context), evidence, candidate.NumericCellCount, isSupportOnly));
        }
        return result;
    }

    // A one-column serial field is support for an overlapping multi-column
    // table, not an independent analytic table. Require every signal below so
    // a real one-column measurement or log remains visible: an index header,
    // at least five non-negative integer observations, a predominantly
    // sequential/reset sequence, and a nearby overlapping primary candidate.
    private static HashSet<CapturedCandidate> DetectSupportIndexCandidates(SqliteConnection connection, IReadOnlyList<CapturedCandidate> candidates)
    {
        var support = new HashSet<CapturedCandidate>();
        foreach (var candidate in candidates)
        {
            if (candidate.FirstColumn != candidate.LastColumn || !IsIndexHeader(connection, candidate) || !HasNearbyPrimaryCandidate(candidate, candidates)) continue;
            var values = ReadDirectNumericValues(connection, candidate);
            if (values.Count < 5 || values.Any(value => value < 0 || Math.Abs(value - Math.Round(value)) > 0.000001d)) continue;
            var sequentialTransitions = values.Skip(1).Zip(values, (current, previous) => IsSequentialIndexStep(previous, current)).Count(value => value);
            if (sequentialTransitions / (double)(values.Count - 1) >= 0.75d) support.Add(candidate);
        }
        return support;
    }

    private static bool IsIndexHeader(SqliteConnection connection, CapturedCandidate candidate)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT text_value FROM captured_text_cells WHERE sheet_id=$sheet AND column_index=$column AND row_index BETWEEN $start AND $end AND trim(text_value) <> '';";
        command.Parameters.AddWithValue("$sheet", candidate.Sheet); command.Parameters.AddWithValue("$column", candidate.FirstColumn);
        command.Parameters.AddWithValue("$start", candidate.HeaderStart); command.Parameters.AddWithValue("$end", candidate.HeaderEnd);
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var normalized = reader.GetString(0).Trim().ToUpperInvariant().Replace(".", string.Empty).Replace(" ", string.Empty);
            if (normalized is "NO" or "NO#" or "S/N" or "STT" or "INDEX" or "SEQ") return true;
        }
        return false;
    }

    private static bool HasNearbyPrimaryCandidate(CapturedCandidate candidate, IReadOnlyList<CapturedCandidate> candidates) => candidates.Any(other =>
        other != candidate && other.Sheet == candidate.Sheet && other.FirstColumn > candidate.LastColumn &&
        other.LastColumn - other.FirstColumn + 1 >= 2 && other.FirstColumn - candidate.LastColumn <= 8 &&
        Math.Max(0, Math.Min(candidate.End, other.End) - Math.Max(candidate.Start, other.Start) + 1) >=
        0.8d * Math.Min(candidate.End - candidate.Start + 1, other.End - other.Start + 1));

    private static List<double> ReadDirectNumericValues(SqliteConnection connection, CapturedCandidate candidate)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT numeric_value, value_text FROM numeric_cells WHERE sheet_id=$sheet AND column_index=$column AND row_index BETWEEN $start AND $end ORDER BY row_index;";
        command.Parameters.AddWithValue("$sheet", candidate.Sheet); command.Parameters.AddWithValue("$column", candidate.FirstColumn);
        command.Parameters.AddWithValue("$start", candidate.Start); command.Parameters.AddWithValue("$end", candidate.End);
        using var reader = command.ExecuteReader(); var values = new List<double>();
        while (reader.Read())
        {
            if (!reader.IsDBNull(0)) values.Add(reader.GetDouble(0));
            else if (double.TryParse(reader.GetString(1), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var value)) values.Add(value);
        }
        return values;
    }

    private static bool IsSequentialIndexStep(double previous, double current) =>
        Math.Abs(current - previous - 1d) < 0.000001d || current is 0d or 1d;

    private static string SourceGridTable(SqliteConnection connection, CapturedCandidate section)
    {
        const int maxRenderedRows = 250;
        const int maxRenderedColumns = 32;
        var dataStart = Math.Max(section.Start, section.HeaderEnd + 1);
        var renderedEnd = Math.Min(section.End, dataStart + maxRenderedRows - 1);
        var renderedLastColumn = Math.Min(section.LastColumn, section.FirstColumn + maxRenderedColumns - 1);
        var texts = ReadTextCells(connection, section.Sheet, section.HeaderStart, renderedEnd);
        var values = ReadNumericCells(connection, section.Sheet, dataStart, renderedEnd);
        var dates = ReadDateCells(connection, section.Sheet, dataStart, renderedEnd);
        var headers = new List<string>();
        for (var column = section.FirstColumn; column <= renderedLastColumn; column++)
        {
            var parts = Enumerable.Range(section.HeaderStart, Math.Max(1, section.HeaderEnd - section.HeaderStart + 1))
                .Select(row => texts.GetValueOrDefault((row, column)))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            headers.Add(parts.Count == 0 ? $"Column {column}" : string.Join(" / ", parts));
        }
        var mergeMap = ReadMergeMap(connection, section.Sheet, dataStart, renderedEnd, section.FirstColumn, renderedLastColumn);
        var rows = new List<(int Row, string[] Cells)>();
        for (var row = dataStart; row <= renderedEnd; row++)
        {
            var cells = new string[headers.Count];
            // Headers and cells must share the same capped column range.  A wide
            // captured table may contain more columns than we render (32), and
            // indexing its uncapped columns into the header/cell arrays would
            // abort the entire batch before the render index is written.
            for (var column = section.FirstColumn; column <= renderedLastColumn; column++)
            {
                var header = headers[column - section.FirstColumn]; var value = texts.GetValueOrDefault((row, column)) ?? dates.GetValueOrDefault((row, column)) ?? values.GetValueOrDefault((row, column));
                cells[column - section.FirstColumn] = FormatSourceValue(value, header);
            }
            rows.Add((row, cells));
        }
        var html = new StringBuilder("<table class='source-table'><thead><tr>");
        foreach (var header in headers) html.Append($"<th>{Escape(header)}</th>");
        html.Append("</tr></thead><tbody>");
        foreach (var row in rows)
        {
            html.Append("<tr>");
            for (var index = 0; index < row.Cells.Length; index++)
            {
                var column = section.FirstColumn + index;
                if (mergeMap.TryGetValue((row.Row, column), out var merge) && (merge.Top != row.Row || merge.Left != column)) continue;
                var cellClass = headers[index].Contains("NG rate", StringComparison.OrdinalIgnoreCase) ? "rate" : string.Empty;
                var spans = mergeMap.TryGetValue((row.Row, column), out merge)
                    ? $" rowspan='{merge.Bottom - merge.Top + 1}' colspan='{merge.Right - merge.Left + 1}'"
                    : string.Empty;
                html.Append($"<td class='{cellClass}'{spans}>{Escape(row.Cells[index])}</td>");
            }
            html.Append("</tr>");
        }
        html.Append("</tbody></table>");
        if (renderedEnd < section.End || renderedLastColumn < section.LastColumn)
            html.Append("<p class='notice'>표가 커서 처음 250행·32열만 표시합니다. 전체 원본은 Excel 보기에서 확인하세요.</p>");
        return html.ToString();
    }

    // Context grid: candidates identify the numeric island, but the visible
    // surface retains its nearby title, header band, label gutter and original
    // fully-contained merge geometry.  This is shared by every catalog profile.
    private static string SourceGridTable(SqliteConnection connection, CapturedCandidate section, CapturedTableContext context)
    {
        var texts = CapturedTableContextResolver.ReadText(connection, section.Sheet, context.TopRow, context.BottomRow);
        var values = CapturedTableContextResolver.ReadNumeric(connection, section.Sheet, context.TopRow, context.BottomRow);
        var dates = CapturedTableContextResolver.ReadDates(connection, section.Sheet, context.TopRow, context.BottomRow);
        var derivedRowsSuppressed = Enumerable.Range(context.DataStartRow, Math.Max(0, context.BottomRow - context.DataStartRow + 1))
            .Any(row => IsDerivedPercentageAuxiliaryRow(row, section, context, texts, values));
        var visibleRows = CompactGridRows(section, context, texts, values, dates);
        var visibleColumns = CompactGridColumns(section, context, visibleRows, texts, values, dates);
        var html = new StringBuilder("<table class='source-table source-grid'><tbody>");
        foreach (var row in visibleRows)
        {
            html.Append("<tr>");
            var rowColumns = row < context.DataStartRow
                ? HeaderColumnsThroughLastContent(row, visibleColumns, context, texts, values, dates)
                : visibleColumns;
            foreach (var column in rowColumns)
            {
                var merge = CapturedTableContextResolver.MergeAt(context.MergeRanges, row, column);
                var mergeRows = merge is { } range
                    ? visibleRows.Where(value => value >= range.Top && value <= range.Bottom).ToList()
                    : [row];
                var mergeColumns = merge is { } span
                    ? visibleColumns.Where(value => value >= span.Left && value <= span.Right).ToList()
                    : [column];
                var displayRow = mergeRows[0];
                var displayColumn = mergeColumns[0];
                if (displayRow != row || displayColumn != column) continue;
                var anchorRow = merge?.Top ?? row;
                var anchorColumn = merge?.Left ?? column;
                var rowSpan = mergeRows.Count;
                var columnSpan = mergeColumns.Count;
                var header = HeaderForGridColumn(section, context, texts, column);
                var raw = CapturedTableContextResolver.EffectiveValue(texts, dates, values, context.MergeRanges, anchorRow, anchorColumn);
                var isHeader = row < context.DataStartRow;
                var isLabel = !isHeader && column < section.FirstColumn;
                var element = isHeader || isLabel ? "th" : "td";
                var classes = string.Join(" ", new[]
                {
                    isHeader ? "grid-header" : string.Empty,
                    isLabel ? "grid-label" : string.Empty,
                    header.Contains("RATE", StringComparison.OrdinalIgnoreCase) ? "rate" : string.Empty
                }.Where(value => value.Length > 0));
                var spans = rowSpan > 1 || columnSpan > 1 ? $" rowspan='{rowSpan}' colspan='{columnSpan}'" : string.Empty;
                html.Append($"<{element} class='{classes}'{spans}>{Escape(FormatGridValue(raw, header))}</{element}>");
            }
            html.Append("</tr>");
        }
        html.Append("</tbody></table>");
        if (derivedRowsSuppressed)
            html.Append("<p class='notice'>Auxiliary percentage rows immediately below their main numeric rows are hidden to preserve the primary source table.</p>");
        return html.ToString();
    }

    // Header/title bands may deliberately stop before the data grid's last
    // metric column. Do not emit a tail of one empty <th> per data column;
    // interior blanks still remain when needed to align a later header cell.
    private static IReadOnlyList<int> HeaderColumnsThroughLastContent(int row, IReadOnlyList<int> visibleColumns, CapturedTableContext context,
        IReadOnlyDictionary<(int Row, int Column), string> texts, IReadOnlyDictionary<(int Row, int Column), string> values,
        IReadOnlyDictionary<(int Row, int Column), string> dates)
    {
        for (var index = visibleColumns.Count - 1; index >= 0; index--)
            if (CellHasContent(row, visibleColumns[index], context, texts, values, dates)) return visibleColumns.Take(index + 1).ToList();
        return [];
    }

    // The rendered surface is intentionally smaller than the capture rectangle:
    // blank spacer axes add no source meaning.  A blank cell within a retained
    // data row/column stays intact, while a wholly blank row/column is removed.
    // Non-empty merged cells keep every covered axis so their HTML span remains
    // a faithful compact representation rather than repeated empty slots.
    private static List<int> CompactGridRows(CapturedCandidate section, CapturedTableContext context,
        IReadOnlyDictionary<(int Row, int Column), string> texts, IReadOnlyDictionary<(int Row, int Column), string> values,
        IReadOnlyDictionary<(int Row, int Column), string> dates)
    {
        var rows = new List<int>();
        for (var row = context.TopRow; row <= context.BottomRow; row++)
        {
            if (row >= context.DataStartRow && IsDerivedPercentageAuxiliaryRow(row, section, context, texts, values)) continue;
            if (Enumerable.Range(context.LeftColumn, context.RightColumn - context.LeftColumn + 1)
                .Any(column => CellHasContent(row, column, context, texts, values, dates))) rows.Add(row);
        }
        return rows;
    }

    private static List<int> CompactGridColumns(CapturedCandidate section, CapturedTableContext context, IReadOnlyList<int> visibleRows,
        IReadOnlyDictionary<(int Row, int Column), string> texts, IReadOnlyDictionary<(int Row, int Column), string> values,
        IReadOnlyDictionary<(int Row, int Column), string> dates)
    {
        var dataRows = visibleRows.Where(row => row >= context.DataStartRow).ToList();
        var lastDataColumn = 0;
        for (var column = section.FirstColumn; column <= context.RightColumn; column++)
            if (dataRows.Any(row => CellHasContent(row, column, context, texts, values, dates))) lastDataColumn = column;

        // A candidate normally contains numeric data. For a structural edge case
        // with no visible data cells, retain the actual header extent but do not
        // let an unrelated title merge force a right-side blank band.
        if (lastDataColumn == 0)
            for (var column = section.FirstColumn; column <= context.RightColumn; column++)
                if (Enumerable.Range(section.HeaderStart, Math.Max(1, section.HeaderEnd - section.HeaderStart + 1))
                    .Any(row => CellHasContent(row, column, context, texts, values, dates))) lastDataColumn = column;

        var right = Math.Max(section.FirstColumn, lastDataColumn);
        var columns = new List<int>();
        for (var column = context.LeftColumn; column <= right; column++)
            if (visibleRows.Any(row => CellHasContent(row, column, context, texts, values, dates))) columns.Add(column);
        return columns;
    }

    private static bool CellHasContent(int row, int column, CapturedTableContext context,
        IReadOnlyDictionary<(int Row, int Column), string> texts, IReadOnlyDictionary<(int Row, int Column), string> values,
        IReadOnlyDictionary<(int Row, int Column), string> dates) =>
        !string.IsNullOrWhiteSpace(CapturedTableContextResolver.EffectiveValue(texts, dates, values, context.MergeRanges, row, column));

    private static string HeaderForGridColumn(CapturedCandidate section, CapturedTableContext context, IReadOnlyDictionary<(int Row, int Column), string> texts, int column)
    {
        var parts = new List<string>();
        for (var row = section.HeaderStart; row <= section.HeaderEnd; row++)
        {
            var value = CapturedTableContextResolver.EffectiveValue(texts, EmptyCells, EmptyCells, context.MergeRanges, row, column).Trim();
            if (value.Length > 0 && !parts.Contains(value, StringComparer.OrdinalIgnoreCase)) parts.Add(value);
        }
        return string.Join(" / ", parts);
    }

    private static string FormatGridValue(string value, string header)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        if (header.Contains("RATE", StringComparison.OrdinalIgnoreCase) && TryReadNumeric(value, out var rate))
            return $"{rate * 100d:F2}%";
        return value;
    }

    private static bool IsDerivedPercentageAuxiliaryRow(int row, CapturedCandidate section, CapturedTableContext context,
        IReadOnlyDictionary<(int Row, int Column), string> texts, IReadOnlyDictionary<(int Row, int Column), string> values)
    {
        if (row <= context.DataStartRow) return false;
        var ownText = Enumerable.Range(context.LeftColumn, context.RightColumn - context.LeftColumn + 1)
            .Any(column => texts.TryGetValue((row, column), out var value) && !string.IsNullOrWhiteSpace(value));
        if (ownText) return false;
        var current = NumericRowValues(row, section, context, values);
        var previous = NumericRowValues(row - 1, section, context, values);
        return current.Count >= 2 && current.All(value => value is >= 0 and <= 1) &&
               previous.Count >= 2 && previous.Any(value => value > 1);
    }

    private static List<double> NumericRowValues(int row, CapturedCandidate section, CapturedTableContext context, IReadOnlyDictionary<(int Row, int Column), string> values)
    {
        var result = new List<double>(); var seen = new HashSet<(int Row, int Column)>();
        for (var column = section.FirstColumn; column <= context.RightColumn; column++)
        {
            var merge = CapturedTableContextResolver.MergeAt(context.MergeRanges, row, column);
            var anchor = merge is { } range ? (range.Top, range.Left) : (row, column);
            if (!seen.Add(anchor) || !values.TryGetValue(anchor, out var raw) || !TryReadNumeric(raw, out var value)) continue;
            result.Add(value);
        }
        return result;
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

    private static Dictionary<(int Row, int Column), MergeRange> ReadMergeMap(SqliteConnection connection, long sheetId, int startRow, int endRow, int startColumn, int endColumn)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT range_ref FROM captured_merge_ranges WHERE sheet_id=$sheet;";
        command.Parameters.AddWithValue("$sheet", sheetId);
        using var reader = command.ExecuteReader();
        var map = new Dictionary<(int, int), MergeRange>();
        while (reader.Read())
        {
            if (!TryParseRange(reader.GetString(0), out var range) || range.Bottom < startRow || range.Top > endRow || range.Right < startColumn || range.Left > endColumn) continue;
            if (range.Top < startRow || range.Left < startColumn || range.Bottom > endRow || range.Right > endColumn) continue;
            for (var row = range.Top; row <= range.Bottom; row++)
                for (var column = range.Left; column <= range.Right; column++) map[(row, column)] = range;
        }
        return map;
    }

    private static bool TryParseRange(string value, out MergeRange range)
    {
        range = default;
        var parts = value.Split(':', 2);
        if (parts.Length != 2 || !TryParseCell(parts[0], out var top, out var left) || !TryParseCell(parts[1], out var bottom, out var right)) return false;
        range = new MergeRange(Math.Min(top, bottom), Math.Min(left, right), Math.Max(top, bottom), Math.Max(left, right));
        return true;
    }

    private static bool TryParseCell(string value, out int row, out int column)
    {
        row = 0; column = 0;
        var index = 0;
        while (index < value.Length && char.IsLetter(value[index])) { column = column * 26 + char.ToUpperInvariant(value[index]) - 'A' + 1; index++; }
        return index > 0 && index < value.Length && int.TryParse(value[index..], out row) && row > 0;
    }

    private static string ColumnName(int column)
    {
        var value = new StringBuilder();
        while (column > 0) { column--; value.Insert(0, (char)('A' + column % 26)); column /= 26; }
        return value.ToString();
    }

    private static Dictionary<(int Row, int Column), string> ReadTextCells(SqliteConnection connection, long sheetId, int start, int end) => ReadCells(connection, "SELECT row_index, column_index, text_value FROM captured_text_cells WHERE sheet_id=$sheet AND row_index BETWEEN $start AND $end;", sheetId, start, end);
    private static Dictionary<(int Row, int Column), string> ReadNumericCells(SqliteConnection connection, long sheetId, int start, int end) => ReadCells(connection, "SELECT row_index, column_index, value_text FROM numeric_cells WHERE sheet_id=$sheet AND row_index BETWEEN $start AND $end;", sheetId, start, end);
    private static Dictionary<(int Row, int Column), string> ReadDateCells(SqliteConnection connection, long sheetId, int start, int end) => ReadCells(connection, "SELECT row_index, column_index, date_value FROM date_cells WHERE sheet_id=$sheet AND row_index BETWEEN $start AND $end;", sheetId, start, end);
    private static Dictionary<(int Row, int Column), string> ReadCells(SqliteConnection connection, string sql, long sheetId, int start, int end)
    {
        using var command = connection.CreateCommand(); command.CommandText = sql; command.Parameters.AddWithValue("$sheet", sheetId); command.Parameters.AddWithValue("$start", start); command.Parameters.AddWithValue("$end", end); using var reader = command.ExecuteReader(); var values = new Dictionary<(int, int), string>();
        while (reader.Read()) values[(reader.GetInt32(0), reader.GetInt32(1))] = reader.GetString(2); return values;
    }
    private static string TextAt(SqliteConnection connection, long sheetId, int row, int column)
    {
        using var command = connection.CreateCommand(); command.CommandText = "SELECT text_value FROM captured_text_cells WHERE sheet_id=$sheet AND row_index=$row AND column_index=$column;"; command.Parameters.AddWithValue("$sheet", sheetId); command.Parameters.AddWithValue("$row", row); command.Parameters.AddWithValue("$column", column); return command.ExecuteScalar() as string ?? string.Empty;
    }
    private static string FindSectionTitle(SqliteConnection connection, long sheetId, int headerStart)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT text_value FROM captured_text_cells WHERE sheet_id=$sheet AND row_index BETWEEN $first AND $last AND length(trim(text_value)) >= 4 ORDER BY row_index DESC, length(text_value) DESC LIMIT 1;";
        command.Parameters.AddWithValue("$sheet", sheetId);
        command.Parameters.AddWithValue("$first", Math.Max(1, headerStart - 3));
        command.Parameters.AddWithValue("$last", Math.Max(1, headerStart - 1));
        return command.ExecuteScalar() as string ?? string.Empty;
    }
    private static string FormatSourceValue(string? value, string header)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        if (header.Contains("NG rate", StringComparison.OrdinalIgnoreCase) &&
            double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var rate) &&
            rate >= 0 && rate <= 1)
            return $"{rate * 1_000_000:N0} ppm ({rate:P1})";
        return value;
    }
    private static string Page(string path, CatalogGroup group, GroupProfile profile, string status, string body)
    {
        var source = Escape(Path.GetFileName(path));
        var summary = $"<section class='summary-card'><div><strong>Form group</strong><code>{Escape(group.Id)}</code></div><div><strong>Renderer</strong><code>{Escape(profile.Version)}</code></div><div><strong>State</strong><span class='status {(status == "PENDING_PROFILE" ? "review" : "good")}'>{Escape(status)}</span></div></section>";
        const string css = """
            html,body{width:100%;margin:0;padding:0;background:#f4f6fa;color:#17202e;font-family:Segoe UI,Malgun Gothic,Arial,sans-serif;line-height:1.45}body{overflow-x:hidden}header{padding:18px 22px 14px;border-bottom:1px solid #d8e0ea;background:#ffffff}h1{margin:0;font-size:23px;word-wrap:break-word}header p{margin:5px 0 0;color:#667085;font-size:12px;word-wrap:break-word}main{width:96%;max-width:1500px;margin:0 auto;padding:16px 0 30px}.summary-card{width:100%;margin-bottom:16px;overflow:hidden}.summary-card>div{float:left;width:31%;min-height:72px;margin-right:2%;padding:12px;border:1px solid #d8e0ea;background:#ffffff}.summary-card>div:last-child{margin-right:0}.summary-card strong,.summary-card code{display:block;word-wrap:break-word}.summary-card strong{font-size:11px;color:#667085;margin-bottom:5px}.summary-card:after{content:'';display:block;clear:both}.table-wrap{width:100%;overflow-x:auto;overflow-y:hidden;border:1px solid #d8e0ea;background:#ffffff}table{width:100%;border-collapse:collapse;font-size:12px}.source-table{min-width:900px}caption{padding:9px 10px;border-bottom:1px solid #d8e0ea;color:#344054;background:#f8fafc;text-align:left;font-weight:bold}th,td{padding:8px 9px;border-right:1px solid #d8e0ea;border-bottom:1px solid #d8e0ea;vertical-align:top;word-wrap:break-word}th:last-child,td:last-child{border-right:0}thead th{color:#475467;background:#eef2f6;text-align:center}tbody tr:nth-child(even){background:#f8fbff}.source-grid .grid-header{background:#eef2f6;color:#475467;text-align:center;font-weight:700}.source-grid .grid-label{background:#f7f9fc;font-weight:600}.num,.rate{text-align:right;white-space:nowrap;font-weight:bold}.status{display:inline-block;padding:3px 7px;font-size:10px;font-weight:bold;white-space:nowrap}.good{color:#067647;background:#ecfdf3}.review{color:#b54708;background:#fffaeb}.neutral{color:#175cd3;background:#eff4ff}.notice{padding:12px;background:#fffaeb;border:1px solid #f1dfad}.section-card{margin:0 0 14px;border:1px solid #d8e0ea;background:#ffffff}.section-title{padding:11px 14px;font-weight:bold;color:#183b67;background:#edf3fb;word-wrap:break-word}.section-card .table-wrap{width:calc(100% - 24px);margin:12px}@media(max-width:760px){main{width:98%}.summary-card>div{float:none;width:100%;margin:0 0 8px}.source-table{min-width:720px}}
            """;
        return $"<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>{source}</title><style>{css}</style></head><body><header><h1>{source}</h1><p>AI group renderer · {Escape(group.Id)} · {Escape(profile.Key)}</p></header><main>{summary}{body}</main></body></html>";
    }
    private static string Section(string title, string table) => $"<section><h2>{Escape(title)}</h2>{table}</section>";
    private static string Table(SqliteConnection connection, long workbookId, string sql, IReadOnlyList<string> headers)
    {
        using var command = connection.CreateCommand(); command.CommandText = sql; command.Parameters.AddWithValue("$id", workbookId);
        using var reader = command.ExecuteReader(); var rows = new StringBuilder(); var count = 0;
        while (reader.Read()) { count++; rows.Append("<tr>"); for (var index = 0; index < reader.FieldCount; index++) rows.Append($"<td class='{CellClass(headers[index])}'>{Escape(FormatCell(reader.IsDBNull(index) ? null : reader.GetValue(index), headers[index]))}</td>"); rows.Append("</tr>"); }
        var head = string.Concat(headers.Select(header => $"<th>{Escape(header)}</th>"));
        return count == 0 ? "<p class='notice'>No extracted facts.</p>" : $"<div class='table-wrap'><table><thead><tr>{head}</tr></thead><tbody>{rows}</tbody></table></div>";
    }
    private static string RepeatedBlocksTable(SqliteConnection connection, long workbookId)
    {
        using var command = connection.CreateCommand(); command.CommandText = "SELECT block_label, condition_label, input_value, total_ng_value, computed_ng_rate, fact_status FROM repeated_defect_block_facts WHERE workbook_id=$id ORDER BY table_id, block_key, row_index;"; command.Parameters.AddWithValue("$id", workbookId);
        using var reader = command.ExecuteReader(); var groups = new Dictionary<string, List<object?[]>>(StringComparer.Ordinal);
        while (reader.Read()) { var block = reader.IsDBNull(0) ? "Unlabeled block" : reader.GetString(0); if (!groups.TryGetValue(block, out var rows)) groups[block] = rows = []; rows.Add(Enumerable.Range(1, reader.FieldCount - 1).Select(index => reader.IsDBNull(index) ? null : reader.GetValue(index)).ToArray()); }
        var headers = new[] { "Condition", "Input", "Total NG", "NG rate", "Status" };
        return string.Concat(groups.Select(group => $"<div class='block'><h3>{Escape(group.Key)} · {group.Value.Count:N0} rows</h3>{RenderedRows(headers, group.Value)}</div>"));
    }
    private static string RenderedRows(IReadOnlyList<string> headers, IReadOnlyList<object?[]> values)
    {
        if (values.Count == 0) return "<p class='notice'>No extracted facts.</p>";
        var head = string.Concat(headers.Select(header => $"<th>{Escape(header)}</th>")); var rows = new StringBuilder();
        foreach (var row in values) { rows.Append("<tr>"); for (var index = 0; index < row.Length; index++) rows.Append($"<td class='{CellClass(headers[index])}'>{Escape(FormatCell(row[index], headers[index]))}</td>"); rows.Append("</tr>"); }
        return $"<div class='table-wrap'><table><thead><tr>{head}</tr></thead><tbody>{rows}</tbody></table></div>";
    }
    private static string CellClass(string header) => header == "NG rate" ? "rate" : header == "Status" ? "status" : string.Empty;
    private static string FormatCell(object? value, string header)
    {
        if (value is null) return string.Empty;
        if (header == "NG rate" && double.TryParse(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var rate)) return $"{rate * 1_000_000:N0} ppm ({rate:P3})";
        if (header == "Status" && string.Equals(Convert.ToString(value), "NEEDS_REVIEW", StringComparison.Ordinal)) return "Review needed";
        return Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
    }
    private static string StableId(string path) => Convert.ToHexString(SHA256.HashData(new UTF8Encoding(false).GetBytes(path.ToLowerInvariant()))).ToLowerInvariant()[..20];
    private static string Escape(string value) => System.Net.WebUtility.HtmlEncode(value);

    private sealed record GroupProfile(string Key, string Version, IReadOnlyList<string> RequiredComponents, bool IsImplemented);
    private sealed record ComponentDescriptor(string Key, string Title);
    private sealed record CapturedCandidate(long Sheet, string SheetName, int Start, int End, int FirstColumn, int LastColumn, int HeaderStart, int HeaderEnd, string ComponentKey, int NumericCellCount);
    private sealed record CapturedSection(string ComponentKey, string Title, string Table, string Evidence, int NumericCellCount, bool IsSupportOnly);
    private sealed record SourceSection(string ComponentKey, string Title, string Table, string Evidence, int NumericCellCount);
    private sealed record ComponentRender(
        IReadOnlyList<DashboardSection> Sections,
        IReadOnlyList<SourceSection> SourceSections,
        IReadOnlyDictionary<string, int> ComponentCounts,
        IReadOnlyList<string> MissingComponents,
        IReadOnlyList<string> UnexpectedComponents)
    {
        internal bool IsContractSatisfied => MissingComponents.Count == 0 && UnexpectedComponents.Count == 0;
        internal string Status => IsContractSatisfied ? "COMPONENTS_RENDERED" : "CONTRACT_MISMATCH";
    }
    private sealed record CatalogGroup(
        string Id,
        string Name,
        string RendererKey,
        IReadOnlyList<string> ComponentRecipe,
        IReadOnlyList<string> ImportantData,
        string ExtractionRule,
        string HtmlRule,
        IReadOnlyList<string> OpenQuestions);
    private readonly record struct MergeRange(int Top, int Left, int Bottom, int Right);
    private sealed record Catalog(IReadOnlyDictionary<string, CatalogGroup> Assignments, IReadOnlyDictionary<string, CatalogGroup> Groups)
    {
        private static readonly IReadOnlyDictionary<string, string> RendererAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["needs-review"] = "renderNeedsReview",
        };

        public static Catalog Load(string path)
        {
            using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8)); var root = document.RootElement;
            var groups = root.GetProperty("groups").EnumerateArray()
                .Select(value =>
                {
                    var rendererKey = value.GetProperty("rendererKey").GetString()!;
                    var recipe = value.GetProperty("componentRecipe").EnumerateArray()
                        .Select(component => NormalizeComponentKey(component.GetString() ?? string.Empty))
                        .Where(component => !string.IsNullOrWhiteSpace(component))
                        .Distinct(StringComparer.Ordinal)
                        .ToList();
                    if (!IsSupportedComponentRecipe(recipe))
                        throw new InvalidOperationException($"Catalog group '{value.GetProperty("id").GetString()}' has an unsupported component recipe.");
                    return new CatalogGroup(
                        value.GetProperty("id").GetString()!,
                        value.GetProperty("name").GetString() ?? value.GetProperty("id").GetString()!,
                        RendererAliases.GetValueOrDefault(rendererKey, rendererKey),
                        recipe,
                        value.GetProperty("importantData").EnumerateArray().Select(item => item.GetString() ?? string.Empty).Where(item => item.Length > 0).ToList(),
                        value.GetProperty("extractionRule").GetString() ?? string.Empty,
                        value.GetProperty("htmlRule").GetString() ?? string.Empty,
                        value.GetProperty("openQuestions").EnumerateArray().Select(item => item.GetString() ?? string.Empty).Where(item => item.Length > 0).ToList());
                })
                .ToDictionary(group => group.Id, StringComparer.Ordinal);
            var assignments = root.GetProperty("fileAssignments").EnumerateArray().ToDictionary(value => value.GetProperty("relativePath").GetString()!, value => groups[value.GetProperty("groupId").GetString()!], StringComparer.OrdinalIgnoreCase);
            return new Catalog(assignments, groups);
        }
    }
}

internal sealed record GroupCatalogRendererRequest(string ServiceDirectory, string StructureBatchId);
internal sealed record GroupCatalogRendererRunResult(string BatchDirectory, bool CatalogFound, GroupCatalogRendererSummary? Summary);
internal sealed record GroupCatalogRendererSummary(string SchemaVersion, string RendererVersion, string GeneratedAt, int WorkbookCount, IReadOnlyDictionary<string, int> StatusCounts, string Html, string Json);
internal sealed record GroupRenderRow(string RelativePath, string GroupId, string RendererKey, string ProfileVersion, string Status, string ReportPath, IReadOnlyDictionary<string, int> ComponentCounts);
