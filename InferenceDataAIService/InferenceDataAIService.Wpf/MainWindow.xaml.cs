using System.Collections.ObjectModel;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Data.Sqlite;
using Forms = System.Windows.Forms;
using WpfDataFormats = System.Windows.DataFormats;
using WpfDragDropEffects = System.Windows.DragDropEffects;
using WpfDragEventArgs = System.Windows.DragEventArgs;

namespace InferenceDataAIService.Wpf;

public partial class MainWindow : Window
{
    private const int HomeWorkspaceIndex = 8;
    private const int MaxPendingVisualLogLines = 4_000;
    private const int MaxVisualLogCharacters = 160_000;
    private const int RetainedVisualLogCharacters = 120_000;
    private static readonly UTF8Encoding Utf8WithBom = new(encoderShouldEmitUTF8Identifier: true);
    private static readonly HashSet<string> ExcelExtensions = new(
        [".xlsx", ".xlsm", ".xlsb", ".xls"],
        StringComparer.OrdinalIgnoreCase);

    private readonly ObservableCollection<ExcelItem> _items = [];
    private readonly ObservableCollection<LearnedResultBatch> _learnedBatches = [];
    private readonly ObservableCollection<LearnedReportRow> _learnedReports = [];
    private readonly string _discoveredServiceDirectory;
    private AppPathSettings _pathSettings = null!;
    private string _serviceDirectory = string.Empty;
    private string _databasePath = string.Empty;
    private string _runLogDirectory = string.Empty;
    private readonly StartupAnalysisOptions _startupAnalysis;
    private readonly ConcurrentQueue<string> _pendingLogs = new();
    private readonly DispatcherTimer _logFlushTimer;
    private MemoryStream? _activeHtmlStream;
    private int _pendingVisualLogLineCount;
    private int _droppedVisualLogLineCount;
    private bool _startupAnalysisStarted;
    private bool _numericReviewRunning;
    private bool _loadingLearnedResults;
    private bool _isInsightWide;
    private bool _isInsightSuppressed;
    private GridLength _normalWorkspacePaneWidth = new(46, GridUnitType.Star);
    private GridLength _normalInsightPaneWidth = new(54, GridUnitType.Star);

    internal MainWindow(StartupAnalysisOptions? startupAnalysis = null)
    {
        InitializeComponent();
        _startupAnalysis = startupAnalysis ?? new StartupAnalysisOptions([], false);
        WorkspaceTabs.SelectedIndex = _startupAnalysis.HasStartupWork
            ? 0
            : HomeWorkspaceIndex;
        _discoveredServiceDirectory = FindServiceDirectory();
        _pathSettings = AppPathSettingsStore.Load(
            _discoveredServiceDirectory);
        try
        {
            _pathSettings.Validate();
        }
        catch
        {
            _pathSettings = AppPathSettings.CreateDefaults(
                _discoveredServiceDirectory);
        }
        ApplyPathSettings(_pathSettings);
        FilesGrid.ItemsSource = _items;
        LearnedBatchSelector.ItemsSource = _learnedBatches;
        LearnedReportsGrid.ItemsSource = _learnedReports;
        InitializeCanonicalEvidenceUi();
        LoadPathSettingsUi();
        _logFlushTimer = new DispatcherTimer(DispatcherPriority.Background) { Interval = TimeSpan.FromMilliseconds(120) };
        _logFlushTimer.Tick += (_, _) => FlushPendingLogs();
        _logFlushTimer.Start();
        Closed += (_, _) =>
        {
            _activeHtmlStream?.Dispose();
            _activeHtmlStream = null;
        };
        Log($"서비스 폴더: {_serviceDirectory}");
        Log($"설정 파일: {AppPathSettingsStore.SettingsFilePath}");
        LoadStoredFiles();
        LoadLearnedResults();
        ApplyWorkspaceSelection(WorkspaceTabs.SelectedIndex);
        Loaded += (_, _) => Dispatcher.BeginInvoke(
            DispatcherPriority.Loaded,
            new Action(ShowWorkspaceHome));
    }

    private void ApplyPathSettings(AppPathSettings settings)
    {
        _pathSettings = settings;
        _serviceDirectory = settings.ServiceDirectory;
        _databasePath = settings.PrimaryDatabasePath;
        _runLogDirectory = settings.RunLogDirectory;
        AppRuntimePaths.Apply(settings);
    }

    private void ShowWorkspaceHome()
    {
        if (!_startupAnalysis.HasStartupWork)
        {
            SetInsightWideMode(false);
            WorkspaceTabs.SelectedIndex = HomeWorkspaceIndex;
        }
    }

    private void WorkspaceNavigation_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: string rawIndex } ||
            !int.TryParse(rawIndex, out var index) ||
            index < 0 ||
            index >= WorkspaceTabs.Items.Count)
        {
            return;
        }

        SetInsightWideMode(false);
        WorkspaceTabs.SelectedIndex = index;
    }

    private void ToggleInsightWide_Click(object sender, RoutedEventArgs e) =>
        SetInsightWideMode(!_isInsightWide);

    private void SetInsightPaneSuppressed(bool suppressed)
    {
        if (_isInsightSuppressed == suppressed) return;
        if (suppressed)
        {
            if (_isInsightWide) SetInsightWideMode(false);
            _normalWorkspacePaneWidth = WorkspacePaneColumn.Width;
            _normalInsightPaneWidth = InsightPaneColumn.Width;
            _isInsightSuppressed = true;
            WorkspaceInsightSplitter.Visibility = Visibility.Collapsed;
            InsightPane.Visibility = Visibility.Collapsed;
            WorkspacePaneColumn.MinWidth = 0;
            InsightPaneColumn.MinWidth = 0;
            WorkspacePaneColumn.Width = new GridLength(
                1,
                GridUnitType.Star);
            WorkspaceSplitterColumn.Width = new GridLength(0);
            InsightPaneColumn.Width = new GridLength(0);
            return;
        }

        _isInsightSuppressed = false;
        InsightPane.Visibility = Visibility.Visible;
        WorkspaceInsightSplitter.Visibility = Visibility.Visible;
        WorkspacePaneColumn.MinWidth = 520;
        InsightPaneColumn.MinWidth = 430;
        WorkspacePaneColumn.Width = _normalWorkspacePaneWidth;
        WorkspaceSplitterColumn.Width = new GridLength(5);
        InsightPaneColumn.Width = _normalInsightPaneWidth;
    }

    private void SetInsightWideMode(bool isWide)
    {
        if (_isInsightSuppressed) return;
        if (_isInsightWide == isWide) return;
        _isInsightWide = isWide;
        if (isWide)
        {
            _normalWorkspacePaneWidth = WorkspacePaneColumn.Width;
            _normalInsightPaneWidth = InsightPaneColumn.Width;
            WorkspaceTabs.Visibility = Visibility.Collapsed;
            WorkspaceInsightSplitter.Visibility = Visibility.Collapsed;
            WorkspacePaneColumn.MinWidth = 0;
            InsightPaneColumn.MinWidth = 0;
            WorkspacePaneColumn.Width = new GridLength(0);
            WorkspaceSplitterColumn.Width = new GridLength(0);
            InsightPaneColumn.Width = new GridLength(1, GridUnitType.Star);
            InsightWideButton.Content = "분할 보기";
            InsightWideButton.ToolTip = "질문·근거 목록과 결과 표를 함께 봅니다.";
            return;
        }

        WorkspacePaneColumn.MinWidth = 520;
        InsightPaneColumn.MinWidth = 430;
        WorkspacePaneColumn.Width = _normalWorkspacePaneWidth;
        WorkspaceSplitterColumn.Width = new GridLength(5);
        InsightPaneColumn.Width = _normalInsightPaneWidth;
        WorkspaceTabs.Visibility = Visibility.Visible;
        WorkspaceInsightSplitter.Visibility = Visibility.Visible;
        InsightWideButton.Content = "넓게 보기";
        InsightWideButton.ToolTip = "왼쪽 작업 영역을 접고 결과 표를 창 너비로 봅니다.";
    }

    private void WorkspaceTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (e.Source == WorkspaceTabs)
        {
            ApplyWorkspaceSelection(WorkspaceTabs.SelectedIndex);
        }
    }

    private void ApplyWorkspaceSelection(int index)
    {
        if (WorkspaceNavigationPanel is null ||
            WorkspacePageTitleText is null ||
            WorkspaceSubtitleText is null ||
            WorkspaceStatusText is null)
        {
            return;
        }

        var (title, subtitle) = index switch
        {
            0 => ("Excel 분석 목록", "파일·폴더를 끌어놓고 단건 또는 배치 분석 결과를 관리합니다."),
            1 => ("저장된 분석", "완료된 배치와 검증 결과를 다시 열어봅니다."),
            2 => ("질문으로 자료 찾기", "DB에서 질문과 관련된 보고서와 원본 근거를 찾습니다."),
            3 => ("원본 값 대조", "DB에 저장된 값과 원본 Excel 셀이 같은지 읽기 전용으로 확인합니다."),
            4 => ("보류 항목 판정", "AI가 확정하지 못한 비교를 사람이 승인·거절·제외합니다."),
            5 => ("개념 사전", "공정·재료·설비 명칭과 별칭을 표준 개념으로 정리합니다."),
            6 => ("신규 Excel 수집", "폴더를 DB와 비교하고 없는 Excel만 로컬 보관함으로 복사합니다."),
            7 => ("DRM Excel 전체 처리", "선택한 Excel을 COM 추출부터 AI 분석·DB 반영까지 처리합니다."),
            8 => ("홈", "Excel 수집·처리와 질문 검색, 품질 관리 작업을 선택합니다."),
            9 => ("설정", "DB와 모든 작업 경로를 저장하고 다음 실행 작업부터 즉시 적용합니다."),
            10 => ("Excel COM 사전 분석", "보관함 Excel을 읽기 전용 COM으로 추출하고 기존 양식과 먼저 비교합니다."),
            11 => ("양식군 검토·등록", "구조가 같은 유사·신규 양식을 대표본 AI 분석과 사람 승인으로 등록합니다."),
            _ => ("Inference Data", "검증된 표 근거를 탐색합니다.")
        };

        WorkspacePageTitleText.Text = title;
        WorkspaceSubtitleText.Text = subtitle;
        WorkspaceStatusText.Text = $"{title} · Ready";
        SetInsightPaneSuppressed(
            index is 6 or 7 or 9 or 10 or 11
            || index == HomeWorkspaceIndex);
        if (index == 2)
            ShowQuestionWorkspaceInsight();

        foreach (var button in WorkspaceNavigationPanel.Children.OfType<System.Windows.Controls.Button>())
        {
            var selected = button.Tag is string rawIndex &&
                           int.TryParse(rawIndex, out var buttonIndex) &&
                           buttonIndex == index;
            button.Background = selected
                ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(54, 54, 54))
                : System.Windows.Media.Brushes.Transparent;
            button.Foreground = selected
                ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(240, 240, 240))
                : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(184, 184, 184));
            button.FontWeight = selected ? FontWeights.SemiBold : FontWeights.Normal;
        }
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            ToggleWindowState();
            return;
        }

        if (e.LeftButton == MouseButtonState.Pressed)
        {
            DragMove();
        }
    }

    private void MinimizeButton_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

    private void MaximizeButton_Click(object sender, RoutedEventArgs e) => ToggleWindowState();

    private void CloseWindowButton_Click(object sender, RoutedEventArgs e) => Close();

    private void MainWindow_StateChanged(object? sender, EventArgs e)
    {
        if (MaximizeGlyphText is not null)
        {
            MaximizeGlyphText.Text = WindowState == WindowState.Maximized ? "\uE923" : "\uE922";
        }
    }

    private void ToggleWindowState() =>
        WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;

    internal void ScheduleStartupAnalysis() =>
        Dispatcher.BeginInvoke(async () =>
        {
            try
            {
                await RunStartupAnalysisAsync();
            }
            catch (Exception exception)
            {
                // Dispatcher async delegates otherwise lose their Task exception
                // without a batch-local diagnostic artifact, leaving the WPF window
                // responsive but apparently stalled.
                WriteStartupFailure(exception);
                SetPipelineRunState("AI grouping failed; see the batch startup failure log.");
                Log($"Startup pipeline failed: {exception.Message}");
            }
        }, DispatcherPriority.ApplicationIdle);

    private void WriteStartupFailure(Exception exception)
    {
        var batchId = _startupAnalysis.AiGroupAnalysisBatchId ?? _startupAnalysis.NumericReviewBatchId ?? _startupAnalysis.ResumeBatchId ?? _startupAnalysis.BatchId;
        var directory = string.IsNullOrWhiteSpace(batchId)
            ? _runLogDirectory
            : Path.Combine(
                _pathSettings.BatchRootDirectory,
                batchId,
                "logs");
        try
        {
            Directory.CreateDirectory(directory);
            File.AppendAllText(Path.Combine(directory, "startup-failure.log"), $"[{DateTimeOffset.UtcNow:O}] {exception}\n", Utf8WithBom);
        }
        catch (Exception loggingException)
        {
            Log($"Could not write startup failure log: {loggingException.Message}");
        }
    }

    private async Task RunStartupAnalysisAsync()
    {
        if (_startupAnalysisStarted || !_startupAnalysis.HasStartupWork) return;
        _startupAnalysisStarted = true;

        if (_startupAnalysis.IsNumericReview)
        {
            await RunStartupNumericReviewAsync();
            return;
        }

        if (_startupAnalysis.IsAiGroupAnalysis)
        {
            await RunAiGroupAnalysisAsync();
            return;
        }

        if (_startupAnalysis.IsBatchScan)
        {
            await RunStartupBatchScanAsync();
            return;
        }

        AddFiles(_startupAnalysis.ExcelPaths);
        var selected = new List<ExcelItem>();
        FilesGrid.SelectedItems.Clear();
        foreach (var path in _startupAnalysis.ExcelPaths)
        {
            var item = _items.FirstOrDefault(candidate => string.Equals(candidate.FullPath, path, StringComparison.OrdinalIgnoreCase));
            if (item is null) continue;
            selected.Add(item);
            FilesGrid.SelectedItems.Add(item);
        }
        if (selected.Count == 0)
        {
            Log("명령줄로 받은 Excel 파일을 목록에서 선택하지 못했습니다.");
            return;
        }

        Log(_startupAnalysis.Force ? "명령줄 강제 재분석을 시작합니다." : "명령줄 분석을 시작합니다.");
        await AnalyzeItemsAsync(selected, forceAiDraft: _startupAnalysis.Force);
    }

    private Task RunAiGroupAnalysisAsync() => RunAiGroupAnalysisAsync(_startupAnalysis.AiGroupAnalysisBatchId!);

    // Retained temporarily as source history for the old coarse semantic-plan path.
    // All call sites use the validated two-level method below instead.
    private async Task RunLegacyAiGroupAnalysisAsync(string batchId)
    {
        SetPipelineRunState("AI 그룹 일괄 분석: 준비 중");
        var batchDirectory = Path.Combine(
            _pathSettings.BatchRootDirectory,
            batchId);
        var classification = Path.Combine(batchDirectory, "classification.csv");
        if (!File.Exists(classification)) throw new FileNotFoundException("Run the structure scan before AI group analysis.", classification);
        var plan = Path.Combine(batchDirectory, "group-plan.json");
        var output = Path.Combine(batchDirectory, "group-catalog.json");
        var summaryPath = Path.Combine(batchDirectory, "summary.json");
        if (File.Exists(plan))
        {
            SetPipelineRunState("AI 그룹 규칙 확인 완료 · 그룹별 HTML 렌더링 중");
            Log("기존 AI 그룹 계획을 이어서 배정·렌더링합니다.");
            LoadBatchPipeline(batchDirectory, aiState: "완료");
            MaterializeFileAssignments(plan, output, classification, summaryPath, batchId);
            using var resumedRenderLog = CreateRunLog($"batch_{batchId}.group-render-resume", Path.Combine(batchDirectory, "logs"));
            var resumedRender = await RunGroupCatalogRendererAsync(resumedRenderLog, batchId);
            resumedRenderLog.FlushAndThrowIfWriteFailed();
            LoadBatchPipeline(batchDirectory, aiState: "완료");
            SetPipelineRunState(RenderCompletionText(resumedRender));
            ShowGroupCatalog(output, batchId);
            return;
        }
        // A restart must never present the previous run's group assignment as live state.
        foreach (var stale in new[] { "group-plan.json", "group-catalog.json", "group-render-index.json", "group-render-index.html" })
            File.Delete(Path.Combine(batchDirectory, stale));
        LoadBatchPipeline(batchDirectory, aiState: "진행 중");
        SetPipelineRunState("AI 그룹 일괄 분석 1/3: 전체 Excel 데이터 캡처 중");
        using (var summary = JsonDocument.Parse(File.ReadAllText(summaryPath, Encoding.UTF8)))
        {
            var counts = summary.RootElement.GetProperty("statusCounts");
            if (counts.TryGetProperty("DEFERRED", out var deferred) && deferred.GetInt32() > 0)
                throw new InvalidOperationException($"AI group analysis requires a complete structure scan. Deferred files: {deferred.GetInt32()}.");
        }
        Log("AI 그룹 분석 전 전체 문서 좌표 인벤토리를 생성합니다.");
        await Task.Run(() => NumericCaptureEngine.Run(
            new NumericCaptureRequest(_serviceDirectory, batchId, Force: true),
            line => QueueLog($"ENGINE: {line}")));
        LoadBatchPipeline(batchDirectory, aiState: "진행 중");
        SetPipelineRunState("AI 그룹 일괄 분석 2/3: AI가 양식 그룹과 규칙을 생성 중");
        var inventory = DocumentInventoryEngine.Write(batchDirectory);
        Log($"AI 문서 인벤토리 생성: {inventory.WorkbookCount}개 파일, {inventory.SignatureCount}개 레이아웃 서명");
        var schema = Path.Combine(_serviceDirectory, "group-plan.schema.json");
        using var semanticDocument = JsonDocument.Parse(File.ReadAllText(inventory.SemanticSummaryPath, Encoding.UTF8));
        var semanticSummary = string.Join("; ", semanticDocument.RootElement.GetProperty("categories").EnumerateArray().Select(category => $"{category.GetProperty("category").GetString()}={category.GetProperty("fileCount").GetInt32()} files"));
        var prompt = $"Return one final JSON plan only. The complete semantic categories are: {semanticSummary}. Each category suffix encodes section-count and ordered table-pattern. Assign every listed category exactly once using selector.semanticCategories; split categories with different section-patterns into separate renderer groups unless they have the same component recipe. Do not use fallback unless a category explicitly says empty or incomplete (none do). Use only these registered renderer keys: acoustic-dashboard-v1, quality-dashboard-v1, measurement-dashboard-v1, function-process-dashboard-v1, tension-dashboard-v1, general-table-dashboard-v1, renderNeedsReview. Do not access files, web, GitHub, shell commands, scripts, collaboration, subagents, or wait tools. Every group needs data components, extraction rules, HTML dashboard rules, and variants.";
        var codex = _pathSettings.CodexExecutable;
        var info = new ProcessStartInfo(codex) { WorkingDirectory = _serviceDirectory, UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true, StandardInputEncoding = Encoding.UTF8, StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8 };
        info.ArgumentList.Add("exec"); info.ArgumentList.Add("--ephemeral"); info.ArgumentList.Add("--sandbox"); info.ArgumentList.Add("read-only");
        var model = Environment.GetEnvironmentVariable("INFERENCE_DATA_AI_CODEX_MODEL");
        info.ArgumentList.Add("--model"); info.ArgumentList.Add(string.IsNullOrWhiteSpace(model) ? "gpt-5.6-sol" : model);
        // Group classification is the decision point for every downstream renderer.
        // Use Sol's deepest configured reasoning level for this batch-wide task.
        info.ArgumentList.Add("--config"); info.ArgumentList.Add("model_reasoning_effort=xhigh");
        info.ArgumentList.Add("--output-schema"); info.ArgumentList.Add(schema); info.ArgumentList.Add("-o"); info.ArgumentList.Add(plan); info.ArgumentList.Add("-");
        Log($"Starting Codex CLI group analysis for {batchId}.");
        using var process = Process.Start(info) ?? throw new InvalidOperationException("Could not start Codex CLI.");
        await process.StandardInput.WriteAsync(prompt);
        process.StandardInput.Close();
        process.OutputDataReceived += (_, value) => { if (value.Data is not null) QueueLog($"CODEX: {value.Data}"); };
        process.ErrorDataReceived += (_, value) => { if (value.Data is not null) QueueLog($"CODEX: {value.Data}"); };
        process.BeginOutputReadLine(); process.BeginErrorReadLine();
        using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(10));
        await process.WaitForExitAsync(timeout.Token);
        if (process.ExitCode != 0) throw new InvalidOperationException($"Codex CLI exited with code {process.ExitCode}.");
        if (!File.Exists(plan)) throw new FileNotFoundException("Codex CLI completed without writing the AI group plan.", plan);
        MaterializeFileAssignments(plan, output, classification, summaryPath, batchId);
        Log($"AI group catalog created: {output}");
        using (var groupRenderLog = CreateRunLog($"batch_{batchId}.group-render", Path.Combine(batchDirectory, "logs")))
        {
            SetPipelineRunState("AI 그룹 일괄 분석 3/3: 그룹별 HTML 렌더링 중");
            var render = await RunGroupCatalogRendererAsync(groupRenderLog, batchId);
            groupRenderLog.FlushAndThrowIfWriteFailed();
            SetPipelineRunState(RenderCompletionText(render));
        }
        LoadBatchPipeline(batchDirectory, aiState: "완료");
        ShowGroupCatalog(output, batchId);
    }

    private async Task RunAiGroupAnalysisAsync(string batchId)
    {
        var batchDirectory = Path.Combine(
            _pathSettings.BatchRootDirectory,
            batchId);
        var classification = Path.Combine(batchDirectory, "classification.csv");
        var output = Path.Combine(batchDirectory, "group-catalog.json");
        using var groupingLog = CreateRunLog($"batch_{batchId}.group-analysis", Path.Combine(batchDirectory, "logs"));
        groupingLog.WriteLine("WPF", $"Starting validated two-level AI grouping for {batchId}.");
        if (!File.Exists(classification))
            throw new FileNotFoundException("Run the structure scan before AI group analysis.", classification);

        if (TwoLevelGroupAnalysisEngine.HasValidatedCatalog(batchDirectory))
        {
            SetPipelineRunState("AI grouping validation is already complete; rendering the validated catalog.");
            Log("A validated two-level catalog already exists; source Excel files will not be recaptured.");
            using var resumeLog = CreateRunLog($"batch_{batchId}.group-render-resume", Path.Combine(batchDirectory, "logs"));
            var resumed = await RunGroupCatalogRendererAsync(resumeLog, batchId);
            resumeLog.FlushAndThrowIfWriteFailed();
            LoadBatchPipeline(batchDirectory, aiState: "complete");
            SetPipelineRunState(RenderCompletionText(resumed));
            ShowGroupCatalog(output, batchId);
            return;
        }

        // Preserve capture/inventory evidence. An old coarse plan and every derived
        // rendering artifact are intentionally invalidated before refinement.
        foreach (var stale in new[] { "group-plan.json", "group-catalog.json", "group-validation-report.json", "group-render-index.json", "group-render-index.html" })
            File.Delete(Path.Combine(batchDirectory, stale));
        var refinementDirectory = Path.Combine(batchDirectory, "group-refinement");
        if (Directory.Exists(refinementDirectory)) Directory.Delete(refinementDirectory, recursive: true);

        LoadBatchPipeline(batchDirectory, aiState: "running");
        SetPipelineRunState("AI grouping 1/4: reusing the complete capture database; no Excel recapture.");
        Log("Starting two-level AI grouping from persisted numeric-capture.sqlite.");
        TwoLevelGroupAnalysisRunResult grouping;
        try
        {
            grouping = await TwoLevelGroupAnalysisEngine.RunAsync(
                new TwoLevelGroupAnalysisRequest(_serviceDirectory, batchId),
                line => WriteProcessLog(groupingLog, "GROUPING", line));
            groupingLog.WriteLine("WPF", $"Two-level validation passed. Plan={grouping.PlanPath}; Catalog={grouping.CatalogPath}; Validation={grouping.ValidationPath}");
            groupingLog.FlushAndThrowIfWriteFailed();
        }
        catch (Exception exception)
        {
            groupingLog.WriteLine("WPF", $"Two-level grouping failed: {exception}");
            try { groupingLog.FlushAndThrowIfWriteFailed(); }
            catch (Exception loggingException) { Log($"AI grouping log failure: {loggingException.Message}"); }
            throw;
        }
        LoadBatchPipeline(batchDirectory, aiState: "validated");
        SetPipelineRunState($"AI grouping 3/4: full validation passed ({grouping.GroupCount} second-level groups; fallback {grouping.FallbackCount}/{grouping.ScannedCount}).");
        Log($"Validated AI group catalog created: {grouping.CatalogPath}");

        using (var renderLog = CreateRunLog($"batch_{batchId}.group-render", Path.Combine(batchDirectory, "logs")))
        {
            SetPipelineRunState("AI grouping 4/4: rendering only the validated group catalog.");
            var render = await RunGroupCatalogRendererAsync(renderLog, batchId);
            renderLog.FlushAndThrowIfWriteFailed();
            SetPipelineRunState(RenderCompletionText(render));
        }
        LoadBatchPipeline(batchDirectory, aiState: "complete");
        ShowGroupCatalog(output, batchId);
    }

    private static void MaterializeFileAssignments(string planPath, string catalogPath, string classificationPath, string summaryPath, string batchId)
    {
        var root = JsonNode.Parse(File.ReadAllText(planPath, Encoding.UTF8))?.AsObject()
            ?? throw new InvalidOperationException("AI group plan is not a JSON object.");
        var groups = root["groups"]?.AsArray() ?? throw new InvalidOperationException("AI group catalog has no groups.");
        // The AI plans assignable semantic categories.  The exception route is a
        // pipeline invariant, so add its empty default once rather than asking the
        // model to contradict the "no normal fallback" instruction.
        if (!groups.Any(group => group?["selector"]?["fallback"]?.GetValue<bool>() == true))
        {
            groups.Add(new JsonObject
            {
                ["id"] = "needs-review-fallback",
                ["name"] = "Scanner-proven empty or incomplete fallback",
                ["rendererKey"] = "renderNeedsReview",
                ["selector"] = new JsonObject { ["semanticCategories"] = new JsonArray(), ["fallback"] = true },
                ["representativeFiles"] = new JsonArray(),
                ["importantData"] = new JsonArray("Reserved for EMPTY_LAYOUT or CAPTURE_INCOMPLETE only."),
                ["extractionRule"] = "Do not infer missing source data.",
                ["htmlRule"] = "Render a review notice only.",
                ["openQuestions"] = new JsonArray()
            });
        }
        var rules = groups.Select(group => GroupRule.Parse(group?.AsObject() ?? throw new InvalidOperationException("Invalid group entry."))).ToList();
        if (rules.Select(rule => rule.Id).Distinct(StringComparer.OrdinalIgnoreCase).Count() != rules.Count || rules.Any(rule => !Regex.IsMatch(rule.Id, "^[a-z0-9]+(?:[._-][a-z0-9]+)*$")))
            throw new InvalidOperationException("AI group IDs must be unique stable slugs.");
        var rendererKeys = groups.Select(group => group?["rendererKey"]?.GetValue<string>()).ToList();
        // Multiple form groups legitimately share the generic needs-review renderer.
        // A renderer key selects an implementation; it is not a group identifier.
        if (rendererKeys.Any(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("AI renderer keys must be non-empty.");
        var fallbacks = rules.Where(rule => rule.Fallback).ToList();
        if (fallbacks.Count != 1) throw new InvalidOperationException("AI group plan requires exactly one fallback needs-review group.");
        var fallback = fallbacks[0];
        var clusterByPath = ReadLayoutClusters(Path.Combine(Path.GetDirectoryName(planPath)!, "layout-clusters.json"));
        var categoryByCluster = ReadSemanticCategories(Path.Combine(Path.GetDirectoryName(planPath)!, "layout-semantic-summary.json"));
        var assignments = new JsonArray();
        foreach (var row in ReadClassificationRows(classificationPath).Where(row => string.Equals(row.Status, "SCANNED", StringComparison.OrdinalIgnoreCase)))
        {
            if (!clusterByPath.TryGetValue(row.RelativePath, out var clusterId)) throw new InvalidOperationException($"No layout cluster for '{row.RelativePath}'.");
            if (!categoryByCluster.TryGetValue(clusterId, out var category)) throw new InvalidOperationException($"No semantic category for '{clusterId}'.");
            var matches = rules.Where(rule => rule.Matches(category)).ToList();
            if (matches.Count > 1) throw new InvalidOperationException($"AI group rules overlap for '{row.RelativePath}': {string.Join(", ", matches.Select(match => match.Id))}.");
            var match = matches.SingleOrDefault();
            if (match is null)
                throw new InvalidOperationException($"AI group plan left an assignable Excel ungrouped: '{row.RelativePath}' (cluster {clusterId}). Reclassification is required; fallback is reserved for scanner-proven exceptions.");
            assignments.Add(new JsonObject { ["relativePath"] = row.RelativePath, ["layoutClusterId"] = clusterId, ["semanticCategory"] = category, ["groupId"] = match.Id });
        }
        foreach (var group in groups.Select(node => node!.AsObject()))
        {
            group["memberSelectionRule"] = GroupRule.ToDisplayRule(group["selector"]!.AsObject());
            group.Remove("selector");
        }
        root["fileAssignments"] = assignments;
        var scanned = JsonDocument.Parse(File.ReadAllText(summaryPath, Encoding.UTF8)).RootElement.GetProperty("statusCounts").GetProperty("SCANNED").GetInt32();
        var paths = assignments.Select(node => node?["relativePath"]?.GetValue<string>()).ToList();
        if (assignments.Count != scanned || paths.Any(string.IsNullOrWhiteSpace) || paths.Distinct(StringComparer.OrdinalIgnoreCase).Count() != assignments.Count)
            throw new InvalidOperationException($"AI rule application must assign each scanned Excel exactly once. Scanned={scanned}, assignments={assignments.Count}.");
        var assignedGroupByPath = assignments.ToDictionary(node => node!["relativePath"]!.GetValue<string>(), node => node!["groupId"]!.GetValue<string>(), StringComparer.OrdinalIgnoreCase);
        var normalizedRepresentativeCount = 0;
        foreach (var group in groups.Select(node => node!.AsObject()))
        {
            var rule = rules.Single(candidate => string.Equals(candidate.Id, group["id"]!.GetValue<string>(), StringComparison.OrdinalIgnoreCase));
            var members = assignedGroupByPath.Where(entry => string.Equals(entry.Value, rule.Id, StringComparison.OrdinalIgnoreCase)).Select(entry => entry.Key).ToList();
            if (members.Count == 0)
            {
                if (rule.Fallback) continue;
                // The AI can propose a structurally valid but absent selector. It is
                // not a pipeline failure; retain it as inactive evidence and do not
                // prevent all matched files from reaching their renderer.
                group["inactive"] = true;
                group["inactiveReason"] = "No scanned workbook matched this selector in the current batch.";
                group["representativeFiles"] = new JsonArray();
                continue;
            }
            var representatives = group["representativeFiles"]!.AsArray();
            for (var index = 0; index < representatives.Count; index++)
            {
                var representative = representatives[index]?.GetValue<string>();
                if (!string.IsNullOrWhiteSpace(representative) && assignedGroupByPath.TryGetValue(representative, out var groupId) && string.Equals(groupId, rule.Id, StringComparison.OrdinalIgnoreCase)) continue;
                representatives[index] = members[Math.Min(index, members.Count - 1)];
                normalizedRepresentativeCount++;
            }
        }
        root["provenance"] = new JsonObject { ["batchId"] = batchId, ["planFile"] = Path.GetFileName(planPath), ["classificationFile"] = Path.GetFileName(classificationPath), ["materializedUtc"] = DateTimeOffset.UtcNow.ToString("O"), ["scannedCount"] = scanned, ["normalizedRepresentativeCount"] = normalizedRepresentativeCount };
        var temporary = catalogPath + ".tmp";
        File.WriteAllText(temporary, root.ToJsonString(new JsonSerializerOptions { WriteIndented = true }), Utf8WithBom);
        File.Move(temporary, catalogPath, overwrite: true);
    }

    private static IEnumerable<ClassificationRow> ReadClassificationRows(string path)
    {
        using var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        var header = ParseCsvRow(reader.ReadLine() ?? string.Empty);
        var columns = header.Select((name, index) => (name, index)).ToDictionary(value => value.name, value => value.index, StringComparer.OrdinalIgnoreCase);
        while (reader.ReadLine() is { } line)
        {
            var values = ParseCsvRow(line);
            string Value(string name) => columns.TryGetValue(name, out var index) && index < values.Count ? values[index] : string.Empty;
            yield return new ClassificationRow(Value("relativePath"), Value("status"), Value("primaryStructure"), Value("structuralTypes"));
        }
    }

    private static List<string> ParseCsvRow(string line)
    {
        var values = new List<string>();
        var value = new StringBuilder();
        var quoted = false;
        for (var index = 0; index < line.Length; index++)
        {
            var character = line[index];
            if (character == '"' && quoted && index + 1 < line.Length && line[index + 1] == '"') { value.Append(character); index++; continue; }
            if (character == '"') { quoted = !quoted; continue; }
            if (character == ',' && !quoted) { values.Add(value.ToString()); value.Clear(); continue; }
            value.Append(character);
        }
        values.Add(value.ToString());
        return values;
    }

    private sealed record ClassificationRow(string RelativePath, string Status, string PrimaryStructure, string StructuralTypes)
    {
        public HashSet<string> Types { get; } = StructuralTypes.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private sealed record BatchClassificationRow(string RelativePath, string Status, string Reason);
    private sealed record PipelineStage(string Status, string Reason);
    private sealed record RenderStage(string Status, string ReportPath, string Reason);

    private static Dictionary<string, string> ReadLayoutClusters(string path)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
        return document.RootElement.GetProperty("clusters").EnumerateArray()
            .SelectMany(cluster => cluster.GetProperty("MemberPaths").EnumerateArray().Select(member => (Path: member.GetString() ?? string.Empty, Cluster: cluster.GetProperty("LayoutClusterId").GetString() ?? string.Empty)))
            .ToDictionary(value => value.Path, value => value.Cluster, StringComparer.OrdinalIgnoreCase);
    }

    private static Dictionary<string, string> ReadSemanticCategories(string path)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
        return document.RootElement.GetProperty("categories").EnumerateArray()
            .SelectMany(category => category.GetProperty("layoutClusterIds").EnumerateArray().Select(cluster => (Cluster: cluster.GetString() ?? string.Empty, Category: category.GetProperty("category").GetString() ?? string.Empty)))
            .ToDictionary(value => value.Cluster, value => value.Category, StringComparer.OrdinalIgnoreCase);
    }

    private sealed record GroupRule(string Id, HashSet<string> SemanticCategories, bool Fallback)
    {
        public static GroupRule Parse(JsonObject group)
        {
            var id = group["id"]?.GetValue<string>();
            var selector = group["selector"]?.AsObject();
            var fallback = selector?["fallback"]?.GetValue<bool>() ?? false;
            if (string.IsNullOrWhiteSpace(id)) throw new InvalidOperationException("Every AI group requires id and structured selector.");
            var categories = selector?["semanticCategories"]?.AsArray().Select(node => node?.GetValue<string>() ?? string.Empty).ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];
            if (categories.Any(string.IsNullOrWhiteSpace) || (!fallback && categories.Count == 0)) throw new InvalidOperationException($"Group '{id}' has an invalid semantic-category selector.");
            return new GroupRule(id, categories, fallback);
        }

        public bool Matches(string category) => SemanticCategories.Contains(category);

        public static string ToDisplayRule(JsonObject selector)
        {
            var categories = selector["semanticCategories"]?.AsArray().Select(node => node?.GetValue<string>() ?? string.Empty).Where(value => !string.IsNullOrWhiteSpace(value)).ToList() ?? [];
            return selector["fallback"]?.GetValue<bool>() == true ? "Fallback: empty, incomplete, or newly unsupported layouts." : $"semantic categories = {string.Join("; ", categories)}";
        }
    }

    private void ShowGroupCatalog(string path, string batchId)
    {
        var json = File.ReadAllText(path, Encoding.UTF8);
        using var document = JsonDocument.Parse(json);
        var cards = new StringBuilder();
        foreach (var group in document.RootElement.GetProperty("groups").EnumerateArray())
        {
            string Text(string name) => group.GetProperty(name).GetString() ?? string.Empty;
            string List(string name) => string.Concat(group.GetProperty(name).EnumerateArray().Select(value => $"<li>{System.Net.WebUtility.HtmlEncode(value.GetString())}</li>"));
            cards.Append($"<section><h2>{System.Net.WebUtility.HtmlEncode(Text("name"))}</h2><p><code>{System.Net.WebUtility.HtmlEncode(Text("id"))}</code></p><h3>대표 파일</h3><ul>{List("representativeFiles")}</ul><h3>핵심 데이터</h3><ul>{List("importantData")}</ul><h3>추출 규칙</h3><p>{System.Net.WebUtility.HtmlEncode(Text("extractionRule"))}</p><h3>HTML 규칙</h3><p>{System.Net.WebUtility.HtmlEncode(Text("htmlRule"))}</p><h3>확인 필요</h3><ul>{List("openQuestions")}</ul></section>");
        }
        NavigateUtf8Html($"<!doctype html><html><head><meta charset='utf-8'><style>body{{font-family:'Segoe UI',sans-serif;margin:24px;color:#162d50}}section{{background:#fff;border:1px solid #d5dfeb;border-radius:8px;padding:16px;margin:14px 0}}h1{{margin-top:0}}h2{{margin-bottom:4px}}h3{{font-size:14px;margin:14px 0 4px}}code{{color:#53657a}}</style></head><body><h1>AI 양식 그룹 카탈로그</h1><p>배치: {System.Net.WebUtility.HtmlEncode(batchId)}</p>{cards}</body></html>");
        ResultTitle.Text = $"AI 그룹 분석 결과 — {batchId}";
    }

    private void RefreshLearnedResults_Click(object sender, RoutedEventArgs e) => LoadLearnedResults();

    private void LearnedBatchSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loadingLearnedResults || LearnedBatchSelector.SelectedItem is not LearnedResultBatch batch) return;
        LoadLearnedReports(batch);
    }

    private void ShowLearnedRenderIndex_Click(object sender, RoutedEventArgs e)
    {
        if (LearnedBatchSelector.SelectedItem is LearnedResultBatch batch) ShowLearnedRenderIndex(batch);
    }

    private void ShowLearnedCatalog_Click(object sender, RoutedEventArgs e)
    {
        if (LearnedBatchSelector.SelectedItem is LearnedResultBatch batch) ShowLearnedCatalog(batch);
    }

    private void LoadLearnedResults()
    {
        var previousBatchId = (LearnedBatchSelector.SelectedItem as LearnedResultBatch)?.BatchId;
        var batchesRoot = _pathSettings.BatchRootDirectory;
        var discovered = new List<LearnedResultBatch>();
        if (Directory.Exists(batchesRoot))
        {
            foreach (var directory in Directory.EnumerateDirectories(batchesRoot))
            {
                var catalogPath = Path.Combine(directory, "group-catalog.json");
                var validationPath = Path.Combine(directory, "group-validation-report.json");
                var indexPath = Path.Combine(directory, "group-render-index.json");
                var batchPath = Path.Combine(directory, "batch.json");
                if (!File.Exists(catalogPath) || !File.Exists(validationPath) || !File.Exists(indexPath)) continue;
                try
                {
                    using var validation = JsonDocument.Parse(File.ReadAllText(validationPath, Encoding.UTF8));
                    if (!validation.RootElement.TryGetProperty("isValid", out var isValid) || !isValid.GetBoolean()) continue;

                    var sourceRootPath = string.Empty;
                    if (File.Exists(batchPath))
                    {
                        using var batch = JsonDocument.Parse(File.ReadAllText(batchPath, Encoding.UTF8));
                        sourceRootPath = batch.RootElement.TryGetProperty("rootPath", out var rootPath) ? rootPath.GetString() ?? string.Empty : string.Empty;
                    }
                    using var catalog = JsonDocument.Parse(File.ReadAllText(catalogPath, Encoding.UTF8));
                    using var index = JsonDocument.Parse(File.ReadAllText(indexPath, Encoding.UTF8));
                    var groups = catalog.RootElement.TryGetProperty("groups", out var groupArray) ? groupArray.GetArrayLength() : 0;
                    var summary = index.RootElement.GetProperty("summary");
                    var statuses = summary.TryGetProperty("StatusCounts", out var statusCounts) ? statusCounts : default;
                    discovered.Add(new LearnedResultBatch(
                        Path.GetFileName(directory),
                        directory,
                        sourceRootPath,
                        File.GetLastWriteTime(directory),
                        summary.TryGetProperty("WorkbookCount", out var workbookCount) ? workbookCount.GetInt32() : 0,
                        groups,
                        ReadStatusCount(statuses, "COMPONENTS_RENDERED"),
                        ReadStatusCount(statuses, "STRUCTURE_READY"),
                        ReadStatusCount(statuses, "CONTRACT_MISMATCH")));
                }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException or KeyNotFoundException)
                {
                    Log($"기존 학습 결과를 읽지 못했습니다: {Path.GetFileName(directory)} ({exception.Message})");
                }
            }
        }

        var selected = discovered
            .OrderByDescending(batch => batch.LastUpdated)
            .FirstOrDefault(batch => string.Equals(batch.BatchId, previousBatchId, StringComparison.OrdinalIgnoreCase))
            ?? discovered.OrderByDescending(batch => batch.LastUpdated).FirstOrDefault();

        _loadingLearnedResults = true;
        try
        {
            _learnedBatches.Clear();
            foreach (var batch in discovered.OrderByDescending(batch => batch.LastUpdated)) _learnedBatches.Add(batch);
            LearnedBatchSelector.SelectedItem = selected;
        }
        finally
        {
            _loadingLearnedResults = false;
        }

        if (selected is null)
        {
            LearnedResultsTitle.Text = "저장된 검증 학습 결과가 없습니다.";
            LearnedResultsSummary.Text = "유효한 그룹 검증 보고서와 렌더 인덱스가 있는 배치만 표시합니다.";
            _learnedReports.Clear();
            NavigateUtf8Html("<html><body style='font-family:Segoe UI;padding:24px'><h2>표시할 기존 학습 결과가 없습니다.</h2><p>AI 그룹화와 렌더링이 완료된 배치가 생성되면 이 탭에서 바로 확인할 수 있습니다.</p></body></html>");
            return;
        }

        LoadLearnedReports(selected);
    }

    private void ShowLearnedRenderIndex(LearnedResultBatch batch)
    {
        var indexPath = Path.Combine(batch.BatchDirectory, "group-render-index.html");
        LearnedResultsTitle.Text = $"기존 학습 결과 — {batch.BatchId}";
        LearnedResultsSummary.Text = $"검증된 그룹 {batch.GroupCount:N0}개 · 보고서 {batch.WorkbookCount:N0}개 · 컴포넌트 렌더 {batch.ComponentsRendered:N0}개 · 구조 예외 {batch.StructureReady:N0}개 · 계약 불일치 {batch.ContractMismatch:N0}개";
        if (!File.Exists(indexPath))
        {
            NavigateUtf8Html("<html><body style='font-family:Segoe UI;padding:24px'><h2>렌더 인덱스를 찾을 수 없습니다.</h2></body></html>");
            return;
        }

        ResultTitle.Text = $"기존 학습 렌더 인덱스 — {batch.BatchId}";
        NavigateResultHtmlFile(indexPath);
    }

    private void ShowLearnedCatalog(LearnedResultBatch batch)
    {
        var catalogPath = Path.Combine(batch.BatchDirectory, "group-catalog.json");
        if (!File.Exists(catalogPath))
        {
            NavigateUtf8Html("<html><body style='font-family:Segoe UI;padding:24px'><h2>그룹 카탈로그를 찾을 수 없습니다.</h2></body></html>");
            return;
        }

        using var catalog = JsonDocument.Parse(File.ReadAllText(catalogPath, Encoding.UTF8));
        var root = catalog.RootElement;
        var assignments = root.TryGetProperty("fileAssignments", out var assignmentArray)
            ? assignmentArray.EnumerateArray().Select(item => item.GetProperty("groupId").GetString() ?? string.Empty).GroupBy(id => id, StringComparer.OrdinalIgnoreCase).ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var rows = new StringBuilder();
        foreach (var group in root.GetProperty("groups").EnumerateArray())
        {
            var id = group.GetProperty("id").GetString() ?? string.Empty;
            var name = group.TryGetProperty("name", out var nameValue) ? nameValue.GetString() ?? id : id;
            var renderer = group.TryGetProperty("rendererKey", out var rendererValue) ? rendererValue.GetString() ?? string.Empty : string.Empty;
            var recipe = group.TryGetProperty("componentRecipe", out var recipeArray)
                ? string.Join(" → ", recipeArray.EnumerateArray().Select(value => value.GetString() ?? string.Empty))
                : "(구성 요소 정보 없음)";
            var representatives = group.TryGetProperty("representativeFiles", out var files)
                ? string.Join("<br>", files.EnumerateArray().Take(2).Select(value => System.Net.WebUtility.HtmlEncode(value.GetString() ?? string.Empty)))
                : string.Empty;
            rows.Append($"<tr><td><b>{System.Net.WebUtility.HtmlEncode(name)}</b><br><code>{System.Net.WebUtility.HtmlEncode(id)}</code></td><td>{assignments.GetValueOrDefault(id):N0}</td><td><code>{System.Net.WebUtility.HtmlEncode(renderer)}</code></td><td>{System.Net.WebUtility.HtmlEncode(recipe)}</td><td>{representatives}</td></tr>");
        }
        LearnedResultsTitle.Text = $"AI 그룹 카탈로그 — {batch.BatchId}";
        ResultTitle.Text = $"AI 그룹 카탈로그 — {batch.BatchId}";
        NavigateUtf8Html($"<!doctype html><html><head><meta charset='utf-8'><style>body{{font-family:'Segoe UI','Malgun Gothic',sans-serif;margin:24px;background:#f5f7fa;color:#172b4d}}table{{width:100%;border-collapse:collapse;background:#fff;font-size:13px}}th,td{{border:1px solid #d8e2f0;padding:10px;vertical-align:top;text-align:left}}th{{background:#eaf1f8}}code{{font-family:Consolas;font-size:11px}}</style></head><body><h1>검증된 AI 그룹 카탈로그</h1><p>배치: <code>{System.Net.WebUtility.HtmlEncode(batch.BatchId)}</code> · 그룹 {batch.GroupCount:N0}개</p><table><thead><tr><th>그룹</th><th>파일 수</th><th>렌더러</th><th>데이터 구성 요소</th><th>대표 양식</th></tr></thead><tbody>{rows}</tbody></table></body></html>");
    }

    private static int ReadStatusCount(JsonElement statuses, string status) =>
        statuses.ValueKind == JsonValueKind.Object && statuses.TryGetProperty(status, out var count) && count.TryGetInt32(out var value) ? value : 0;

    private void LoadLearnedReports(LearnedResultBatch batch)
    {
        var indexPath = Path.Combine(batch.BatchDirectory, "group-render-index.json");
        var reports = new List<LearnedReportRow>();
        try
        {
            using var index = JsonDocument.Parse(File.ReadAllText(indexPath, Encoding.UTF8));
            foreach (var row in index.RootElement.GetProperty("rows").EnumerateArray())
            {
                var relativePath = row.GetProperty("RelativePath").GetString() ?? string.Empty;
                var reportPath = row.GetProperty("ReportPath").GetString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(relativePath) || string.IsNullOrWhiteSpace(reportPath)) continue;
                var components = row.TryGetProperty("ComponentCounts", out var componentCounts) && componentCounts.ValueKind == JsonValueKind.Object
                    ? string.Join(" · ", componentCounts.EnumerateObject().OrderBy(component => component.Name, StringComparer.Ordinal).Select(component => $"{component.Name} {component.Value.GetInt32():N0}"))
                    : "구조 예외";
                reports.Add(new LearnedReportRow(
                    batch.BatchId,
                    batch.BatchDirectory,
                    batch.SourceRootPath,
                    relativePath,
                    row.GetProperty("GroupId").GetString() ?? string.Empty,
                    row.GetProperty("RendererKey").GetString() ?? string.Empty,
                    row.GetProperty("Status").GetString() ?? string.Empty,
                    reportPath,
                    components));
            }
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException or KeyNotFoundException)
        {
            LearnedResultsTitle.Text = $"기존 학습 결과를 읽을 수 없습니다 — {batch.BatchId}";
            LearnedResultsSummary.Text = exception.Message;
            _learnedReports.Clear();
            return;
        }

        _loadingLearnedResults = true;
        try
        {
            _learnedReports.Clear();
            foreach (var report in reports.OrderBy(report => report.FileName, StringComparer.OrdinalIgnoreCase)) _learnedReports.Add(report);
            LearnedReportsGrid.SelectedIndex = -1;
        }
        finally
        {
            _loadingLearnedResults = false;
        }

        LearnedResultsTitle.Text = $"기존 학습 보고서 — {batch.BatchId}";
        LearnedResultsSummary.Text = $"목록에서 Excel을 선택하면 오른쪽에 해당 분석 HTML이 열립니다. 보고서 {reports.Count:N0}개 · 컴포넌트 렌더 {batch.ComponentsRendered:N0}개 · 구조 예외 {batch.StructureReady:N0}개 · 계약 불일치 {batch.ContractMismatch:N0}개";
        if (reports.Count > 0) LearnedReportsGrid.SelectedIndex = 0;
    }

    private void LearnedReportsGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loadingLearnedResults || LearnedReportsGrid.SelectedItem is not LearnedReportRow report) return;
        ShowLearnedReport(report);
    }

    private void OpenLearnedSourceExcel_Click(object sender, RoutedEventArgs e)
    {
        if (LearnedReportsGrid.SelectedItem is not LearnedReportRow report) return;
        if (!File.Exists(report.FullSourcePath))
        {
            Log($"원본 Excel을 찾을 수 없습니다: {report.RelativePath}");
            return;
        }
        Process.Start(new ProcessStartInfo(report.FullSourcePath) { UseShellExecute = true });
    }

    private void OpenLearnedReportHtml_Click(object sender, RoutedEventArgs e)
    {
        if (LearnedReportsGrid.SelectedItem is not LearnedReportRow report) return;
        if (!File.Exists(report.FullReportPath))
        {
            Log($"분석 HTML을 찾을 수 없습니다: {report.RelativePath}");
            return;
        }
        Process.Start(new ProcessStartInfo(report.FullReportPath) { UseShellExecute = true });
    }

    private void ShowLearnedReport(LearnedReportRow report)
    {
        ResultTitle.Text = $"기존 학습 보고서 — {report.FileName}";
        PipelineFileTitle.Text = report.FileName;
        PipelineBatchText.Text = $"배치: {report.BatchId} · 그룹: {report.GroupId}";
        PipelineStagesText.Text = string.Join(Environment.NewLine,
            $"AI 그룹    {report.GroupId}",
            $"렌더러     {report.RendererKey}",
            $"렌더 상태  {report.Status}",
            $"구성 요소  {report.ComponentSummary}");
        if (!File.Exists(report.FullReportPath))
        {
            NavigateUtf8Html("<html><body style='font-family:Segoe UI;padding:24px'><h2>선택한 보고서 HTML을 찾을 수 없습니다.</h2></body></html>");
            return;
        }
        NavigateResultHtmlFile(report.FullReportPath);
    }

    private void NavigateResultHtmlFile(string htmlPath)
    {
        var baseUri = new Uri(Path.GetFullPath(Path.GetDirectoryName(htmlPath)!) + Path.DirectorySeparatorChar).AbsoluteUri;
        var html = File.ReadAllText(htmlPath, Encoding.UTF8)
            .Replace("<head>", $"<head><base href=\"{baseUri}\">", StringComparison.OrdinalIgnoreCase);
        NavigateUtf8Html(html);
    }

    private sealed record LearnedResultBatch(
        string BatchId,
        string BatchDirectory,
        string SourceRootPath,
        DateTime LastUpdated,
        int WorkbookCount,
        int GroupCount,
        int ComponentsRendered,
        int StructureReady,
        int ContractMismatch)
    {
        public string DisplayName => $"{BatchId} · {LastUpdated:yyyy-MM-dd HH:mm} · 보고서 {WorkbookCount:N0}개";
    }

    private sealed record LearnedReportRow(
        string BatchId,
        string BatchDirectory,
        string SourceRootPath,
        string RelativePath,
        string GroupId,
        string RendererKey,
        string Status,
        string ReportPath,
        string ComponentSummary)
    {
        public string FileName => Path.GetFileName(RelativePath);
        public string FullSourcePath => string.IsNullOrWhiteSpace(SourceRootPath) ? string.Empty : Path.GetFullPath(Path.Combine(SourceRootPath, RelativePath));
        public string FullReportPath => Path.GetFullPath(Path.Combine(BatchDirectory, ReportPath));
    }

    private async Task RunStartupNumericReviewAsync()
    {
        var batchId = _startupAnalysis.NumericReviewBatchId!;
        var batchLogDirectory = Path.Combine(
            _pathSettings.BatchRootDirectory,
            batchId,
            "logs");
        using var runLog = CreateRunLog($"batch_{batchId}.numeric-review", batchLogDirectory);
        FilesGrid.IsEnabled = false;
        SourceToolbar.IsEnabled = false;
        _numericReviewRunning = true;
        try
        {
            SetNumericReviewState("숫자 검토: 원본 숫자 표 적재 중… Excel/COM을 사용하지 않습니다.");
            Log("명령줄 숫자 표 원본 적재를 시작합니다. Excel/COM은 사용하지 않습니다.");
            Log($"실행 로그: {runLog.FilePath}");
            await RunNumericReviewBatchAsync(batchId, runLog);
            ShowNumericReviewIndex(batchId);
            Log("숫자 표 원본 적재, Test–Normal 검토, HTML 생성이 완료되었습니다.");
        }
        catch (Exception ex)
        {
            runLog.WriteLine("WPF", $"Numeric capture or review failed: {ex.Message}");
            try { runLog.FlushAndThrowIfWriteFailed(); }
            catch (Exception logException) { Log($"숫자 검토 로그 오류: {logException.Message}"); }
            Log($"숫자 검토 실패: {ex.Message}");
            System.Windows.MessageBox.Show(ex.Message, "숫자 검토 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _numericReviewRunning = false;
            FilesGrid.IsEnabled = true;
            SourceToolbar.IsEnabled = true;
        }
    }

    private async Task RunStartupBatchScanAsync()
    {
        var batchName = _startupAnalysis.IsBatchFolderScan
            ? Path.GetFileName(Path.TrimEndingDirectorySeparator(_startupAnalysis.BatchFolder!))
            : _startupAnalysis.ResumeBatchId!;
        // Give first-time command-line batches a stable ID before starting the child.
        // This keeps both scanner artifacts and the WPF execution log below the same
        // batch-scoped output directory instead of touching the regular run-log area.
        var effectiveBatchId = _startupAnalysis.IsBatchFolderScan
            ? _startupAnalysis.BatchId ?? $"structure-scan-{DateTime.UtcNow:yyyyMMddTHHmmssZ}"
            : _startupAnalysis.ResumeBatchId!;
        var batchLogDirectory = Path.Combine(
            _pathSettings.BatchRootDirectory,
            effectiveBatchId,
            "logs");
        using var runLog = CreateRunLog($"batch_{batchName}.scan", batchLogDirectory);
        try
        {
            if (_startupAnalysis.IsBatchFolderScan)
            {
                var files = await Task.Run(() => Directory.EnumerateFiles(_startupAnalysis.BatchFolder!, "*.*", SearchOption.AllDirectories)
                    .Where(path => new[] { ".xlsx", ".xlsm", ".xls", ".xlsb" }.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase) && !Path.GetFileName(path).StartsWith("~$"))
                    .ToList());
                AddFiles(files);
            }
            Log(_startupAnalysis.IsBatchFolderScan ? "명령줄 구조 사전 스캔을 시작합니다." : "명령줄 배치 구조 사전 스캔을 재개합니다.");
            Log($"실행 로그: {runLog.FilePath}");
            await RunStructureScanAsync(
                runLog,
                new StructureScanRequest(
                    _serviceDirectory,
                    _startupAnalysis.BatchFolder,
                    _startupAnalysis.ResumeBatchId,
                    _startupAnalysis.IsBatchFolderScan ? effectiveBatchId : null,
                    Pilot: string.Equals(_startupAnalysis.BatchLimitArgument, "--pilot", StringComparison.OrdinalIgnoreCase) ? _startupAnalysis.BatchLimit ?? 0 : 0,
                    Limit: string.Equals(_startupAnalysis.BatchLimitArgument, "--limit", StringComparison.OrdinalIgnoreCase) ? _startupAnalysis.BatchLimit ?? 0 : 0,
                    RetryFailed: _startupAnalysis.RetryFailed));

            // A command-line batch starts with an empty manually-added list.  Re-read
            // the completed scanner artifacts now so the UI immediately shows the
            // per-file SCANNED/TRUNCATED state instead of leaving every row at 대기.
            LoadBatchPipeline(
                Path.Combine(
                    _pathSettings.BatchRootDirectory,
                    effectiveBatchId),
                aiState: "대기");
            if (_startupAnalysis.RunAiGroupAnalysis) await RunAiGroupAnalysisAsync(effectiveBatchId);
            runLog.WriteLine("WPF", "C# batch structure scan completed successfully.");
            runLog.FlushAndThrowIfWriteFailed();
            Log("명령줄 구조 사전 스캔이 완료되었습니다.");
        }
        catch (Exception ex)
        {
            runLog.WriteLine("WPF", $"Batch structure scan failed: {ex.Message}");
            try { runLog.FlushAndThrowIfWriteFailed(); }
            catch (Exception logException) { Log($"구조 사전 스캔 로그 오류: {logException.Message}"); }
            Log($"구조 사전 스캔 실패: {ex.Message}");
            System.Windows.MessageBox.Show(ex.Message, "구조 사전 스캔 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string FindServiceDirectory()
    {
        var configured = Environment.GetEnvironmentVariable("INFERENCE_DATA_AI_SERVICE_DIR");
        if (!string.IsNullOrWhiteSpace(configured) && File.Exists(Path.Combine(configured, "inference_data_ai_cli.py"))) return configured;
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
            if (File.Exists(Path.Combine(directory.FullName, "inference_data_ai_cli.py"))) return directory.FullName;
        throw new DirectoryNotFoundException("inference_data_ai_cli.py가 있는 서비스 폴더를 찾지 못했습니다.");
    }

    private async void AddFolder_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new Forms.FolderBrowserDialog { Description = "Excel 파일이 있는 폴더 선택" };
        if (dialog.ShowDialog() != Forms.DialogResult.OK) return;
        Log("Excel 파일을 비동기로 검색 중입니다.");
        var files = await Task.Run(
            () => ExpandExcelInputPaths([dialog.SelectedPath]));
        AddFiles(files);
        SetPipelineRunState(
            $"입력 완료 · Excel {files.Count:N0}개 · 다음: 단건은 목록 우클릭, 배치는 ② AI 그룹 분석");
    }

    private void FilesDropSurface_PreviewDragEnter(
        object sender,
        WpfDragEventArgs e) =>
        UpdateFilesDropFeedback(e);

    private void FilesDropSurface_PreviewDragOver(
        object sender,
        WpfDragEventArgs e) =>
        UpdateFilesDropFeedback(e);

    private void FilesDropSurface_DragLeave(
        object sender,
        WpfDragEventArgs e)
    {
        FilesDropIndicator.Visibility = Visibility.Collapsed;
        e.Handled = true;
    }

    private async void FilesDropSurface_Drop(
        object sender,
        WpfDragEventArgs e)
    {
        FilesDropIndicator.Visibility = Visibility.Collapsed;
        e.Handled = true;
        if (!TryGetDroppedPaths(e, out var droppedPaths))
            return;

        SetPipelineRunState("끌어놓은 Excel 파일·폴더를 검색 중");
        Log(
            $"끌어놓기 입력: {droppedPaths.Length:N0}개 경로를 확인합니다.");
        try
        {
            var files = await Task.Run(
                () => ExpandExcelInputPaths(droppedPaths));
            if (files.Count == 0)
            {
                SetPipelineRunState(
                    "추가할 Excel 파일 없음 · 지원 형식: .xlsx/.xlsm/.xlsb/.xls");
                return;
            }

            var before = _items.Count;
            AddFiles(files);
            var added = _items.Count - before;
            SetPipelineRunState(
                $"끌어놓기 완료 · 발견 {files.Count:N0}개 · 새로 추가 {added:N0}개"
                + " · 다음: 단건은 목록 우클릭, 배치는 ② AI 그룹 분석");
            var first = files
                .Select(path => _items.FirstOrDefault(item =>
                    string.Equals(
                        item.FullPath,
                        path,
                        StringComparison.OrdinalIgnoreCase)))
                .FirstOrDefault(item => item is not null);
            if (first is not null)
                FilesGrid.SelectedItem = first;
        }
        catch (Exception exception)
        {
            SetPipelineRunState("끌어놓기 실패 · 개발자 로그 확인");
            Log($"끌어놓기 입력 실패: {exception.Message}");
            System.Windows.MessageBox.Show(
                exception.Message,
                "Excel 끌어놓기 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void UpdateFilesDropFeedback(WpfDragEventArgs e)
    {
        var canDrop = TryGetDroppedPaths(e, out _);
        e.Effects = canDrop
            ? WpfDragDropEffects.Copy
            : WpfDragDropEffects.None;
        FilesDropIndicator.Visibility = canDrop
            ? Visibility.Visible
            : Visibility.Collapsed;
        e.Handled = true;
    }

    private static bool TryGetDroppedPaths(
        WpfDragEventArgs e,
        out string[] paths)
    {
        paths = e.Data.GetDataPresent(WpfDataFormats.FileDrop)
            && e.Data.GetData(WpfDataFormats.FileDrop) is string[] dropped
            ? dropped
                .Where(path =>
                    File.Exists(path) || Directory.Exists(path))
                .ToArray()
            : [];
        return paths.Length > 0;
    }

    private static List<string> ExpandExcelInputPaths(
        IEnumerable<string> inputPaths)
    {
        var files = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase);
        var enumerationOptions = new EnumerationOptions
        {
            RecurseSubdirectories = true,
            IgnoreInaccessible = true,
            ReturnSpecialDirectories = false,
        };
        foreach (var input in inputPaths)
        {
            var fullPath = Path.GetFullPath(input);
            if (File.Exists(fullPath))
            {
                if (IsExcelInputFile(fullPath))
                    files.Add(fullPath);
                continue;
            }
            if (!Directory.Exists(fullPath))
                continue;
            foreach (var file in Directory.EnumerateFiles(
                         fullPath,
                         "*",
                         enumerationOptions))
                if (IsExcelInputFile(file))
                    files.Add(Path.GetFullPath(file));
        }
        return files
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsExcelInputFile(string path) =>
        ExcelExtensions.Contains(Path.GetExtension(path))
        && !Path.GetFileName(path).StartsWith(
            "~$",
            StringComparison.Ordinal);

    private async void NumericReviewFolder_Click(object sender, RoutedEventArgs e)
    {
        if (_numericReviewRunning)
        {
            Log("숫자 검토 배치가 이미 실행 중입니다.");
            return;
        }

        using var dialog = new Forms.FolderBrowserDialog
        {
            Description = "숫자 검토할 Excel 파일 폴더를 선택하세요 (읽기 전용, Excel/COM 미사용)"
        };
        if (dialog.ShowDialog() != Forms.DialogResult.OK) return;
        await RunNumericReviewFolderAsync(dialog.SelectedPath);
    }

    private async Task RunNumericReviewFolderAsync(string folderPath)
    {
        var sourceFolder = Path.GetFullPath(folderPath);
        var batchId = $"numeric-review-{DateTime.UtcNow:yyyyMMddTHHmmssfffZ}";
        var batchLogDirectory = Path.Combine(
            _pathSettings.BatchRootDirectory,
            batchId,
            "logs");
        using var runLog = CreateRunLog($"batch_{batchId}.numeric-review", batchLogDirectory);
        _numericReviewRunning = true;
        FilesGrid.IsEnabled = false;
        SourceToolbar.IsEnabled = false;

        try
        {
            SetNumericReviewState("숫자 검토 1/3: Excel 구조와 숫자 표 후보를 읽는 중… (Non-COM)");
            Log($"숫자 검토 폴더: {sourceFolder}");
            Log($"실행 로그: {runLog.FilePath}");
            await RunStructureScanAsync(
                runLog,
                new StructureScanRequest(
                    _serviceDirectory,
                    sourceFolder,
                    ResumeBatchId: null,
                    BatchId: batchId));

            await RunNumericReviewBatchAsync(batchId, runLog);
            ShowNumericReviewIndex(batchId);
            Log("숫자 검토 배치가 완료되었습니다. 우측에 배치 HTML 목록을 표시했습니다.");
        }
        catch (Exception ex)
        {
            runLog.WriteLine("WPF", $"Numeric folder review failed: {ex.Message}");
            try { runLog.FlushAndThrowIfWriteFailed(); }
            catch (Exception logException) { Log($"숫자 검토 로그 오류: {logException.Message}"); }
            SetNumericReviewState("숫자 검토: 실패 — 실행 로그를 확인하세요.");
            Log($"숫자 검토 실패: {ex.Message}");
            System.Windows.MessageBox.Show(ex.Message, "숫자 검토 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _numericReviewRunning = false;
            FilesGrid.IsEnabled = true;
            SourceToolbar.IsEnabled = true;
        }
    }

    private async Task RunNumericReviewBatchAsync(string batchId, RunLog runLog)
    {
        SetNumericReviewState("숫자 검토 2/3: 원본 숫자 표를 전용 DB에 적재 중…");
        await RunNumericCaptureAsync(runLog, batchId);

        SetNumericReviewState("숫자 검토 3/3: 같은 날짜 Test–Normal 비교 및 HTML 생성 중…");
        Log("동일 날짜 Test–Normal 숫자 검토를 생성합니다.");
        await RunNumericReviewAsync(runLog, batchId);
        await RunNumericRendererAsync(runLog, batchId);
        await RunGroupCatalogRendererAsync(runLog, batchId);
        runLog.WriteLine("WPF", "Numeric capture, same-date Test-Normal review, and HTML rendering completed successfully.");
        runLog.FlushAndThrowIfWriteFailed();
    }

    private Task<StructureScanRunResult> RunStructureScanAsync(RunLog runLog, StructureScanRequest request)
    {
        runLog.WriteLine("WPF", $"Starting C# StructureScanEngine for {(request.BatchFolder ?? request.ResumeBatchId ?? request.BatchId ?? "batch")}");
        return Task.Run(() => StructureScanEngine.Run(request, line => WriteProcessLog(runLog, "ENGINE", line)));
    }

    private Task<NumericCaptureRunResult> RunNumericCaptureAsync(RunLog runLog, string batchId)
    {
        runLog.WriteLine("WPF", $"Starting C# NumericCaptureEngine for {batchId}");
        return Task.Run(() => NumericCaptureEngine.Run(
            new NumericCaptureRequest(_serviceDirectory, batchId),
            line => WriteProcessLog(runLog, "ENGINE", line)));
    }

    private Task<NumericReviewRunResult> RunNumericReviewAsync(RunLog runLog, string batchId)
    {
        runLog.WriteLine("WPF", $"Starting C# NumericReviewEngine for {batchId}");
        return Task.Run(() => NumericReviewEngine.Run(
            new NumericReviewRequest(_serviceDirectory, batchId),
            line => WriteProcessLog(runLog, "ENGINE", line)));
    }

    private Task<NumericRendererRunResult> RunNumericRendererAsync(RunLog runLog, string batchId)
    {
        runLog.WriteLine("WPF", $"Starting C# NumericRendererEngine for {batchId}");
        return Task.Run(() => NumericRendererEngine.Run(
            new NumericRendererRequest(_serviceDirectory, batchId),
            line => WriteProcessLog(runLog, "ENGINE", line)));
    }

    private Task<GroupCatalogRendererRunResult> RunGroupCatalogRendererAsync(RunLog runLog, string batchId)
    {
        runLog.WriteLine("WPF", $"Starting GroupCatalogRendererEngine for {batchId}");
        return Task.Run(() => GroupCatalogRendererEngine.Run(
            new GroupCatalogRendererRequest(_serviceDirectory, batchId),
            line => WriteProcessLog(runLog, "ENGINE", line)));
    }

    private void ShowNumericReviewIndex(string batchId)
    {
        var indexPath = Path.Combine(
            _pathSettings.BatchRootDirectory,
            batchId,
            "numeric-report-index.html");
        if (!File.Exists(indexPath))
            throw new FileNotFoundException("숫자 검토 HTML 색인을 찾을 수 없습니다.", indexPath);

        // The browser receives a stream so its UTF-8 encoding is deterministic.
        // A stream has no file URL of its own, therefore add an explicit base URL
        // so the per-workbook links in the index still resolve within this batch.
        var baseUri = new Uri(Path.GetFullPath(Path.GetDirectoryName(indexPath)!) + Path.DirectorySeparatorChar).AbsoluteUri;
        var indexHtml = File.ReadAllText(indexPath, Encoding.UTF8)
            .Replace("<head>", $"<head><base href=\"{baseUri}\">", StringComparison.OrdinalIgnoreCase);
        NavigateUtf8Html(indexHtml);
        ResultTitle.Text = $"숫자 검토 결과 — {batchId}";
        SetNumericReviewState($"숫자 검토: 완료 — 전용 DB와 HTML 보고서가 배치에 저장되었습니다. ({batchId})");
    }

    private void SetNumericReviewState(string text) => NumericReviewState.Text = text;

    private void SetPipelineRunState(string text) => PipelineRunState.Text = $"파이프라인: {text}";

    private static string RenderCompletionText(GroupCatalogRendererRunResult run)
    {
        var summary = run.Summary;
        if (summary is null) return "AI 그룹 일괄 분석 실패: 렌더 결과가 생성되지 않았습니다.";
        var componentRendered = summary.StatusCounts.GetValueOrDefault("COMPONENTS_RENDERED");
        var mismatch = summary.StatusCounts.GetValueOrDefault("CONTRACT_MISMATCH");
        var review = summary.StatusCounts.GetValueOrDefault("STRUCTURE_READY") + summary.StatusCounts.GetValueOrDefault("REVIEW_RENDERED") + summary.StatusCounts.GetValueOrDefault("NEEDS_REVIEW") + mismatch;
        return $"AI 그룹 일괄 분석 완료 · HTML {summary.WorkbookCount:N0}개 생성 · 컴포넌트 계약 렌더 {componentRendered:N0}개 · 검토/계약 불일치 {review:N0}개";
    }

    private void Refresh_Click(object sender, RoutedEventArgs e) => LoadStoredFiles();

    // Batch-wide operation: capture, AI grouping, catalog assignment, and HTML rendering.
    // It deliberately uses the batch id carried by the visible list, never only the selection.
    private async void StartAiGroupAnalysis_Click(object sender, RoutedEventArgs e)
    {
        var batchId = _items.Select(item => item.BatchId)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (string.IsNullOrWhiteSpace(batchId))
        {
            System.Windows.MessageBox.Show(
                "먼저 --batch-folder로 구조 스캔한 배치를 열어야 합니다.",
                "AI 그룹 일괄 분석",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        AiGroupAnalysisButton.IsEnabled = false;
        FilesGrid.IsEnabled = false;
        try
        {
            await RunAiGroupAnalysisAsync(batchId);
        }
        catch (Exception ex)
        {
            Log($"AI 그룹 일괄 분석 실패: {ex.Message}");
            System.Windows.MessageBox.Show(ex.Message, "AI 그룹 일괄 분석 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            AiGroupAnalysisButton.IsEnabled = true;
            FilesGrid.IsEnabled = true;
        }
    }

    private void ShowGroupStatus_Click(object sender, RoutedEventArgs e)
    {
        var batchId = _items.Select(item => item.BatchId).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (string.IsNullOrWhiteSpace(batchId))
        {
            System.Windows.MessageBox.Show("현재 표시된 배치가 없습니다.", "AI 그룹 현황", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        ShowGroupAnalysisStatus(batchId);
    }

    private void ShowGroupAnalysisStatus(string batchId)
    {
        var batchDirectory = Path.Combine(
            _pathSettings.BatchRootDirectory,
            batchId);
        var catalogPath = Path.Combine(batchDirectory, "group-catalog.json");
        ResultTitle.Text = $"AI 그룹 현황 — {batchId}";
        PipelineFileTitle.Text = "현재 배치의 그룹 분석 상태";
        PipelineBatchText.Text = $"배치: {batchId} · 구조 스캔: {_items.Count:N0}개 파일";
        if (!File.Exists(catalogPath))
        {
            PipelineStagesText.Text = "AI 그룹: 아직 생성되지 않음\n다음 작업: AI 그룹 일괄 분석을 먼저 실행하세요.";
            NavigateUtf8Html($$"""<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:'Segoe UI','Malgun Gothic',sans-serif;margin:32px;background:#f5f7fa;color:#172b4d}.card{max-width:900px;background:#fff;border:1px solid #d8e2f0;border-radius:10px;padding:28px}.warn{padding:14px 16px;background:#fff7e6;border-left:5px solid #d97706;border-radius:4px}</style></head><body><div class="card"><h1>AI 그룹 분석이 아직 필요합니다</h1><p>구조 스캔은 <b>{{_items.Count:N0}}개</b> 파일에 대해 완료됐지만, AI 그룹 규칙과 렌더러 배정은 아직 없습니다.</p><div class="warn"><b>다음 단계</b><br>상단의 <b>AI 그룹 일괄 분석</b>을 실행하세요. 전체 배치를 기준으로 유사 양식을 묶고, 각 그룹의 데이터 추출·HTML 규칙을 만든 뒤 렌더링합니다.</div><p>단건 <code>analysis_report_*.html</code>은 이 배치 그룹 분석 결과가 아닙니다.</p></div></body></html>""");
            return;
        }

        using var catalog = JsonDocument.Parse(File.ReadAllText(catalogPath, Encoding.UTF8));
        var root = catalog.RootElement;
        var counts = root.TryGetProperty("fileAssignments", out var assignmentArray)
            ? assignmentArray.EnumerateArray().Select(item => item.GetProperty("groupId").GetString() ?? string.Empty).GroupBy(id => id, StringComparer.OrdinalIgnoreCase).ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var groups = root.TryGetProperty("groups", out var groupArray) ? groupArray.EnumerateArray().ToList() : [];
        PipelineStagesText.Text = $"AI 그룹: 완료 ({groups.Count}개)\n배정: {counts.Values.Sum():N0}개 파일";
        var rows = new StringBuilder();
        foreach (var group in groups)
        {
            var id = group.GetProperty("id").GetString() ?? string.Empty;
            var name = group.TryGetProperty("name", out var nameValue) ? nameValue.GetString() ?? id : id;
            var renderer = group.TryGetProperty("rendererKey", out var rendererValue) ? rendererValue.GetString() ?? string.Empty : string.Empty;
            var representatives = group.TryGetProperty("representativeFiles", out var reps) ? string.Join("<br>", reps.EnumerateArray().Take(3).Select(value => System.Net.WebUtility.HtmlEncode(value.GetString() ?? string.Empty))) : string.Empty;
            rows.Append($"<tr><td><b>{System.Net.WebUtility.HtmlEncode(name)}</b><br><code>{System.Net.WebUtility.HtmlEncode(id)}</code></td><td>{counts.GetValueOrDefault(id):N0}</td><td><code>{System.Net.WebUtility.HtmlEncode(renderer)}</code></td><td>{representatives}</td></tr>");
        }
        NavigateUtf8Html($$"""<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:'Segoe UI','Malgun Gothic',sans-serif;margin:24px;background:#f5f7fa;color:#172b4d}table{width:100%;border-collapse:collapse;background:#fff;font-size:13px}th,td{border:1px solid #d8e2f0;padding:10px;vertical-align:top;text-align:left}th{background:#eaf1f8}code{font-family:Consolas;font-size:11px}</style></head><body><h1>AI 그룹 현황</h1><p>배치: <code>{{System.Net.WebUtility.HtmlEncode(batchId)}}</code> · 그룹 {{groups.Count}}개 · 배정 {{counts.Values.Sum():N0}}개 파일</p><table><thead><tr><th>그룹</th><th>파일 수</th><th>렌더러</th><th>대표 양식</th></tr></thead><tbody>{{rows}}</tbody></table></body></html>""");
    }

    private void AddFiles(IEnumerable<string> files)
    {
        foreach (var file in files.Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase))
            if (_items.All(item => !string.Equals(item.FullPath, file, StringComparison.OrdinalIgnoreCase))) _items.Add(ReadStatus(file));
        Log($"목록: {_items.Count:N0}개 Excel 파일");
    }

    private void LoadStoredFiles()
    {
        var existing = _items.Select(item => item.FullPath).ToList();
        _items.Clear();
        AddFiles(existing);
    }

    // Argument-driven batch runs have no manually added files. Rebuild the left list
    // from the batch root and join only batch-local artifacts by relative path.
    private void LoadBatchPipeline(string batchDirectory, string aiState)
    {
        var batchPath = Path.Combine(batchDirectory, "batch.json");
        var classificationPath = Path.Combine(batchDirectory, "classification.csv");
        if (!File.Exists(batchPath) || !File.Exists(classificationPath)) return;

        using var batch = JsonDocument.Parse(File.ReadAllText(batchPath, Encoding.UTF8));
        var rootPath = batch.RootElement.GetProperty("rootPath").GetString();
        if (string.IsNullOrWhiteSpace(rootPath)) return;

        var capture = ReadCaptureStages(Path.Combine(batchDirectory, "numeric-capture.sqlite"));
        var groups = ReadGroupAssignments(Path.Combine(batchDirectory, "group-catalog.json"));
        var rendered = ReadRenderStages(Path.Combine(batchDirectory, "group-render-index.json"));
        _items.Clear();
        foreach (var row in ReadBatchClassificationRows(classificationPath))
        {
            var item = new ExcelItem(Path.GetFullPath(Path.Combine(rootPath, row.RelativePath)))
            {
                BatchId = Path.GetFileName(batchDirectory),
                RelativePath = row.RelativePath,
                StructureStatus = row.Status,
                StructureReason = row.Reason
            };
            if (capture.TryGetValue(row.RelativePath, out var captured))
            {
                item.CaptureStatus = captured.Status;
                item.CaptureReason = captured.Reason;
            }
            else item.CaptureStatus = string.Equals(row.Status, "SCANNED", StringComparison.OrdinalIgnoreCase) ? "대기" : "건너뜀";

            if (groups.TryGetValue(row.RelativePath, out var group))
            {
                item.GroupId = group;
                item.AiGroupStatus = group;
                item.AiReason = "AI 그룹 배정 완료";
            }
            else
            {
                item.AiGroupStatus = string.Equals(aiState, "진행 중", StringComparison.Ordinal) ? "진행 중" : "대기";
                item.AiReason = string.Equals(row.Status, "SCANNED", StringComparison.OrdinalIgnoreCase) ? "캡처 완료 파일만 AI 계획에 배정됩니다." : row.Reason;
            }

            if (rendered.TryGetValue(row.RelativePath, out var render))
            {
                item.RenderStatus = render.Status;
                item.RenderReason = render.Reason;
                item.ResultHtmlPath = Path.Combine(batchDirectory, render.ReportPath.Replace('/', Path.DirectorySeparatorChar));
            }
            else item.RenderStatus = string.Equals(aiState, "진행 중", StringComparison.Ordinal) ? "대기" : "대기";
            _items.Add(item);
        }
        Log($"배치 목록: {_items.Count:N0}개 Excel 파일 ({Path.GetFileName(batchDirectory)})");
    }

    private static IEnumerable<BatchClassificationRow> ReadBatchClassificationRows(string path)
    {
        using var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        var header = ParseCsvRow(reader.ReadLine() ?? string.Empty);
        var columns = header.Select((name, index) => (name, index)).ToDictionary(value => value.name, value => value.index, StringComparer.OrdinalIgnoreCase);
        while (reader.ReadLine() is { } line)
        {
            var values = ParseCsvRow(line);
            string Value(string name) => columns.TryGetValue(name, out var index) && index < values.Count ? values[index] : string.Empty;
            yield return new BatchClassificationRow(Value("relativePath"), Value("status"), Value("warningOrError"));
        }
    }

    private static Dictionary<string, PipelineStage> ReadCaptureStages(string databasePath)
    {
        var result = new Dictionary<string, PipelineStage>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(databasePath)) return result;
        using var connection = new SqliteConnection($"Data Source={databasePath};Mode=ReadOnly");
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT relative_path, capture_status, COALESCE(error_text, '') FROM capture_workbooks;";
        using var reader = command.ExecuteReader();
        while (reader.Read()) result[reader.GetString(0)] = new PipelineStage(reader.GetString(1), reader.GetString(2));
        return result;
    }

    private static Dictionary<string, string> ReadGroupAssignments(string catalogPath)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(catalogPath)) return result;
        using var catalog = JsonDocument.Parse(File.ReadAllText(catalogPath, Encoding.UTF8));
        if (!catalog.RootElement.TryGetProperty("fileAssignments", out var assignments)) return result;
        foreach (var assignment in assignments.EnumerateArray())
        {
            var path = assignment.GetProperty("relativePath").GetString();
            var group = assignment.GetProperty("groupId").GetString();
            if (!string.IsNullOrWhiteSpace(path) && !string.IsNullOrWhiteSpace(group)) result[path] = group;
        }
        return result;
    }

    private static Dictionary<string, RenderStage> ReadRenderStages(string indexPath)
    {
        var result = new Dictionary<string, RenderStage>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(indexPath)) return result;
        using var index = JsonDocument.Parse(File.ReadAllText(indexPath, Encoding.UTF8));
        if (!index.RootElement.TryGetProperty("rows", out var rows)) return result;
        foreach (var row in rows.EnumerateArray())
        {
            var path = row.GetProperty("RelativePath").GetString();
            var status = row.GetProperty("Status").GetString();
            var report = row.GetProperty("ReportPath").GetString();
            if (!string.IsNullOrWhiteSpace(path) && !string.IsNullOrWhiteSpace(status) && !string.IsNullOrWhiteSpace(report))
                result[path] = new RenderStage(status, report, row.TryGetProperty("ProfileVersion", out var profile) ? profile.GetString() ?? string.Empty : string.Empty);
        }
        return result;
    }

    private ExcelItem ReadStatus(string fullPath)
    {
        var item = new ExcelItem(fullPath);
        if (!File.Exists(_databasePath)) return item;
        try
        {
            using var connection = new SqliteConnection($"Data Source={_databasePath};Mode=ReadOnly");
            connection.Open();
            using var command = connection.CreateCommand();
            command.CommandText = """
                SELECT w.status, ar.overall_status, ar.overall_decision, ar.dashboard_html_path, ar.manifest_path
                FROM workbooks w
                LEFT JOIN analysis_reports ar ON ar.workbook_id=w.workbook_id AND ar.overall_status <> 'STALE'
                WHERE w.dataset='InputDataFinish' AND w.source_path=$source
                ORDER BY ar.analysis_report_id DESC LIMIT 1;
                """;
            command.Parameters.AddWithValue("$source", fullPath);
            using var reader = command.ExecuteReader();
            if (!reader.Read()) return item;
            item.DbStatus = reader.GetString(0);
            item.AnalysisStatus = reader.IsDBNull(1) ? "미분석" : $"{reader.GetString(1)} / {reader.GetString(2)}";
            item.ResultHtmlPath = ResolvePreferredDashboardPath(
                fullPath,
                reader.IsDBNull(3) ? null : reader.GetString(3),
                reader.IsDBNull(4) ? null : reader.GetString(4));
            item.Progress = reader.IsDBNull(1) ? "DB 적재 완료" : "분석 완료";
        }
        catch (Exception ex) { item.Progress = "상태 조회 오류"; Log(ex.Message); }
        return item;
    }

    private async void AnalyzeSelected_Click(object sender, RoutedEventArgs e)
    {
        await AnalyzeSelectedAsync(forceAiDraft: false);
    }

    private async void ForceAnalyzeSelected_Click(object sender, RoutedEventArgs e)
    {
        if (FilesGrid.SelectedItems.Count == 0) return;
        var confirmation = System.Windows.MessageBox.Show(
            "선택한 Excel 파일을 강제 재분석하시겠습니까?\n\n"
            + "• com-index --force로 DB를 다시 적재합니다.\n"
            + "• 새 분석 초안을 생성합니다.\n"
            + "• 기존 큐레이션 CLI 보고서와 산출물은 보존됩니다.\n"
            + "• 다만 DB 재적재로 기존 분석 상태는 STALE로 바뀔 수 있습니다.",
            "강제 재분석 확인",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Warning);
        if (confirmation != MessageBoxResult.OK) return;
        await AnalyzeSelectedAsync(forceAiDraft: true);
    }

    private async Task AnalyzeSelectedAsync(bool forceAiDraft)
    {
        var selected = FilesGrid.SelectedItems.Cast<ExcelItem>().ToList();
        await AnalyzeItemsAsync(selected, forceAiDraft);
    }

    private async Task AnalyzeItemsAsync(IReadOnlyList<ExcelItem> selected, bool forceAiDraft)
    {
        if (selected.Count == 0) return;
        FilesGrid.IsEnabled = false;
        try
        {
            foreach (var item in selected)
            {
                using var runLog = CreateRunLog(item.FullPath);
                try
                {
                    Log($"실행 로그: {runLog.FilePath}");
                    item.Progress = "DB 적재 중"; FilesGrid.Items.Refresh();
                    if (forceAiDraft)
                        await RunPythonAsync(runLog, "inference_data_ai_cli.py", "com-index", "--input", item.FullPath, "--dataset", "InputDataFinish", "--db", _databasePath, "--covered-cell-mode", "blank", "--verify-after-import", "--include-hidden", "--force");
                    else
                        await RunPythonAsync(runLog, "inference_data_ai_cli.py", "com-index", "--input", item.FullPath, "--dataset", "InputDataFinish", "--db", _databasePath, "--covered-cell-mode", "blank", "--verify-after-import", "--include-hidden");
                    item.Progress = "분석 보고서 생성 중"; FilesGrid.Items.Refresh();
                    if (forceAiDraft)
                        await RunPythonAsync(runLog, "inference_data_ai_analysis_runner.py", "--service-dir", _serviceDirectory, "--db", _databasePath, "--source", item.FullPath, "--dataset", "InputDataFinish", "--replace-auto-draft", "--force-ai-draft");
                    else
                        await RunPythonAsync(runLog, "inference_data_ai_analysis_runner.py", "--service-dir", _serviceDirectory, "--db", _databasePath, "--source", item.FullPath, "--dataset", "InputDataFinish", "--replace-auto-draft");
                    runLog.WriteLine("WPF", "Workbook analysis completed successfully.");
                    runLog.FlushAndThrowIfWriteFailed();
                    ReplaceItem(item);
                }
                catch (Exception ex)
                {
                    runLog.WriteLine("WPF", $"Workbook analysis failed: {ex.Message}");
                    throw;
                }
            }
            Log(forceAiDraft ? "선택한 Excel 강제 재분석이 완료되었습니다." : "선택한 Excel 분석이 완료되었습니다.");
        }
        catch (Exception ex) { Log($"분석 실패: {ex.Message}"); System.Windows.MessageBox.Show(ex.Message, "분석 실패", MessageBoxButton.OK, MessageBoxImage.Error); }
        finally { FilesGrid.IsEnabled = true; }
    }

    private async Task RunPythonAsync(RunLog runLog, string script, params string[] arguments)
    {
        var executable = _pathSettings.PythonExecutable;
        runLog.WriteLine("WPF", $"Starting {script} {string.Join(' ', arguments)}");
        var info = new ProcessStartInfo(executable)
        {
            WorkingDirectory = _serviceDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        // StandardOutputEncoding only controls how this process decodes pipe
        // bytes. Force the Python child to emit matching UTF-8 bytes as well;
        // otherwise a Windows locale code page can produce Korean mojibake.
        info.Environment["PYTHONUTF8"] = "1";
        info.Environment["PYTHONIOENCODING"] = "utf-8";
        info.ArgumentList.Add(Path.Combine(_serviceDirectory, script));
        foreach (var argument in arguments) info.ArgumentList.Add(argument);
        using var process = Process.Start(info) ?? throw new InvalidOperationException("Python 실행을 시작하지 못했습니다.");
        var standardOutputClosed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var standardErrorClosed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        // Never synchronously invoke the UI for each subprocess output line: it can
        // produce thousands of lines and starve repaint/input while analysing.
        process.OutputDataReceived += (_, e) =>
        {
            if (e.Data is null) standardOutputClosed.TrySetResult();
            else WriteProcessLog(runLog, "OUT", e.Data);
        };
        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data is null) standardErrorClosed.TrySetResult();
            else WriteProcessLog(runLog, "ERR", e.Data);
        };
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();
        await process.WaitForExitAsync();
        await Task.WhenAll(standardOutputClosed.Task, standardErrorClosed.Task);
        runLog.WriteLine("WPF", $"{script} exited with code {process.ExitCode}.");
        runLog.FlushAndThrowIfWriteFailed();
        if (process.ExitCode != 0) throw new InvalidOperationException($"{script} 실행 실패: {process.ExitCode}");
    }

    private void WriteProcessLog(RunLog runLog, string stream, string line)
    {
        runLog.WriteLine(stream, line);
        // The delegate is called from the process pipe reader, so it must not
        // synchronously marshal every line to the Dispatcher.
        QueueLog($"{stream}: {line}");
    }

    private RunLog CreateRunLog(string workbookPath, string? directory = null)
    {
        var logDirectory = directory ?? _runLogDirectory;
        Directory.CreateDirectory(logDirectory);
        var workbookName = Path.GetFileNameWithoutExtension(workbookPath);
        var safeWorkbookName = string.Concat(workbookName.Select(character =>
            Path.GetInvalidFileNameChars().Contains(character) ? '_' : character));
        if (string.IsNullOrWhiteSpace(safeWorkbookName)) safeWorkbookName = "workbook";
        if (safeWorkbookName.Length > 80) safeWorkbookName = safeWorkbookName[..80];
        var fileName = $"{DateTime.UtcNow:yyyyMMddTHHmmssfffZ}_{safeWorkbookName}_{Guid.NewGuid():N}.log";
        return new RunLog(Path.Combine(logDirectory, fileName));
    }

    private string? ResolvePreferredDashboardPath(string sourcePath, string? generatedDashboardPath, string? latestManifestPath)
    {
        // Curated CLI dashboards are more complete than the generic renderer.
        // Use one only when the latest DB report is curated and its manifest
        // names this exact workbook path. A runner draft must display its own
        // newly-generated dashboard after a forced reanalysis.
        if (IsCuratedManifest(latestManifestPath))
        {
            var curatedPath = FindManifestCuratedDashboard(sourcePath);
            if (curatedPath is not null) return curatedPath;
        }
        return !string.IsNullOrWhiteSpace(generatedDashboardPath) && File.Exists(generatedDashboardPath)
            ? generatedDashboardPath
            : null;
    }

    private static bool IsCuratedManifest(string? manifestPath) =>
        !string.IsNullOrWhiteSpace(manifestPath) &&
        !Path.GetFileNameWithoutExtension(manifestPath).StartsWith("workbook_", StringComparison.OrdinalIgnoreCase);

    private string? FindManifestCuratedDashboard(string sourcePath)
    {
        var manifestsDirectory =
            _pathSettings.AnalysisManifestDirectory;
        if (!Directory.Exists(manifestsDirectory)) return null;

        try
        {
            foreach (var manifestPath in Directory.EnumerateFiles(manifestsDirectory, "*.json", SearchOption.TopDirectoryOnly)
                         .OrderByDescending(File.GetLastWriteTimeUtc)
                         .ThenBy(path => path, StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    using var manifest = JsonDocument.Parse(File.ReadAllText(manifestPath, Encoding.UTF8));
                    var root = manifest.RootElement;
                    if (!TryGetString(root, out var manifestSourcePath, "source", "sourcePath") ||
                        !string.Equals(manifestSourcePath, sourcePath, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!TryGetString(root, out var htmlPath, "report", "artifacts", "html")) continue;

                    var resolvedHtmlPath = ResolveManifestArtifactPath(htmlPath);
                    if (resolvedHtmlPath is not null) return resolvedHtmlPath;
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
                catch (JsonException) { }
            }
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
        return null;
    }

    private string? ResolveManifestArtifactPath(string artifactPath)
    {
        if (string.IsNullOrWhiteSpace(artifactPath)) return null;
        var outputsDirectory = Path.GetFullPath(
            _pathSettings.OutputRootDirectory);
        var resolvedPath = Path.GetFullPath(Path.IsPathRooted(artifactPath)
            ? artifactPath
            : Path.Combine(_serviceDirectory, artifactPath));
        if (!resolvedPath.StartsWith(outputsDirectory + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
            !File.Exists(resolvedPath)) return null;
        return resolvedPath;
    }

    private static bool TryGetString(JsonElement element, out string value, params string[] path)
    {
        foreach (var propertyName in path)
        {
            if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(propertyName, out element))
            {
                value = string.Empty;
                return false;
            }
        }
        value = element.ValueKind == JsonValueKind.String ? element.GetString() ?? string.Empty : string.Empty;
        return !string.IsNullOrWhiteSpace(value);
    }

    private void ReplaceItem(ExcelItem item)
    {
        var index = _items.IndexOf(item); _items[index] = ReadStatus(item.FullPath); FilesGrid.SelectedItem = _items[index];
    }

    private void FilesGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (FilesGrid.SelectedItem is ExcelItem item) ShowResult(item);
    }

    private void FilesGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (FilesGrid.SelectedItem is not ExcelItem item) return;
        if (!File.Exists(item.FullPath))
        {
            Log($"원본 Excel 파일을 찾을 수 없습니다: {item.FullPath}");
            return;
        }
        Process.Start(new ProcessStartInfo(item.FullPath) { UseShellExecute = true });
        Log($"원본 Excel 열기: {item.FileName}");
    }

    private void OpenResult_Click(object sender, RoutedEventArgs e)
    {
        if (FilesGrid.SelectedItem is ExcelItem item) ShowResult(item);
    }

    private void OpenSourceExcel_Click(object sender, RoutedEventArgs e)
    {
        if (FilesGrid.SelectedItem is not ExcelItem item) return;
        if (!File.Exists(item.FullPath)) { Log($"원본 Excel 파일을 찾을 수 없습니다: {item.FullPath}"); return; }
        Process.Start(new ProcessStartInfo(item.FullPath) { UseShellExecute = true });
    }

    private void OpenResultHtml_Click(object sender, RoutedEventArgs e)
    {
        if (FilesGrid.SelectedItem is not ExcelItem item) return;
        if (string.IsNullOrWhiteSpace(item.ResultHtmlPath) || !File.Exists(item.ResultHtmlPath)) { Log($"생성된 HTML 보고서가 없습니다: {item.FileName}"); return; }
        Process.Start(new ProcessStartInfo(item.ResultHtmlPath) { UseShellExecute = true });
    }

    private void ShowResult(ExcelItem item)
    {
        ResultTitle.Text = $"분석 결과 — {item.FileName}";
        PipelineFileTitle.Text = item.FileName;
        PipelineBatchText.Text = string.IsNullOrWhiteSpace(item.BatchId) ? "개별 분석 목록" : $"배치: {item.BatchId} · 그룹: {item.AiGroupStatus}";
        PipelineStagesText.Text = string.Join(Environment.NewLine,
            $"1. 구조 스캔  {item.StructureStatus}  {item.StructureReason}",
            $"2. 전체 캡처  {item.CaptureStatus}  {item.CaptureReason}",
            $"3. AI 그룹    {item.AiGroupStatus}  {item.AiReason}",
            $"4. 추출 규칙  {item.ExtractionStatus}  {item.ExtractionReason}",
            $"5. HTML 렌더  {item.RenderStatus}  {item.RenderReason}");
        if (!string.IsNullOrWhiteSpace(item.ResultHtmlPath) && File.Exists(item.ResultHtmlPath))
        {
            NavigateUtf8Html(File.ReadAllText(item.ResultHtmlPath, Encoding.UTF8));
        }
        else NavigateUtf8Html("<html><body style='font-family:Segoe UI;padding:24px'><h2>아직 분석 결과가 없습니다.</h2><p>목록에서 우클릭한 뒤 ‘선택 파일 분석 시작’을 선택하세요.</p></body></html>");
    }

    private void NavigateUtf8Html(string html)
    {
        // WPF's NavigateToString internally serializes text with StreamWriter's
        // default UTF-8 encoding without a BOM. The hosted legacy WebBrowser can
        // then choose the Windows ANSI code page before honoring a page's meta
        // charset. Supply an explicit UTF-8 BOM in the document stream instead.
        html = ApplyWorkHubDocumentTheme(html);
        var documentStream = new MemoryStream();
        using (var writer = new StreamWriter(documentStream, Utf8WithBom, 16_384, leaveOpen: true))
            writer.Write(html);
        documentStream.Position = 0;
        var previousStream = _activeHtmlStream;
        try
        {
            ResultBrowser.Visibility = Visibility.Hidden;
            ResultBrowserPlaceholder.Visibility = Visibility.Visible;
            ResultBrowser.NavigateToStream(documentStream);
            _activeHtmlStream = documentStream;
            previousStream?.Dispose();
        }
        catch
        {
            documentStream.Dispose();
            throw;
        }
    }

    private void ResultBrowser_LoadCompleted(
        object sender,
        System.Windows.Navigation.NavigationEventArgs e)
    {
        ResultBrowserPlaceholder.Visibility = Visibility.Collapsed;
        ResultBrowser.Visibility = Visibility.Visible;
    }

    private static string ApplyWorkHubDocumentTheme(string html)
    {
        const string marker = "data-workhub-theme";
        if (html.Contains(marker, StringComparison.OrdinalIgnoreCase))
        {
            return html;
        }

        const string theme = """
            <style data-workhub-theme>
            html{color-scheme:dark;background:#202020!important}
            body{font-family:'Segoe UI','Malgun Gothic',sans-serif!important;background:#202020!important;color:#d6d6d6!important}
            h1,h2,h3,h4,strong,b{color:#eeeeee!important}
            a{color:#c4b5fd!important}
            section,article,.card,.panel,.summary,.answer-card,.finding,.intent-grid,pre{background:#252525!important;color:#d6d6d6!important;border-color:#3a3a3a!important;box-shadow:none!important}
            table{background:#252525!important;color:#d6d6d6!important;border-color:#3a3a3a!important}
            th{background:#303030!important;color:#eeeeee!important;border-color:#3a3a3a!important}
            td{background:#252525!important;color:#d6d6d6!important;border-color:#3a3a3a!important}
            tbody tr:nth-child(even) td{background:#292929!important}
            code{color:#c4b5fd!important}
            .muted,.note,.eyebrow,small{color:#8e8e8e!important}
            .chip{background:#3c2a58!important;color:#d8c7ff!important}
            </style>
            """;
        var headEnd = html.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
        if (headEnd >= 0)
        {
            return html.Insert(headEnd, theme);
        }

        var bodyStart = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
        return bodyStart >= 0 ? html.Insert(bodyStart, theme) : theme + html;
    }

    private void Log(string text)
    {
        QueueLog(text);
    }

    private void QueueLog(string text)
    {
        if (Interlocked.Increment(ref _pendingVisualLogLineCount) <= MaxPendingVisualLogLines)
        {
            _pendingLogs.Enqueue(text);
            return;
        }

        Interlocked.Decrement(ref _pendingVisualLogLineCount);
        Interlocked.Increment(ref _droppedVisualLogLineCount);
    }

    private void FlushPendingLogs()
    {
        var droppedLineCount = Interlocked.Exchange(ref _droppedVisualLogLineCount, 0);
        if (_pendingLogs.IsEmpty && droppedLineCount == 0) return;

        var buffer = new StringBuilder();
        if (droppedLineCount > 0)
            buffer.Append('[').Append(DateTime.Now.ToString("HH:mm:ss")).Append("] UI 로그 ")
                .Append(droppedLineCount.ToString("N0")).AppendLine("줄을 생략했습니다. 전체 출력은 실행 로그 파일에 보존됩니다.");

        for (var count = 0; count < 150 && _pendingLogs.TryDequeue(out var line); count++)
        {
            Interlocked.Decrement(ref _pendingVisualLogLineCount);
            buffer.Append('[').Append(DateTime.Now.ToString("HH:mm:ss")).Append("] ").AppendLine(line);
        }
        if (buffer.Length == 0) return;

        LogText.AppendText(buffer.ToString());
        TrimVisualLog();
        LogText.ScrollToEnd();
    }

    private void TrimVisualLog()
    {
        var currentText = LogText.Text;
        if (currentText.Length <= MaxVisualLogCharacters) return;

        var keepFrom = Math.Max(0, currentText.Length - RetainedVisualLogCharacters);
        var nextLine = currentText.IndexOf('\n', keepFrom);
        if (nextLine >= 0) keepFrom = nextLine + 1;
        else if (keepFrom < currentText.Length && char.IsLowSurrogate(currentText[keepFrom])) keepFrom++;

        LogText.Text = "[이전 UI 로그는 화면 크기 제한으로 정리되었습니다. 전체 출력은 실행 로그 파일에 보존됩니다.]\r\n"
            + currentText[keepFrom..];
    }

    private sealed class RunLog : IDisposable
    {
        private readonly object _sync = new();
        private FileStream? _stream;
        private StreamWriter? _writer;
        private Exception? _writeFailure;

        public RunLog(string filePath)
        {
            FilePath = filePath;
            _stream = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write, FileShare.Read, 16_384, FileOptions.SequentialScan);
            _writer = new StreamWriter(_stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), 16_384);
            try
            {
                WriteLine("WPF", "Run log created.");
                ThrowIfWriteFailed();
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        public string FilePath { get; }

        public void WriteLine(string source, string text)
        {
            lock (_sync)
            {
                if (_writeFailure is not null || _writer is null) return;
                try
                {
                    _writer.Write('[');
                    _writer.Write(DateTime.UtcNow.ToString("O"));
                    _writer.Write("] [");
                    _writer.Write(source);
                    _writer.Write("] ");
                    _writer.WriteLine(text);
                }
                catch (Exception ex)
                {
                    _writeFailure ??= ex;
                }
            }
        }

        public void ThrowIfWriteFailed()
        {
            lock (_sync)
            {
                if (_writeFailure is not null)
                    throw new IOException($"실행 로그를 쓸 수 없습니다: {FilePath}", _writeFailure);
            }
        }

        public void FlushAndThrowIfWriteFailed()
        {
            lock (_sync)
            {
                if (_writeFailure is null)
                {
                    try
                    {
                        _writer?.Flush();
                        _stream?.Flush(flushToDisk: true);
                    }
                    catch (Exception ex)
                    {
                        _writeFailure = ex;
                    }
                }
                if (_writeFailure is not null)
                    throw new IOException($"실행 로그를 쓸 수 없습니다: {FilePath}", _writeFailure);
            }
        }

        public void Dispose()
        {
            lock (_sync)
            {
                try
                {
                    _writer?.Flush();
                    _stream?.Flush(flushToDisk: true);
                }
                catch (Exception ex)
                {
                    _writeFailure ??= ex;
                }
                finally
                {
                    _writer?.Dispose();
                    _writer = null;
                    _stream?.Dispose();
                    _stream = null;
                }
            }
        }
    }
}

public sealed class ExcelItem(string fullPath)
{
    public string FullPath { get; } = fullPath;
    public string FileName => Path.GetFileName(FullPath);
    public string DbStatus { get; set; } = "미적재";
    public string AnalysisStatus { get; set; } = "미분석";
    public string Progress { get; set; } = "대기";
    public string? ResultHtmlPath { get; set; }
    public string BatchId { get; set; } = string.Empty;
    public string RelativePath { get; set; } = string.Empty;
    public string StructureStatus { get; set; } = "대기";
    public string StructureReason { get; set; } = string.Empty;
    public string CaptureStatus { get; set; } = "대기";
    public string CaptureReason { get; set; } = string.Empty;
    public string AiGroupStatus { get; set; } = "대기";
    public string AiReason { get; set; } = string.Empty;
    public string GroupId { get; set; } = string.Empty;
    public string ExtractionStatus { get; set; } = "대기";
    public string ExtractionReason { get; set; } = "AI 그룹 규칙 생성 후 실행";
    public string RenderStatus { get; set; } = "대기";
    public string RenderReason { get; set; } = string.Empty;
}
