using System.Collections.ObjectModel;
using System.Diagnostics;
using System.ComponentModel;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;
using Microsoft.Data.Sqlite;
using Brush = System.Windows.Media.Brush;
using Color = System.Windows.Media.Color;
using SolidColorBrush = System.Windows.Media.SolidColorBrush;
using Forms = System.Windows.Forms;
using WpfMessageBox = System.Windows.MessageBox;

namespace InferenceDataAIService.Wpf;

public partial class MainWindow
{
    private readonly ObservableCollection<EvidenceCitationRow>
        _evidenceCitations = [];
    private readonly ObservableCollection<ReviewQueueRow>
        _reviewQueue = [];
    private readonly ObservableCollection<ReviewEvidenceRow>
        _reviewEvidence = [];
    private readonly ObservableCollection<WorkbookComparisonSummary>
        _workbookComparisons = [];
    private readonly ObservableCollection<ObservationComparisonRow>
        _workbookComparisonRows = [];
    private readonly ObservableCollection<ConceptCandidateRow>
        _conceptCandidates = [];
    private readonly ObservableCollection<CanonicalConceptRow>
        _canonicalConcepts = [];
    private readonly ObservableCollection<IngestStageViewModel>
        _ingestStages = [];
    private readonly ObservableCollection<ExcelFolderSearchRow>
        _excelFolderSearchRows = [];
    private readonly ObservableCollection<FormPreflightRow>
        _formPreflightRows = [];
    private readonly ObservableCollection<FormFamilyGroupRow>
        _formFamilyGroups = [];
    private CanonicalEvidenceClient _canonicalEvidenceClient = null!;
    private WorkbookComparisonClient _workbookComparisonClient = null!;
    private EvidenceDetailDocument? _activeEvidenceDetail;
    private ReviewDetailDocument? _activeReviewDetail;
    private WorkbookComparisonDocument? _activeWorkbookComparison;
    private IngestWorkbookResult? _activeIngestResult;
    private RelatedStudiesDocument? _activeIngestRelated;
    private FormPreflightDocument? _activeFormPreflight;
    private FormGroupReviewDocument? _activeFormGroupReview;
    private QuestionInsightSnapshot? _questionInsightSnapshot;
    private bool _canonicalEvidenceBusy;
    private bool _workbookComparisonBusy;
    private bool _excelFolderSearchBusy;
    private bool _formPreflightRunning;
    private bool _excelFolderSearchCopyCompleted;
    private bool _canonicalConceptListCurrent;
    private bool _selectedIngestHasFailedJournal;
    private int _folderIngestTotal;
    private string _excelFolderSearchRoot = string.Empty;
    private readonly Dictionary<int, HashSet<string>>
        _folderStageCompleted = Enumerable.Range(1, 7).ToDictionary(
            number => number,
            _ => new HashSet<string>(
                StringComparer.OrdinalIgnoreCase));
    private readonly Dictionary<int, HashSet<string>>
        _folderStageFailed = Enumerable.Range(1, 7).ToDictionary(
            number => number,
            _ => new HashSet<string>(
                StringComparer.OrdinalIgnoreCase));
    private string _loadedCanonicalConceptKind = string.Empty;

    private void InitializeCanonicalEvidenceUi()
    {
        _canonicalEvidenceClient = new CanonicalEvidenceClient(
            _pathSettings);
        _workbookComparisonClient = new WorkbookComparisonClient(
            _pathSettings);
        CanonicalDbPathText.Text = _databasePath;
        ExcelArchivePathText.Text =
            _pathSettings.ExcelArchiveDirectory;
        EvidenceCitationsGrid.ItemsSource = _evidenceCitations;
        ReviewQueueGrid.ItemsSource = _reviewQueue;
        ReviewEvidenceGrid.ItemsSource = _reviewEvidence;
        WorkbookComparisonGrid.ItemsSource = _workbookComparisons;
        WorkbookComparisonValuesGrid.ItemsSource =
            _workbookComparisonRows;
        ConceptCandidatesGrid.ItemsSource = _conceptCandidates;
        CanonicalConceptsGrid.ItemsSource = _canonicalConcepts;
        ExcelFolderSearchGrid.ItemsSource =
            _excelFolderSearchRows;
        FormPreflightGrid.ItemsSource = _formPreflightRows;
        FormFamilyGroupsGrid.ItemsSource = _formFamilyGroups;
        FormPreflightArchivePathText.Text =
            ExcelLocalCopyService.GetLocalCopyBase(
                _pathSettings.ExcelArchiveDirectory);
        RefreshFormPreflightUi();
        FormFamilyReviewerText.Text = Environment.UserName;
        LoadCachedFormGroupReview();
        InitializeIngestStages();
        ConfigureArchiveIngestSource(resetStages: true);
        ReviewReviewerText.Text = Environment.UserName;
        ConceptReviewerText.Text = Environment.UserName;
        UpdateConceptDecisionAvailability();
    }

    private void InitializeIngestStages()
    {
        _ingestStages.Clear();
        var titles = new[]
        {
            "사전 분석 결과 확인",
            "COM 추출 결과 재사용",
            "구조 패킷 생성",
            "AI 시험·비교군 분석",
            "근거 검증·DB 적재",
            "DB 무결성·질문 반영",
            "AI 문의 준비",
        };
        for (var index = 0; index < titles.Length; index++)
            _ingestStages.Add(new IngestStageViewModel(
                index + 1,
                titles[index]));
        IngestStagesList.ItemsSource = _ingestStages;
    }

    private async void RefreshWorkbookComparison_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_workbookComparisonBusy || _canonicalEvidenceBusy)
            return;
        if (!File.Exists(_databasePath))
        {
            WpfMessageBox.Show(
                $"검토 DB를 찾을 수 없습니다:\n{_databasePath}",
                "Excel ↔ DB 검수",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        WorkbookComparisonSummary? firstAvailable = null;
        SetWorkbookComparisonBusy(true);
        WorkbookComparisonState.Text =
            "대표 30건과 현재 DB 상태를 연결하는 중...";
        try
        {
            var rows = await _workbookComparisonClient.ListAsync(
                _databasePath);
            _workbookComparisons.Clear();
            _workbookComparisonRows.Clear();
            _activeWorkbookComparison = null;
            foreach (var row in rows)
                _workbookComparisons.Add(row);

            firstAvailable = rows.FirstOrDefault(
                row => row.IsAvailable);
            var availableCount = rows.Count(
                row => row.IsAvailable);
            var missingCount = rows.Count - availableCount;
            WorkbookComparisonState.Text =
                $"대표 {rows.Count:N0}건 · DB 연결 "
                + $"{availableCount:N0}건"
                + (missingCount == 0
                    ? " · 파일을 선택하세요."
                    : $" · DB 없음 {missingCount:N0}건");
            Log(
                $"대표 검수 목록 로드: {rows.Count}건, "
                + $"DB 연결 {availableCount}건");
        }
        catch (Exception exception)
        {
            WorkbookComparisonState.Text = "대표 30건 조회 실패";
            Log($"대표 검수 목록 조회 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "Excel ↔ DB 검수 조회 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetWorkbookComparisonBusy(false);
        }

        if (firstAvailable is not null)
            WorkbookComparisonGrid.SelectedItem = firstAvailable;
    }

    private async void WorkbookComparisonGrid_SelectionChanged(
        object sender,
        SelectionChangedEventArgs e)
    {
        if (_workbookComparisonBusy || _canonicalEvidenceBusy)
            return;
        if (WorkbookComparisonGrid.SelectedItem
            is not WorkbookComparisonSummary summary)
            return;

        if (!summary.IsAvailable)
        {
            _activeWorkbookComparison = null;
            _workbookComparisonRows.Clear();
            WorkbookComparisonState.Text =
                $"#{summary.BenchmarkNumber} {summary.FileName} · DB 분석 없음";
            return;
        }

        SetWorkbookComparisonBusy(true);
        WorkbookComparisonState.Text =
            $"#{summary.BenchmarkNumber} {summary.FileName}을 읽는 중...";
        try
        {
            var document = await _workbookComparisonClient.LoadAsync(
                _databasePath,
                summary.PublicAnalysisId);
            _activeWorkbookComparison = document;
            _workbookComparisonRows.Clear();
            foreach (var row in document.Observations)
                _workbookComparisonRows.Add(row);

            var matched = document.Observations.Count(
                row => row.ComparisonStatus == "일치");
            var problemCount = document.Observations.Count - matched;
            WorkbookComparisonState.Text =
                $"#{summary.BenchmarkNumber} · 전체 "
                + $"{document.Observations.Count:N0}건 · 문제 없음 "
                + $"{matched:N0}건 · 문제 {problemCount:N0}건";
            ResultTitle.Text = "Excel ↔ DB 값 검수";
            PipelineFileTitle.Text = summary.FileName;
            PipelineBatchText.Text =
                $"대표 #{summary.BenchmarkNumber} · "
                + summary.PublicAnalysisId;
            PipelineStagesText.Text =
                "왼쪽의 문제, Excel 실제값, DB 인식값 세 열만 확인합니다.";
            NavigateUtf8Html(
                $$"""
                 <!doctype html>
                 <html lang="ko"><head><meta charset="utf-8">
                 <style>
                 body{font-family:'Segoe UI','Malgun Gothic',sans-serif;padding:28px;color:#172033;background:#f8fafc}
                 .card{padding:22px;background:white;border:1px solid #d5dfeb;border-radius:8px}
                 </style></head><body><div class="card">
                 <h1>{{System.Net.WebUtility.HtmlEncode(summary.FileName)}}</h1>
                 <p>왼쪽 표의 세 열만 확인하세요.</p>
                 <p><strong>전체 {{document.Observations.Count:N0}}건 · 문제 없음 {{matched:N0}}건 · 문제 {{problemCount:N0}}건</strong></p>
                 <p>문제가 있는 행을 더블클릭하면 원본 Excel 셀로 이동합니다.</p>
                 </div></body></html>
                 """);
            Log(
                $"검수 상세 로드: {summary.PublicAnalysisId}, "
                + $"Study {document.Studies.Count}, "
                + $"값 {document.Observations.Count}");
        }
        catch (Exception exception)
        {
            WorkbookComparisonState.Text = "검수 상세 조회 실패";
            Log($"검수 상세 조회 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "Excel ↔ DB 검수 상세 조회 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetWorkbookComparisonBusy(false);
        }
    }

    private async void OpenWorkbookComparisonExcel_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_workbookComparisonBusy || _canonicalEvidenceBusy)
            return;
        if (!TryGetSelectedWorkbookComparisonRow(out var row))
            return;
        if (string.IsNullOrWhiteSpace(row.Sheet)
            || string.IsNullOrWhiteSpace(row.Range))
        {
            WpfMessageBox.Show(
                "이 값에는 원본 Excel 위치가 연결되어 있지 않습니다.",
                "원본 Excel 셀",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        SetWorkbookComparisonBusy(true);
        WorkbookComparisonState.Text =
            $"Excel {row.Sheet}!{row.Range}을 여는 중...";
        try
        {
            await ExcelRangeNavigator.OpenReadOnlyAsync(
                row.SourcePath,
                row.Sheet,
                row.Range);
            WorkbookComparisonState.Text =
                $"원본 Excel · {row.Sheet}!{row.Range}";
            Log(
                $"검수 원본 Excel 열기: {row.SourcePath} / "
                + $"{row.Sheet}!{row.Range}");
        }
        catch (Exception exception)
        {
            WorkbookComparisonState.Text = "원본 Excel 열기 실패";
            Log($"검수 원본 Excel 열기 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "원본 Excel 셀 열기 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetWorkbookComparisonBusy(false);
        }
    }

    private void WorkbookComparisonValuesGrid_MouseDoubleClick(
        object sender,
        MouseButtonEventArgs e) =>
        OpenWorkbookComparisonExcel_Click(sender, e);

    private bool TryGetSelectedWorkbookComparisonRow(
        out ObservationComparisonRow row)
    {
        if (WorkbookComparisonValuesGrid.SelectedItem
            is ObservationComparisonRow selected)
        {
            row = selected;
            return true;
        }

        row = null!;
        WpfMessageBox.Show(
            "먼저 아래 수치 대조 행을 선택하세요.",
            "Excel ↔ DB 검수",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
        return false;
    }

    private void SetWorkbookComparisonBusy(bool busy)
    {
        _workbookComparisonBusy = busy;
        var enabled = !busy && !_canonicalEvidenceBusy;
        RefreshWorkbookComparisonButton.IsEnabled = enabled;
        OpenWorkbookComparisonExcelButton.IsEnabled = enabled;
        WorkbookComparisonGrid.IsEnabled = enabled;
        WorkbookComparisonValuesGrid.IsEnabled = enabled;
    }

    private async void AskEvidence_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy) return;
        var question = EvidenceQuestionText.Text.Trim();
        if (question.Length == 0)
        {
            WpfMessageBox.Show(
                "질문을 입력하세요. 특정 도메인이나 VP+CD로 제한되지 않습니다.",
                "검토 DB 질문",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        var queryMode = (
            EvidenceQueryModeSelector.SelectedItem as ComboBoxItem
        )?.Tag?.ToString() ?? "canonical";
        var historyDatabasePath = _canonicalEvidenceClient.HistoryDatabasePath();
        var selectedDatabasePath = string.Equals(
            queryMode,
            "history",
            StringComparison.OrdinalIgnoreCase)
            ? historyDatabasePath
            : _databasePath;
        if (!File.Exists(selectedDatabasePath))
        {
            WpfMessageBox.Show(
                "선택한 질문 DB를 찾을 수 없습니다:\n"
                + selectedDatabasePath,
                "검토 DB 질문",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        SetCanonicalEvidenceBusy(true);
        DeveloperConsoleExpander.IsExpanded = true;
        EvidenceQueryState.Text =
            "DB 후보를 수집한 뒤 AI가 질문에 필요한 보고서만 선별 중...";
        Log($"[AI 문서 관련성] 질문: {question}");
        Log("[AI 문서 관련성] DB 후보 수집 및 관련 보고서 판정을 시작합니다.");
        try
        {
            var queryTask = _canonicalEvidenceClient.AskAsync(
                _databasePath,
                question,
                queryMode);
            var elapsed = Stopwatch.StartNew();
            while (!queryTask.IsCompleted)
            {
                var completed = await Task.WhenAny(
                    queryTask,
                    Task.Delay(TimeSpan.FromSeconds(5)));
                if (completed == queryTask) break;
                var progressMessage =
                    $"[2/4] DB 후보 수집·AI 문서 관련성 판정 중 · "
                    + $"경과 {elapsed.Elapsed:mm\\:ss}";
                EvidenceQueryState.Text = progressMessage;
                Log($"[AI 진행] {progressMessage}");
            }
            var session = await queryTask;
            elapsed.Stop();
            Log(
                "[3/4] AI 관련성 응답 수신 및 스키마 검증 완료 · "
                + $"경과 {elapsed.Elapsed:mm\\:ss}");
            _evidenceCitations.Clear();
            foreach (var citation in session.Citations)
                _evidenceCitations.Add(citation);
            _activeEvidenceDetail = null;
            var isHistoryAnswer = session.AnswerStatus.Contains(
                "HISTORY",
                StringComparison.OrdinalIgnoreCase);
            var isContextualAnswer = session.IsContextual;
            var isRelevanceAnswer = session.IsRelevance;
            EvidenceQueryState.Text = isRelevanceAnswer
                ? $"{session.AnswerStatus} · 관련 보고서 "
                  + $"{session.RelevantStudyCount:N0}건 · 원본 근거 "
                  + $"{session.Citations.Count:N0}건"
                : isContextualAnswer
                ? $"{session.AnswerStatus} · 관련 Study "
                  + $"{session.RelevantStudyCount:N0}건 · 관련 원본 근거 "
                  + $"{session.Citations.Count:N0}건"
                : isHistoryAnswer
                ? $"{session.AnswerStatus} · 관련 이력 Study "
                  + $"{session.RelevantStudyCount:N0}건"
                : $"{session.AnswerStatus} · 관련 DATA "
                  + $"{session.RelevantStudyCount:N0}건 · 정량 효과 "
                  + $"{session.EligibleEffectCount:N0}건";
            ResultTitle.Text = isRelevanceAnswer
                ? "질문 관련 보고서"
                : isContextualAnswer
                ? "문맥 AI 근거 답변"
                : isHistoryAnswer
                ? "전체 시험 이력 근거 답변"
                : "검토 DB 근거 답변";
            PipelineFileTitle.Text = question;
            PipelineBatchText.Text = isRelevanceAnswer
                ? "DB 후보 수집 → AI 문서 관련성 판정 → 관련 Study·원본 범위 취합"
                : isContextualAnswer
                ? "키워드 후보 수집 → AI 문맥·관계 판정 → 수치 근거 검증 → 간결한 답변"
                : isHistoryAnswer
                ? "전체 table-first Study 검색 → 시간순 이력 → 근거 제한 답변"
                : "범용 의미 검색 → 비교 적격성 검증 → 결정론적 한국어 답변";
            PipelineStagesText.Text = isRelevanceAnswer
                ? "AI는 질문에 필요한 보고서인지 여부만 판단합니다. 결과 수치의 의미·증감·우열·효과·원인은 판단하지 않습니다.\n"
                  + "선택된 보고서의 시험 조건·수집 지표·원본 Excel 시트와 범위를 사람이 확인할 수 있게 취합합니다."
                : isContextualAnswer
                ? "단어가 겹치는 자료를 그대로 답변에 넣지 않습니다. AI가 대상·조건·지표의 관계를 확인한 Study 전체와 원본 근거를 표시합니다.\n"
                  + "답변의 숫자는 검증된 핵심 fact만 사용하며, 추이·인과 조건이 부족하면 한계를 별도로 명시합니다."
                : isHistoryAnswer
                ? "전체 분석 이력은 검색하되 NEEDS_REVIEW 자료로 승인된 수치 효과나 인과 결론을 만들지 않습니다.\n"
                  + "각 이력에는 원본 파일·시트·범위 TF-EVD 근거가 연결됩니다."
                : "수치 관계 문장은 VERIFIED 대조군·비교군 효과와 직접 EVD 근거가 있을 때만 생성됩니다.\n"
                  + "서술 자료와 제외 자료는 별도로 표시됩니다. 이미지 분석은 하지 않습니다.";
            NavigateQuestionInsight(
                EvidenceHtmlRenderer.RenderAnswer(session));
            if (isRelevanceAnswer)
            {
                SetInsightWideMode(true);
                DeveloperConsoleExpander.IsExpanded = false;
                LogRelevanceAiSelection(session);
            }
            Log("[4/4] 관련 보고서 표와 원본 근거 목록 표시 완료");
            Log(
                $"근거 DB 질문 완료: {session.AnswerStatus}, "
                + $"DATA {session.RelevantStudyCount}, "
                + $"효과 {session.EligibleEffectCount}");
        }
        catch (Exception exception)
        {
            EvidenceQueryState.Text = "조회 실패";
            Log($"근거 DB 조회 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "검토 DB 조회 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private void LoadEvidenceHistory_Click(
        object sender,
        RoutedEventArgs e) =>
        LoadLatestRelevanceHistory(showMissingMessage: true);

    private void LoadLatestRelevanceHistory(bool showMissingMessage)
    {
        if (_canonicalEvidenceBusy) return;
        var session = _canonicalEvidenceClient.LoadLatestRelevanceSession();
        if (session is null)
        {
            if (showMissingMessage)
            {
                WpfMessageBox.Show(
                    "저장된 관련 보고서 이력이 없습니다.",
                    "기존 이력 보기",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            return;
        }
        using var document = JsonDocument.Parse(session.RawJson);
        var root = document.RootElement;
        var question = RelevanceJsonString(root, "question");
        EvidenceQuestionText.Text = question;
        _evidenceCitations.Clear();
        foreach (var citation in session.Citations)
            _evidenceCitations.Add(citation);
        _activeEvidenceDetail = null;
        EvidenceQueryState.Text =
            $"기존 이력 · 관련 보고서 {session.RelevantStudyCount:N0}건 · "
            + $"원본 근거 {session.Citations.Count:N0}건";
        ResultTitle.Text = "질문 관련 보고서 · 기존 이력";
        PipelineFileTitle.Text = question;
        PipelineBatchText.Text =
            "저장된 AI 문서 관련성 판정 → 원본 수치·셀 범위 복원";
        PipelineStagesText.Text =
            "AI를 다시 호출하지 않고 저장된 관련 보고서 이력을 불러왔습니다.\n"
            + "수량·불량 건수·불량률은 원본 캡처 표시값이며 AI 결과 판단은 포함하지 않습니다.";
        NavigateQuestionInsight(
            EvidenceHtmlRenderer.RenderAnswer(session));
        SetInsightWideMode(true);
        DeveloperConsoleExpander.IsExpanded = false;
        Log(
            $"[기존 이력] {Path.GetFileName(session.AnswerPath)} · "
            + $"관련 Study {session.RelevantStudyCount:N0}건 · "
            + $"원본 근거 {session.Citations.Count:N0}건");
        LogRelevanceAiSelection(session);
    }

    private void LogRelevanceAiSelection(EvidenceAnswerSession session)
    {
        using var document = JsonDocument.Parse(session.RawJson);
        var root = document.RootElement;
        var interpretation = root.GetProperty("queryInterpretation");
        var coverage = root.GetProperty("coverage");
        Log(
            "[AI 질문 해석] 필요한 문서: "
            + RelevanceJsonString(interpretation, "documentNeed"));
        Log(
            "[AI 질문 해석] 대상: "
            + RelevanceArrayText(interpretation, "subjects"));
        Log(
            "[AI 질문 해석] 조건: "
            + RelevanceArrayText(interpretation, "conditions"));
        Log(
            "[AI 질문 해석] 지표: "
            + RelevanceArrayText(interpretation, "metrics"));
        Log(
            $"[AI 관련성 결과] DB 후보 "
            + $"{coverage.GetProperty("candidateStudyCount").GetInt32():N0}건 → "
            + $"관련 Study {coverage.GetProperty("relevantStudyCount").GetInt32():N0}건 → "
            + $"원본 근거 {coverage.GetProperty("citationCount").GetInt32():N0}건");

        var selectedStudies = root.GetProperty("studies");
        var selectionNumber = 0;
        foreach (var study in selectedStudies.EnumerateArray())
        {
            selectionNumber++;
            Log(
                $"[AI 선택 {selectionNumber:00}] "
                + RelevanceJsonString(study, "fileName"));
            Log(
                "    Study: "
                + RelevanceJsonString(study, "studyGroup"));
            Log(
                "    질문 연결: "
                + RelevanceArrayText(study, "matchedAspects"));
            Log(
                "    선택 이유: "
                + RelevanceJsonString(study, "relevanceReason"));
        }
        Log($"[AI 관련성 JSON] {session.AnswerPath}");
    }

    private static string RelevanceJsonString(
        JsonElement element,
        string property) =>
        element.TryGetProperty(property, out var value)
        && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;

    private static string RelevanceArrayText(
        JsonElement element,
        string property)
    {
        if (!element.TryGetProperty(property, out var values)
            || values.ValueKind != JsonValueKind.Array)
            return "없음";
        var text = string.Join(
            ", ",
            values.EnumerateArray()
                .Where(value => value.ValueKind == JsonValueKind.String)
                .Select(value => value.GetString())
                .Where(value => !string.IsNullOrWhiteSpace(value)));
        return string.IsNullOrWhiteSpace(text) ? "없음" : text;
    }

    private async void OpenEvidencePreview_Click(
        object sender,
        RoutedEventArgs e) =>
        await ShowSelectedEvidenceDetailAsync();

    private async void EvidenceCitationsGrid_MouseDoubleClick(
        object sender,
        MouseButtonEventArgs e) =>
        await ShowSelectedEvidenceDetailAsync();

    private async Task<EvidenceDetailDocument?>
        ShowSelectedEvidenceDetailAsync()
    {
        if (_canonicalEvidenceBusy) return null;
        if (EvidenceCitationsGrid.SelectedItem
            is not EvidenceCitationRow citation)
        {
            WpfMessageBox.Show(
                "먼저 EVD 근거 행을 선택하세요.",
                "근거 표",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return null;
        }

        SetCanonicalEvidenceBusy(true);
        EvidenceQueryState.Text =
            $"{citation.EvidenceId}의 고정 revision 표를 읽는 중...";
        try
        {
            var detail = await _canonicalEvidenceClient.DetailAsync(
                _databasePath,
                citation.EvidenceId);
            _activeEvidenceDetail = detail;
            EvidenceQueryState.Text =
                $"{detail.EvidenceId} · {detail.TrustStatus}";
            ResultTitle.Text = $"근거 표 — {detail.EvidenceId}";
            PipelineFileTitle.Text = Path.GetFileName(detail.SourcePath);
            PipelineBatchText.Text =
                $"{detail.Sheet}!{detail.Range} · {detail.TrustStatus}";
            PipelineStagesText.Text =
                "현재 canonical revision과 정확히 연결된 Capture v2 셀만 표시합니다.\n"
                + "병합·수식·캐시값·표시값·서식·숨김 행/열을 보존하며 이미지는 제외합니다.";
            NavigateQuestionInsight(
                EvidenceHtmlRenderer.RenderDetail(detail));
            return detail;
        }
        catch (Exception exception)
        {
            EvidenceQueryState.Text = "근거 표 조회 실패";
            Log($"근거 표 조회 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "근거 표 조회 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return null;
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private async void OpenEvidenceExcel_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy) return;
        if (EvidenceCitationsGrid.SelectedItem
            is not EvidenceCitationRow citation)
        {
            WpfMessageBox.Show(
                "먼저 EVD 근거 행을 선택하세요.",
                "원본 Excel 범위",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        var detail = _activeEvidenceDetail is not null
            && string.Equals(
                _activeEvidenceDetail.EvidenceId,
                citation.EvidenceId,
                StringComparison.OrdinalIgnoreCase)
                ? _activeEvidenceDetail
                : await ShowSelectedEvidenceDetailAsync();
        if (detail is null) return;
        try
        {
            EvidenceQueryState.Text =
                $"{detail.EvidenceId} 원본 Excel 범위를 여는 중...";
            await ExcelRangeNavigator.OpenReadOnlyAsync(
                detail.SourcePath,
                detail.Sheet,
                detail.Range);
            EvidenceQueryState.Text =
                $"{detail.EvidenceId} · Excel {detail.Sheet}!{detail.Range}";
            Log(
                $"원본 Excel 근거 열기: {detail.SourcePath} / "
                + $"{detail.Sheet}!{detail.Range}");
        }
        catch (Exception exception)
        {
            EvidenceQueryState.Text = "Excel 범위 열기 실패";
            Log($"Excel 범위 열기 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "원본 Excel 범위 열기 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private async void ResultBrowser_Navigating(
        object sender,
        NavigatingCancelEventArgs e)
    {
        if (e.Uri is null
            || !string.Equals(
                e.Uri.Scheme,
                "inference-excel",
                StringComparison.OrdinalIgnoreCase))
            return;
        e.Cancel = true;
        var parameters = ParseNavigationQuery(e.Uri.Query);
        parameters.TryGetValue("source", out var sourcePath);
        parameters.TryGetValue("sheet", out var sheet);
        parameters.TryGetValue("range", out var range);
        if (string.IsNullOrWhiteSpace(sourcePath)
            || string.IsNullOrWhiteSpace(sheet)
            || string.IsNullOrWhiteSpace(range))
        {
            WpfMessageBox.Show(
                "Excel 원본 경로, 시트 또는 범위 정보가 없습니다.",
                "원본 Excel 열기",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }
        try
        {
            EvidenceQueryState.Text = $"Excel {sheet}!{range} 여는 중...";
            await ExcelRangeNavigator.OpenReadOnlyAsync(
                sourcePath,
                sheet,
                range);
            EvidenceQueryState.Text = $"Excel {sheet}!{range}";
            Log($"비교표에서 원본 Excel 열기: {sourcePath} / {sheet}!{range}");
        }
        catch (Exception exception)
        {
            EvidenceQueryState.Text = "Excel 범위 열기 실패";
            Log($"비교표 Excel 열기 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "원본 Excel 열기 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private static Dictionary<string, string> ParseNavigationQuery(
        string query)
    {
        var result = new Dictionary<string, string>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var part in query.TrimStart('?').Split(
                     '&',
                     StringSplitOptions.RemoveEmptyEntries))
        {
            var separator = part.IndexOf('=');
            var key = separator < 0 ? part : part[..separator];
            var value = separator < 0 ? string.Empty : part[(separator + 1)..];
            result[System.Net.WebUtility.UrlDecode(key)] =
                System.Net.WebUtility.UrlDecode(value);
        }
        return result;
    }

    private async void RefreshReviewQueue_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy) return;
        SetCanonicalEvidenceBusy(true);
        try
        {
            await RefreshReviewQueueCoreAsync();
        }
        catch (Exception exception)
        {
            ReviewQueueState.Text = "검토 큐 조회 실패";
            Log($"사람 검토 큐 조회 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "사람 검토 큐 조회 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private async Task RefreshReviewQueueCoreAsync()
    {
        ReviewQueueState.Text = "현재 revision의 검토 대기 비교 조회 중...";
        var queue = await _canonicalEvidenceClient.ReviewQueueAsync(
            _databasePath);
        _reviewQueue.Clear();
        foreach (var item in queue.Items) _reviewQueue.Add(item);
        _reviewEvidence.Clear();
        _activeReviewDetail = null;
        ReviewQueueState.Text = $"검토 대기 {_reviewQueue.Count:N0}건";
        Log($"사람 검토 큐 조회 완료: {_reviewQueue.Count}건");
    }

    private async void ReviewQueueGrid_SelectionChanged(
        object sender,
        System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (_canonicalEvidenceBusy
            || ReviewQueueGrid.SelectedItem is not ReviewQueueRow item)
            return;
        SetCanonicalEvidenceBusy(true);
        ReviewQueueState.Text = $"{item.PublicComparisonId} 근거 조회 중...";
        try
        {
            var detail = await _canonicalEvidenceClient.ReviewDetailAsync(
                _databasePath,
                item.PublicComparisonId);
            _activeReviewDetail = detail;
            _reviewEvidence.Clear();
            foreach (var evidence in detail.Evidence)
                _reviewEvidence.Add(evidence);
            SetComboValue(
                ReviewStudyComparabilityCombo,
                detail.StudyComparabilityStatus);
            SetComboValue(
                ReviewStudyConfoundingCombo,
                detail.StudyConfoundingStatus);
            SetComboValue(
                ReviewComparisonValidityCombo,
                detail.ComparisonValidityStatus);
            SetComboValue(
                ReviewComparisonConfoundingCombo,
                detail.ComparisonConfoundingStatus);
            ReviewMatchingBasisText.Text = detail.MatchingBasis;
            ReviewReasonText.Text = string.Empty;
            ReviewQueueState.Text =
                $"{detail.PublicComparisonId} · 근거 {_reviewEvidence.Count:N0}개"
                + $" · 현재 승인 준비 {detail.ApprovalReady}";
            ResultTitle.Text = $"사람 검토 {detail.PublicComparisonId}";
            PipelineFileTitle.Text = Path.GetFileName(detail.SourcePath);
            PipelineBatchText.Text =
                $"{detail.PublicDataId} · {detail.ComparedArmLabel} vs "
                + detail.ControlArmLabel;
            PipelineStagesText.Text =
                "원본 revision과 SHA-256이 현재 상태인지 확인했습니다.\n"
                + "승인 전 비교가능성·교란·matching basis를 사람이 판정해야 합니다.";
            NavigateUtf8Html(
                EvidenceHtmlRenderer.RenderReviewDetail(detail));
        }
        catch (Exception exception)
        {
            ReviewQueueState.Text = "비교 상세 조회 실패";
            Log($"사람 검토 상세 조회 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "사람 검토 상세 조회 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private static void SetComboValue(
        System.Windows.Controls.ComboBox comboBox,
        string value) =>
        comboBox.SelectedValue = value;

    private static string SelectedComboValue(
        System.Windows.Controls.ComboBox comboBox) =>
        Convert.ToString(comboBox.SelectedValue)?.Trim()
        ?? string.Empty;

    private async void OpenReviewEvidencePreview_Click(
        object sender,
        RoutedEventArgs e) =>
        await ShowSelectedReviewEvidenceAsync(openExcel: false);

    private async void ReviewEvidenceGrid_MouseDoubleClick(
        object sender,
        MouseButtonEventArgs e) =>
        await ShowSelectedReviewEvidenceAsync(openExcel: false);

    private async void OpenReviewEvidenceExcel_Click(
        object sender,
        RoutedEventArgs e) =>
        await ShowSelectedReviewEvidenceAsync(openExcel: true);

    private async Task ShowSelectedReviewEvidenceAsync(bool openExcel)
    {
        if (_canonicalEvidenceBusy) return;
        if (ReviewEvidenceGrid.SelectedItem
            is not ReviewEvidenceRow evidence)
        {
            WpfMessageBox.Show(
                "먼저 검토 근거 행을 선택하세요.",
                "사람 검토 근거",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        SetCanonicalEvidenceBusy(true);
        try
        {
            var detail = await _canonicalEvidenceClient.DetailAsync(
                _databasePath,
                evidence.EvidenceId);
            if (openExcel)
            {
                await ExcelRangeNavigator.OpenReadOnlyAsync(
                    detail.SourcePath,
                    detail.Sheet,
                    detail.Range);
                ReviewQueueState.Text =
                    $"{detail.EvidenceId} · Excel "
                    + $"{detail.Sheet}!{detail.Range}";
            }
            else
            {
                ResultTitle.Text = $"검토 근거 {detail.EvidenceId}";
                PipelineFileTitle.Text = Path.GetFileName(
                    detail.SourcePath);
                PipelineBatchText.Text =
                    $"{detail.Sheet}!{detail.Range} · {detail.TrustStatus}";
                PipelineStagesText.Text =
                    "현재 revision에 고정된 Capture v2 표입니다. "
                    + "이미지 분석은 포함되지 않습니다.";
                NavigateUtf8Html(
                    EvidenceHtmlRenderer.RenderDetail(detail));
            }
        }
        catch (Exception exception)
        {
            Log($"사람 검토 근거 열기 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "사람 검토 근거 열기 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private async void ApproveReview_Click(
        object sender,
        RoutedEventArgs e) =>
        await DecideSelectedReviewAsync("APPROVE");

    private async void RejectReview_Click(
        object sender,
        RoutedEventArgs e) =>
        await DecideSelectedReviewAsync("REJECT");

    private async void ExcludeReview_Click(
        object sender,
        RoutedEventArgs e) =>
        await DecideSelectedReviewAsync("EXCLUDE");

    private async void ReturnReview_Click(
        object sender,
        RoutedEventArgs e) =>
        await DecideSelectedReviewAsync("RETURN_TO_REVIEW");

    private async Task DecideSelectedReviewAsync(string decision)
    {
        if (_canonicalEvidenceBusy) return;
        if (_activeReviewDetail is null
            || ReviewQueueGrid.SelectedItem is not ReviewQueueRow selected
            || !string.Equals(
                selected.PublicComparisonId,
                _activeReviewDetail.PublicComparisonId,
                StringComparison.OrdinalIgnoreCase))
        {
            WpfMessageBox.Show(
                "먼저 검토할 CMP 행과 상세 근거를 선택하세요.",
                "사람 검토 결정",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        var reviewer = ReviewReviewerText.Text.Trim();
        var reason = ReviewReasonText.Text.Trim();
        if (reviewer.Length == 0 || reason.Length == 0)
        {
            WpfMessageBox.Show(
                "검토자와 판정 이유를 모두 입력하세요.",
                "사람 검토 결정",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        ReviewAssessment? assessment = null;
        if (decision == "APPROVE")
        {
            assessment = new ReviewAssessment(
                SelectedComboValue(ReviewStudyComparabilityCombo),
                SelectedComboValue(ReviewStudyConfoundingCombo),
                SelectedComboValue(ReviewComparisonValidityCombo),
                SelectedComboValue(ReviewComparisonConfoundingCombo),
                ReviewMatchingBasisText.Text.Trim());
            if (assessment.StudyComparabilityStatus != "VALID"
                || assessment.StudyConfoundingStatus != "NONE"
                || assessment.ComparisonValidityStatus != "VALID"
                || assessment.ComparisonConfoundingStatus != "NONE"
                || assessment.MatchingBasis.Length == 0)
            {
                WpfMessageBox.Show(
                    "승인하려면 Study 비교가능성=VALID, Study 교란=NONE, "
                    + "비교 유효성=VALID, 비교 교란=NONE, 그리고 구체적인 "
                    + "matching basis가 필요합니다.",
                    "승인 조건 미충족",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }
        }

        var prompt = decision == "APPROVE"
            ? "현재 원본의 직접 셀 근거, 시험군·대조군 pairing, 표본 차이와 "
              + "다른 변경 요인을 직접 확인했습니까?\n\n승인하면 검증된 "
              + "관측값에서 효과를 결정론적으로 계산해 집계 가능 상태로 만듭니다."
            : $"{decision} 결정을 저장하시겠습니까?\n\n집계 가능한 효과는 비활성화됩니다.";
        if (WpfMessageBox.Show(
                $"{_activeReviewDetail.PublicComparisonId}\n\n{prompt}",
                "사람 검토 결정 확인",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning)
            != MessageBoxResult.Yes)
            return;

        SetCanonicalEvidenceBusy(true);
        ReviewQueueState.Text = $"{decision} 저장 중...";
        try
        {
            var result = await _canonicalEvidenceClient.DecideReviewAsync(
                _databasePath,
                _activeReviewDetail.PublicComparisonId,
                decision,
                reviewer,
                reason,
                assessment);
            ResultTitle.Text =
                $"사람 검토 결과 {result.PublicComparisonId}";
            PipelineBatchText.Text =
                $"{result.Decision} · 집계 가능 "
                + result.AggregationEligible;
            PipelineStagesText.Text =
                $"생성·갱신 효과 {result.EffectPublicIds.Count:N0}건\n"
                + "모든 효과는 현재 revision의 직접 VERIFIED 셀 근거에서 계산했습니다.";
            NavigateUtf8Html(
                EvidenceHtmlRenderer.RenderReviewDecision(result));
            Log(
                $"사람 검토 결정 완료: {result.PublicComparisonId}, "
                + $"{result.Decision}, 효과 {result.EffectPublicIds.Count}");
            await RefreshReviewQueueCoreAsync();
        }
        catch (Exception exception)
        {
            ReviewQueueState.Text = "사람 검토 결정 실패 · DB 변경 없음";
            Log($"사람 검토 결정 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "사람 검토 결정 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private async void RefreshConceptCandidates_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy) return;
        SetCanonicalEvidenceBusy(true);
        try
        {
            await RefreshConceptCandidatesCoreAsync();
        }
        catch (Exception exception)
        {
            ConceptCurationState.Text =
                "OPEN 후보 조회 실패 · 기존 화면 목록 유지";
            Log($"개념 정규화 후보 조회 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "개념 정규화 후보 조회 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private async Task RefreshConceptCandidatesCoreAsync()
    {
        ConceptCurationState.Text = "OPEN 후보 조회 중...";
        var result = await _canonicalEvidenceClient
            .ConceptCandidatesAsync(
                _databasePath,
                ConceptCandidateKindFilterText.Text.Trim(),
                ConceptCandidateQueryText.Text.Trim());

        _conceptCandidates.Clear();
        foreach (var candidate in result.Candidates)
            _conceptCandidates.Add(candidate);
        _canonicalConceptListCurrent = false;
        ConceptCanonicalNameText.Text = string.Empty;
        ConceptAliasText.Text = string.Empty;
        ConceptCurationState.Text =
            $"OPEN 후보 {_conceptCandidates.Count:N0}건"
            + $" · {result.SchemaVersion}";
        Log(
            $"개념 정규화 OPEN 후보 조회 완료: "
            + $"{_conceptCandidates.Count}건");
    }

    private async void ConceptCandidatesGrid_SelectionChanged(
        object sender,
        System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (_canonicalEvidenceBusy) return;
        if (ConceptCandidatesGrid.SelectedItem
            is not ConceptCandidateRow candidate)
        {
            _canonicalConceptListCurrent = false;
            UpdateConceptDecisionAvailability();
            return;
        }

        ConceptCanonicalNameText.Text =
            candidate.SuggestedCanonicalName.Length > 0
                ? candidate.SuggestedCanonicalName
                : candidate.OriginalValue;
        ConceptAliasText.Text = candidate.OriginalValue;
        _canonicalConceptListCurrent = false;
        UpdateConceptDecisionAvailability();
        if (!candidate.IsConceptCandidate)
        {
            ConceptCurationState.Text =
                $"{candidate.CandidateUid} · {candidate.CandidateKind}는 "
                + "CONCEPT:* 후보가 아니므로 이 화면에서 결정할 수 없습니다.";
            return;
        }

        SetCanonicalEvidenceBusy(true);
        try
        {
            await RefreshCanonicalConceptsCoreAsync(
                candidate.ConceptKind);
        }
        catch (Exception exception)
        {
            ConceptCurationState.Text =
                $"{candidate.ConceptKind} ACTIVE 개념 조회 실패"
                + " · 기존 목록은 표시만 유지";
            Log(
                $"동일 kind ACTIVE 개념 조회 실패: "
                + exception.Message);
            WpfMessageBox.Show(
                exception.Message,
                "동일 kind ACTIVE 개념 조회 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private async void RefreshCanonicalConcepts_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy) return;
        if (ConceptCandidatesGrid.SelectedItem
            is not ConceptCandidateRow candidate
            || !candidate.IsConceptCandidate)
        {
            WpfMessageBox.Show(
                "먼저 CONCEPT:* OPEN 후보를 선택하세요.",
                "ACTIVE 개념 조회",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        _canonicalConceptListCurrent = false;
        SetCanonicalEvidenceBusy(true);
        try
        {
            await RefreshCanonicalConceptsCoreAsync(
                candidate.ConceptKind);
        }
        catch (Exception exception)
        {
            ConceptCurationState.Text =
                $"{candidate.ConceptKind} ACTIVE 개념 조회 실패"
                + " · 기존 목록은 표시만 유지";
            Log(
                $"동일 kind ACTIVE 개념 조회 실패: "
                + exception.Message);
            WpfMessageBox.Show(
                exception.Message,
                "동일 kind ACTIVE 개념 조회 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private async Task RefreshCanonicalConceptsCoreAsync(
        string conceptKind)
    {
        ConceptCurationState.Text =
            $"{conceptKind} ACTIVE 개념 조회 중...";
        var result = await _canonicalEvidenceClient.ConceptsAsync(
            _databasePath,
            conceptKind,
            CanonicalConceptQueryText.Text.Trim());

        _canonicalConcepts.Clear();
        foreach (var concept in result.Concepts)
            _canonicalConcepts.Add(concept);
        _loadedCanonicalConceptKind = conceptKind;
        _canonicalConceptListCurrent = true;
        ConceptCurationState.Text =
            $"{conceptKind} ACTIVE 개념 {_canonicalConcepts.Count:N0}건"
            + $" · {result.SchemaVersion}";
        Log(
            $"동일 kind ACTIVE 개념 조회 완료: "
            + $"{conceptKind}, {_canonicalConcepts.Count}건");
    }

    private void CanonicalConceptsGrid_SelectionChanged(
        object sender,
        System.Windows.Controls.SelectionChangedEventArgs e) =>
        UpdateConceptDecisionAvailability();

    private async void CreateConcept_Click(
        object sender,
        RoutedEventArgs e) =>
        await ResolveSelectedConceptAsync("CREATE");

    private async void MergeConcept_Click(
        object sender,
        RoutedEventArgs e) =>
        await ResolveSelectedConceptAsync("MERGE");

    private async void RejectConcept_Click(
        object sender,
        RoutedEventArgs e) =>
        await ResolveSelectedConceptAsync("REJECT");

    private async Task ResolveSelectedConceptAsync(string action)
    {
        if (_canonicalEvidenceBusy) return;
        if (ConceptCandidatesGrid.SelectedItem
                is not ConceptCandidateRow candidate
            || !candidate.IsConceptCandidate
            || !string.Equals(
                candidate.Status,
                "OPEN",
                StringComparison.Ordinal))
        {
            WpfMessageBox.Show(
                "결정 가능한 CONCEPT:* OPEN 후보를 선택하세요. "
                + "UNIT 등 다른 후보는 별도 정규화 경로가 필요합니다.",
                "개념 정규화 결정",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var reviewer = ConceptReviewerText.Text.Trim();
        var note = ConceptNoteText.Text.Trim();
        if (reviewer.Length == 0 || note.Length == 0)
        {
            WpfMessageBox.Show(
                "검토자와 판단 근거를 모두 입력하세요.",
                "개념 정규화 결정",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var canonicalName = ConceptCanonicalNameText.Text.Trim();
        var alias = ConceptAliasText.Text.Trim();
        CanonicalConceptRow? targetConcept = null;
        if (action == "CREATE"
            && (canonicalName.Length == 0 || alias.Length == 0))
        {
            WpfMessageBox.Show(
                "CREATE에는 canonical name과 승인 별칭이 필요합니다.",
                "개념 정규화 결정",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        if (action == "MERGE")
        {
            targetConcept =
                CanonicalConceptsGrid.SelectedItem as CanonicalConceptRow;
            if (!_canonicalConceptListCurrent
                || targetConcept is null
                || !string.Equals(
                    _loadedCanonicalConceptKind,
                    candidate.ConceptKind,
                    StringComparison.OrdinalIgnoreCase)
                || !string.Equals(
                    targetConcept.ConceptKind,
                    candidate.ConceptKind,
                    StringComparison.OrdinalIgnoreCase)
                || !string.Equals(
                    targetConcept.LifecycleStatus,
                    "ACTIVE",
                    StringComparison.Ordinal)
                || alias.Length == 0)
            {
                WpfMessageBox.Show(
                    "MERGE에는 방금 조회한 동일 kind ACTIVE 개념과 "
                    + "승인 별칭이 필요합니다.",
                    "개념 정규화 결정",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }
        }

        var confirmation = new List<string>
        {
            "이 결정은 감사 이력에 영구 저장되며 되돌릴 수 없습니다.",
            string.Empty,
            $"Action: {action}",
            $"Candidate UID: {candidate.CandidateUid}",
            $"Candidate kind: {candidate.CandidateKind}",
            $"Candidate value: {candidate.OriginalValue}",
        };
        if (action == "CREATE")
        {
            confirmation.Add($"Canonical name: {canonicalName}");
            confirmation.Add($"Approved alias: {alias}");
        }
        else if (action == "MERGE" && targetConcept is not null)
        {
            confirmation.Add(
                $"Target concept UID: {targetConcept.ConceptUid}");
            confirmation.Add(
                $"Target concept kind: {targetConcept.ConceptKind}");
            confirmation.Add(
                $"Target canonical name: {targetConcept.CanonicalName}");
            confirmation.Add($"Approved alias: {alias}");
        }
        confirmation.Add(string.Empty);
        confirmation.Add("위 내용으로 저장하시겠습니까?");
        if (WpfMessageBox.Show(
                string.Join("\n", confirmation),
                "되돌릴 수 없는 개념 정규화 결정",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning)
            != MessageBoxResult.Yes)
            return;

        SetCanonicalEvidenceBusy(true);
        ConceptCurationState.Text =
            $"{action} 저장 중 · {candidate.CandidateUid}";
        ConceptResolutionDocument result;
        try
        {
            result = await _canonicalEvidenceClient.ResolveConceptAsync(
                _databasePath,
                candidate.CandidateUid,
                action,
                reviewer,
                note,
                canonicalName: action == "CREATE"
                    ? canonicalName
                    : string.Empty,
                conceptUid: action == "MERGE"
                    ? targetConcept!.ConceptUid
                    : string.Empty,
                alias: action is "CREATE" or "MERGE"
                    ? alias
                    : string.Empty);
        }
        catch (Exception exception)
        {
            ConceptCurationState.Text =
                "저장 실패 또는 결과 확인 실패"
                + " · 목록을 새로고침해 DB 상태 확인 필요";
            Log($"개념 정규화 {action} 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "개념 정규화 결정 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            SetCanonicalEvidenceBusy(false);
            return;
        }

        var replayText = result.IdempotentReplay
            ? "기존 결정을 재확인함"
            : "새 결정을 저장함";
        ConceptCurationState.Text =
            $"DB 저장 완료 · {result.ResolutionUid}"
            + $" · IdempotentReplay={result.IdempotentReplay}"
            + $" ({replayText})";
        Log(
            $"개념 정규화 저장 완료: {result.ResolutionUid}, "
            + $"{result.Action}, "
            + $"IdempotentReplay={result.IdempotentReplay}");

        try
        {
            await RefreshConceptCandidatesCoreAsync();
            await RefreshCanonicalConceptsCoreAsync(
                candidate.ConceptKind);
            ConceptCurationState.Text =
                $"DB 저장 및 화면 새로고침 완료 · "
                + $"{result.ResolutionUid}"
                + $" · IdempotentReplay={result.IdempotentReplay}"
                + $" ({replayText})";
        }
        catch (Exception exception)
        {
            ConceptCurationState.Text =
                $"DB 저장 완료 · 화면 새로고침 실패 · "
                + $"{result.ResolutionUid}"
                + $" · IdempotentReplay={result.IdempotentReplay}";
            Log(
                $"개념 정규화 저장 후 화면 새로고침 실패: "
                + exception.Message);
            WpfMessageBox.Show(
                "DB 저장은 완료되었지만 화면 새로고침에 실패했습니다.\n\n"
                + exception.Message,
                "개념 정규화 저장 완료 · 새로고침 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private void UpdateConceptDecisionAvailability()
    {
        var candidate =
            ConceptCandidatesGrid.SelectedItem as ConceptCandidateRow;
        var eligible = !_canonicalEvidenceBusy
            && candidate is not null
            && candidate.IsConceptCandidate
            && string.Equals(
                candidate.Status,
                "OPEN",
                StringComparison.Ordinal);
        RefreshConceptCandidatesButton.IsEnabled =
            !_canonicalEvidenceBusy;
        RefreshCanonicalConceptsButton.IsEnabled = eligible;
        ConceptCandidatesGrid.IsEnabled = !_canonicalEvidenceBusy;
        CanonicalConceptsGrid.IsEnabled = eligible;
        CreateConceptButton.IsEnabled = eligible;
        RejectConceptButton.IsEnabled = eligible;

        var concept =
            CanonicalConceptsGrid.SelectedItem as CanonicalConceptRow;
        MergeConceptButton.IsEnabled = eligible
            && _canonicalConceptListCurrent
            && concept is not null
            && string.Equals(
                _loadedCanonicalConceptKind,
                candidate!.ConceptKind,
                StringComparison.OrdinalIgnoreCase)
            && string.Equals(
                concept.ConceptKind,
                candidate.ConceptKind,
                StringComparison.OrdinalIgnoreCase)
            && string.Equals(
                concept.LifecycleStatus,
                "ACTIVE",
                StringComparison.Ordinal);
    }

    private void BrowseIngestWorkbook_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (IsIngestFolderMode())
        {
            using var folderDialog = new Forms.FolderBrowserDialog
            {
                Description =
                    "DRM Excel 파일이 들어 있는 폴더를 선택하세요.",
                UseDescriptionForTitle = true,
                ShowNewFolderButton = false,
            };
            if (folderDialog.ShowDialog() == Forms.DialogResult.OK)
            {
                IngestWorkbookPathText.Text = Path.GetFullPath(
                    folderDialog.SelectedPath);
                RefreshIngestRetryAvailability();
            }
            return;
        }
        using var dialog = new Forms.OpenFileDialog
        {
            Title = "Excel COM으로 처리할 DRM 원본 선택",
            Filter = (
                "Excel Workbook (*.xlsx;*.xlsm;*.xlsb;*.xls)|"
                + "*.xlsx;*.xlsm;*.xlsb;*.xls|All files (*.*)|*.*"
            ),
            CheckFileExists = true,
            Multiselect = false,
        };
        if (dialog.ShowDialog() == Forms.DialogResult.OK)
        {
            IngestWorkbookPathText.Text = Path.GetFullPath(dialog.FileName);
            RefreshIngestRetryAvailability();
        }
    }

    private async void ExcelFolderSearch_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy || _excelFolderSearchBusy)
            return;

        using var folderDialog = new Forms.FolderBrowserDialog
        {
            Description =
                "파일을 열거나 변경하지 않고 Excel 파일명만 검색할 폴더를 선택하세요.",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = false,
        };
        if (folderDialog.ShowDialog() != Forms.DialogResult.OK)
            return;

        await SearchExcelFoldersAsync(
            [folderDialog.SelectedPath]);
    }

    private async void ExcelMultiFolderSearch_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy || _excelFolderSearchBusy)
            return;

        var dialog = new ExcelFolderSelectionWindow
        {
            Owner = this,
        };
        if (dialog.ShowDialog() != true)
            return;

        await SearchExcelFoldersAsync(dialog.SelectedFolders);
    }

    private async Task SearchExcelFoldersAsync(
        IReadOnlyList<string> selectedFolders)
    {
        var folders = selectedFolders
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (folders.Length == 0)
            return;

        _excelFolderSearchBusy = true;
        ExcelFolderSearchButton.IsEnabled = false;
        ExcelMultiFolderSearchButton.IsEnabled = false;
        ProcessCollectedExcelButton.IsEnabled = false;
        CopyMissingExcelSearchButton.IsEnabled = false;
        ExcelFolderSearchGrid.IsEnabled = false;
        ExcelFolderSearchPanel.Visibility = Visibility.Visible;
        _excelFolderSearchRows.Clear();
        _excelFolderSearchCopyCompleted = false;
        CopyMissingExcelSearchButton.Content =
            "DB 없음 모두 이 폴더로 복사";
        _excelFolderSearchRoot = folders.Length == 1
            ? folders[0]
            : $"{folders.Length:N0}개 검색 폴더";
        ExcelFolderSearchPathText.Text = folders.Length == 1
            ? folders[0]
            : $"{folders.Length:N0}개 폴더 · {folders[0]} 외 "
              + $"{folders.Length - 1:N0}개";
        ExcelFolderSearchPathText.ToolTip = string.Join(
            Environment.NewLine,
            folders);
        ExcelFolderSearchSummary.Text =
            $"{folders.Length:N0}개 폴더와 하위 폴더의 Excel 파일명만 읽는 중입니다. "
            + "파일 열기·복사·수정·COM 실행은 하지 않습니다.";
        try
        {
            var result = await ExcelFolderSearchService.SearchManyAsync(
                folders,
                _databasePath);
            foreach (var row in result.Rows)
                _excelFolderSearchRows.Add(row);
            _excelFolderSearchRoot = result.RootPath;
            UpdateExcelFolderSearchSummary("파일명만 읽음");
        }
        catch (
            Exception exception
        ) when (
            exception is IOException
            or UnauthorizedAccessException
            or ArgumentException
            or NotSupportedException
            or SqliteException)
        {
            ExcelFolderSearchSummary.Text =
                "검색 실패 · 파일이나 폴더는 변경하지 않았습니다.";
            WpfMessageBox.Show(
                exception.Message,
                "엑셀 검색",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
        finally
        {
            _excelFolderSearchBusy = false;
            ExcelFolderSearchButton.IsEnabled =
                !_canonicalEvidenceBusy;
            ExcelMultiFolderSearchButton.IsEnabled =
                !_canonicalEvidenceBusy;
            ExcelFolderSearchGrid.IsEnabled = true;
            UpdateExcelSearchActionAvailability();
        }
    }

    private async void CopyMissingExcelSearch_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy || _excelFolderSearchBusy)
            return;

        var rows = _excelFolderSearchRows
            .Where(row => !row.ExistsInDatabase)
            .ToArray();
        if (rows.Length == 0)
        {
            WpfMessageBox.Show(
                "DB에 없는 Excel 파일이 없습니다.",
                "신규 Excel 수집",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var localBase = ExcelLocalCopyService.GetLocalCopyBase(
            _pathSettings.ExcelArchiveDirectory);
        var confirmation = WpfMessageBox.Show(
            $"DB 없음 {rows.Length:N0}개를 Excel 파일 보관함에 복사합니다.\n\n"
            + $"보관함 위치: {localBase}\n"
            + "날짜·배치 하위 폴더 없이 이 위치에 직접 저장합니다.\n\n"
            + "복사와 파일 크기 검증까지만 실행합니다. "
            + "AI·Excel COM·DB 전체 처리는 시작하지 않습니다.\n"
            + "네트워크 원본은 수정·이동·삭제하지 않습니다.",
            "DB 없음 Excel 보관함 복사",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Information);
        if (confirmation != MessageBoxResult.OK)
            return;

        _excelFolderSearchBusy = true;
        ExcelFolderSearchButton.IsEnabled = false;
        ExcelMultiFolderSearchButton.IsEnabled = false;
        ProcessCollectedExcelButton.IsEnabled = false;
        CopyMissingExcelSearchButton.IsEnabled = false;
        CopyMissingExcelSearchButton.Content =
            "보관함으로 복사 중...";
        ExcelFolderSearchGrid.IsEnabled = false;
        var progress = new Progress<ExcelLocalCopyProgress>(
            value =>
            {
                ExcelFolderSearchSummary.Text =
                    $"보관함 복사 {value.Current:N0}/{value.Total:N0} · "
                    + value.FileName
                    + " · 네트워크 원본 변경 없음";
            });
        try
        {
            var result = await ExcelLocalCopyService.CopyAsync(
                rows,
                _pathSettings.ExcelArchiveDirectory,
                progress);
            _excelFolderSearchCopyCompleted = true;
            CopyMissingExcelSearchButton.Content =
                "보관함 복사 완료";
            UpdateExcelFolderSearchSummary(
                $"보관함 복사 {result.CopiedPaths.Count:N0}개 완료 · "
                + $"저장 위치 {result.LocalRoot} · 전체 처리 미실행");
            Log(
                $"신규 Excel 수집 완료: {result.CopiedPaths.Count:N0}개 · "
                + result.LocalRoot);
            WpfMessageBox.Show(
                $"DB에 없는 Excel {result.CopiedPaths.Count:N0}개를 "
                + "로컬 보관함으로 복사했습니다.\n\n"
                + result.LocalRoot
                + "\n\n원본과 DB는 변경하지 않았습니다. 이제 이 화면의 "
                + "‘보관함 Excel 전체 처리’를 누르면 별도 파일 선택 없이 "
                + "전체 처리를 시작합니다.",
                "신규 Excel 수집 완료",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (
            Exception exception
        ) when (
            exception is IOException
            or UnauthorizedAccessException
            or ArgumentException
            or NotSupportedException
            or InvalidOperationException)
        {
            ExcelFolderSearchSummary.Text =
                "보관함 복사 실패 · 네트워크 원본은 변경하지 않았습니다.";
            CopyMissingExcelSearchButton.Content =
                "DB 없음 모두 이 폴더로 복사";
            WpfMessageBox.Show(
                exception.Message,
                "신규 Excel 수집",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
        finally
        {
            _excelFolderSearchBusy = false;
            ExcelFolderSearchGrid.IsEnabled = true;
            ExcelFolderSearchButton.IsEnabled =
                !_canonicalEvidenceBusy;
            ExcelMultiFolderSearchButton.IsEnabled =
                !_canonicalEvidenceBusy;
            UpdateExcelSearchActionAvailability();
        }
    }

    private void GoToFormPreflight_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy || _excelFolderSearchBusy)
            return;

        var archivePath = ExcelLocalCopyService.GetLocalCopyBase(
            _pathSettings.ExcelArchiveDirectory);
        var workbooks = GetCollectedExcelFiles(archivePath);
        if (workbooks.Count == 0)
        {
            WpfMessageBox.Show(
                "로컬 보관함에 사전 분석할 Excel이 없습니다.\n\n"
                + archivePath,
                "Excel COM 사전 분석",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            UpdateCollectedExcelProcessingAvailability();
            return;
        }
        FormPreflightArchivePathText.Text = archivePath;
        RefreshFormPreflightUi();
        WorkspaceTabs.SelectedIndex = 10;
    }

    private void RefreshFormPreflight_Click(
        object sender,
        RoutedEventArgs e) =>
        RefreshFormPreflightUi();

    private void RefreshFormPreflightUi()
    {
        var archivePath = ExcelLocalCopyService.GetLocalCopyBase(
            _pathSettings.ExcelArchiveDirectory);
        FormPreflightArchivePathText.Text = archivePath;
        _activeFormPreflight =
            TryLoadMatchingFormPreflight(archivePath);
        _formPreflightRows.Clear();
        if (_activeFormPreflight is null)
        {
            var workbooks = GetCollectedExcelFiles(archivePath);
            PopulatePendingFormPreflightRows(workbooks);
            var count = workbooks.Count;
            var partial = _canonicalEvidenceClient
                .LoadLatestFormPreflight();
            if (partial is not null
                && string.Equals(
                    partial.Status,
                    "CANCELLED",
                    StringComparison.OrdinalIgnoreCase)
                && IsFormPreflightForArchive(
                    partial,
                    archivePath))
            {
                foreach (var row in partial.Items)
                    OverlayFormPreflightRow(row);
                FormPreflightSummaryText.Text =
                    $"사용자 중지 · 완료 {partial.TotalCount:N0}/{count:N0} · "
                    + $"기존 {partial.KnownFormCount:N0} · "
                    + $"유사 {partial.SimilarReviewCount:N0} · "
                    + $"신규 {partial.NewFormCount:N0} · "
                    + $"제외 {partial.ExcludedFormCount:N0} · "
                    + $"실패 {partial.CaptureFailedCount:N0}";
                FormPreflightProgressText.Text =
                    "완료된 캡처와 판정은 저장됐습니다. 다시 시작하면 재사용하고 남은 파일을 계속 분석합니다.";
            }
            else
            {
            FormPreflightSummaryText.Text = count == 0
                ? "사전 분석할 Excel이 없습니다."
                : $"현재 보관함 Excel {count:N0}개 · COM 사전 분석이 필요합니다.";
            FormPreflightProgressText.Text =
                "이전 결과가 없거나 현재 보관함 구성과 다릅니다. 사전 분석을 다시 실행하세요.";
            }
            RunKnownFormsFullProcessingButton.IsEnabled =
                !_canonicalEvidenceBusy && count > 0;
            RunFormPreflightButton.IsEnabled =
                !_canonicalEvidenceBusy && count > 0;
            return;
        }

        foreach (var row in _activeFormPreflight.Items)
            _formPreflightRows.Add(row);
        FormPreflightSummaryText.Text =
            $"판정 완료 {_activeFormPreflight.TotalCount:N0}개 · "
            + $"기존 양식 {_activeFormPreflight.KnownFormCount:N0} · "
            + $"유사 검토 {_activeFormPreflight.SimilarReviewCount:N0} · "
            + $"신규 보류 {_activeFormPreflight.NewFormCount:N0} · "
            + $"사람 제외 {_activeFormPreflight.ExcludedFormCount:N0} · "
            + $"추출 실패 {_activeFormPreflight.CaptureFailedCount:N0}";
        FormPreflightProgressText.Text =
            _activeFormPreflight.HasBlockingItems
                ? "기존 양식만 전체 처리할 수 있습니다. 유사·신규 양식과 COM 추출 실패 파일은 자동 보류됩니다."
                : "모든 파일이 기존 양식으로 확인됐습니다. 전체 처리를 실행할 수 있습니다.";
        RunKnownFormsFullProcessingButton.IsEnabled =
            !_canonicalEvidenceBusy
            && _activeFormPreflight.TotalCount > 0;
        RunFormPreflightButton.IsEnabled = !_canonicalEvidenceBusy;
    }

    private void PopulatePendingFormPreflightRows(
        IReadOnlyList<string> workbooks)
    {
        _formPreflightRows.Clear();
        foreach (var workbook in workbooks)
        {
            _formPreflightRows.Add(
                new FormPreflightRow(
                    "PENDING",
                    Path.GetFileName(workbook),
                    workbook,
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    0,
                    string.Empty,
                    "COM 사전 분석 대기"));
        }
    }

    private void MarkFormPreflightRowRunning(
        string sourcePath)
    {
        for (var index = 0;
             index < _formPreflightRows.Count;
             index++)
        {
            var row = _formPreflightRows[index];
            if (!string.Equals(
                    row.SourcePath,
                    sourcePath,
                    StringComparison.OrdinalIgnoreCase))
                continue;
            _formPreflightRows[index] = row with
            {
                Status = "RUNNING",
                Reason = "Excel COM 추출·구조 판정 중",
            };
            break;
        }
    }

    private void ApplyCompletedFormPreflightRow(
        string sourcePath)
    {
        var partial = _canonicalEvidenceClient
            .LoadLatestFormPreflight();
        var completedRow = partial?.Items.FirstOrDefault(
            row => string.Equals(
                row.SourcePath,
                sourcePath,
                StringComparison.OrdinalIgnoreCase));
        if (completedRow is null)
            return;
        OverlayFormPreflightRow(completedRow);
    }

    private void OverlayFormPreflightRow(
        FormPreflightRow completedRow)
    {
        for (var index = 0;
             index < _formPreflightRows.Count;
             index++)
        {
            if (!string.Equals(
                    _formPreflightRows[index].SourcePath,
                    completedRow.SourcePath,
                    StringComparison.OrdinalIgnoreCase))
                continue;
            _formPreflightRows[index] = completedRow;
            break;
        }
    }

    private static bool IsFormPreflightForArchive(
        FormPreflightDocument document,
        string archivePath)
    {
        try
        {
            return string.Equals(
                Path.GetFullPath(document.SourceRoot).TrimEnd(
                    Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar),
                Path.GetFullPath(archivePath).TrimEnd(
                    Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar),
                StringComparison.OrdinalIgnoreCase);
        }
        catch (
            Exception exception
        ) when (
            exception is ArgumentException
            or NotSupportedException
            or PathTooLongException)
        {
            return false;
        }
    }

    private FormPreflightDocument? TryLoadMatchingFormPreflight(
        string archivePath)
    {
        var latest = _canonicalEvidenceClient.LoadLatestFormPreflight();
        if (latest is null
            || !string.Equals(
                latest.Status,
                "COMPLETED",
                StringComparison.OrdinalIgnoreCase)
            || !File.Exists(latest.KnownFormManifestPath))
            return null;
        try
        {
            var normalizedArchive = Path.GetFullPath(archivePath)
                .TrimEnd(
                    Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar);
            var normalizedSource = Path.GetFullPath(latest.SourceRoot)
                .TrimEnd(
                    Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar);
            if (!string.Equals(
                    normalizedArchive,
                    normalizedSource,
                    StringComparison.OrdinalIgnoreCase))
                return null;
            return GetCollectedExcelFiles(normalizedArchive).Count
                   == latest.TotalCount
                ? latest
                : null;
        }
        catch (
            Exception exception
        ) when (
            exception is ArgumentException
            or NotSupportedException
            or PathTooLongException)
        {
            return null;
        }
    }

    private void LoadCachedFormGroupReview()
    {
        var review = _canonicalEvidenceClient
            .LoadLatestFormGroupReview();
        if (review is null)
        {
            _activeFormGroupReview = null;
            _formFamilyGroups.Clear();
            FormGroupSummaryText.Text =
                "사전 분석 결과에서 양식군을 불러오세요.";
            FormGroupProgressText.Text =
                "유사·신규 양식은 사람 판정 전까지 전체 처리 대상에서 보류됩니다.";
            UpdateFormFamilyDecisionAvailability();
            return;
        }
        ApplyFormGroupReview(review);
    }

    private void ApplyFormGroupReview(
        FormGroupReviewDocument review,
        string selectedFamilyId = "")
    {
        _activeFormGroupReview = review;
        _formFamilyGroups.Clear();
        foreach (var group in review.Groups)
            _formFamilyGroups.Add(group);
        FormGroupSummaryText.Text =
            $"양식군 {review.GroupCount:N0}개 · "
            + $"승인 대기 {review.PendingCount:N0} · "
            + $"등록/연결 {review.ApprovedCount:N0} · "
            + $"제외 {review.ExcludedCount:N0} · "
            + $"대상 Excel {review.WorkbookCount:N0}개";
        FormGroupProgressText.Text =
            string.Equals(
                review.PreflightStatus,
                "COMPLETED",
                StringComparison.OrdinalIgnoreCase)
                ? "완료된 사전 분석 기준입니다. 사람 판정은 통과 manifest에 즉시 반영됩니다."
                : $"사전 분석 상태 {review.PreflightStatus}의 부분 결과입니다. 남은 COM 추출을 재개하면 결정이 같은 구조에 자동 적용됩니다.";
        var selected = _formFamilyGroups.FirstOrDefault(
            group => string.Equals(
                group.FamilyId,
                selectedFamilyId,
                StringComparison.Ordinal));
        FormFamilyGroupsGrid.SelectedItem = selected;
        UpdateFormFamilyDecisionAvailability();
    }

    private bool TryGetFormGroupReport(
        out FormPreflightDocument report,
        bool showMessage = true)
    {
        report = _canonicalEvidenceClient.LoadLatestFormPreflight()!;
        var archivePath = ExcelLocalCopyService.GetLocalCopyBase(
            _pathSettings.ExcelArchiveDirectory);
        if (report is not null
            && IsFormPreflightForArchive(report, archivePath)
            && File.Exists(report.ReportPath))
            return true;
        if (showMessage)
        {
            WpfMessageBox.Show(
                "현재 보관함과 일치하는 COM 사전 분석 결과가 없습니다.\n\n"
                + "먼저 Excel COM 사전 분석을 실행하거나 이전 부분 결과를 불러오세요.",
                "양식군 검토 전 사전 분석 필요",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        report = null!;
        return false;
    }

    private async Task RefreshFormGroupsAsync(
        bool showMessage)
    {
        if (_canonicalEvidenceBusy
            || !TryGetFormGroupReport(
                out var report,
                showMessage))
            return;
        SetCanonicalEvidenceBusy(true);
        FormGroupProgressText.Text =
            "캡처된 Excel의 구조 지문을 양식군으로 묶는 중입니다.";
        try
        {
            var review = await _canonicalEvidenceClient
                .RefreshFormGroupsAsync(
                    _databasePath,
                    report.ReportPath);
            ApplyFormGroupReview(review);
            WorkspaceStatusText.Text =
                $"양식군 검토 · {review.GroupCount:N0}개 그룹";
        }
        catch (Exception exception)
        {
            FormGroupProgressText.Text = exception.Message;
            Log($"[양식군 검토] 목록 생성 실패: {exception.Message}");
            if (showMessage)
            {
                WpfMessageBox.Show(
                    exception.Message,
                    "양식군 목록 생성 실패",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
            UpdateFormFamilyDecisionAvailability();
        }
    }

    private async void RefreshFormGroups_Click(
        object sender,
        RoutedEventArgs e) =>
        await RefreshFormGroupsAsync(showMessage: true);

    private void FormFamilyGroupsGrid_SelectionChanged(
        object sender,
        SelectionChangedEventArgs e)
    {
        if (FormFamilyGroupsGrid.SelectedItem
            is not FormFamilyGroupRow selected)
        {
            UpdateFormFamilyDecisionAvailability();
            return;
        }
        FormFamilyDisplayNameText.Text = selected.DisplayName;
        FormFamilyLinkedSignatureText.Text =
            selected.NearestKnownFormSignatureId;
        FormFamilyNotesText.Text = selected.Notes;
        if (!string.IsNullOrWhiteSpace(selected.Reviewer))
            FormFamilyReviewerText.Text = selected.Reviewer;
        FormGroupProgressText.Text =
            $"선택: {selected.DisplayName} · "
            + $"{selected.MemberCount:N0}개 Excel · "
            + $"대표본 {selected.RepresentativeFile}";
        UpdateFormFamilyDecisionAvailability();
    }

    private void FormFamilyDecisionInput_Changed(
        object sender,
        TextChangedEventArgs e) =>
        UpdateFormFamilyDecisionAvailability();

    private void UpdateFormFamilyDecisionAvailability()
    {
        if (AnalyzeFormFamilyButton is null)
            return;
        var selected = FormFamilyGroupsGrid.SelectedItem
            as FormFamilyGroupRow;
        var available = !_canonicalEvidenceBusy
            && selected is not null
            && !string.IsNullOrWhiteSpace(
                FormFamilyReviewerText.Text);
        RefreshFormGroupsButton.IsEnabled =
            !_canonicalEvidenceBusy;
        AnalyzeFormFamilyButton.IsEnabled =
            available
            && selected!.DecisionStatus
                is "PENDING" or "ANALYZED_PENDING_APPROVAL";
        RegisterNewFormFamilyButton.IsEnabled =
            available && selected!.CanRegisterNew;
        LinkExistingFormFamilyButton.IsEnabled =
            available
            && !string.IsNullOrWhiteSpace(
                FormFamilyLinkedSignatureText.Text);
        ExcludeFormFamilyButton.IsEnabled = available;
    }

    private async void AnalyzeFormFamily_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy
            || FormFamilyGroupsGrid.SelectedItem
            is not FormFamilyGroupRow selected
            || !TryGetFormGroupReport(
                out var report))
            return;
        var confirmation = WpfMessageBox.Show(
            $"양식군 '{selected.DisplayName}'의 대표 Excel 1개와 "
            + $"검증 표본 {Math.Max(0, selected.SampleSources.Count - 1):N0}개를 AI로 분석합니다.\n\n"
            + "AI는 좌표가 포함된 읽기 전용 COM 캡처만 사용하며 원본 Excel을 열거나 변경하지 않습니다.",
            "대표본 AI 추출 계약 분석",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Information);
        if (confirmation != MessageBoxResult.OK)
            return;
        SetCanonicalEvidenceBusy(true);
        FormGroupProgressText.Text =
            $"대표본 AI 분석 중 · {selected.RepresentativeFile}";
        try
        {
            var review = await _canonicalEvidenceClient
                .AnalyzeFormFamilyAsync(
                    _databasePath,
                    report.ReportPath,
                    selected.FamilyId);
            ApplyFormGroupReview(review, selected.FamilyId);
            var updated = _formFamilyGroups.FirstOrDefault(
                group => group.FamilyId == selected.FamilyId);
            FormGroupProgressText.Text =
                updated?.CanRegisterNew == true
                    ? "대표본과 표본 검증이 통과했습니다. 사람 판정 후 신규 양식으로 등록할 수 있습니다."
                    : "AI 분석은 완료됐지만 표본 호환성 검증이 통과하지 않았습니다. 제외 또는 기존 양식 연결을 검토하세요.";
        }
        catch (Exception exception)
        {
            FormGroupProgressText.Text = exception.Message;
            Log($"[양식군 검토] AI 분석 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "대표본 AI 분석 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
            UpdateFormFamilyDecisionAvailability();
        }
    }

    private async Task DecideSelectedFormFamilyAsync(
        string decision,
        string decisionDisplay)
    {
        if (_canonicalEvidenceBusy
            || FormFamilyGroupsGrid.SelectedItem
            is not FormFamilyGroupRow selected
            || !TryGetFormGroupReport(
                out var report))
            return;
        var reviewer = FormFamilyReviewerText.Text.Trim();
        if (reviewer.Length == 0)
        {
            WpfMessageBox.Show(
                "판정자 이름을 입력하세요.",
                "양식군 사람 판정",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        var linkedSignature =
            FormFamilyLinkedSignatureText.Text.Trim();
        if (decision == "LINK_EXISTING"
            && linkedSignature.Length == 0)
        {
            WpfMessageBox.Show(
                "연결할 기존 양식 서명을 입력하세요.",
                "기존 양식 연결",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        if (decision == "REGISTER_NEW"
            && !selected.CanRegisterNew)
        {
            WpfMessageBox.Show(
                "대표본 AI 분석과 모든 선정 표본 검증이 통과해야 신규 양식으로 등록할 수 있습니다.",
                "신규 양식 등록 전 검증 필요",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        var confirmation = WpfMessageBox.Show(
            $"양식군 '{selected.DisplayName}'의 Excel "
            + $"{selected.MemberCount:N0}개를 '{decisionDisplay}'으로 판정합니다.\n\n"
            + "판정 결과는 현재 사전 분석 보고서와 전체 처리 manifest에 즉시 반영됩니다.",
            "양식군 최종 판정",
            MessageBoxButton.OKCancel,
            decision == "EXCLUDE"
                ? MessageBoxImage.Warning
                : MessageBoxImage.Information);
        if (confirmation != MessageBoxResult.OK)
            return;
        SetCanonicalEvidenceBusy(true);
        FormGroupProgressText.Text =
            $"{decisionDisplay} 반영 중 · {selected.DisplayName}";
        try
        {
            var review = await _canonicalEvidenceClient
                .DecideFormFamilyAsync(
                    _databasePath,
                    report.ReportPath,
                    selected.FamilyId,
                    decision,
                    reviewer,
                    FormFamilyDisplayNameText.Text,
                    linkedSignature,
                    FormFamilyNotesText.Text);
            ApplyFormGroupReview(review, selected.FamilyId);
            RefreshFormPreflightUi();
            UpdateCollectedExcelProcessingAvailability();
            FormGroupProgressText.Text =
                $"{decisionDisplay} 판정이 저장되고 사전 분석 통과 manifest가 갱신됐습니다.";
        }
        catch (Exception exception)
        {
            FormGroupProgressText.Text = exception.Message;
            Log($"[양식군 검토] 사람 판정 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "양식군 판정 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
            UpdateFormFamilyDecisionAvailability();
        }
    }

    private async void RegisterNewFormFamily_Click(
        object sender,
        RoutedEventArgs e) =>
        await DecideSelectedFormFamilyAsync(
            "REGISTER_NEW",
            "신규 양식 등록");

    private async void LinkExistingFormFamily_Click(
        object sender,
        RoutedEventArgs e) =>
        await DecideSelectedFormFamilyAsync(
            "LINK_EXISTING",
            "기존 양식 연결");

    private async void ExcludeFormFamily_Click(
        object sender,
        RoutedEventArgs e) =>
        await DecideSelectedFormFamilyAsync(
            "EXCLUDE",
            "전체 처리 제외");

    private bool TryGetAuthDialogSettings(
        out bool inspectAuth,
        out bool dismissAuth,
        out string authTitle,
        out string authClass,
        out string authButton)
    {
        inspectAuth = InspectAuthDialogCheck.IsChecked == true;
        dismissAuth = DismissAuthDialogCheck.IsChecked == true;
        authTitle = AuthDialogTitleText.Text.Trim();
        authClass = AuthDialogClassText.Text.Trim();
        authButton = AuthDialogButtonText.Text.Trim();
        if (inspectAuth && dismissAuth)
        {
            WpfMessageBox.Show(
                "인증창 정보 확인과 자동 버튼 클릭은 동시에 사용할 수 없습니다.",
                "인증창 설정",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return false;
        }
        if (dismissAuth && (
                authTitle.Length == 0
                || authClass.Length == 0
                || authButton.Length == 0))
        {
            WpfMessageBox.Show(
                "자동 클릭에는 인증창 제목, Window class, 버튼 문구가 모두 필요합니다.",
                "인증창 설정",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return false;
        }
        return true;
    }

    private async void RunFormPreflight_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy || _excelFolderSearchBusy)
            return;
        var archivePath = ExcelLocalCopyService.GetLocalCopyBase(
            _pathSettings.ExcelArchiveDirectory);
        var workbooks = GetCollectedExcelFiles(archivePath);
        if (workbooks.Count == 0)
        {
            WpfMessageBox.Show(
                "로컬 보관함에 사전 분석할 Excel이 없습니다.\n\n"
                + archivePath,
                "Excel COM 사전 분석",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            RefreshFormPreflightUi();
            return;
        }
        if (!TryGetAuthDialogSettings(
                out var inspectAuth,
                out var dismissAuth,
                out var authTitle,
                out var authClass,
                out var authButton))
            return;

        var confirmation = WpfMessageBox.Show(
            $"보관함 Excel {workbooks.Count:N0}개를 COM으로 읽어 양식을 판정합니다.\n\n"
            + "원본은 변경하지 않고 AI 분석도 실행하지 않습니다. "
            + "파일 수에 따라 시간이 오래 걸릴 수 있습니다.",
            "Excel COM 사전 분석",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Information);
        if (confirmation != MessageBoxResult.OK)
            return;

        PopulatePendingFormPreflightRows(workbooks);
        _activeFormPreflight = null;
        _formPreflightRunning = true;
        SetCanonicalEvidenceBusy(true);
        FormPreflightSummaryText.Text =
            $"COM 추출·양식 판정 진행 중 · 전체 {workbooks.Count:N0}개";
        FormPreflightProgressText.Text = "첫 Excel을 준비하는 중입니다.";
        var progress = new Progress<IngestProgressEvent>(
            item =>
            {
                if (string.Equals(
                        item.Status,
                        "RUNNING",
                        StringComparison.OrdinalIgnoreCase))
                    MarkFormPreflightRowRunning(item.SourcePath);
                else
                    ApplyCompletedFormPreflightRow(
                        item.SourcePath);
                FormPreflightProgressText.Text = item.Detail;
                WorkspaceStatusText.Text =
                    $"Excel COM 사전 분석 · {item.Detail}";
            });
        try
        {
            var result = await _canonicalEvidenceClient
                .PreflightFormsAsync(
                    _databasePath,
                    archivePath,
                    progress,
                    inspectAuth,
                    dismissAuth,
                    authTitle,
                    authClass,
                    authButton);
            if (string.Equals(
                    result.Status,
                    "CANCELLED",
                    StringComparison.OrdinalIgnoreCase))
            {
                _activeFormPreflight = null;
                FormPreflightSummaryText.Text =
                    $"사용자 중지 · 완료 {result.TotalCount:N0}/{workbooks.Count:N0}";
                FormPreflightProgressText.Text =
                    "현재 Excel COM 작업을 정리했고 완료된 부분 결과를 보존했습니다.";
                return;
            }
            _activeFormPreflight = result;
            _formPreflightRows.Clear();
            foreach (var row in result.Items)
                _formPreflightRows.Add(row);
            FormPreflightSummaryText.Text =
                $"판정 완료 {result.TotalCount:N0}개 · "
                + $"기존 양식 {result.KnownFormCount:N0} · "
                + $"유사 검토 {result.SimilarReviewCount:N0} · "
                + $"신규 보류 {result.NewFormCount:N0} · "
                + $"사람 제외 {result.ExcludedFormCount:N0} · "
                + $"추출 실패 {result.CaptureFailedCount:N0}";
            FormPreflightProgressText.Text =
                "기존 양식만 전체 처리 대상으로 사용할 수 있습니다.";
            WpfMessageBox.Show(
                $"COM 사전 분석이 완료됐습니다.\n\n"
                + $"기존 양식: {result.KnownFormCount:N0}개\n"
                + $"유사 양식 검토: {result.SimilarReviewCount:N0}개\n"
                + $"신규 양식 보류: {result.NewFormCount:N0}개\n"
                + $"사람 판정 제외: {result.ExcludedFormCount:N0}개\n"
                + $"COM 추출 실패: {result.CaptureFailedCount:N0}개\n\n"
                + "기존 양식만 전체 처리할 수 있습니다.",
                "Excel COM 사전 분석 완료",
                MessageBoxButton.OK,
                result.HasBlockingItems
                    ? MessageBoxImage.Warning
                    : MessageBoxImage.Information);
        }
        catch (Exception exception)
        {
            FormPreflightSummaryText.Text =
                "COM 사전 분석 중 오류가 발생했습니다.";
            FormPreflightProgressText.Text = exception.Message;
            Log($"[COM 사전 분석] 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "Excel COM 사전 분석 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            _formPreflightRunning = false;
            StopFormPreflightButton.Content = "추출 중지";
            SetCanonicalEvidenceBusy(false);
            RefreshFormPreflightUi();
            UpdateCollectedExcelProcessingAvailability();
        }
    }

    private void StopFormPreflight_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (!_formPreflightRunning)
            return;
        try
        {
            _canonicalEvidenceClient
                .RequestFormPreflightCancellation();
            StopFormPreflightButton.IsEnabled = false;
            StopFormPreflightButton.Content = "중지 요청됨";
            FormPreflightProgressText.Text =
                "현재 전용 Excel COM 프로세스를 정리하는 중입니다...";
        }
        catch (Exception exception) when (
            exception is IOException
            or UnauthorizedAccessException)
        {
            WpfMessageBox.Show(
                exception.Message,
                "사전 분석 중지 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private async void RunKnownFormsFullProcessing_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy)
            return;
        var archivePath = ExcelLocalCopyService.GetLocalCopyBase(
            _pathSettings.ExcelArchiveDirectory);
        var workbooks = GetCollectedExcelFiles(archivePath);
        if (workbooks.Count == 0)
        {
            WpfMessageBox.Show(
                "로컬 보관함에 전체 분석할 Excel이 없습니다.\n\n"
                + archivePath,
                "Excel 전체 자동 분석",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            RefreshFormPreflightUi();
            return;
        }
        var reviewer = FormFamilyReviewerText.Text.Trim();
        if (reviewer.Length == 0)
            reviewer = Environment.UserName;
        var confirmation = WpfMessageBox.Show(
            $"보관함 Excel {workbooks.Count:N0}개를 끝까지 자동 처리합니다.\n\n"
            + "1. 읽기 전용 COM 캡처와 양식 판정\n"
            + "2. 신규·유사 양식군 대표본 AI 계약 검증\n"
            + "3. 검증 결과에 따른 신규 등록·기존 연결·제외\n"
            + "4. 승인 Excel의 PACKET→AI→DB 적재→무결성 검증\n"
            + "5. 최신 적재 DB를 AI 문의 검색 대상으로 반영\n\n"
            + "원본 Excel과 실패 산출물은 삭제하거나 변경하지 않습니다.",
            "Excel 전체 자동 분석·DB 반영",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Information);
        if (confirmation != MessageBoxResult.OK)
            return;

        ResetIngestStages();
        SetIngestStage(
            1,
            "RUNNING",
            $"보관함 {workbooks.Count:N0}개 · COM 사전 분석 재개");
        SetCanonicalEvidenceBusy(true);
        DeveloperConsoleExpander.IsExpanded = true;
        FormPreflightSummaryText.Text =
            $"전체 자동 분석 진행 중 · Excel {workbooks.Count:N0}개";
        FormPreflightProgressText.Text =
            "COM 사전 분석을 시작합니다.";
        WorkspaceStatusText.Text = "Excel 전체 자동 분석 시작";
        var progress = new Progress<IngestProgressEvent>(
            item =>
            {
                var stage = item.Stage switch
                {
                    "FORM_PREFLIGHT" => "1/5 COM·양식 사전 분석",
                    "FORM_FAMILY_REVIEW" => "2/5 양식군 AI 검증·판정",
                    "CORPUS_INGEST" => "3/5 승인 corpus 준비",
                    "CORPUS" => "4/5 의미 분석·DB 적재·검증",
                    _ => item.Stage,
                };
                FormPreflightProgressText.Text =
                    $"{stage} · {item.Detail}";
                WorkspaceStatusText.Text =
                    $"Excel 전체 자동 분석 · {stage}";
                if (!string.IsNullOrWhiteSpace(item.SourcePath))
                {
                    if (string.Equals(
                            item.Status,
                            "RUNNING",
                            StringComparison.OrdinalIgnoreCase))
                        MarkFormPreflightRowRunning(item.SourcePath);
                    else
                        ApplyCompletedFormPreflightRow(item.SourcePath);
                }
                Log(
                    $"[Excel 전체 자동 분석] {stage} · "
                    + $"{item.Status} · {item.Detail}");
            });
        try
        {
            var result = await _canonicalEvidenceClient
                .CompleteFormPipelineAsync(
                    _databasePath,
                    archivePath,
                    reviewer,
                    progress);
            RefreshFormPreflightUi();
            LoadCachedFormGroupReview();
            ResetIngestStages();
            SetIngestStage(
                1,
                "COMPLETED",
                $"사전 분석 {result.PreflightTotalCount:N0}개");
            SetIngestStage(
                2,
                "COMPLETED",
                $"승인 양식 {result.KnownFormCount:N0}개");
            SetIngestStage(
                3,
                "COMPLETED",
                $"의미 패킷 대상 {result.CorpusSelectedCount:N0}개");
            SetIngestStage(
                4,
                result.AnalysisErrorCount == 0
                    ? "COMPLETED"
                    : "FAILED",
                $"양식군 {result.ApprovedFormGroupCount:N0} 승인 · "
                + $"오류 {result.AnalysisErrorCount:N0}");
            SetIngestStage(
                5,
                result.CorpusFailedCount == 0
                    ? "COMPLETED"
                    : "FAILED",
                $"이번 실행 완료 {result.CorpusCompletedCount:N0} · "
                + $"실패 {result.CorpusFailedCount:N0}");
            SetIngestStage(
                6,
                result.IsComplete ? "COMPLETED" : "FAILED",
                result.IsComplete
                    ? "canonical DB 무결성 확인 완료"
                    : "journal 결과 검토 필요");
            SetIngestStage(
                7,
                result.IsComplete ? "COMPLETED" : "FAILED",
                result.IsComplete
                    ? "최신 적재 DB에서 AI 문의 가능"
                    : "미완료 항목 해결 후 AI 문의 반영");
            FormPreflightSummaryText.Text =
                $"{result.Status} · 사전 분석 "
                + $"{result.PreflightTotalCount:N0} · 승인 "
                + $"{result.KnownFormCount:N0} · 양식군 대기 "
                + $"{result.PendingFormGroupCount:N0}";
            FormPreflightProgressText.Text =
                $"corpus 선택 {result.CorpusSelectedCount:N0} · "
                + $"이번 실행 시도 {result.CorpusAttemptedCount:N0} · "
                + $"완료 {result.CorpusCompletedCount:N0} · "
                + $"실패 {result.CorpusFailedCount:N0}";
            ResultTitle.Text = "Excel 전체 자동 분석 결과";
            PipelineFileTitle.Text = archivePath;
            PipelineBatchText.Text =
                $"{result.Status} · 승인 {result.KnownFormCount:N0}개";
            PipelineStagesText.Text =
                $"Result: {result.ResultPath}\n"
                + $"Manifest: {result.ManifestPath}\n"
                + "완료된 canonical Study는 최신 적재 DB AI 문의에 즉시 포함됩니다.";
            NavigateUtf8Html(
                "<!doctype html><meta charset=\"utf-8\">"
                + "<style>body{font-family:'Malgun Gothic';"
                + "background:#181818;color:#ddd;padding:24px}"
                + "code{color:#c4b5fd}</style>"
                + "<h2>Excel 전체 자동 분석 결과</h2>"
                + $"<p>상태 <b>{System.Net.WebUtility.HtmlEncode(result.Status)}</b>"
                + $" · 승인 {result.KnownFormCount:N0}개"
                + $" · corpus 실패 {result.CorpusFailedCount:N0}개</p>"
                + "<p>COM 캡처 → 양식군 판정 → 의미 패킷 → AI 분석 → "
                + "canonical DB 적재·검증을 같은 journal로 재개할 수 있습니다.</p>"
                + "<p>질문 화면의 <b>최신 적재 DB</b>가 완료된 Study를 "
                + "즉시 검색합니다.</p><code>"
                + System.Net.WebUtility.HtmlEncode(result.ResultPath)
                + "</code>");
            DeveloperConsoleExpander.IsExpanded = false;
            WpfMessageBox.Show(
                result.IsComplete
                    ? "전체 자동 분석과 DB 반영이 완료됐습니다.\n\n"
                      + "질문 화면의 ‘최신 적재 DB’에서 새 Excel 분석을 조회할 수 있습니다."
                    : "전체 실행은 끝났지만 미완료 항목이 있습니다.\n\n"
                      + FormPreflightProgressText.Text,
                "Excel 전체 자동 분석 결과",
                MessageBoxButton.OK,
                result.IsComplete
                    ? MessageBoxImage.Information
                    : MessageBoxImage.Warning);
        }
        catch (Exception exception)
        {
            MarkRunningIngestStageFailed(
                IngestErrorSummary(exception.Message));
            FormPreflightSummaryText.Text =
                "전체 자동 분석 중 오류가 발생했습니다.";
            FormPreflightProgressText.Text = exception.Message;
            Log($"[Excel 전체 자동 분석] 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "Excel 전체 자동 분석 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
            UpdateCollectedExcelProcessingAvailability();
        }
    }

    private IReadOnlyList<string> GetCollectedExcelFiles(
        string archivePath)
    {
        if (!Directory.Exists(archivePath))
            return [];
        try
        {
            return Directory
                .EnumerateFiles(
                    archivePath,
                    "*",
                    SearchOption.TopDirectoryOnly)
                .Where(IsExcelInputFile)
                .Select(Path.GetFullPath)
                .OrderBy(
                    path => path,
                    StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }
        catch (
            Exception exception
        ) when (
            exception is IOException
            or UnauthorizedAccessException
            or ArgumentException
            or NotSupportedException)
        {
            return [];
        }
    }

    private void ConfigureArchiveIngestSource(bool resetStages)
    {
        var archivePath = ExcelLocalCopyService.GetLocalCopyBase(
            _pathSettings.ExcelArchiveDirectory);
        if (IngestSourceModeSelector.SelectedIndex != 1)
            IngestSourceModeSelector.SelectedIndex = 1;
        IngestWorkbookPathText.Text = archivePath;
        if (resetStages)
            ResetIngestStages();
        IngestWorkbookState.Text =
            "신규 Excel 보관함을 자동으로 사용합니다. 파일 선택은 필요하지 않습니다.";
        UpdateCollectedExcelProcessingAvailability();
    }

    private void OpenMatchedDatabaseExcel_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (sender is not FrameworkElement
            {
                DataContext: ExcelFolderSearchRow
                {
                    ExistsInDatabase: true
                } row
            })
        {
            return;
        }

        var databaseSourcePath = row.DatabaseSourcePath;
        if (string.IsNullOrWhiteSpace(databaseSourcePath)
            || !File.Exists(databaseSourcePath))
        {
            WpfMessageBox.Show(
                "DB에 연결된 Excel 경로의 파일을 찾을 수 없습니다.\n\n"
                + databaseSourcePath,
                "DB 연결 Excel 열기",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        Process.Start(
            new ProcessStartInfo(databaseSourcePath)
            {
                UseShellExecute = true,
            });
        Log(
            $"DB 연결 Excel 열기: {row.FileName} → "
            + databaseSourcePath);
    }

    private void UpdateExcelFolderSearchSummary(
        string detail)
    {
        var existing = _excelFolderSearchRows.Count(
            row => row.ExistsInDatabase);
        var missing = _excelFolderSearchRows.Count - existing;
        ExcelFolderSearchSummary.Text =
            $"{_excelFolderSearchRoot} · "
            + $"전체 {_excelFolderSearchRows.Count:N0}개 · "
            + $"DB 있음 {existing:N0}개 · "
            + $"DB 없음 {missing:N0}개 · {detail}";
    }

    private void UpdateExcelSearchActionAvailability()
    {
        CopyMissingExcelSearchButton.IsEnabled =
            !_canonicalEvidenceBusy
            && !_excelFolderSearchBusy
            && !_excelFolderSearchCopyCompleted
            && _excelFolderSearchRows.Any(
                row => !row.ExistsInDatabase);
        UpdateCollectedExcelProcessingAvailability();
    }

    private void UpdateCollectedExcelProcessingAvailability()
    {
        var archivePath = ExcelLocalCopyService.GetLocalCopyBase(
            _pathSettings.ExcelArchiveDirectory);
        var count = GetCollectedExcelFiles(archivePath).Count;
        ProcessCollectedExcelButton.Content = count > 0
            ? $"보관함 {count:N0}개 COM 사전 분석"
            : "Excel COM 사전 분석";
        ProcessCollectedExcelButton.IsEnabled =
            !_canonicalEvidenceBusy
            && !_excelFolderSearchBusy
            && count > 0;
        var latest = TryLoadMatchingFormPreflight(archivePath);
        CollectedExcelProcessingState.Text = count == 0
            ? "사전 분석 대기 · 로컬 보관함에 Excel이 없습니다."
            : latest is null
                ? $"COM 사전 분석 필요 · 로컬 보관함 Excel {count:N0}개"
                : $"최근 판정 · 기존 {latest.KnownFormCount:N0} · "
                  + $"유사 검토 {latest.SimilarReviewCount:N0} · "
                  + $"신규 보류 {latest.NewFormCount:N0} · "
                  + $"제외 {latest.ExcludedFormCount:N0} · "
                  + $"추출 실패 {latest.CaptureFailedCount:N0}";
    }

    private void RefreshIngestRetryAvailability()
    {
        var sourcePath = IngestWorkbookPathText.Text.Trim();
        if (IsIngestFolderMode() || !File.Exists(sourcePath))
        {
            _selectedIngestHasFailedJournal = false;
            RetryIngestButton.IsEnabled = false;
            _activeIngestResult = null;
            _activeIngestRelated = null;
            InspectIngestResultButton.IsEnabled = false;
            return;
        }
        var failedJournalFound = false;
        var artifactRoot =
            _pathSettings.IncrementalIngestDirectory;
        try
        {
            foreach (var journalPath in Directory.EnumerateFiles(
                         artifactRoot,
                         "journal.json",
                         SearchOption.AllDirectories))
            {
                try
                {
                    using var document = JsonDocument.Parse(
                        File.ReadAllText(journalPath));
                    var root = document.RootElement;
                    if (!root.TryGetProperty("status", out var status)
                        || !string.Equals(
                            status.GetString(),
                            "FAILED",
                            StringComparison.OrdinalIgnoreCase)
                        || !root.TryGetProperty("source", out var source)
                        || !source.TryGetProperty(
                            "sourcePath",
                            out var journalSource))
                    {
                        continue;
                    }
                    var journalSourcePath = journalSource.GetString();
                    if (!string.IsNullOrWhiteSpace(journalSourcePath)
                        && string.Equals(
                            Path.GetFullPath(journalSourcePath),
                            Path.GetFullPath(sourcePath),
                            StringComparison.OrdinalIgnoreCase))
                    {
                        failedJournalFound = true;
                        break;
                    }
                }
                catch (
                    Exception exception
                ) when (
                    exception is IOException
                    or UnauthorizedAccessException
                    or JsonException
                    or ArgumentException
                    or NotSupportedException)
                {
                    // Ignore unrelated partial or inaccessible journals.
                }
            }
        }
        catch (
            Exception exception
        ) when (
            exception is IOException
            or UnauthorizedAccessException
            or DirectoryNotFoundException)
        {
            // A missing artifact directory simply means there is no retry.
        }
        _selectedIngestHasFailedJournal = failedJournalFound;
        _activeIngestResult =
            _canonicalEvidenceClient.LoadLatestIngestResult(sourcePath);
        _activeIngestRelated = null;
        RetryIngestButton.IsEnabled = failedJournalFound;
        InspectIngestResultButton.IsEnabled =
            !_canonicalEvidenceBusy && _activeIngestResult is not null;
        IngestWorkbookState.Text = failedJournalFound
            ? "기존 실패 journal 발견 · ‘실패 단계 재시도’로 이어서 처리하세요."
            : "DRM Excel 원본 확인 완료 · ‘전체 실행’을 누르세요.";
    }

    private void IngestSourceModeSelector_SelectionChanged(
        object sender,
        SelectionChangedEventArgs e)
    {
        if (BrowseIngestSourceButton is null) return;
        BrowseIngestSourceButton.Content = IsIngestFolderMode()
            ? "폴더 선택"
            : "Excel 선택";
        _selectedIngestHasFailedJournal = false;
        _activeIngestResult = null;
        _activeIngestRelated = null;
        InspectIngestResultButton.IsEnabled = false;
        IngestWorkbookPathText.Text = string.Empty;
        IngestWorkbookState.Text = IsIngestFolderMode()
            ? "DRM Excel 폴더를 선택하세요."
            : "DRM Excel 원본을 선택하세요.";
        ResetIngestStages();
    }

    private async void IngestWorkbook_Click(
        object sender,
        RoutedEventArgs e)
    {
        await RunIngestWorkflowAsync(retryFailed: false);
    }

    private async void RetryIngest_Click(
        object sender,
        RoutedEventArgs e)
    {
        await RunIngestWorkflowAsync(retryFailed: true);
    }

    private async void InspectIngestResult_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy || IsIngestFolderMode()) return;
        var sourcePath = IngestWorkbookPathText.Text.Trim();
        var result = _activeIngestResult;
        if (result is null
            || !string.Equals(
                Path.GetFullPath(result.SourcePath),
                Path.GetFullPath(sourcePath),
                StringComparison.OrdinalIgnoreCase))
        {
            result = _canonicalEvidenceClient.LoadLatestIngestResult(
                sourcePath);
            _activeIngestResult = result;
            _activeIngestRelated = null;
        }
        if (result is null)
        {
            WpfMessageBox.Show(
                "선택한 Excel의 완료된 처리 결과를 찾을 수 없습니다.",
                "처리 내용 검증",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            InspectIngestResultButton.IsEnabled = false;
            return;
        }

        SetCanonicalEvidenceBusy(true);
        IngestWorkbookState.Text = "저장된 Study와 원본 근거를 불러오는 중...";
        try
        {
            var related = _activeIngestRelated
                ?? await _canonicalEvidenceClient.RelatedAsync(
                    _databasePath,
                    result.RevisionUid);
            _activeIngestRelated = related;
            var html = IngestVerificationHtmlRenderer.Render(
                result,
                related);
            var window = new IngestVerificationWindow(
                result,
                html,
                _pathSettings.TemporaryDirectory)
            {
                Owner = this,
            };
            window.Show();
            IngestWorkbookState.Text =
                $"{result.Status} · Study {result.StudyCount:N0}건 · "
                + "검증 창 열림";
            Log(
                $"처리 내용 검증 창 열기: {result.SourcePath} / "
                + $"{result.PublicAnalysisId}");
        }
        catch (Exception exception)
        {
            IngestWorkbookState.Text = "처리 내용 검증 창 열기 실패";
            Log($"처리 내용 검증 창 열기 실패: {exception.Message}");
            WpfMessageBox.Show(
                exception.Message,
                "처리 내용 검증 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private void GoToEvidenceQuestion_Click(
        object sender,
        RoutedEventArgs e)
    {
        WorkspaceTabs.SelectedIndex = 2;
        EvidenceQueryModeSelector.SelectedIndex = 0;
        EvidenceQuestionText.Focus();
    }

    private bool IsIngestFolderMode() =>
        (IngestSourceModeSelector.SelectedItem as ComboBoxItem)
            ?.Tag?.ToString() == "folder";

    private async Task RunIngestWorkflowAsync(bool retryFailed)
    {
        if (_canonicalEvidenceBusy) return;
        var sourcePath = IngestWorkbookPathText.Text.Trim();
        var folderMode = IsIngestFolderMode();
        var validSource = folderMode
            ? Directory.Exists(sourcePath)
            : File.Exists(sourcePath)
              && new[] { ".xlsx", ".xlsm", ".xlsb", ".xls" }
                  .Contains(
                      Path.GetExtension(sourcePath),
                      StringComparer.OrdinalIgnoreCase);
        if (!validSource)
        {
            WpfMessageBox.Show(
                folderMode
                    ? "처리할 Excel 폴더를 선택하세요."
                    : "처리할 .xlsx/.xlsm/.xlsb/.xls 파일을 선택하세요.",
                "DRM Excel 전체 처리",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        var effectiveRetry = retryFailed
            || (!folderMode && _selectedIngestHasFailedJournal);
        FormPreflightDocument? preflight = null;
        if (folderMode)
        {
            preflight = TryLoadMatchingFormPreflight(sourcePath);
            if (preflight is null || preflight.KnownFormCount == 0)
            {
                WpfMessageBox.Show(
                    "전체 처리 전에 현재 보관함의 Excel COM 사전 분석을 완료해야 합니다.\n\n"
                    + "사전 분석에서 기존 양식으로 판정된 파일만 전체 처리할 수 있습니다.",
                    "Excel COM 사전 분석 필요",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                WorkspaceTabs.SelectedIndex = 10;
                RefreshFormPreflightUi();
                return;
            }
            if (!effectiveRetry)
            {
                var held = preflight.SimilarReviewCount
                           + preflight.NewFormCount
                           + preflight.ExcludedFormCount
                           + preflight.CaptureFailedCount;
                var confirmation = WpfMessageBox.Show(
                    $"사전 분석을 통과한 기존 양식 {preflight.KnownFormCount:N0}개를 전체 처리합니다.\n\n"
                    + $"자동 보류: {held:N0}개\n"
                    + "사전 COM 추출 결과를 재사용해 AI 분석과 DB 반영을 진행합니다.",
                    "기존 양식 전체 처리",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Information);
                if (confirmation != MessageBoxResult.OK)
                    return;
            }
        }
        if (!TryGetAuthDialogSettings(
                out var inspectAuth,
                out var dismissAuth,
                out var authTitle,
                out var authClass,
                out var authButton))
            return;

        ResetIngestStages();
        _activeIngestResult = null;
        _activeIngestRelated = null;
        InspectIngestResultButton.IsEnabled = false;
        SetIngestStage(
            1,
            "COMPLETED",
            folderMode
                ? $"사전 분석 통과 {preflight!.KnownFormCount:N0}개"
                : $"{Path.GetFileName(sourcePath)} · 읽기 전용 처리");
        SetCanonicalEvidenceBusy(true);
        RetryIngestButton.IsEnabled = false;
        DeveloperConsoleExpander.IsExpanded = true;
        IngestWorkbookState.Text =
            effectiveRetry
                ? "journal에서 실패 단계부터 다시 시작합니다."
                : folderMode
                    ? "사전 COM 추출 결과를 재사용해 전체 처리를 시작합니다."
                    : "Excel COM 추출을 시작합니다.";
        Log(
            $"[전체 처리] {(folderMode ? "폴더" : "파일")} 시작: "
            + sourcePath);
        var progress = new Progress<IngestProgressEvent>(
            HandleIngestProgress);
        try
        {
            if (folderMode)
            {
                var result = await _canonicalEvidenceClient
                    .IngestCorpusAsync(
                        _databasePath,
                        sourcePath,
                        effectiveRetry,
                        progress,
                        inspectAuth,
                        dismissAuth,
                        authTitle,
                        authClass,
                        authButton,
                        sourceManifestPath:
                            preflight!.KnownFormManifestPath);
                if (result.FailedCount > 0)
                    throw new InvalidOperationException(
                        $"폴더 처리 중 {result.FailedCount:N0}개 파일이 "
                        + "실패했습니다. 실패 단계 재시도를 사용하세요. "
                        + $"Journal: {result.JournalPath}");
                CompletePendingIngestStages(
                    "journal 재사용 또는 폴더 처리 완료");
                SetIngestStage(
                    7,
                    "COMPLETED",
                    "최신 적재 DB 질문 모드 사용 가능");
                IngestWorkbookState.Text =
                    $"{result.Status} · 선택 {result.SelectedCount:N0} · "
                    + $"완료 {result.CompletedCount:N0}";
                ResultTitle.Text = "DRM Excel 폴더 처리 결과";
                PipelineFileTitle.Text = sourcePath;
                PipelineBatchText.Text =
                    $"완료 {result.CompletedCount:N0} · "
                    + $"실패 {result.FailedCount:N0}";
                PipelineStagesText.Text =
                    $"Journal: {result.JournalPath}\n"
                    + "질문 화면의 ‘최신 적재 DB’가 새 자료를 즉시 포함합니다.";
                NavigateUtf8Html(
                    "<!doctype html><meta charset=\"utf-8\">"
                    + "<style>body{font-family:'Malgun Gothic';"
                    + "background:#181818;color:#ddd;padding:24px}"
                    + "code{color:#c4b5fd}</style>"
                    + "<h2>DRM Excel 폴더 처리 완료</h2>"
                    + $"<p>완료 {result.CompletedCount:N0}개 · "
                    + $"실패 {result.FailedCount:N0}개</p>"
                    + "<p>질문 화면에서 <b>최신 적재 DB</b>를 선택하면 "
                    + "새로 반영한 자료를 바로 조회합니다.</p>"
                    + "<code>"
                    + System.Net.WebUtility.HtmlEncode(
                        result.JournalPath)
                    + "</code>");
            }
            else
            {
                var result = await _canonicalEvidenceClient.IngestAsync(
                    _databasePath,
                    sourcePath,
                    effectiveRetry,
                    progress,
                    inspectAuth,
                    dismissAuth,
                    authTitle,
                    authClass,
                    authButton);
                var related = await _canonicalEvidenceClient.RelatedAsync(
                    _databasePath,
                    result.RevisionUid);
                _activeIngestResult = result;
                _activeIngestRelated = related;
                InspectIngestResultButton.IsEnabled = true;
                SetIngestStage(
                    7,
                    "COMPLETED",
                    "관련 자료 연결 완료 · AI 문의 가능");
                IngestWorkbookState.Text =
                    $"{result.Status} · {result.WorkbookStatus} · "
                    + $"Study {result.StudyCount:N0}건";
                ResultTitle.Text = "DRM Excel 전체 처리 결과";
                PipelineFileTitle.Text = Path.GetFileName(sourcePath);
                PipelineBatchText.Text =
                    $"{result.RevisionUid} · {result.Status}";
                PipelineStagesText.Text =
                    $"Journal: {result.JournalPath}\n"
                    + "질문 화면의 ‘최신 적재 DB’가 새 자료를 즉시 포함합니다.";
                NavigateUtf8Html(
                    EvidenceHtmlRenderer.RenderIngest(result, related));
                Log(
                    $"[전체 처리] 완료: {result.Status}, "
                    + $"{result.WorkbookStatus}, "
                    + $"Study {result.StudyCount}");
            }
            DeveloperConsoleExpander.IsExpanded = false;
        }
        catch (Exception exception)
        {
            var errorSummary = IngestErrorSummary(exception.Message);
            IngestWorkbookState.Text =
                "실패 · journal을 유지했습니다. ‘실패 단계 재시도’를 누르세요.";
            RetryIngestButton.IsEnabled = true;
            MarkRunningIngestStageFailed(errorSummary);
            Log($"[전체 처리] 실패: {exception.Message}");
            WpfMessageBox.Show(
                "전체 처리 중 오류가 발생했습니다.\n\n"
                + errorSummary
                + "\n\n상세 Traceback은 Developer console에 남겼습니다.",
                "DRM Excel 전체 처리 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetCanonicalEvidenceBusy(false);
        }
    }

    private void HandleIngestProgress(IngestProgressEvent progress)
    {
        var folderMode = IsIngestFolderMode();
        if (folderMode && progress.Stage == "CORPUS")
        {
            HandleCorpusProgress(progress);
            return;
        }
        var stageNumber = progress.Stage switch
        {
            "CAPTURE" or "AUTH_DIALOG" => 2,
            "PACKET" => 3,
            "LOCATOR" or "DRAFT" => 4,
            "IMPORT" => 5,
            "VERIFY" => 6,
            "WORKFLOW" => 7,
            _ => 0,
        };
        if (stageNumber == 0) return;
        var effectiveStatus = progress.Status;
        var detail = progress.Stage switch
        {
            "AUTH_DIALOG" =>
                "인증창 감지 · 개발자 로그에서 정확한 정보를 확인하세요.",
            "LOCATOR" when progress.Status == "RUNNING" =>
                "4-1/2 관련 구간 탐색 중",
            "LOCATOR" when progress.Status == "COMPLETED" =>
                "4-1/2 관련 구간 탐색 완료 · Study 구성 대기",
            "DRAFT" when progress.Status == "RUNNING" =>
                "4-2/2 Study·시험군·비교군 구성 중",
            "DRAFT" when progress.Status == "COMPLETED" =>
                "4-2/2 Study·시험군·비교군 구성 완료",
            _ when string.IsNullOrWhiteSpace(progress.SourcePath) =>
                progress.Detail,
            _ => Path.GetFileName(progress.SourcePath),
        };
        if (progress.Stage == "LOCATOR"
            && progress.Status == "COMPLETED")
        {
            effectiveStatus = "RUNNING";
        }
        if (folderMode)
        {
            HandleFolderStageProgress(
                progress,
                stageNumber,
                effectiveStatus,
                detail);
            return;
        }
        SetIngestStage(stageNumber, effectiveStatus, detail);
        var state = effectiveStatus switch
        {
            "RUNNING" => "처리 중",
            "COMPLETED" => "완료",
            "FAILED" => "실패",
            "WAITING" => "사용자 확인 대기",
            _ => progress.Status,
        };
        IngestWorkbookState.Text =
            $"{stageNumber}/7 {_ingestStages[stageNumber - 1].Title} · "
            + state;
        Log(
            $"[전체 처리 {stageNumber}/7] "
            + $"{_ingestStages[stageNumber - 1].Title} · {state}"
            + (
                string.IsNullOrWhiteSpace(detail)
                    ? string.Empty
                    : $" · {detail}"
            ));
    }

    private void HandleCorpusProgress(IngestProgressEvent progress)
    {
        try
        {
            using var document = JsonDocument.Parse(progress.Detail);
            var root = document.RootElement;
            if (progress.Status == "RUNNING")
            {
                var selected = JsonInt(root, "selected");
                var eligible = JsonInt(root, "eligible");
                var skippedCompleted = JsonInt(
                    root,
                    "skippedCompleted");
                var skippedFailed = JsonInt(root, "skippedFailed");
                var reconciledExisting = JsonInt(
                    root,
                    "reconciledExisting");
                _folderIngestTotal = eligible;
                foreach (var stage in _ingestStages)
                    stage.SetProgress(0, 0, eligible);
                _ingestStages[0].SetProgress(
                    eligible,
                    0,
                    eligible);
                _ingestStages[0].SetState(
                    "COMPLETED",
                    $"선택 {selected:N0} · 이번 처리 {eligible:N0}"
                    + (
                        skippedCompleted > 0
                            ? $" · 기존 완료 {skippedCompleted:N0}"
                            : string.Empty
                    )
                    + (
                        skippedFailed > 0
                            ? $" · 재시도 제외 {skippedFailed:N0}"
                            : string.Empty
                    )
                    + (
                        reconciledExisting > 0
                            ? $" · DB 기존 파일명 {reconciledExisting:N0}개 자동 스킵"
                            : string.Empty
                    ));
                IngestWorkbookState.Text =
                    $"파이프라인 시작 · 처리 대상 {eligible:N0}개";
                Log(
                    "[전체 처리] bounded pipeline 시작 · "
                    + $"대상 {eligible:N0}개 · "
                    + "COM 1 / 패킷 3 / AI 3 / DB 1");
                return;
            }

            var failed = JsonInt(root, "failedThisRun");
            var completed = JsonInt(root, "completedThisRun");
            IngestWorkbookState.Text =
                $"폴더 처리 집계 · 완료 {completed:N0} · 실패 {failed:N0}";
            Log(
                "[전체 처리] bounded pipeline 종료 · "
                + $"완료 {completed:N0} · 실패 {failed:N0}");
        }
        catch (JsonException)
        {
            Log(
                "[전체 처리] corpus 진행률 JSON을 읽지 못했습니다: "
                + progress.Detail);
        }
    }

    private void HandleFolderStageProgress(
        IngestProgressEvent progress,
        int stageNumber,
        string effectiveStatus,
        string detail)
    {
        var sourceKey = string.IsNullOrWhiteSpace(progress.SourcePath)
            ? progress.SourcePath
            : Path.GetFullPath(progress.SourcePath);
        var terminalCompleted = progress.Status == "COMPLETED"
            && progress.Stage != "LOCATOR";
        if (terminalCompleted)
        {
            _folderStageFailed[stageNumber].Remove(sourceKey);
            _folderStageCompleted[stageNumber].Add(sourceKey);
        }
        else if (progress.Status == "FAILED")
        {
            _folderStageCompleted[stageNumber].Remove(sourceKey);
            _folderStageFailed[stageNumber].Add(sourceKey);
        }

        var completed = _folderStageCompleted[stageNumber].Count;
        var failed = _folderStageFailed[stageNumber].Count;
        var finished = completed + failed;
        var aggregateStatus = effectiveStatus;
        if (_folderIngestTotal > 0
            && finished >= _folderIngestTotal)
        {
            aggregateStatus = failed > 0
                ? "FAILED"
                : "COMPLETED";
        }
        else if (terminalCompleted || progress.Status == "FAILED")
        {
            aggregateStatus = "RUNNING";
        }
        var fileName = string.IsNullOrWhiteSpace(progress.SourcePath)
            ? string.Empty
            : Path.GetFileName(progress.SourcePath);
        var aggregateDetail = _folderIngestTotal > 0
            ? $"{finished:N0}/{_folderIngestTotal:N0}"
              + (
                  failed > 0
                      ? $" · 실패 {failed:N0}"
                      : string.Empty
              )
              + (
                  string.IsNullOrWhiteSpace(fileName)
                      ? string.Empty
                      : $" · {fileName}"
              )
              + (
                  string.IsNullOrWhiteSpace(detail)
                  || string.Equals(
                      detail,
                      fileName,
                      StringComparison.Ordinal)
                      ? string.Empty
                      : $" · {detail}"
              )
            : detail;
        var stage = _ingestStages[stageNumber - 1];
        stage.SetProgress(
            completed,
            failed,
            _folderIngestTotal);
        stage.SetState(aggregateStatus, aggregateDetail);
        var state = aggregateStatus switch
        {
            "RUNNING" => "처리 중",
            "COMPLETED" => "완료",
            "FAILED" => "실패",
            "WAITING" => "사용자 확인 대기",
            _ => aggregateStatus,
        };
        IngestWorkbookState.Text =
            $"{stageNumber}/7 {stage.Title} · {state}"
            + (
                _folderIngestTotal > 0
                    ? $" · {finished:N0}/{_folderIngestTotal:N0}"
                    : string.Empty
            );
        Log(
            $"[전체 처리 {stageNumber}/7] {stage.Title} · "
            + $"{state} · {aggregateDetail}");
    }

    private static int JsonInt(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value)
        && value.TryGetInt32(out var result)
            ? result
            : 0;

    private void ResetIngestStages()
    {
        _folderIngestTotal = 0;
        foreach (var values in _folderStageCompleted.Values)
            values.Clear();
        foreach (var values in _folderStageFailed.Values)
            values.Clear();
        foreach (var stage in _ingestStages)
        {
            stage.SetState("PENDING", string.Empty);
            stage.SetProgress(0, 0, 0);
        }
        RetryIngestButton.IsEnabled = false;
    }

    private void SetIngestStage(
        int stageNumber,
        string status,
        string detail)
    {
        if (stageNumber < 1 || stageNumber > _ingestStages.Count)
            return;
        _ingestStages[stageNumber - 1].SetState(status, detail);
    }

    private void MarkRunningIngestStageFailed(string detail)
    {
        var failed = _ingestStages.LastOrDefault(
            stage => stage.Status == "FAILED");
        if (failed is not null)
        {
            failed.SetState("FAILED", detail);
            return;
        }
        var running = _ingestStages.LastOrDefault(
            stage => stage.Status == "RUNNING");
        (running
         ?? _ingestStages.FirstOrDefault(
             stage => stage.Status == "PENDING")
         ?? _ingestStages.LastOrDefault())
            ?.SetState("FAILED", detail);
    }

    private static string IngestErrorSummary(string message)
    {
        var lines = message
            .Split(
                ['\r', '\n'],
                StringSplitOptions.RemoveEmptyEntries
                | StringSplitOptions.TrimEntries)
            .Where(line =>
                !line.All(character => character == '^'))
            .ToList();
        var coverageLine = lines.LastOrDefault(line =>
            line.Contains(
                "ContentCoverageError:",
                StringComparison.Ordinal));
        if (coverageLine is not null)
        {
            const string prefix =
                "Source content coverage is incomplete; ";
            var prefixIndex = coverageLine.IndexOf(
                prefix,
                StringComparison.Ordinal);
            var detail = prefixIndex >= 0
                ? coverageLine[(prefixIndex + prefix.Length)..]
                : coverageLine[
                    (coverageLine.IndexOf(
                        ':',
                        StringComparison.Ordinal) + 1)..].Trim();
            var coverageSummary =
                "AI 초안이 원본 근거를 모두 포함하지 못했습니다: "
                + detail
                + ". ‘실패 단계 재시도’를 누르면 누락 근거를 보정합니다.";
            return coverageSummary.Length <= 320
                ? coverageSummary
                : coverageSummary[..317] + "...";
        }
        var summary = lines.LastOrDefault()
            ?? "알 수 없는 처리 오류입니다.";
        return summary.Length <= 260
            ? summary
            : summary[..257] + "...";
    }

    private void CompletePendingIngestStages(string detail)
    {
        foreach (var stage in _ingestStages.Where(
                     stage => stage.Status == "PENDING"))
            stage.SetState("COMPLETED", detail);
    }

    private void SetCanonicalEvidenceBusy(bool busy)
    {
        _canonicalEvidenceBusy = busy;
        AskEvidenceButton.IsEnabled = !busy;
        LoadEvidenceHistoryButton.IsEnabled = !busy;
        AskEvidenceButton.Content = busy
            ? "관련 보고서 찾는 중..."
            : "관련 보고서 찾기";
        IngestWorkbookButton.IsEnabled = !busy;
        BrowseIngestSourceButton.IsEnabled = !busy;
        ExcelFolderSearchButton.IsEnabled =
            !busy && !_excelFolderSearchBusy;
        ExcelMultiFolderSearchButton.IsEnabled =
            !busy && !_excelFolderSearchBusy;
        UpdateExcelSearchActionAvailability();
        RefreshFormPreflightButton.IsEnabled = !busy;
        RunFormPreflightButton.IsEnabled =
            !busy
            && GetCollectedExcelFiles(
                ExcelLocalCopyService.GetLocalCopyBase(
                    _pathSettings.ExcelArchiveDirectory)).Count > 0;
        RunKnownFormsFullProcessingButton.IsEnabled =
            !busy
            && GetCollectedExcelFiles(
                ExcelLocalCopyService.GetLocalCopyBase(
                    _pathSettings.ExcelArchiveDirectory)).Count > 0;
        StopFormPreflightButton.IsEnabled =
            busy && _formPreflightRunning;
        RefreshFormGroupsButton.IsEnabled = !busy;
        FormFamilyGroupsGrid.IsEnabled = !busy;
        FormFamilyReviewerText.IsEnabled = !busy;
        FormFamilyDisplayNameText.IsEnabled = !busy;
        FormFamilyLinkedSignatureText.IsEnabled = !busy;
        FormFamilyNotesText.IsEnabled = !busy;
        UpdateFormFamilyDecisionAvailability();
        // The grid is read-only, so keep scrolling and row inspection
        // available while the external COM preflight is running.
        FormPreflightGrid.IsEnabled = true;
        IngestSourceModeSelector.IsEnabled = !busy;
        InspectAuthDialogCheck.IsEnabled = !busy;
        DismissAuthDialogCheck.IsEnabled = !busy;
        AuthDialogTitleText.IsEnabled = !busy;
        AuthDialogClassText.IsEnabled = !busy;
        AuthDialogButtonText.IsEnabled = !busy;
        EvidenceQueryModeSelector.IsEnabled = !busy;
        InspectIngestResultButton.IsEnabled =
            !busy && _activeIngestResult is not null;
        EvidenceCitationsGrid.IsEnabled = !busy;
        RefreshReviewQueueButton.IsEnabled = !busy;
        ReviewQueueGrid.IsEnabled = !busy;
        ReviewEvidenceGrid.IsEnabled = !busy;
        ApproveReviewButton.IsEnabled = !busy;
        RejectReviewButton.IsEnabled = !busy;
        ExcludeReviewButton.IsEnabled = !busy;
        ReturnReviewButton.IsEnabled = !busy;
        var comparisonEnabled = !busy && !_workbookComparisonBusy;
        RefreshWorkbookComparisonButton.IsEnabled = comparisonEnabled;
        OpenWorkbookComparisonExcelButton.IsEnabled =
            comparisonEnabled;
        WorkbookComparisonGrid.IsEnabled = comparisonEnabled;
        WorkbookComparisonValuesGrid.IsEnabled = comparisonEnabled;
        UpdateConceptDecisionAvailability();
    }

    private void NavigateQuestionInsight(string html)
    {
        _questionInsightSnapshot = new QuestionInsightSnapshot(
            ResultTitle.Text,
            PipelineFileTitle.Text,
            PipelineBatchText.Text,
            PipelineStagesText.Text,
            html);
        NavigateUtf8Html(html);
    }

    private void ShowQuestionWorkspaceInsight()
    {
        if (_questionInsightSnapshot is not null)
        {
            ResultTitle.Text = _questionInsightSnapshot.ResultTitle;
            PipelineFileTitle.Text =
                _questionInsightSnapshot.FileTitle;
            PipelineBatchText.Text =
                _questionInsightSnapshot.BatchText;
            PipelineStagesText.Text =
                _questionInsightSnapshot.StagesText;
            NavigateUtf8Html(_questionInsightSnapshot.Html);
            return;
        }

        ResultTitle.Text = "질문 결과";
        PipelineFileTitle.Text =
            "질문을 입력하면 관련 보고서가 여기에 표시됩니다.";
        PipelineBatchText.Text =
            "질문 입력 → DB 후보 검색 → AI 관련성 판정 → 원본 근거 연결";
        PipelineStagesText.Text =
            "AI는 보고서 관련성만 판단합니다. "
            + "수치의 의미·증감·원인 판단은 하지 않습니다.";
        NavigateUtf8Html(
            """
            <!doctype html>
            <html lang="ko">
            <head>
              <meta charset="utf-8">
              <style>
                body{margin:0;padding:36px;font-family:'Segoe UI','Malgun Gothic',sans-serif;background:#181818;color:#d8d8d8}
                .eyebrow{font-size:11px;font-weight:700;letter-spacing:.08em;color:#a78bfa}
                h1{margin:10px 0 8px;font-size:25px;color:#f0f0f0}
                p{margin:0;color:#9a9a9a;line-height:1.7}
                .steps{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:26px}
                .step{padding:16px;border:1px solid #383838;border-radius:8px;background:#232323}
                .number{font-size:11px;font-weight:700;color:#c4b5fd}
                .title{margin-top:8px;font-weight:700;color:#e8e8e8}
                .detail{margin-top:5px;font-size:12px;line-height:1.5;color:#858585}
              </style>
            </head>
            <body>
              <div class="eyebrow">QUESTION WORKSPACE</div>
              <h1>찾고 싶은 내용을 질문해 주세요.</h1>
              <p>관련 보고서와 Study를 선별하고, 사람이 확인할 수 있는 원본 Excel 범위를 연결합니다.</p>
              <div class="steps">
                <div class="step"><div class="number">01</div><div class="title">질문 입력</div><div class="detail">대상·조건·지표를 자연어로 작성합니다.</div></div>
                <div class="step"><div class="number">02</div><div class="title">보고서 선별</div><div class="detail">DB 후보에서 질문과 직접 관련된 자료만 찾습니다.</div></div>
                <div class="step"><div class="number">03</div><div class="title">원본 확인</div><div class="detail">선택된 보고서의 시트와 정확한 셀 범위를 엽니다.</div></div>
              </div>
            </body>
            </html>
            """);
    }

    private sealed record QuestionInsightSnapshot(
        string ResultTitle,
        string FileTitle,
        string BatchText,
        string StagesText,
        string Html);
}

internal sealed class IngestStageViewModel : INotifyPropertyChanged
{
    private static readonly Brush PendingBackground =
        CreateBrush(0x20, 0x20, 0x20);
    private static readonly Brush RunningBackground =
        CreateBrush(0x2C, 0x21, 0x3E);
    private static readonly Brush CompletedBackground =
        CreateBrush(0x1D, 0x31, 0x2B);
    private static readonly Brush FailedBackground =
        CreateBrush(0x3A, 0x20, 0x24);
    private static readonly Brush PendingBorder =
        CreateBrush(0x3A, 0x3A, 0x3A);
    private static readonly Brush RunningBrush =
        CreateBrush(0xA7, 0x8B, 0xFA);
    private static readonly Brush CompletedBrush =
        CreateBrush(0x5E, 0xD1, 0xA7);
    private static readonly Brush FailedBrush =
        CreateBrush(0xFF, 0x7B, 0x88);
    private static readonly Brush PendingBrush =
        CreateBrush(0x68, 0x68, 0x68);

    internal IngestStageViewModel(int number, string title)
    {
        Number = number;
        Title = title;
        SetState("PENDING", string.Empty);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public int Number { get; }
    public string Title { get; }
    public string Status { get; private set; } = "PENDING";
    public string StatusText { get; private set; } = "대기";
    public string Detail { get; private set; } = string.Empty;
    public Brush Background { get; private set; } = PendingBackground;
    public Brush BorderBrush { get; private set; } = PendingBorder;
    public Brush MarkerBrush { get; private set; } = PendingBrush;
    public Brush StatusBrush { get; private set; } = PendingBrush;
    public int ProgressMaximum { get; private set; } = 1;
    public int ProgressValue { get; private set; }
    public string ProgressText { get; private set; } = string.Empty;
    public Visibility ProgressVisibility { get; private set; } =
        Visibility.Collapsed;

    internal void SetState(string status, string detail)
    {
        Status = status.Trim().ToUpperInvariant();
        var compactDetail = string.Join(
            " ",
            detail.Split(
                ['\r', '\n', '\t'],
                StringSplitOptions.RemoveEmptyEntries
                | StringSplitOptions.TrimEntries));
        Detail = compactDetail.Length <= 180
            ? compactDetail
            : compactDetail[..177] + "...";
        (StatusText, Background, BorderBrush, MarkerBrush, StatusBrush) =
            Status switch
            {
                "RUNNING" => (
                    "진행 중",
                    RunningBackground,
                    RunningBrush,
                    RunningBrush,
                    RunningBrush),
                "COMPLETED" => (
                    "완료",
                    CompletedBackground,
                    CompletedBrush,
                    CompletedBrush,
                    CompletedBrush),
                "FAILED" => (
                    "실패",
                    FailedBackground,
                    FailedBrush,
                    FailedBrush,
                    FailedBrush),
                "WAITING" => (
                    "확인 대기",
                    RunningBackground,
                    RunningBrush,
                    RunningBrush,
                    RunningBrush),
                _ => (
                    "대기",
                    PendingBackground,
                    PendingBorder,
                    PendingBrush,
                    PendingBrush),
            };
        PropertyChanged?.Invoke(
            this,
            new PropertyChangedEventArgs(string.Empty));
    }

    internal void SetProgress(
        int completed,
        int failed,
        int total)
    {
        var normalizedTotal = Math.Max(0, total);
        var normalizedCompleted = Math.Max(0, completed);
        var normalizedFailed = Math.Max(0, failed);
        ProgressMaximum = Math.Max(1, normalizedTotal);
        ProgressValue = Math.Min(
            ProgressMaximum,
            normalizedCompleted + normalizedFailed);
        ProgressText = normalizedTotal > 0
            ? $"{ProgressValue:N0}/{normalizedTotal:N0}"
              + (
                  normalizedFailed > 0
                      ? $" · 실패 {normalizedFailed:N0}"
                      : string.Empty
              )
            : string.Empty;
        ProgressVisibility = normalizedTotal > 1
            ? Visibility.Visible
            : Visibility.Collapsed;
        PropertyChanged?.Invoke(
            this,
            new PropertyChangedEventArgs(string.Empty));
    }

    private static Brush CreateBrush(
        byte red,
        byte green,
        byte blue)
    {
        var brush = new SolidColorBrush(Color.FromRgb(red, green, blue));
        brush.Freeze();
        return brush;
    }
}
