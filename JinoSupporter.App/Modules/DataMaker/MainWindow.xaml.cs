using DataMaker.Logger;
using DataMaker.R6;
using DataMaker.R6.FetchDataBMES;
using DataMaker.R6.Grouping;
using DataMaker.R6.PreProcessor;
using DataMaker.R6.Report;
using DataMaker.R6.SaveClass;
using DataMaker.R6.SQLProcess;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using WorkbenchHost.Infrastructure;

namespace DataMaker
{
    public partial class MainWindow : Window
    {
        public event Action? WebModuleSnapshotChanged;

        private const int SummaryBlankLeadingColumns = 2;
        private const int SummaryLeadingColumnsToSkip = 4;
        private const string SummarySeparatorDayWeek = "__SUMMARY_SEPARATOR_DAY_WEEK__";
        private const string SummarySeparatorWeekMonth = "__SUMMARY_SEPARATOR_WEEK_MONTH__";
        private const int LargeScanProgressLogInterval = 100_000;
        private const string OriginalTableMetaTableName = "__DataMakerMeta";
        private const string OriginalTableUpdatedAtKey = "OriginalTableUpdatedAt";
        private static readonly HashSet<string> SummaryBlankHeaderValues = new(StringComparer.OrdinalIgnoreCase)
        {
            "BLANK_AFTER_PROCESS",
            "BLANK_1",
            "BLANK_2"
        };

        private sealed class ReportExportResult
        {
            public required List<clModelGroupData> GroupsData { get; init; }
            public required string BaseName { get; init; }
            public required DataTable DetailReport { get; init; }
            public required DataTable WorstReport { get; init; }
            public required DataTable WorstReasonReport { get; init; }
            public required DataTable WorstProcessReport { get; init; }
            public DataTable? FirstGroupData { get; init; }
        }

        private sealed class ReportOrderConfig
        {
            public List<string> JsonFiles { get; set; } = new();
            public string OutputPath { get; set; } = string.Empty;
        }

        private sealed class LoadedGroupDisplay
        {
            public required string GroupName { get; init; }
            public required List<string> SubGroups { get; init; }
            public int Count => SubGroups.Count;
            public string DisplayName => $"{GroupName} ({Count})";
        }

        private sealed class MappingScanResult
        {
            public required IReadOnlyList<IReadOnlyDictionary<string, string>> MissingRoutingItems { get; init; }
            public required IReadOnlyList<IReadOnlyDictionary<string, string>> MissingReasonItems { get; init; }
            public int TotalOriginalRows { get; init; }
            public int RoutingTableRows { get; init; }
            public int ReasonTableRows { get; init; }
            public int AppliedRoutingRows { get; init; }
            public int MissingRoutingRows { get; init; }
            public int AppliedReasonRows { get; init; }
            public int MissingReasonRows { get; init; }
        }

        private sealed class AiTrainingRecord
        {
            public required string ModelName { get; init; }
            public required string CodeName { get; init; }
            public required string ProcessName { get; init; }
            public required string Label { get; init; }
        }

        private sealed class AiInferenceRecord
        {
            public required string ModelName { get; init; }
            public required string CodeName { get; init; }
            public required string ProcessName { get; init; }
            public required string NormalizedModelName { get; init; }
            public required string NormalizedCodeName { get; init; }
            public required string NormalizedProcessName { get; init; }
        }

        private sealed class AiReasonTrainingRecord
        {
            public required string ProcessName { get; init; }
            public required string NgName { get; init; }
            public required string Reason { get; init; }
        }

        private sealed class AiReasonInferenceRecord
        {
            public required string ProcessName { get; init; }
            public required string NgName { get; init; }
            public required string NormalizedProcessName { get; init; }
            public required string NormalizedNgName { get; init; }
        }

        private sealed class AiReasonPredictionResult
        {
            public required string ProcessName { get; init; }
            public required string NgName { get; init; }
            public required string PredictedReason { get; init; }
            public required double Confidence { get; init; }
            public required string Source { get; init; }
        }

        private sealed class AiPredictionResult
        {
            public required string ModelName { get; init; }
            public required string CodeName { get; init; }
            public required string ProcessName { get; init; }
            public required string PredictedLabel { get; init; }
            public required double Confidence { get; init; }
            public required string Source { get; init; }
        }

        private sealed class AiModelPayload
        {
            public DateTime TrainedAtUtc { get; set; }
            public int TrainingRowCount { get; set; }
            public Dictionary<string, int> LabelCounts { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, Dictionary<string, int>> TokenLabelCounts { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> KeywordDictionary { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> ExactLookup { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> NormalizedLookup { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        }

        private List<clModelGroupData> ListModelGroupData;
        List<string> ListUniqueModelName;

        private string _processTypeCsvPath = @"D:\000. MyWorks\003. BMES\Routing.txt";
        private string _reasonCsvPath = @"D:\000. MyWorks\003. BMES\reason.txt";
        private string PathSelectedDB = string.Empty;
        private readonly Dictionary<string, string> _debugGroupDisplayToTable = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private string _cachedDebugGroupTableName = string.Empty;
        private DataTable? _cachedDebugGroupTable;
        private bool _isUpdatingDebugCombo = false;
        private readonly List<string> _pendingReportJsonFiles = new();
        private string _pendingReportOutputPath = string.Empty;
        private string _currentProgressTitle = "Processing";
        private string _currentProgressMessage = "Please wait...";
        private int _currentTaskProgressPercent;
        private int _currentTotalProgressPercent;
        private readonly List<LoadedGroupDisplay> _loadedGroupDisplays = new();
        private readonly List<string> _loadedGroupSourceJsonFiles = new();
        private MappingTableKind _currentMappingTableKind = MappingTableKind.None;
        private const int MappingPreviewRowLimit = 300;
        private int _mappingTotalRowCount;
        private DataTable? _currentMappingTableData;
        private IReadOnlyList<IReadOnlyDictionary<string, string>> _cachedMissingRoutingItems = Array.Empty<IReadOnlyDictionary<string, string>>();
        private IReadOnlyList<IReadOnlyDictionary<string, string>> _cachedMissingReasonItems = Array.Empty<IReadOnlyDictionary<string, string>>();
        private bool _hasScannedMissingRoutingItems;
        private bool _hasScannedMissingReasonItems;
        private MappingScanResult? _mappingScanResult;
        private int _lastLoadDbLogBucket = -1;
        private string _lastLoadDbLogMessage = string.Empty;
        private string _cachedSummaryDbPath = string.Empty;
        private DateTime _cachedSummaryLastWriteTimeUtc = DateTime.MinValue;
        private string _cachedOriginalTableDateRange = string.Empty;
        private string _cachedDbLastUpdatedAt = string.Empty;
        private string _aiTrainingFilePath = string.Empty;
        private string _aiInferenceFilePath = string.Empty;
        private string _aiOutputFilePath = string.Empty;
        private string _aiFeedbackFilePath = string.Empty;
        private string _aiModelZipPath = string.Empty;
        private string _aiStatusMessage = "Select a training file to begin.";
        private int _aiTrainingRowCount;
        private int _aiInferenceRowCount;
        private int _aiOutputRowCount;
        private int _aiFeedbackRowCount;
        private DateTime? _aiLastTrainedAtUtc;
        private Dictionary<string, int> _aiLabelCounts = new(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, Dictionary<string, int>> _aiTokenLabelCounts = new(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> _aiKeywordDictionary = new(StringComparer.OrdinalIgnoreCase)
        {
            ["VIS"] = "VISUAL",
            ["VSL"] = "VISUAL",
            ["FUNC"] = "FUNCTION"
        };
        private Dictionary<string, string> _aiExactLookup = new(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> _aiNormalizedLookup = new(StringComparer.OrdinalIgnoreCase);
        private List<AiPredictionResult> _aiPredictionResults = new();
        private Window? _aiInferencePreviewWindow;
        private string _aiReasonTrainingFilePath = string.Empty;
        private string _aiReasonOutputFilePath = string.Empty;
        private string _aiReasonFeedbackFilePath = string.Empty;
        private string _aiReasonModelZipPath = string.Empty;
        private string _aiReasonStatusMessage = "Select a reason training file to begin.";
        private int _aiReasonTrainingRowCount;
        private int _aiReasonInferenceRowCount;
        private int _aiReasonOutputRowCount;
        private int _aiReasonFeedbackRowCount;
        private DateTime? _aiReasonLastTrainedAtUtc;
        private Dictionary<string, int> _aiReasonLabelCounts = new(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, Dictionary<string, int>> _aiReasonTokenLabelCounts = new(StringComparer.OrdinalIgnoreCase);
        private List<AiReasonPredictionResult> _aiReasonPredictionResults = new();
        private Dictionary<string, string> _aiReasonExactLookup = new(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> _aiReasonNormalizedLookup = new(StringComparer.OrdinalIgnoreCase);
        private Window? _aiReasonInferencePreviewWindow;

        private enum MappingTableKind
        {
            None,
            Routing,
            Reason
        }

        string[] keywords = new[] { "A2", "B2", "C2", "D2", "E2" };

        public MainWindow()
        {
            InitializeComponent();

            FormSettingBMESWindow.EnsureLoadedInfo();

            // DatePicker defaults
            CT_DATE_START.SelectedDate = DateTime.Today;
            CT_DATE_END.SelectedDate = DateTime.Today;
            CT_CB_DEBUG_DATE.Text = DateTime.Today.ToString("yyyy-MM-dd");

            // 로그 설정
            clLogger.WpfTextBox = CT_TB_LOG;
            clLogger.InitializeFileLogging();

            string logPath = clLogger.GetLogFilePath();
            if (!string.IsNullOrEmpty(logPath))
            {
                clLogger.LogImportant($"Log file: {logPath}");
            }

            SetReportAvailability(false);
            ShowEmptyHome();
            ResetLoadMappingStatus();
            TryAutoLoadAiModel();
            TryAutoLoadAiReasonModel();
        }

        private void ShowEmptyHome()
        {
            CT_EMPTY_HOME.Visibility = Visibility.Visible;
            CT_LOAD_CONTENT.Visibility = Visibility.Collapsed;
            CT_REPORT_SETUP_CONTENT.Visibility = Visibility.Collapsed;
            CT_MAIN_CONTENT_SCROLL.Visibility = Visibility.Collapsed;
            CT_MAPPING_CONTENT.Visibility = Visibility.Collapsed;
            SetMainContentSectionsVisibility(showMappingOnly: false);
        }

        private void ShowMainContent()
        {
            CT_EMPTY_HOME.Visibility = Visibility.Collapsed;
            CT_LOAD_CONTENT.Visibility = Visibility.Collapsed;
            CT_REPORT_SETUP_CONTENT.Visibility = Visibility.Collapsed;
            CT_MAIN_CONTENT_SCROLL.Visibility = Visibility.Visible;
            CT_MAPPING_CONTENT.Visibility = Visibility.Collapsed;
            SetMainContentSectionsVisibility(showMappingOnly: false);
        }

        private void ShowLoadContent()
        {
            CT_EMPTY_HOME.Visibility = Visibility.Collapsed;
            CT_LOAD_CONTENT.Visibility = Visibility.Visible;
            CT_REPORT_SETUP_CONTENT.Visibility = Visibility.Collapsed;
            CT_MAIN_CONTENT_SCROLL.Visibility = Visibility.Collapsed;
            CT_MAPPING_CONTENT.Visibility = Visibility.Collapsed;
            SetMainContentSectionsVisibility(showMappingOnly: false);
        }

        private void ShowReportSetupContent()
        {
            CT_EMPTY_HOME.Visibility = Visibility.Collapsed;
            CT_LOAD_CONTENT.Visibility = Visibility.Collapsed;
            CT_REPORT_SETUP_CONTENT.Visibility = Visibility.Visible;
            CT_MAIN_CONTENT_SCROLL.Visibility = Visibility.Collapsed;
            CT_MAPPING_CONTENT.Visibility = Visibility.Collapsed;
            SetMainContentSectionsVisibility(showMappingOnly: false);
        }

        private void ShowMappingContent()
        {
            CT_EMPTY_HOME.Visibility = Visibility.Collapsed;
            CT_LOAD_CONTENT.Visibility = Visibility.Collapsed;
            CT_REPORT_SETUP_CONTENT.Visibility = Visibility.Collapsed;
            CT_MAIN_CONTENT_SCROLL.Visibility = Visibility.Visible;
            CT_MAPPING_CONTENT.Visibility = Visibility.Visible;
            SetMainContentSectionsVisibility(showMappingOnly: true);
        }

        private void SetMainContentSectionsVisibility(bool showMappingOnly)
        {
            foreach (UIElement child in CT_MAIN_CONTENT.Children)
            {
                if (ReferenceEquals(child, CT_MAPPING_CONTENT))
                {
                    child.Visibility = showMappingOnly ? Visibility.Visible : Visibility.Collapsed;
                    continue;
                }

                child.Visibility = showMappingOnly ? Visibility.Collapsed : Visibility.Visible;
            }
        }

        private void CT_BT_SHOW_GETDATA_Click(object sender, RoutedEventArgs e)
        {
            ShowMainContent();
            CT_MAPPING_CONTENT.Visibility = Visibility.Collapsed;
        }

        private void CT_BT_CLEAR_LOG_Click(object sender, RoutedEventArgs e)
        {
            CT_TB_LOG.Document.Blocks.Clear();
        }

        private void SetReportAvailability(bool isEnabled)
        {
            CT_BT_GETREPORT.IsEnabled = isEnabled;
            CT_BT_GETREPORT.Content = "Get Report";
        }

        private void PopulateReportOrderUi()
        {
            int selectedIndex = CT_LB_REPORT_JSON_ORDER.SelectedIndex;

            CT_TB_REPORT_OUTPUT_PATH.Text = string.IsNullOrWhiteSpace(_pendingReportOutputPath)
                ? "Selected when Start is pressed"
                : $"{_pendingReportOutputPath} (can change on Start)";
            CT_LB_REPORT_JSON_ORDER.ItemsSource = null;
            CT_LB_REPORT_JSON_ORDER.ItemsSource = _pendingReportJsonFiles
                .Select(Path.GetFileName)
                .ToList();

            if (_pendingReportJsonFiles.Count == 0)
            {
                CT_LB_REPORT_JSON_ORDER.SelectedIndex = -1;
                return;
            }

            if (selectedIndex < 0)
            {
                selectedIndex = 0;
            }

            CT_LB_REPORT_JSON_ORDER.SelectedIndex = Math.Min(selectedIndex, _pendingReportJsonFiles.Count - 1);
        }

        private void ClearPendingReportSetup()
        {
            _pendingReportJsonFiles.Clear();
            _pendingReportOutputPath = string.Empty;
            CT_TB_REPORT_OUTPUT_PATH.Text = "Selected when Start is pressed";
            CT_LB_REPORT_JSON_ORDER.ItemsSource = null;
        }

        private void SetReportSetupButtonsEnabled(bool isEnabled)
        {
            CT_BT_REPORT_MOVE_UP.IsEnabled = isEnabled;
            CT_BT_REPORT_MOVE_DOWN.IsEnabled = isEnabled;
            CT_BT_REPORT_LOAD_SAVED.IsEnabled = isEnabled;
            CT_BT_REPORT_SAVE.IsEnabled = isEnabled;
            CT_BT_REPORT_CANCEL.IsEnabled = isEnabled;
            CT_BT_REPORT_START.IsEnabled = isEnabled;
        }

        private async Task LoadProcess()
        {
            ShowLoadContent();

            var openFileDialog = new OpenFileDialog
            {
                Filter = "db Files|*.db",
                Title = "Open BMES Data"
            };
            if (ShowCommonDialog(openFileDialog))
            {
                SetReportAvailability(false);
                PathSelectedDB = openFileDialog.FileName;
                InvalidateDatabaseSummaryCache();
                CT_TB_DB_PATH.Text = PathSelectedDB;
                InvalidateDebugGroupTableCache();
                RefreshDebugGroupDropdown();

                try
                {
                    CT_BT_LOAD.IsEnabled = false;
                    CT_TB_LOG.Document.Blocks.Clear();
                    ResetLoadDbProgressLogging();
                    SetProgressPopupState("Load DB", "Preparing load...");
                    UpdateLoadDbProgress(0, 0, "Preparing selected DB...");
                    clLogger.LogImportant("Started DB load and processing.");

                    UpdateLoadDbProgress(10, 10, "Processing DB and preparing OriginalTable...");
                    var dbProgress = new Progress<(int Percent, string Message)>(state =>
                    {
                        UpdateLoadDbProgress(state.Percent, state.Percent, state.Message);
                    });
                    await Task.Run(() => RunDataProcessing(PathSelectedDB, _processTypeCsvPath, _reasonCsvPath, dbProgress));

                    UpdateLoadDbProgress(70, 70, "Refreshing model list...");
                    await GetUniqueModelsAsync();
                    ResetLoadedGroupJsonUi();

                    UpdateLoadDbProgress(85, 85, "Calculating mapping status...");
                    await RefreshLoadMappingStatusAsync();

                    UpdateLoadDbProgress(100, 100, "Load DB completed.");

                    SetReportAvailability(true);
                    clLogger.LogImportant("DB load and OriginalTable preparation completed successfully.");
                }
                catch (Exception ex)
                {
                    SetReportAvailability(false);
                    UpdateProgressBars(0, 0);
                    ShowErrorMessage($"Error: {ex.Message}", "Error");
                    clLogger.LogException(ex, "UI: Error in button click handler");
                }
                finally
                {
                    CT_BT_LOAD.IsEnabled = true;
                }
            }
        }

        private async void CT_BT_BROWSE_DB_Click(object sender, RoutedEventArgs e)
        {
            await LoadProcess();
        }

        private void ResetLoadedGroupJsonUi()
        {
            _loadedGroupDisplays.Clear();
            _loadedGroupSourceJsonFiles.Clear();
            CT_LB_GROUP_JSON_PATH.Text = "Load a group JSON file after DB load.";
            CT_LB_GROUPS.ItemsSource = null;
            CT_LB_SUBGROUPS.ItemsSource = null;
            CT_LB_SUBGROUP_HEADER.Text = "Sub Groups";
        }

        private void PopulateLoadedGroupJsonUi(IEnumerable<clModelGroupData> groups, string jsonPath)
        {
            _loadedGroupDisplays.Clear();

            foreach (var group in groups)
            {
                string groupName = string.IsNullOrWhiteSpace(group.GroupName)
                    ? string.Join(", ", group.ModelList?.Take(3) ?? Enumerable.Empty<string>())
                    : group.GroupName.Trim();

                if (string.IsNullOrWhiteSpace(groupName))
                {
                    groupName = "Unnamed Group";
                }

                var subGroups = (group.ModelList ?? new List<string>())
                    .Where(model => !string.IsNullOrWhiteSpace(model))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(model => model)
                    .ToList();

                _loadedGroupDisplays.Add(new LoadedGroupDisplay
                {
                    GroupName = groupName,
                    SubGroups = subGroups
                });
            }

            CT_LB_GROUP_JSON_PATH.Text = jsonPath;
            CT_LB_GROUPS.ItemsSource = _loadedGroupDisplays;
            CT_LB_GROUPS.DisplayMemberPath = nameof(LoadedGroupDisplay.DisplayName);

            if (_loadedGroupDisplays.Count == 0)
            {
                CT_LB_SUBGROUPS.ItemsSource = null;
                CT_LB_SUBGROUP_HEADER.Text = "Sub Groups";
                return;
            }

            CT_LB_GROUPS.SelectedIndex = 0;
        }

        private void CT_BT_LOAD_GROUP_JSON_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(PathSelectedDB) || !File.Exists(PathSelectedDB))
            {
                ShowWarningMessage("Load DB first.", "Warning");
                return;
            }

            var dlg = new OpenFileDialog
            {
                Filter = "JSON Files|*.json",
                Title = "Load Model Groups"
            };

            if (!ShowCommonDialog(dlg))
            {
                return;
            }

            try
            {
                ResetLoadedGroupJsonUi();
                var sourceJsonFiles = ResolveReportJsonInputFiles(new[] { dlg.FileName });
                var loadedGroups = LoadFromJson(dlg.FileName);
                _loadedGroupSourceJsonFiles.AddRange(sourceJsonFiles);
                ListModelGroupData = loadedGroups;
                PopulateLoadedGroupJsonUi(loadedGroups, dlg.FileName);
                clLogger.Log($"Loaded group JSON into main view: {dlg.FileName}");
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "CT_BT_LOAD_GROUP_JSON_Click");
                ShowErrorMessage(BuildJsonLoadErrorMessage(ex), "Error");
            }
        }

        private void GroupJsonList_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[]? files = e.Data.GetData(DataFormats.FileDrop) as string[];
                bool hasJson = files?.Any(path => string.Equals(Path.GetExtension(path), ".json", StringComparison.OrdinalIgnoreCase)) == true;
                e.Effects = hasJson ? DragDropEffects.Copy : DragDropEffects.None;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = true;
        }

        private void GroupJsonList_Drop(object sender, DragEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(PathSelectedDB) || !File.Exists(PathSelectedDB))
            {
                MessageBox.Show("Load DB first.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            string[]? files = e.Data.GetData(DataFormats.FileDrop) as string[];
            string? jsonFile = files?
                .FirstOrDefault(path => string.Equals(Path.GetExtension(path), ".json", StringComparison.OrdinalIgnoreCase));

            if (string.IsNullOrWhiteSpace(jsonFile))
            {
                return;
            }

            try
            {
                ResetLoadedGroupJsonUi();
                var sourceJsonFiles = ResolveReportJsonInputFiles(new[] { jsonFile });
                var loadedGroups = LoadFromJson(jsonFile);
                _loadedGroupSourceJsonFiles.AddRange(sourceJsonFiles);
                ListModelGroupData = loadedGroups;
                PopulateLoadedGroupJsonUi(loadedGroups, jsonFile);
                clLogger.Log($"Loaded group JSON by drag-and-drop: {jsonFile}");
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "GroupJsonList_Drop");
                MessageBox.Show(BuildJsonLoadErrorMessage(ex), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CT_LB_GROUPS_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CT_LB_GROUPS.SelectedItem is not LoadedGroupDisplay selectedGroup)
            {
                CT_LB_SUBGROUPS.ItemsSource = null;
                CT_LB_SUBGROUP_HEADER.Text = "Sub Groups";
                return;
            }

            CT_LB_SUBGROUP_HEADER.Text = $"{selectedGroup.GroupName} Sub Groups";
            CT_LB_SUBGROUPS.ItemsSource = selectedGroup.SubGroups;
        }

        private void CT_BT_BMESSETTING_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new FormSettingBMESWindow();
            ShowModalWindow(dlg, sender);
        }

        private async void CT_BT_GETBMESDATA_Click(object sender, RoutedEventArgs e)
        {
            var result = ShowQuestionMessage(
                "기존 DB에 데이터를 추가하시겠습니까?\n\n" +
                "YES: 기존 DB 선택 후 데이터 추가\n" +
                "NO: 새 DB 파일 생성",
                "Add to Existing DB?",
                MessageBoxButton.YesNoCancel
            );

            if (result == MessageBoxResult.Cancel)
                return;

            bool isAdd = (result == MessageBoxResult.Yes);

            SetReportAvailability(false);
            CT_BT_GETBMESDATA.IsEnabled = false;
            CT_BT_GETBMESDATA.Content = "Fetching...";
            UpdateProgressBars(0, 0);

            try
            {
                clFetchBMESNGDATA BMES = new clFetchBMESNGDATA();
                var credentials = FormSettingBMESWindow.EnsureLoadedInfo();
                BMESOption option = new BMESOption
                {
                    StartDate = CT_DATE_START.SelectedDate ?? DateTime.Today,
                    EndDate = CT_DATE_END.SelectedDate ?? DateTime.Today,
                    InfoId = credentials
                };

                var totalStopwatch = Stopwatch.StartNew();

                // Step 1: fetch from BMES web.
                UpdateProgressBars(15, 10);
                var fetchStopwatch = Stopwatch.StartNew();
                await BMES.GetDataFromWebAsync(option);
                fetchStopwatch.Stop();
                clLogger.LogImportant($"BMES fetch completed in {fetchStopwatch.Elapsed.TotalSeconds:F2}s");
                UpdateProgressBars(55, 45);

                // Step 2: save/merge DB.
                var saveStopwatch = Stopwatch.StartNew();
                string dbPath = await BMES.SaveBMESDataToDB(isAdd);
                saveStopwatch.Stop();
                clLogger.LogImportant($"BMES save/merge completed in {saveStopwatch.Elapsed.TotalSeconds:F2}s");
                UpdateProgressBars(75, 65);

                if (!string.IsNullOrEmpty(dbPath))
                {
                    try
                    {
                        clLogger.Log($"Auto-loading DB: {dbPath}");
                        PathSelectedDB = dbPath;
                        InvalidateDatabaseSummaryCache();
                        InvalidateDebugGroupTableCache();
                        RefreshDebugGroupDropdown();

                        // Step 3: post-processing (OriginalTable preparation).
                        UpdateProgressBars(85, 80);
                        var processingStopwatch = Stopwatch.StartNew();
                        await Task.Run(() => RunDataProcessing(PathSelectedDB, _processTypeCsvPath, _reasonCsvPath));
                        processingStopwatch.Stop();
                        clLogger.LogImportant($"Post-processing completed in {processingStopwatch.Elapsed.TotalSeconds:F2}s");
                        SetProgressPopupState("Load DB", "Refreshing model list...");
                        var modelRefreshStopwatch = Stopwatch.StartNew();
                        await GetUniqueModelsAsync();
                        modelRefreshStopwatch.Stop();
                        clLogger.LogImportant($"Model list refresh completed in {modelRefreshStopwatch.Elapsed.TotalSeconds:F2}s");
                        ResetLoadedGroupJsonUi();
                        UpdateProgressBars(100, 100);
                        totalStopwatch.Stop();

                        SetReportAvailability(true);
                        clLogger.LogImportant($"BMES Get Data flow completed in {totalStopwatch.Elapsed.TotalSeconds:F2}s");
                        clLogger.Log("OriginalTable preparation completed after BMES Fetch");
                        ShowInfoMessage(
                            "BMES Data saved and processed successfully!\n\n" +
                            $"Database: {Path.GetFileName(dbPath)}\n" +
                            "OriginalTable is ready for model filtering.",
                            "Success"
                        );

                        MakeGroupUI();
                    }
                    catch (Exception ex)
                    {
                        SetReportAvailability(false);
                        clLogger.LogException(ex, "Error during auto-load after BMES Fetch");
                        ShowWarningMessage(
                            $"BMES Data saved but error during processing:\n{ex.Message}\n\n" +
                            "Please use LOAD button to manually load the database.",
                            "Warning"
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                SetReportAvailability(false);
                UpdateProgressBars(0, 0);
                clLogger.LogException(ex, "Error during BMES Get Data");
                ShowErrorMessage($"Error during Get Data: {ex.Message}", "Error");
            }
            finally
            {
                CT_BT_GETBMESDATA.IsEnabled = true;
                CT_BT_GETBMESDATA.Content = "Get Data";
            }
        }

        private void RemoveOrgTableFromNewTable(DataTable org, DataTable add)
        {
            if (org == null || add == null) return;
            if (org.Columns.Count < 6 || add.Columns.Count < 6)
                throw new ArgumentException("열 개수가 6개 이상이어야 합니다.");

            var keySet = new HashSet<string>();
            foreach (DataRow row in add.Rows)
            {
                string key = $"{row[CONSTANT.PRODUCTION_LINE.NEW]}|{row[CONSTANT.PROCESSNAME.NEW]}|{row[CONSTANT.NGNAME.NEW]}|{row[CONSTANT.PRODUCT_DATE.NEW]}|{row[CONSTANT.MATERIALNAME.NEW]}";
                keySet.Add(key);
            }

            int n = 0;
            for (int i = org.Rows.Count - 1; i >= 0; i--)
            {
                string key = $"{org.Rows[i][CONSTANT.PRODUCTION_LINE.NEW]}|{org.Rows[i][CONSTANT.PROCESSNAME.NEW]}|{org.Rows[i][CONSTANT.NGNAME.NEW]}|{org.Rows[i][CONSTANT.PRODUCT_DATE.NEW]}|{org.Rows[i][CONSTANT.MATERIALNAME.NEW]}";
                if (keySet.Contains(key))
                {
                    n++;
                    org.Rows.RemoveAt(i);
                }
            }

            org.Merge(add);
            clLogger.Log("Removed " + n + " rows from Org table that exist in the new data.");
        }

        private async void CT_BT_LOAD_Click(object sender, RoutedEventArgs e)
        {
            await LoadProcess();
        }

        private void CT_BT_GETREPORT_Click(object sender, RoutedEventArgs e)
        {
            _ = StartGetReportAsync();
        }

        private async Task StartGetReportAsync()
        {
            if (_loadedGroupSourceJsonFiles.Count > 0)
            {
                _pendingReportJsonFiles.Clear();
                _pendingReportJsonFiles.AddRange(_loadedGroupSourceJsonFiles);
                _pendingReportOutputPath = string.Empty;
            }
            else if (!TrySelectReportInputJsonFiles())
            {
                return;
            }

            if (_pendingReportJsonFiles.Count == 0)
            {
                ShowInfoMessage("No valid report JSON files were found.", "Info");
                return;
            }

            PopulateReportOrderUi();

            var folderDialog = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "Select folder to save the report CSV files",
                FolderName = _pendingReportOutputPath
            };

            if (!ShowCommonDialog(folderDialog))
            {
                return;
            }

            _pendingReportOutputPath = folderDialog.FolderName;
            CT_BT_GETREPORT.IsEnabled = false;
            CT_BT_GETREPORT.Content = "Processing...";
            UpdateProgressBars(0, 0);

            try
            {
                await GetReport(_pendingReportJsonFiles.ToArray(), _pendingReportOutputPath);
                ClearPendingReportSetup();
                ShowEmptyHome();
            }
            finally
            {
                CT_BT_GETREPORT.IsEnabled = true;
                CT_BT_GETREPORT.Content = "Get Report";
            }
        }

        private async Task GetReport(string[] jsonFiles, string outputPath)
        {
            try
            {
                ShowProgressPopup("Get Report", "Preparing report generation...");
                clLogger.Log($"=== Starting Report Generation for {jsonFiles.Length} JSON file(s) ===");

                var allGroupsData = new List<clModelGroupData>();
                var reportResults = new List<ReportExportResult>();

                for (int i = 0; i < jsonFiles.Length; i++)
                {
                    string jsonPath = jsonFiles[i];
                    string baseName = Path.GetFileNameWithoutExtension(jsonPath);
                    SetProgressPopupState("Get Report", $"Processing {baseName} ({i + 1}/{jsonFiles.Length})");

                    clLogger.Log($"");
                    clLogger.Log($"========================================");
                    clLogger.Log($"Processing [{i + 1}/{jsonFiles.Length}]: {baseName}");
                    clLogger.Log($"========================================");

                    try
                    {
                        var reportResult = await ProcessSingleJsonFile(jsonPath, i + 1, jsonFiles.Length);
                        reportResults.Add(reportResult);

                        if (reportResult.GroupsData.Count > 0 && reportResult.FirstGroupData != null)
                        {
                            var group = reportResult.GroupsData[0];
                            group.Index = allGroupsData.Count;
                            allGroupsData.Add(group);
                        }
                    }
                    catch (Exception ex)
                    {
                        clLogger.LogError($"Error processing {baseName}: {ex.Message} -> SKIP");
                        clLogger.LogImportant($"SKIP: {baseName}");
                    }
                }

                if (allGroupsData.Count > 0)
                {
                    SetProgressPopupState("Get Report", "Creating combined AllGroupsDetail report...");
                    clLogger.Log($"");
                    clLogger.Log($"========================================");
                    clLogger.Log($"Creating combined AllGroupsDetailReport from {allGroupsData.Count} JSON files (index 0 groups only)");
                    clLogger.Log($"========================================");

                    var allGroupsPreloadedData = BuildAllGroupsPreloadedData(reportResults, allGroupsData);
                    clLogger.Log($"Generating AllGroupsDetail report using preloaded first-group data...");
                    clAllGroupsDetailReportMaker allGroupsDetailReportMaker = new clAllGroupsDetailReportMaker(
                        PathSelectedDB,
                        allGroupsData,
                        allGroupsPreloadedData);
                    var allGroupsDetailReport = await allGroupsDetailReportMaker.CreateReport();

                    clSaveDataCSV csvSave = new clSaveDataCSV();
                    string allGroupsDetailPath = GetAvailableOutputFilePath(Path.Combine(outputPath, "AllGroupsDetail.csv"));
                    await csvSave.Save(allGroupsDetailReport, allGroupsDetailPath);
                    clLogger.Log($"Saved {Path.GetFileName(allGroupsDetailPath)} with {allGroupsData.Count} groups (index 0 from each JSON)");
                }

                await SaveSummaryReportAsync(reportResults, outputPath);
                await SaveReportsByModelAsync(reportResults, outputPath);

                UpdateProgressBars(100, 100);
                clLogger.Log($"");
                clLogger.Log($"=== All {jsonFiles.Length} JSON file(s) processed successfully! ===");

                clLogger.Log($"");
                clLogger.Log($"Checking for unmapped items...");
                await Task.Run(() => SaveUnmappedItems(outputPath));

                int generatedCsvCount = reportResults.Count + 1 + (allGroupsData.Count > 0 ? 1 : 0);

                ShowInfoMessage(
                    $"모든 리포트 생성 완료!\n\n" +
                    $"처리된 JSON 파일: {jsonFiles.Length}개\n" +
                    $"생성된 CSV 파일: {generatedCsvCount}개\n" +
                    $"출력 폴더: {outputPath}",
                    "Complete"
                );
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "GetReport");
                ShowErrorMessage($"Error during report generation: {ex.Message}", "Error");
            }
            finally
            {
                CloseProgressPopup();
                CT_BT_GETREPORT.IsEnabled = true;
                CT_BT_GETREPORT.Content = "Get Report";
                UpdateProgressBars(0, 0);
            }
        }

        private async Task<ReportExportResult> ProcessSingleJsonFile(string jsonPath, int currentFileIndex, int totalFiles)
        {
            string baseName = Path.GetFileNameWithoutExtension(jsonPath);

            int fileStartProgress = (int)(((currentFileIndex - 1) / (double)totalFiles) * 100);
            int fileEndProgress = (int)((currentFileIndex / (double)totalFiles) * 100);

            clLogger.Log($"Loading JSON: {baseName}");
            ListModelGroupData = LoadModelGroupDataFromJson(jsonPath);

            UpdateProgressBars(0, fileStartProgress);
            SetProgressPopupState("Group Load", $"Creating ProcTable for {baseName}...");
            clLogger.Log($"Step 1/3: Creating ProcTable for {baseName}...");

            await Task.Run(() => GetProcTable(ListModelGroupData));

            UpdateProgressBars(100, fileStartProgress + (int)((fileEndProgress - fileStartProgress) * 0.25));
            clLogger.Log($"ProcTable created successfully");

            UpdateProgressBars(0, fileStartProgress + (int)((fileEndProgress - fileStartProgress) * 0.25));
            SetProgressPopupState("Group Load", $"Creating group tables for {baseName}...");
            clLogger.Log($"Step 2/3: Creating Group tables for {baseName}...");

            clGroupTableMaker groupTableMaker = new clGroupTableMaker(PathSelectedDB);
            await Task.Run(() => groupTableMaker.CreateGroupTables(ListModelGroupData));
            InvalidateDebugGroupTableCache();

            UpdateProgressBars(100, fileStartProgress + (int)((fileEndProgress - fileStartProgress) * 0.50));
            clLogger.Log($"Group tables created successfully");

            UpdateProgressBars(0, fileStartProgress + (int)((fileEndProgress - fileStartProgress) * 0.50));
            SetProgressPopupState("Get Report", $"Creating reports for {baseName}...");
            clLogger.Log($"Step 3/3: Creating Reports for {baseName}...");

            var preloadedGroupDataList = await Task.Run(() => LoadGroupDataForReport(
                ListModelGroupData,
                (taskProgress, message) =>
                {
                    int subProgress = fileStartProgress + (int)((fileEndProgress - fileStartProgress) * (0.50 + taskProgress / 100.0 * 0.08));
                    SetProgressPopupState("Get Report", message);
                    UpdateProgressBars(taskProgress, subProgress);
                }));

            clLogger.Log($"  Creating DetailReport...");
            clDetailReportMakerVer1 detailReportMaker = new clDetailReportMakerVer1(PathSelectedDB, ListModelGroupData, preloadedGroupDataList);
            detailReportMaker.ProgressChanged += (s, args) =>
            {
                int subProgress = fileStartProgress + (int)((fileEndProgress - fileStartProgress) * (0.50 + args.TotalProgress / 100.0 * 0.10));
                UpdateProgressBars(args.TaskProgress, subProgress);
            };
            var detailReport = await detailReportMaker.CreateReport();
            clLogger.Log($"  DetailReport completed");

            clLogger.Log($"  Creating WorstReport...");
            clWorstReportMakerVer2 worstReportMaker = new clWorstReportMakerVer2(PathSelectedDB, ListModelGroupData, preloadedGroupDataList, 0, CONSTANT.OPTION.TopNGCount);
            worstReportMaker.ProgressChanged += (s, args) =>
            {
                int subProgress = fileStartProgress + (int)((fileEndProgress - fileStartProgress) * (0.60 + args.TotalProgress / 100.0 * 0.10));
                UpdateProgressBars(args.TaskProgress, subProgress);
            };
            var worstReport = await worstReportMaker.CreateReport();
            clLogger.Log($"  WorstReport completed");

            clLogger.Log($"  Creating WorstReasonReport...");
            clWorstReasonReportMakerVer1 worstReasonReportMaker = new clWorstReasonReportMakerVer1(PathSelectedDB, ListModelGroupData, preloadedGroupDataList, 0, 20, CONSTANT.OPTION.TopDefectsPerReason);
            worstReasonReportMaker.ProgressChanged += (s, args) =>
            {
                int subProgress = fileStartProgress + (int)((fileEndProgress - fileStartProgress) * (0.70 + args.TotalProgress / 100.0 * 0.10));
                UpdateProgressBars(args.TaskProgress, subProgress);
            };
            var worstReasonReport = await worstReasonReportMaker.CreateReport();
            clLogger.Log($"  WorstReasonReport completed");

            clLogger.Log($"  Creating WorstProcessReport...");
            clWorstProcessReportMaker worstProcessReportMaker = new clWorstProcessReportMaker(PathSelectedDB, ListModelGroupData, preloadedGroupDataList, 0, CONSTANT.OPTION.TopProcessCount);
            worstProcessReportMaker.ProgressChanged += (s, args) =>
            {
                int subProgress = fileStartProgress + (int)((fileEndProgress - fileStartProgress) * (0.80 + args.TotalProgress / 100.0 * 0.10));
                UpdateProgressBars(args.TaskProgress, subProgress);
            };
            var worstProcessReport = await worstProcessReportMaker.CreateReport();
            clLogger.Log($"  WorstProcessReport completed");

            UpdateProgressBars(100, fileEndProgress);
            clLogger.Log($"All reports completed for {baseName}");

            clLogger.Log($"[{currentFileIndex}/{totalFiles}] {baseName} completed!");

            return new ReportExportResult
            {
                BaseName = baseName,
                GroupsData = ListModelGroupData,
                DetailReport = detailReport,
                WorstReport = worstReport,
                WorstReasonReport = worstReasonReport,
                WorstProcessReport = worstProcessReport,
                FirstGroupData = preloadedGroupDataList.FirstOrDefault().Data
            };
        }

        private static List<(int GroupIndex, DataTable Data)> BuildAllGroupsPreloadedData(
            List<ReportExportResult> reportResults,
            List<clModelGroupData> allGroupsData)
        {
            var preloadedData = new List<(int GroupIndex, DataTable Data)>();

            for (int i = 0; i < reportResults.Count && i < allGroupsData.Count; i++)
            {
                DataTable? firstGroupData = reportResults[i].FirstGroupData;
                if (firstGroupData == null)
                {
                    continue;
                }

                preloadedData.Add((allGroupsData[i].Index, firstGroupData));
            }

            return preloadedData;
        }

        private List<clModelGroupData> LoadModelGroupDataFromJson(string jsonPath)
        {
            string json = File.ReadAllText(jsonPath);
            var groupData = Newtonsoft.Json.JsonConvert.DeserializeObject<List<clModelGroupData>>(json);
            return groupData ?? new List<clModelGroupData>();
        }

        private List<(int GroupIndex, DataTable Data)> LoadGroupDataForReport(
            List<clModelGroupData> groupedModels,
            Action<int, string>? progressCallback = null)
        {
            var groupDataList = new List<(int GroupIndex, DataTable Data)>();

            using var sql = new clSQLFileIO(PathSelectedDB);

            for (int i = 0; i < groupedModels.Count; i++)
            {
                var group = groupedModels[i];
                string tableName = CONSTANT.GetGroupTableName(group);
                clLogger.Log($"Loading {tableName}...");

                var data = sql.LoadTable(tableName);
                if (data != null && data.Rows.Count > 0)
                {
                    groupDataList.Add((group.Index, data));
                    clLogger.Log($"  Loaded {data.Rows.Count} rows from {tableName}");
                }
                else
                {
                    clLogger.Log($"  WARNING: {tableName} is empty or not found");
                }

                int taskProgress = groupedModels.Count == 0
                    ? 100
                    : (int)Math.Round((i + 1) * 100d / groupedModels.Count);
                string groupName = string.IsNullOrWhiteSpace(group.GroupName)
                    ? tableName
                    : group.GroupName;
                progressCallback?.Invoke(taskProgress, $"Loading group {i + 1}/{groupedModels.Count} · {groupName}");
            }

            return groupDataList;
        }

        private async Task SaveCombinedReportSectionsAsync(List<(string Name, DataTable Table)> sections, string outputPath, bool includeColumnHeader = true)
        {
            if (sections.Count == 0)
            {
                return;
            }

            string finalPath = GetAvailableOutputFilePath(outputPath);

            await Task.Run(() =>
            {
                using var writer = new StreamWriter(finalPath, false, Encoding.UTF8);

                for (int i = 0; i < sections.Count; i++)
                {
                    WriteDataTableAsCsvSection(writer, sections[i].Table, includeColumnHeader);

                    if (i < sections.Count - 1)
                    {
                        writer.WriteLine();
                    }
                }
            });

            clLogger.Log($"Saved combined report: {Path.GetFileName(finalPath)} ({sections.Count} sections)");
        }

        private async Task SaveReportsByModelAsync(List<ReportExportResult> reportResults, string outputPath)
        {
            foreach (var report in reportResults)
            {
                string reportFilePath = GetAvailableOutputFilePath(Path.Combine(outputPath, $"{SanitizePathSegment(report.BaseName)}.csv"));

                await SaveSingleModelReportAsync(report, reportFilePath);
                clLogger.LogImportant($"Saved merged report for model: {Path.GetFileName(reportFilePath)}");
            }
        }

        private async Task SaveSummaryReportAsync(List<ReportExportResult> reportResults, string outputPath)
        {
            var summaryTables = new List<DataTable>();

            foreach (var report in reportResults)
            {
                var summaryTable = ExtractSummaryTable(report);
                if (summaryTable.Rows.Count > 0)
                {
                    summaryTables.Add(summaryTable);
                }
            }

            if (summaryTables.Count == 0)
            {
                return;
            }

            DataTable alignedSummary = BuildAlignedSummaryTable(summaryTables);
            string summaryPath = GetAvailableOutputFilePath(Path.Combine(outputPath, "Summary.csv"));

            await Task.Run(() =>
            {
                using var writer = new StreamWriter(summaryPath, false, Encoding.UTF8);
                WriteDataTableAsCsvSection(writer, alignedSummary, includeColumnHeader: false);
            });

            clLogger.LogImportant($"Saved {Path.GetFileName(summaryPath)} ({summaryTables.Count} sections)");
        }

        private async Task SaveSingleModelReportAsync(ReportExportResult report, string outputPath)
        {
            await Task.Run(() =>
            {
                using var writer = new StreamWriter(outputPath, false, Encoding.UTF8);

                WriteDataTableAsCsvSection(writer, report.WorstProcessReport, includeColumnHeader: true);
                writer.WriteLine();
                WriteDataTableAsCsvSection(writer, report.WorstReport, includeColumnHeader: true);
                writer.WriteLine();
                WriteDataTableAsCsvSection(writer, report.WorstReasonReport, includeColumnHeader: true);
                writer.WriteLine();
                WriteDataTableAsCsvSection(writer, report.DetailReport, includeColumnHeader: false);
            });
        }

        private static string GetAvailableOutputFilePath(string originalPath)
        {
            if (!File.Exists(originalPath))
            {
                return originalPath;
            }

            string directory = Path.GetDirectoryName(originalPath) ?? string.Empty;
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalPath);
            string extension = Path.GetExtension(originalPath);

            for (int counter = 1; counter <= 1000; counter++)
            {
                string candidateName = $"{fileNameWithoutExtension}-{counter}{extension}";
                string candidatePath = string.IsNullOrEmpty(directory)
                    ? candidateName
                    : Path.Combine(directory, candidateName);

                if (!File.Exists(candidatePath))
                {
                    return candidatePath;
                }
            }

            throw new IOException($"Unable to resolve an available output file name for '{originalPath}'.");
        }

        private static DataTable ExtractSummaryTable(ReportExportResult report)
        {
            var detail = report.DetailReport;
            var summary = new DataTable();

            if (detail.Columns.Count == 0 || detail.Rows.Count == 0 || report.GroupsData.Count == 0)
            {
                return summary;
            }

            string groupZeroName = report.GroupsData[0].GroupName;
            if (string.IsNullOrWhiteSpace(groupZeroName))
            {
                groupZeroName = string.Join("_", report.GroupsData[0].ModelList ?? new List<string>());
            }

            int groupCount = report.GroupsData.Count;
            int fixedColumnCount = Math.Min(5, detail.Columns.Count);

            summary.Columns.Add("Group0", typeof(object));

            for (int i = 0; i < fixedColumnCount; i++)
            {
                summary.Columns.Add(detail.Columns[i].ColumnName, typeof(object));
            }

            int currentIndex = fixedColumnCount;
            AppendFirstGroupPeriodColumns(detail, summary, ref currentIndex, groupCount);

            if (currentIndex < detail.Columns.Count && IsSeparatorColumn(detail.Columns[currentIndex].ColumnName))
            {
                summary.Columns.Add(SummarySeparatorDayWeek, typeof(object));
                currentIndex++;
            }

            AppendFirstGroupPeriodColumns(detail, summary, ref currentIndex, groupCount);

            if (currentIndex < detail.Columns.Count && IsSeparatorColumn(detail.Columns[currentIndex].ColumnName))
            {
                summary.Columns.Add(SummarySeparatorWeekMonth, typeof(object));
                currentIndex++;
            }

            AppendFirstGroupPeriodColumns(detail, summary, ref currentIndex, groupCount);

            int startRow = FindFirstRowIndex(detail, "NG PPM");
            if (startRow < 0)
            {
                return summary;
            }

            // Detail section keeps header rows above NG PPM.
            // Include the date row, but skip the extra row directly under it.
            if (startRow > 1)
            {
                startRow -= 2;
            }
            else if (startRow > 0)
            {
                startRow--;
            }

            int endRowExclusive = FindNextProcessNameHeaderRow(detail, startRow);
            if (endRowExclusive < 0)
            {
                endRowExclusive = detail.Rows.Count;
            }

            for (int rowIndex = startRow; rowIndex < endRowExclusive; rowIndex++)
            {
                if (rowIndex == startRow + 1)
                {
                    continue;
                }

                var sourceRow = detail.Rows[rowIndex];
                var newRow = summary.NewRow();
                newRow[0] = groupZeroName;

                for (int colIndex = 1; colIndex < summary.Columns.Count; colIndex++)
                {
                    string columnName = summary.Columns[colIndex].ColumnName;
                    newRow[colIndex] = detail.Columns.Contains(columnName)
                        ? NormalizeSummaryCellValue(sourceRow[columnName])
                        : DBNull.Value;
                }

                summary.Rows.Add(newRow);
            }

            return summary;
        }

        private static int FindFirstRowIndex(DataTable table, string processNameValue)
        {
            if (!table.Columns.Contains("ProcessName"))
            {
                return -1;
            }

            for (int i = 0; i < table.Rows.Count; i++)
            {
                string value = table.Rows[i]["ProcessName"]?.ToString()?.Trim() ?? string.Empty;
                if (string.Equals(value, processNameValue, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private static int FindNextProcessNameHeaderRow(DataTable table, int startIndex)
        {
            if (!table.Columns.Contains("ProcessName"))
            {
                return -1;
            }

            for (int i = startIndex + 1; i < table.Rows.Count; i++)
            {
                string value = table.Rows[i]["ProcessName"]?.ToString()?.Trim() ?? string.Empty;
                if (string.Equals(value, "Process Name", StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private static void AppendFirstGroupPeriodColumns(DataTable source, DataTable destination, ref int currentIndex, int groupCount)
        {
            while (currentIndex < source.Columns.Count && !IsSeparatorColumn(source.Columns[currentIndex].ColumnName))
            {
                destination.Columns.Add(source.Columns[currentIndex].ColumnName, typeof(object));
                currentIndex += Math.Max(groupCount, 1);
            }
        }

        private static bool IsSeparatorColumn(string columnName)
        {
            return string.IsNullOrWhiteSpace(columnName);
        }

        private static DataTable BuildAlignedSummaryTable(List<DataTable> summaryTables)
        {
            var result = CreateSummaryTemplateTable(summaryTables);

            foreach (DataTable table in summaryTables)
            {
                AppendAlignedSummarySection(result, table);
                result.Rows.Add(result.NewRow());
            }

            if (result.Rows.Count > 0)
            {
                result.Rows.RemoveAt(result.Rows.Count - 1);
            }

            return result;
        }

        private static DataTable CreateSummaryTemplateTable(List<DataTable> summaryTables)
        {
            DataTable templateSource = summaryTables.First();
            var template = new DataTable();

            for (int i = 0; i < SummaryBlankLeadingColumns; i++)
            {
                template.Columns.Add(string.Empty, typeof(object));
            }

            int fixedColumnStart = Math.Min(SummaryLeadingColumnsToSkip, templateSource.Columns.Count);
            int fixedColumnCount = Math.Max(0, Math.Min(6, templateSource.Columns.Count) - fixedColumnStart);
            for (int i = fixedColumnStart; i < fixedColumnStart + fixedColumnCount; i++)
            {
                template.Columns.Add(templateSource.Columns[i].ColumnName, typeof(object));
            }

            foreach (string header in GetOrderedSummaryHeaders(summaryTables))
            {
                template.Columns.Add(header, typeof(object));
            }

            return template;
        }

        private static void AppendAlignedSummarySection(DataTable destination, DataTable source)
        {
            var destinationHeaderMap = BuildSummaryHeaderMap(destination);
            var sourceHeaderMap = BuildSummaryHeaderMap(source);
            int sourceFixedColumnStart = Math.Min(SummaryLeadingColumnsToSkip, source.Columns.Count);
            int sourceFixedColumnCount = Math.Max(0, Math.Min(6, source.Columns.Count) - sourceFixedColumnStart);

            for (int rowIndex = 0; rowIndex < source.Rows.Count; rowIndex++)
            {
                var newRow = destination.NewRow();

                for (int colIndex = 0; colIndex < Math.Min(sourceFixedColumnCount, destination.Columns.Count - SummaryBlankLeadingColumns); colIndex++)
                {
                    int sourceColumnIndex = sourceFixedColumnStart + colIndex;
                    if (sourceColumnIndex < source.Columns.Count)
                    {
                        newRow[SummaryBlankLeadingColumns + colIndex] = NormalizeSummaryCellValue(source.Rows[rowIndex][sourceColumnIndex]);
                    }
                }

                foreach (var sourceEntry in sourceHeaderMap)
                {
                    if (!destinationHeaderMap.TryGetValue(sourceEntry.Key, out int destinationColumnIndex))
                    {
                        continue;
                    }

                    newRow[destinationColumnIndex] = NormalizeSummaryCellValue(source.Rows[rowIndex][sourceEntry.Value]);
                }

                destination.Rows.Add(newRow);
            }
        }

        private static List<string> GetOrderedSummaryHeaders(List<DataTable> summaryTables)
        {
            var orderedHeaders = new List<string>();
            var seenHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (DataTable table in summaryTables.OrderByDescending(GetSummaryPeriodColumnCount))
            {
                foreach (string header in EnumerateSummaryHeaders(table))
                {
                    if (seenHeaders.Add(header))
                    {
                        orderedHeaders.Add(header);
                    }
                }
            }

            return orderedHeaders;
        }

        private static int GetSummaryPeriodColumnCount(DataTable table)
        {
            return EnumerateSummaryHeaders(table).Count();
        }

        private static IEnumerable<string> EnumerateSummaryHeaders(DataTable table)
        {
            if (table.Rows.Count == 0)
            {
                yield break;
            }

            int startIndex = GetSummaryPeriodStartIndex(table);
            for (int i = startIndex; i < table.Columns.Count; i++)
            {
                string columnName = table.Columns[i].ColumnName;
                if (IsSummarySeparatorColumnName(columnName))
                {
                    yield return columnName;
                    continue;
                }

                string headerValue = table.Rows[0][i]?.ToString()?.Trim() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(headerValue))
                {
                    yield return headerValue;
                }
            }
        }

        private static Dictionary<string, int> BuildSummaryHeaderMap(DataTable table)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int startIndex = GetSummaryPeriodStartIndex(table);
            if (table.Rows.Count == 0)
            {
                for (int i = startIndex; i < table.Columns.Count; i++)
                {
                    string columnName = table.Columns[i].ColumnName;
                    if (!string.IsNullOrWhiteSpace(columnName) && !map.ContainsKey(columnName))
                    {
                        map[columnName] = i;
                    }
                }

                return map;
            }

            foreach (string header in EnumerateSummaryHeaders(table))
            {
                if (IsSummarySeparatorColumnName(header))
                {
                    for (int i = startIndex; i < table.Columns.Count; i++)
                    {
                        if (string.Equals(table.Columns[i].ColumnName, header, StringComparison.OrdinalIgnoreCase) &&
                            !map.ContainsKey(header))
                        {
                            map[header] = i;
                            break;
                        }
                    }

                    continue;
                }

                for (int i = startIndex; i < table.Columns.Count; i++)
                {
                    string cellValue = table.Rows[0][i]?.ToString()?.Trim() ?? string.Empty;
                    if (string.Equals(cellValue, header, StringComparison.OrdinalIgnoreCase) && !map.ContainsKey(header))
                    {
                        map[header] = i;
                        break;
                    }
                }
            }

            return map;
        }

        private static int GetSummaryPeriodStartIndex(DataTable table)
        {
            int fixedColumnStart = Math.Min(SummaryLeadingColumnsToSkip, table.Columns.Count);
            int fixedColumnCount = Math.Max(0, Math.Min(6, table.Columns.Count) - fixedColumnStart);
            return SummaryBlankLeadingColumns + fixedColumnCount;
        }

        private static bool IsSummarySeparatorColumnName(string columnName)
        {
            return string.Equals(columnName, SummarySeparatorDayWeek, StringComparison.Ordinal) ||
                   string.Equals(columnName, SummarySeparatorWeekMonth, StringComparison.Ordinal);
        }

        private static object NormalizeSummaryCellValue(object value)
        {
            string text = value?.ToString()?.Trim() ?? string.Empty;
            return SummaryBlankHeaderValues.Contains(text) ? string.Empty : value;
        }

        private static void WriteDataTableAsCsvSection(StreamWriter writer, DataTable table, bool includeColumnHeader)
        {
            if (includeColumnHeader)
            {
                var headers = table.Columns.Cast<DataColumn>()
                    .Select(column => EscapeCsvField(column.ColumnName));
                writer.WriteLine(string.Join(",", headers));
            }

            foreach (DataRow row in table.Rows)
            {
                var fields = row.ItemArray
                    .Select(field => EscapeCsvField(field?.ToString() ?? string.Empty));
                writer.WriteLine(string.Join(",", fields));
            }
        }

        private static string SanitizePathSegment(string value)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            string sanitized = new string(value.Select(ch => invalidChars.Contains(ch) ? '_' : ch).ToArray()).Trim();
            return string.IsNullOrWhiteSpace(sanitized) ? "UnnamedModel" : sanitized;
        }

        private void CT_BT_REPORT_MOVE_UP_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = CT_LB_REPORT_JSON_ORDER.SelectedIndex;
            if (selectedIndex <= 0 || selectedIndex >= _pendingReportJsonFiles.Count)
            {
                return;
            }

            (_pendingReportJsonFiles[selectedIndex - 1], _pendingReportJsonFiles[selectedIndex]) =
                (_pendingReportJsonFiles[selectedIndex], _pendingReportJsonFiles[selectedIndex - 1]);

            PopulateReportOrderUi();
            CT_LB_REPORT_JSON_ORDER.SelectedIndex = selectedIndex - 1;
        }

        private void CT_BT_REPORT_MOVE_DOWN_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = CT_LB_REPORT_JSON_ORDER.SelectedIndex;
            if (selectedIndex < 0 || selectedIndex >= _pendingReportJsonFiles.Count - 1)
            {
                return;
            }

            (_pendingReportJsonFiles[selectedIndex + 1], _pendingReportJsonFiles[selectedIndex]) =
                (_pendingReportJsonFiles[selectedIndex], _pendingReportJsonFiles[selectedIndex + 1]);

            PopulateReportOrderUi();
            CT_LB_REPORT_JSON_ORDER.SelectedIndex = selectedIndex + 1;
        }

        private void CT_BT_REPORT_CANCEL_Click(object sender, RoutedEventArgs e)
        {
            ClearPendingReportSetup();
            ShowEmptyHome();
        }

        private async void CT_BT_REPORT_START_Click(object sender, RoutedEventArgs e)
        {
            if (_pendingReportJsonFiles.Count == 0)
            {
                return;
            }

            var folderDialog = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "Select folder to save the report CSV files",
                FolderName = _pendingReportOutputPath
            };

            if (!ShowCommonDialog(folderDialog))
            {
                return;
            }

            _pendingReportOutputPath = folderDialog.FolderName;
            PopulateReportOrderUi();

            CT_BT_GETREPORT.IsEnabled = false;
            CT_BT_GETREPORT.Content = "Processing...";
            SetReportSetupButtonsEnabled(false);
            CT_PROG_TASK.Value = 0;
            CT_PROG_TOTAL.Value = 0;

            try
            {
                await GetReport(_pendingReportJsonFiles.ToArray(), _pendingReportOutputPath);
                ClearPendingReportSetup();
                ShowEmptyHome();
            }
            finally
            {
                SetReportSetupButtonsEnabled(true);
            }
        }

        private void CT_BT_REPORT_SAVE_Click(object sender, RoutedEventArgs e)
        {
            if (_pendingReportJsonFiles.Count == 0)
            {
                ShowInfoMessage("There is no report order to save.", "Info");
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                Title = "Save report order configuration",
                Filter = "JSON files (*.json)|*.json",
                DefaultExt = ".json",
                FileName = "ReportOrder.json"
            };

            if (!ShowCommonDialog(saveFileDialog))
                return;

            var config = new ReportOrderConfig
            {
                JsonFiles = _pendingReportJsonFiles.ToList(),
                OutputPath = _pendingReportOutputPath
            };

            string json = JsonSerializer.Serialize(config, new JsonSerializerOptions
            {
                WriteIndented = true
            });

            File.WriteAllText(saveFileDialog.FileName, json, Encoding.UTF8);
            clLogger.LogImportant($"Saved report order config: {saveFileDialog.FileName}");
        }

        private void CT_BT_REPORT_LOAD_SAVED_Click(object sender, RoutedEventArgs e)
        {
            if (!TryLoadSavedReportOrder())
            {
                return;
            }

            PopulateReportOrderUi();
            ShowReportSetupContent();
        }

        private bool TryLoadSavedReportOrder()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Load report order configuration",
                Filter = "JSON files (*.json)|*.json",
                Multiselect = false
            };

            if (!ShowCommonDialog(openFileDialog))
                return false;

            try
            {
                string json = File.ReadAllText(openFileDialog.FileName);
                var config = JsonSerializer.Deserialize<ReportOrderConfig>(json);

                if (config == null || config.JsonFiles.Count == 0)
                {
                    ShowErrorMessage("The selected configuration file is empty.", "Error");
                    return false;
                }

                var missingFiles = config.JsonFiles.Where(path => !File.Exists(path)).ToList();
                var validFiles = config.JsonFiles.Where(File.Exists).ToList();

                if (validFiles.Count == 0)
                {
                    ShowErrorMessage("None of the JSON files in the saved configuration could be found.", "Error");
                    return false;
                }

                _pendingReportJsonFiles.Clear();
                _pendingReportJsonFiles.AddRange(validFiles);
                _pendingReportOutputPath = config.OutputPath;
                clLogger.LogImportant($"Loaded report order config: {openFileDialog.FileName}");

                foreach (string missingFile in missingFiles)
                {
                    clLogger.LogWarning($"Missing saved JSON file skipped: {missingFile}");
                }

                return true;
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Load saved report order");
                ShowErrorMessage($"Error loading saved report order: {ex.Message}", "Error");
                return false;
            }
        }

        private bool TrySelectReportInputJsonFiles()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select report or model JSON file(s)",
                Filter = "JSON files (*.json)|*.json",
                Multiselect = true
            };

            if (!ShowCommonDialog(openFileDialog))
            {
                return false;
            }

            try
            {
                _pendingReportJsonFiles.Clear();
                _pendingReportJsonFiles.AddRange(ResolveReportJsonInputFiles(openFileDialog.FileNames));
                _pendingReportOutputPath = string.Empty;
                return true;
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "TrySelectReportInputJsonFiles");
                ShowErrorMessage(BuildJsonLoadErrorMessage(ex), "Error");
                _pendingReportJsonFiles.Clear();
                return false;
            }
        }

        private static List<string> ResolveReportJsonInputFiles(IEnumerable<string> selectedPaths)
        {
            var resolvedFiles = new List<string>();
            var seenFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (string selectedPath in selectedPaths)
            {
                if (string.IsNullOrWhiteSpace(selectedPath) || !File.Exists(selectedPath))
                {
                    continue;
                }

                foreach (string resolvedPath in ResolveSingleReportJsonInputFile(selectedPath))
                {
                    if (seenFiles.Add(resolvedPath))
                    {
                        resolvedFiles.Add(resolvedPath);
                    }
                }
            }

            return resolvedFiles;
        }

        private static IEnumerable<string> ResolveSingleReportJsonInputFile(string filePath)
        {
            string json = File.ReadAllText(filePath);

            using JsonDocument document = JsonDocument.Parse(json);
            JsonElement root = document.RootElement;

            if (root.ValueKind == JsonValueKind.Array)
            {
                return new[] { filePath };
            }

            if (root.ValueKind == JsonValueKind.Object &&
                root.TryGetProperty("JsonFiles", out JsonElement jsonFilesElement) &&
                jsonFilesElement.ValueKind == JsonValueKind.Array)
            {
                var config = JsonSerializer.Deserialize<ReportOrderConfig>(json) ?? new ReportOrderConfig();
                string baseDirectory = Path.GetDirectoryName(filePath) ?? string.Empty;

                return (config.JsonFiles ?? new List<string>())
                    .Where(path => !string.IsNullOrWhiteSpace(path))
                    .Select(path => Path.IsPathRooted(path) ? path : Path.Combine(baseDirectory, path))
                    .Where(File.Exists)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            throw new JsonException($"Unsupported report JSON format: {filePath}");
        }

        private static string EscapeCsvField(string value)
        {
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\r") || value.Contains("\n"))
            {
                return $"\"{value.Replace("\"", "\"\"")}\"";
            }

            return value;
        }

        private void GetProcTable(List<clModelGroupData> ListModels)
        {
            var allModels = new HashSet<string>();
            foreach (var EachModel in ListModels)
            {
                foreach (var model in EachModel.ModelList)
                {
                    allModels.Add(model);
                }
            }

            ValidateAndPrepareOriginalTable();

            clMakeProcTable c = new clMakeProcTable(PathSelectedDB);
            c.Run(allModels.ToList());
            InvalidateDebugGroupTableCache();
        }

        private static void RunDataProcessing(
            string dbPath,
            string processTypeCsvPath,
            string reasonCsvPath,
            IProgress<(int Percent, string Message)>? progress = null)
        {
            var dataProcessor = new clDataProcessor(
                dbPath,
                processTypeCsvPath,
                reasonCsvPath,
                (percent, message) => progress?.Report((percent, message)));
            try
            {
                dataProcessor.ProcessData();
            }
            finally
            {
                dataProcessor.Dispose();
            }
        }

        private void UpdateProgressBars(int taskProgress, int totalProgress)
        {
            int safeTask = Math.Clamp(taskProgress, 0, 100);
            int safeTotal = Math.Clamp(totalProgress, 0, 100);
            _currentTaskProgressPercent = safeTask;
            _currentTotalProgressPercent = safeTotal;
            NotifyWebModuleSnapshotChanged();

            if (Dispatcher.CheckAccess())
            {
                CT_PROG_TASK.Value = safeTask;
                CT_PROG_TOTAL.Value = safeTotal;
                CT_LB_TASK_PERCENT.Text = $"{safeTask}%";
                CT_LB_TOTAL_PERCENT.Text = $"{safeTotal}%";
                return;
            }

            _ = Dispatcher.BeginInvoke(new Action(() =>
            {
                CT_PROG_TASK.Value = safeTask;
                CT_PROG_TOTAL.Value = safeTotal;
                CT_LB_TASK_PERCENT.Text = $"{safeTask}%";
                CT_LB_TOTAL_PERCENT.Text = $"{safeTotal}%";
            }));
        }

        private void NotifyWebModuleSnapshotChanged()
        {
            WebModuleSnapshotChanged?.Invoke();
        }

        private void InvalidateDebugGroupTableCache()
        {
            _cachedDebugGroupTableName = string.Empty;
            _cachedDebugGroupTable = null;
        }

        private void ValidateAndPrepareOriginalTable()
        {
            bool needsPreparation = false;

            using (var sql = new clSQLFileIO(PathSelectedDB))
            {
                if (sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ORG))
                {
                    var columns = sql.GetColumns(CONSTANT.OPTION_TABLE_NAME.ORG);
                    needsPreparation =
                        !columns.Contains(CONSTANT.NGCODE.NEW) ||
                        !columns.Contains(CONSTANT.ModelName_WithLineShift.NEW);
                }
                else
                {
                    throw new InvalidOperationException("OriginalTable does not exist. Please load the DB first.");
                }
            }

            if (needsPreparation)
            {
                clLogger.LogWarning("OriginalTable is missing required columns for Proc generation. Preparing now...");

                using (var sql = new clSQLFileIO(PathSelectedDB))
                {
                    sql.Processor.SetLineShiftColumnValue(CONSTANT.OPTION_TABLE_NAME.ORG);
                }

                using (var verifySQL = new clSQLFileIO(PathSelectedDB))
                {
                    var verifyColumns = verifySQL.GetColumns(CONSTANT.OPTION_TABLE_NAME.ORG);
                    if (verifyColumns.Contains(CONSTANT.NGCODE.NEW) &&
                        verifyColumns.Contains(CONSTANT.ModelName_WithLineShift.NEW))
                    {
                        clLogger.Log("OriginalTable prepared successfully for Proc generation");
                    }
                    else
                    {
                        clLogger.LogWarning("Warning: OriginalTable preparation may not have completed properly");
                        clLogger.Log($"  -> Current columns: {string.Join(", ", verifyColumns.Take(10))}");
                    }
                }
            }
        }

        private void CT_BT_MODELGROUPS_Click(object sender, RoutedEventArgs e)
        {
            if (ListUniqueModelName == null || ListUniqueModelName.Count == 0)
            {
                ShowWarningMessage("Please load BMES data first.", "Warning");
                return;
            }

            var win = new ModelGroupWindow(ListUniqueModelName);
            ShowModalWindow(win, sender);
        }

        private void CT_BT_SETTINGS_Click(object sender, RoutedEventArgs e)
        {
            var win = new SettingsWindow(_processTypeCsvPath, _reasonCsvPath);
            if (ShowModalWindow(win, sender) == true)
            {
                _processTypeCsvPath = win.RoutingPath;
                _reasonCsvPath = win.ReasonPath;
                clLogger.Log("Settings updated.");
            }
        }

        private bool? ShowModalWindow(Window dialog, object? sender)
        {
            Window? owner = ResolveDialogOwner(sender);
            if (owner is not null && owner.IsVisible && !ReferenceEquals(owner, dialog))
            {
                dialog.Owner = owner;
            }

            return dialog.ShowDialog();
        }

        private bool ShowInfoMessage(string message, string title)
        {
            Window? owner = ResolveDialogOwner(null);
            return owner is not null
                ? MessageBox.Show(owner, message, title, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK
                : MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information) == MessageBoxResult.OK;
        }

        private bool ShowWarningMessage(string message, string title)
        {
            Window? owner = ResolveDialogOwner(null);
            return owner is not null
                ? MessageBox.Show(owner, message, title, MessageBoxButton.OK, MessageBoxImage.Warning) == MessageBoxResult.OK
                : MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Warning) == MessageBoxResult.OK;
        }

        private bool ShowErrorMessage(string message, string title)
        {
            Window? owner = ResolveDialogOwner(null);
            return owner is not null
                ? MessageBox.Show(owner, message, title, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK
                : MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Error) == MessageBoxResult.OK;
        }

        private MessageBoxResult ShowQuestionMessage(string message, string title, MessageBoxButton buttons)
        {
            Window? owner = ResolveDialogOwner(null);
            return owner is not null
                ? MessageBox.Show(owner, message, title, buttons, MessageBoxImage.Question)
                : MessageBox.Show(message, title, buttons, MessageBoxImage.Question);
        }

        private bool ShowCommonDialog(CommonDialog dialog)
        {
            try
            {
                Window? owner = ResolveDialogOwner(null);
                if (owner?.IsVisible == true)
                {
                    return dialog.ShowDialog(owner) == true;
                }

                return dialog.ShowDialog() == true;
            }
            catch (InvalidOperationException ex)
            {
                clLogger.LogException(ex, "Dialog owner was invalid. Retrying without owner.");
                return dialog.ShowDialog() == true;
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error showing common dialog");
                ShowWarningMessage($"Failed to open dialog.\n{ex.Message}", "Dialog Error");
                return false;
            }
        }

        private Window? ResolveDialogOwner(object? sender)
        {
            if (sender is DependencyObject dependencyObject)
            {
                Window? senderWindow = Window.GetWindow(dependencyObject);
                if (senderWindow?.IsVisible == true)
                {
                    return senderWindow;
                }
            }

            if (IsVisible)
            {
                return this;
            }

            Window? activeWindow = Application.Current?.Windows
                .OfType<Window>()
                .FirstOrDefault(window => window.IsActive && window.IsVisible);

            if (activeWindow is not null)
            {
                return activeWindow;
            }

            Window? mainWindow = Application.Current?.MainWindow;
            return mainWindow?.IsVisible == true ? mainWindow : null;
        }

        public static void SaveToJson(List<clModelGroupData> selectedModels, string filePath)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                string json = JsonSerializer.Serialize(selectedModels, options);
                File.WriteAllText(filePath, json);

                clLogger.Log($"Model groups saved to {filePath}");
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error saving model groups to JSON");
                throw;
            }
        }

        public static List<clModelGroupData> LoadFromJson(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    clLogger.LogWarning($"File not found: {filePath}");
                    return new List<clModelGroupData>();
                }

                string json = File.ReadAllText(filePath);
                string trimmedJson = json.TrimStart();
                if (string.IsNullOrWhiteSpace(trimmedJson))
                {
                    clLogger.LogWarning($"Empty JSON file: {filePath}");
                    return new List<clModelGroupData>();
                }

                List<clModelGroupData> data;
                char firstChar = trimmedJson[0];

                if (firstChar == '[')
                {
                    data = LoadGroupArrayJsonWithRepair(filePath);
                }
                else if (firstChar == '{')
                {
                    using JsonDocument document = JsonDocument.Parse(json);
                    JsonElement root = document.RootElement;

                    if (root.TryGetProperty("JsonFiles", out JsonElement jsonFilesElement) &&
                        jsonFilesElement.ValueKind == JsonValueKind.Array)
                    {
                        var config = JsonSerializer.Deserialize<ReportOrderConfig>(json) ?? new ReportOrderConfig();
                        data = LoadGroupsFromReportOrderConfig(config, filePath);
                    }
                    else
                    {
                        throw new JsonException("Unsupported object JSON format. Expected report-order object with JsonFiles.");
                    }
                }
                else
                {
                    throw new JsonException("Unsupported JSON format. Expected group array or report-order object.");
                }

                clLogger.Log($"Model groups loaded from {filePath}");
                return data;
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error loading model groups from JSON");
                throw;
            }
        }

        public async void MakeGroupUI()
        {
            SetProgressPopupState("Load Group", "Loading model groups...");
            try
            {
                await GetUniqueModelsAsync();
                UpdateProgressBars(100, 100);
            }
            finally
            {
                CloseProgressPopup();
            }
        }

        private static List<clModelGroupData> LoadGroupsFromReportOrderConfig(ReportOrderConfig config, string sourceFilePath)
        {
            var mergedGroups = new List<clModelGroupData>();
            string baseDirectory = Path.GetDirectoryName(sourceFilePath) ?? string.Empty;

            foreach (string jsonFile in config.JsonFiles ?? new List<string>())
            {
                if (string.IsNullOrWhiteSpace(jsonFile))
                {
                    continue;
                }

                string resolvedPath = jsonFile;
                if (!Path.IsPathRooted(resolvedPath))
                {
                    resolvedPath = Path.Combine(baseDirectory, resolvedPath);
                }

                if (!File.Exists(resolvedPath))
                {
                    clLogger.LogWarning($"Referenced group JSON not found: {resolvedPath}");
                    continue;
                }

                List<clModelGroupData> childGroups = LoadGroupArrayJsonWithRepair(resolvedPath);

                foreach (var group in childGroups)
                {
                    if (string.IsNullOrWhiteSpace(group.GroupName))
                    {
                        string fileLabel = Path.GetFileNameWithoutExtension(resolvedPath);
                        string modelLabel = string.Join("_", (group.ModelList ?? new List<string>()).OrderBy(m => m));
                        group.GroupName = string.IsNullOrWhiteSpace(modelLabel) ? fileLabel : $"{fileLabel} - {modelLabel}";
                    }

                    mergedGroups.Add(group);
                }
            }

            return mergedGroups;
        }

        private static List<clModelGroupData> LoadGroupArrayJsonWithRepair(string filePath)
        {
            string json = File.ReadAllText(filePath);

            try
            {
                return JsonSerializer.Deserialize<List<clModelGroupData>>(json) ?? new List<clModelGroupData>();
            }
            catch (JsonException firstException)
            {
                string repairedJson = RepairGroupJson(json);

                try
                {
                    var repairedGroups = JsonSerializer.Deserialize<List<clModelGroupData>>(repairedJson) ?? new List<clModelGroupData>();
                    clLogger.LogWarning($"Recovered malformed group JSON automatically: {filePath} ({FormatJsonExceptionDetails(firstException)})");
                    return repairedGroups;
                }
                catch (JsonException secondException)
                {
                    throw new JsonException(
                        $"Invalid group JSON file: {filePath}. Original parse error: {FormatJsonExceptionDetails(firstException)}. Repair parse error: {FormatJsonExceptionDetails(secondException)}",
                        secondException);
                }
            }
        }

        private static string BuildJsonLoadErrorMessage(Exception ex)
        {
            if (ex is JsonException jsonEx)
            {
                return $"Failed to load JSON.\n{FormatJsonExceptionDetails(jsonEx)}";
            }

            return $"Failed to load JSON.\n{ex.Message}";
        }

        private static string FormatJsonExceptionDetails(JsonException ex)
        {
            string message = ex.Message ?? "Unknown JSON parse error.";
            string lineText = ex.LineNumber.HasValue ? $"line {ex.LineNumber.Value + 1}" : "line unknown";
            string byteText = ex.BytePositionInLine.HasValue ? $"byte {ex.BytePositionInLine.Value + 1}" : "byte unknown";

            return $"{message} ({lineText}, {byteText})";
        }

        private static string RepairGroupJson(string json)
        {
            string repaired = json;

            repaired = Regex.Replace(repaired, ",\\s*(\\]|\\})", "$1");
            repaired = Regex.Replace(repaired, "(\\}|\\])\\s*(\\{)", "$1,$2");

            return repaired;
        }

        private async Task GetUniqueModelsAsync()
        {
            if (string.IsNullOrWhiteSpace(PathSelectedDB) || !File.Exists(PathSelectedDB))
            {
                ListUniqueModelName = new List<string>();
                return;
            }

            UpdateLoadDbProgress(75, 75, "Refreshing model list...");
            ListUniqueModelName = await Task.Run(() =>
            {
                using var sql = new clSQLFileIO(PathSelectedDB);
                sql.Processor.SetLineShiftColumnValue(CONSTANT.OPTION_TABLE_NAME.ORG);
                var list = sql.GetUniqueValues(CONSTANT.OPTION_TABLE_NAME.ORG, CONSTANT.ModelName_WithLineShift.NEW);
                list.Sort();
                return list;
            });
            UpdateLoadDbProgress(84, 84, $"Model list refreshed: {ListUniqueModelName.Count:N0} items");
        }

        private void ShowProgressPopup(string title, string message)
        {
            SetProgressPopupState(title, message);
        }

        private void SetProgressPopupState(string title, string message)
        {
            _currentProgressTitle = title;
            _currentProgressMessage = message;
            NotifyWebModuleSnapshotChanged();

            if (!Dispatcher.CheckAccess())
            {
                _ = Dispatcher.BeginInvoke(new Action(() => SetProgressPopupState(title, message)));
                return;
            }
        }

        private void CloseProgressPopup()
        {
            if (!Dispatcher.CheckAccess())
            {
                _ = Dispatcher.BeginInvoke(new Action(CloseProgressPopup));
                return;
            }

            _currentProgressTitle = "Processing";
            _currentProgressMessage = "Please wait...";
        }

        private async void CT_BT_MAPPING_Click(object sender, RoutedEventArgs e)
        {
            await ShowMappingDashboardAsync(MappingTableKind.Routing, forceScan: true);
        }

        private async void CT_BT_MAPPING_SHOW_ROUTING_Click(object sender, RoutedEventArgs e)
        {
            await ShowMappingDashboardAsync(MappingTableKind.Routing, forceScan: _mappingScanResult == null);
        }

        private async void CT_BT_MAPPING_SHOW_REASON_Click(object sender, RoutedEventArgs e)
        {
            await ShowMappingDashboardAsync(MappingTableKind.Reason, forceScan: _mappingScanResult == null);
        }

        private async Task ShowMappingDashboardAsync(MappingTableKind tableKind, bool forceScan)
        {
            if (string.IsNullOrWhiteSpace(PathSelectedDB))
            {
                ShowWarningMessage("Please load BMES data first.", "Warning");
                return;
            }

            _currentMappingTableKind = tableKind;
            ShowMappingContent();
            UpdateMappingSelectorState();

            if (forceScan || _mappingScanResult == null)
            {
                await ScanAllMissingMappingsAsync(showProgressPopup: true);
            }

            await RefreshMappingTableViewAsync();
        }

        private async Task RefreshMappingTableViewAsync()
        {
            if (_currentMappingTableKind == MappingTableKind.None)
            {
                CT_MAPPING_CONTENT.Visibility = Visibility.Collapsed;
                return;
            }

            CT_LB_MAPPING_TITLE.Text = _currentMappingTableKind == MappingTableKind.Routing
                ? "UNMAPPED ROUTING"
                : "UNMAPPED REASON";
            CT_LB_MAPPING_PATH.Text = $"DB: {Path.GetFileName(PathSelectedDB)}";
            UpdateMappingSelectorState();
            UpdateMappingSummaryText();
            _mappingTotalRowCount = _currentMappingTableKind == MappingTableKind.Routing
                ? _cachedMissingRoutingItems.Count
                : _cachedMissingReasonItems.Count;
            CT_DG_MAPPING.ItemsSource = null;
            CT_LB_MAPPING_EMPTY.Visibility = Visibility.Collapsed;
            SetMappingButtonsEnabled(false);

            _currentMappingTableData = await Task.Run(BuildCurrentMissingMappingTable);
            BindMappingGrid();
            SetMappingButtonsEnabled(true);
        }

        private void BindMappingGrid()
        {
            if (_currentMappingTableData == null || _currentMappingTableData.Columns.Count == 0)
            {
                CT_DG_MAPPING.ItemsSource = null;
                CT_LB_MAPPING_EMPTY.Text = "No table schema to display.";
                CT_LB_MAPPING_EMPTY.Visibility = Visibility.Visible;
                return;
            }

            if (_currentMappingTableData.Rows.Count == 0)
            {
                CT_DG_MAPPING.ItemsSource = _currentMappingTableData.DefaultView;
                CT_LB_MAPPING_EMPTY.Text = "No rows to display.";
                CT_LB_MAPPING_EMPTY.Visibility = Visibility.Visible;
                return;
            }

            CT_DG_MAPPING.ItemsSource = _currentMappingTableData.DefaultView;
            CT_LB_MAPPING_EMPTY.Visibility = Visibility.Collapsed;
        }

        private string GetCurrentMappingTableName()
        {
            return _currentMappingTableKind == MappingTableKind.Routing
                ? CONSTANT.OPTION_TABLE_NAME.ROUTING
                : CONSTANT.OPTION_TABLE_NAME.REASON;
        }

        private string GetCurrentMappingFilePath()
        {
            return _currentMappingTableKind == MappingTableKind.Routing
                ? _processTypeCsvPath
                : _reasonCsvPath;
        }

        private async Task ScanAllMissingMappingsAsync(bool showProgressPopup)
        {
            if (string.IsNullOrWhiteSpace(PathSelectedDB))
            {
                return;
            }

            if (showProgressPopup)
            {
                ShowProgressPopup("Mapping Scan", "Scanning OriginalTable for unmapped Routing and Reason...");
                UpdateProgressBars(0, 0);
            }

            try
            {
                var progress = new Progress<(int Percent, string Message)>(state =>
                {
                    if (showProgressPopup)
                    {
                        SetProgressPopupState("Mapping Scan", state.Message);
                        UpdateProgressBars(state.Percent, state.Percent);
                        return;
                    }

                    int loadDbTotalProgress = 85 + (int)Math.Round(state.Percent * 0.11);
                    int loadDbTaskProgress = Math.Max(1, state.Percent);
                    if (_currentProgressTitle == "Load DB")
                    {
                        UpdateLoadDbProgress(loadDbTaskProgress, loadDbTotalProgress, state.Message);
                        return;
                    }
                });

                MappingScanResult result = await Task.Run(() => BuildMissingMappingScanResultOptimized(progress));
                _mappingScanResult = result;
                _cachedMissingRoutingItems = result.MissingRoutingItems;
                _cachedMissingReasonItems = result.MissingReasonItems;
                _hasScannedMissingRoutingItems = true;
                _hasScannedMissingReasonItems = true;
                UpdateLoadMappingStatusUi();

                if (showProgressPopup)
                {
                    UpdateProgressBars(100, 100);
                }
            }
            finally
            {
                if (showProgressPopup)
                {
                    CloseProgressPopup();
                }
            }
        }

        private IReadOnlyList<IReadOnlyDictionary<string, string>> BuildMissingItemsForKind(
            MappingTableKind kind,
            IProgress<(int Percent, string Message)>? progress = null)
        {
            using var sql = new clSQLFileIO(PathSelectedDB);
            if (!sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ORG))
            {
                return Array.Empty<IReadOnlyDictionary<string, string>>();
            }

            progress?.Report((10, "Loading OriginalTable..."));
            DataTable orgTable = sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ORG);

            if (kind == MappingTableKind.Routing)
            {
                progress?.Report((25, "Loading RoutingTable..."));
                DataTable routingTable = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                    ? sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                    : new DataTable();
                string routingModelColumnName = GetRoutingModelColumnName(routingTable);

                var routingKeys = new HashSet<(string ModelName, string ProcessCode, string ProcessName)>();
                foreach (DataRow row in routingTable.Rows)
                {
                    routingKeys.Add((
                        NormalizeCompareValue(row[routingModelColumnName]?.ToString()),
                        NormalizeCompareValue(row["ProcessCode"]?.ToString()),
                        NormalizeCompareValue(row["ProcessName"]?.ToString(), normalize: true)));
                }

                var result = new List<IReadOnlyDictionary<string, string>>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int total = Math.Max(orgTable.Rows.Count, 1);

                for (int i = 0; i < orgTable.Rows.Count; i++)
                {
                    DataRow row = orgTable.Rows[i];
                    string modelName = row.Field<string>(CONSTANT.MATERIALNAME.NEW) ?? string.Empty;
                    string processCode = row.Field<string>(CONSTANT.PROCESSCODE.NEW) ?? string.Empty;
                    string processName = row.Field<string>(CONSTANT.PROCESSNAME.NEW) ?? string.Empty;

                    string modelKey = NormalizeCompareValue(modelName);
                    string processCodeKey = NormalizeCompareValue(processCode);
                    string processNameKey = NormalizeCompareValue(processName, normalize: true);

                    if (!routingKeys.Contains((modelKey, processCodeKey, processNameKey)))
                    {
                        string dedupeKey = $"{modelKey}|{processCodeKey}|{processNameKey}";
                        if (seen.Add(dedupeKey))
                        {
                            result.Add(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                            {
                                ["모델명"] = modelName,
                                ["ProcessCode"] = processCode,
                                ["ProcessName"] = processName,
                                ["ProcessType"] = string.Empty
                            });
                        }
                    }

                    if (i == orgTable.Rows.Count - 1 || (i + 1) % LargeScanProgressLogInterval == 0)
                    {
                        int percent = 35 + (int)(65d * (i + 1) / total);
                        progress?.Report((percent, $"Scanning Routing candidates... {i + 1:N0}/{orgTable.Rows.Count:N0}"));
                    }
                }

                return result
                    .OrderBy(item => item["모델명"])
                    .ThenBy(item => item["ProcessCode"])
                    .ThenBy(item => item["ProcessName"])
                    .ToList();
            }

            progress?.Report((25, "Loading Reason table..."));
            DataTable reasonTable = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.REASON)
                ? sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.REASON)
                : new DataTable();

            var reasonKeys = new HashSet<(string ProcessName, string NgName)>();
            foreach (DataRow row in reasonTable.Rows)
            {
                reasonKeys.Add((
                    NormalizeCompareValue(row["processName"]?.ToString(), normalize: true),
                    NormalizeCompareValue(row["NgName"]?.ToString(), normalize: true)));
            }

            var reasonResult = new List<IReadOnlyDictionary<string, string>>();
            var reasonSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int reasonTotal = Math.Max(orgTable.Rows.Count, 1);

            for (int i = 0; i < orgTable.Rows.Count; i++)
            {
                DataRow row = orgTable.Rows[i];
                string processName = row.Field<string>(CONSTANT.PROCESSNAME.NEW) ?? string.Empty;
                string ngName = row.Field<string>(CONSTANT.NGNAME.NEW) ?? string.Empty;

                string processNameKey = NormalizeCompareValue(processName, normalize: true);
                string ngNameKey = NormalizeCompareValue(ngName, normalize: true);

                if (!reasonKeys.Contains((processNameKey, ngNameKey)))
                {
                    string dedupeKey = $"{processNameKey}|{ngNameKey}";
                    if (reasonSeen.Add(dedupeKey))
                    {
                        reasonResult.Add(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["processName"] = processName,
                            ["NgName"] = ngName,
                            ["Reason"] = string.Empty
                        });
                    }
                }

                if (i == orgTable.Rows.Count - 1 || (i + 1) % LargeScanProgressLogInterval == 0)
                {
                    int percent = 35 + (int)(65d * (i + 1) / reasonTotal);
                    progress?.Report((percent, $"Scanning Reason candidates... {i + 1:N0}/{orgTable.Rows.Count:N0}"));
                }
            }

            return reasonResult
                .OrderBy(item => item["processName"])
                .ThenBy(item => item["NgName"])
                .ToList();
        }

        private void UpdateMappingPathText()
        {
            string tableName = GetCurrentMappingTableName();
            if (_currentMappingTableKind == MappingTableKind.None)
            {
                CT_LB_MAPPING_PATH.Text = $"Table: {tableName}";
                return;
            }

            string countText = _mappingTotalRowCount > 0
                ? $"Rows: {_mappingTotalRowCount:N0}"
                : "Rows: counting...";

            string previewText = _currentMappingTableData == null
                ? "Preview: loading..."
                : _currentMappingTableData.Rows.Count >= MappingPreviewRowLimit
                    ? $"Preview: first {MappingPreviewRowLimit:N0} rows"
                    : $"Preview: all {_currentMappingTableData.Rows.Count:N0} rows";

            CT_LB_MAPPING_PATH.Text = $"Table: {tableName}\n{countText}\n{previewText}";
        }

        private MappingScanResult BuildMissingMappingScanResult(IProgress<(int Percent, string Message)>? progress = null)
        {
            using var sql = new clSQLFileIO(PathSelectedDB);
            if (!sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ORG))
            {
                return new MappingScanResult
                {
                    MissingRoutingItems = Array.Empty<IReadOnlyDictionary<string, string>>(),
                    MissingReasonItems = Array.Empty<IReadOnlyDictionary<string, string>>(),
                    TotalOriginalRows = 0,
                    RoutingTableRows = 0,
                    ReasonTableRows = 0,
                    AppliedRoutingRows = 0,
                    MissingRoutingRows = 0,
                    AppliedReasonRows = 0,
                    MissingReasonRows = 0
                };
            }

            progress?.Report((10, "Loading OriginalTable..."));
            DataTable orgTable = sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ORG);

            int routingTableRows = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                ? sql.GetRowCount(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                : 0;
            int reasonTableRows = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.REASON)
                ? sql.GetRowCount(CONSTANT.OPTION_TABLE_NAME.REASON)
                : 0;

            progress?.Report((20, "Loading RoutingTable..."));
            DataTable routingTable = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                ? sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                : new DataTable();

            progress?.Report((25, "Loading Reason table..."));
            DataTable reasonTable = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.REASON)
                ? sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.REASON)
                : new DataTable();

            string routingModelColumnName = GetRoutingModelColumnName(routingTable);
            var routingKeys = new HashSet<(string ModelName, string ProcessCode, string ProcessName)>();
            foreach (DataRow row in routingTable.Rows)
            {
                routingKeys.Add((
                    NormalizeCompareValue(row[routingModelColumnName]?.ToString()),
                    NormalizeCompareValue(row["ProcessCode"]?.ToString()),
                    NormalizeCompareValue(row["ProcessName"]?.ToString(), normalize: true)));
            }

            var reasonKeys = new HashSet<(string ProcessName, string NgName)>();
            foreach (DataRow row in reasonTable.Rows)
            {
                reasonKeys.Add((
                    NormalizeCompareValue(row["processName"]?.ToString(), normalize: true),
                    NormalizeCompareValue(row["NgName"]?.ToString(), normalize: true)));
            }

            var missingRoutingItems = new List<IReadOnlyDictionary<string, string>>();
            var missingReasonItems = new List<IReadOnlyDictionary<string, string>>();
            var seenRouting = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenReason = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            int appliedRoutingRows = 0;
            int missingRoutingRows = 0;
            int appliedReasonRows = 0;
            int missingReasonRows = 0;
            int total = Math.Max(orgTable.Rows.Count, 1);

            for (int i = 0; i < orgTable.Rows.Count; i++)
            {
                DataRow row = orgTable.Rows[i];
                string modelName = row.Field<string>(CONSTANT.MATERIALNAME.NEW) ?? string.Empty;
                string processCode = row.Field<string>(CONSTANT.PROCESSCODE.NEW) ?? string.Empty;
                string processName = row.Field<string>(CONSTANT.PROCESSNAME.NEW) ?? string.Empty;
                string ngName = row.Field<string>(CONSTANT.NGNAME.NEW) ?? string.Empty;

                string modelKey = NormalizeCompareValue(modelName);
                string processCodeKey = NormalizeCompareValue(processCode);
                string processNameKey = NormalizeCompareValue(processName, normalize: true);
                string ngNameKey = NormalizeCompareValue(ngName, normalize: true);

                bool hasRouting = routingKeys.Contains((modelKey, processCodeKey, processNameKey));
                bool hasReason = reasonKeys.Contains((processNameKey, ngNameKey));

                if (hasRouting)
                {
                    appliedRoutingRows++;
                }
                else
                {
                    missingRoutingRows++;
                    string routingDedupeKey = $"{modelKey}|{processCodeKey}|{processNameKey}";
                    if (seenRouting.Add(routingDedupeKey))
                    {
                        missingRoutingItems.Add(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["ëª¨ë¸ëª…"] = modelName,
                            ["ProcessCode"] = processCode,
                            ["ProcessName"] = processName,
                            ["ProcessType"] = string.Empty
                        });
                    }
                }

                if (hasReason)
                {
                    appliedReasonRows++;
                }
                else
                {
                    missingReasonRows++;
                    string reasonDedupeKey = $"{processNameKey}|{ngNameKey}";
                    if (seenReason.Add(reasonDedupeKey))
                    {
                        missingReasonItems.Add(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["processName"] = processName,
                            ["NgName"] = ngName,
                            ["Reason"] = string.Empty
                        });
                    }
                }

                if (i == orgTable.Rows.Count - 1 || (i + 1) % LargeScanProgressLogInterval == 0)
                {
                    int percent = 35 + (int)(65d * (i + 1) / total);
                    progress?.Report((percent, $"Scanning Routing/Reason candidates... {i + 1:N0}/{orgTable.Rows.Count:N0}"));
                }
            }

            return new MappingScanResult
            {
                MissingRoutingItems = missingRoutingItems
                    .OrderBy(item => item["ëª¨ë¸ëª…"])
                    .ThenBy(item => item["ProcessCode"])
                    .ThenBy(item => item["ProcessName"])
                    .ToList(),
                MissingReasonItems = missingReasonItems
                    .OrderBy(item => item["processName"])
                    .ThenBy(item => item["NgName"])
                    .ToList(),
                TotalOriginalRows = orgTable.Rows.Count,
                RoutingTableRows = routingTableRows,
                ReasonTableRows = reasonTableRows,
                AppliedRoutingRows = appliedRoutingRows,
                MissingRoutingRows = missingRoutingRows,
                AppliedReasonRows = appliedReasonRows,
                MissingReasonRows = missingReasonRows
            };
        }

        private MappingScanResult BuildMissingMappingScanResultOptimized(IProgress<(int Percent, string Message)>? progress = null)
        {
            using var sql = new clSQLFileIO(PathSelectedDB);
            if (!sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ORG))
            {
                return new MappingScanResult
                {
                    MissingRoutingItems = Array.Empty<IReadOnlyDictionary<string, string>>(),
                    MissingReasonItems = Array.Empty<IReadOnlyDictionary<string, string>>(),
                    TotalOriginalRows = 0,
                    RoutingTableRows = 0,
                    ReasonTableRows = 0,
                    AppliedRoutingRows = 0,
                    MissingRoutingRows = 0,
                    AppliedReasonRows = 0,
                    MissingReasonRows = 0
                };
            }

            progress?.Report((10, "Loading OriginalTable..."));
            DataTable orgTable = sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ORG);

            int routingTableRows = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                ? sql.GetRowCount(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                : 0;
            int reasonTableRows = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.REASON)
                ? sql.GetRowCount(CONSTANT.OPTION_TABLE_NAME.REASON)
                : 0;

            progress?.Report((20, "Loading RoutingTable..."));
            DataTable routingTable = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                ? sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                : new DataTable();

            progress?.Report((25, "Loading Reason table..."));
            DataTable reasonTable = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.REASON)
                ? sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.REASON)
                : new DataTable();

            string routingModelColumnName = GetRoutingModelColumnName(routingTable);
            var routingKeys = new HashSet<(string ModelName, string ProcessCode, string ProcessName)>();
            foreach (DataRow row in routingTable.Rows)
            {
                routingKeys.Add((
                    NormalizeCompareValue(row[routingModelColumnName]?.ToString()),
                    NormalizeCompareValue(row["ProcessCode"]?.ToString()),
                    NormalizeCompareValue(row["ProcessName"]?.ToString(), normalize: true)));
            }

            var reasonKeys = new HashSet<(string ProcessName, string NgName)>();
            foreach (DataRow row in reasonTable.Rows)
            {
                reasonKeys.Add((
                    NormalizeCompareValue(row["processName"]?.ToString(), normalize: true),
                    NormalizeCompareValue(row["NgName"]?.ToString(), normalize: true)));
            }

            var missingRoutingItems = new List<IReadOnlyDictionary<string, string>>();
            var missingReasonItems = new List<IReadOnlyDictionary<string, string>>();
            var seenRouting = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenReason = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            int appliedRoutingRows = 0;
            int missingRoutingRows = 0;
            int appliedReasonRows = 0;
            int missingReasonRows = 0;
            int total = Math.Max(orgTable.Rows.Count, 1);

            for (int i = 0; i < orgTable.Rows.Count; i++)
            {
                DataRow row = orgTable.Rows[i];
                string modelName = row.Field<string>(CONSTANT.MATERIALNAME.NEW) ?? string.Empty;
                string processCode = row.Field<string>(CONSTANT.PROCESSCODE.NEW) ?? string.Empty;
                string processName = row.Field<string>(CONSTANT.PROCESSNAME.NEW) ?? string.Empty;
                string ngName = row.Field<string>(CONSTANT.NGNAME.NEW) ?? string.Empty;

                string modelKey = NormalizeCompareValue(modelName);
                string processCodeKey = NormalizeCompareValue(processCode);
                string processNameKey = NormalizeCompareValue(processName, normalize: true);
                string ngNameKey = NormalizeCompareValue(ngName, normalize: true);

                bool hasRouting = routingKeys.Contains((modelKey, processCodeKey, processNameKey));
                bool hasReason = reasonKeys.Contains((processNameKey, ngNameKey));

                if (hasRouting)
                {
                    appliedRoutingRows++;
                }
                else
                {
                    missingRoutingRows++;
                    string routingDedupeKey = $"{modelKey}|{processCodeKey}|{processNameKey}";
                    if (seenRouting.Add(routingDedupeKey))
                    {
                        missingRoutingItems.Add(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            [CONSTANT.MATERIALNAME.NEW] = modelName,
                            ["ProcessCode"] = processCode,
                            ["ProcessName"] = processName,
                            ["ProcessType"] = string.Empty
                        });
                    }
                }

                if (hasReason)
                {
                    appliedReasonRows++;
                }
                else
                {
                    missingReasonRows++;
                    string reasonDedupeKey = $"{processNameKey}|{ngNameKey}";
                    if (seenReason.Add(reasonDedupeKey))
                    {
                        missingReasonItems.Add(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["processName"] = processName,
                            ["NgName"] = ngName,
                            ["Reason"] = string.Empty
                        });
                    }
                }

                if (i == orgTable.Rows.Count - 1 || (i + 1) % LargeScanProgressLogInterval == 0)
                {
                    int percent = 35 + (int)(65d * (i + 1) / total);
                    progress?.Report((percent, $"Scanning Routing/Reason candidates... {i + 1:N0}/{orgTable.Rows.Count:N0}"));
                }
            }

            return new MappingScanResult
            {
                MissingRoutingItems = missingRoutingItems
                    .OrderBy(item => item.TryGetValue(CONSTANT.MATERIALNAME.NEW, out string? currentModelName) ? currentModelName : string.Empty)
                    .ThenBy(item => item["ProcessCode"])
                    .ThenBy(item => item["ProcessName"])
                    .ToList(),
                MissingReasonItems = missingReasonItems
                    .OrderBy(item => item["processName"])
                    .ThenBy(item => item["NgName"])
                    .ToList(),
                TotalOriginalRows = orgTable.Rows.Count,
                RoutingTableRows = routingTableRows,
                ReasonTableRows = reasonTableRows,
                AppliedRoutingRows = appliedRoutingRows,
                MissingRoutingRows = missingRoutingRows,
                AppliedReasonRows = appliedReasonRows,
                MissingReasonRows = missingReasonRows
            };
        }

        private void ResetLoadMappingStatus()
        {
            CT_LB_LOAD_ROUTING_TABLE_COUNT.Text = "0";
            CT_LB_LOAD_REASON_TABLE_COUNT.Text = "0";
            CT_LB_LOAD_MISSING_ROUTING_COUNT.Text = "0";
            CT_LB_LOAD_MISSING_REASON_COUNT.Text = "0";
            CT_LB_LOAD_MAPPING_STATUS_SUMMARY.Text = "Mapping status is not available yet.";
            CT_LB_LOAD_MAPPING_STATUS_HINT.Text = "Load DB 완료 후 자동으로 Mapping 상태를 계산합니다.";
        }

        private async Task RefreshLoadMappingStatusAsync()
        {
            ResetLoadMappingStatus();

            if (string.IsNullOrWhiteSpace(PathSelectedDB) || !File.Exists(PathSelectedDB))
            {
                return;
            }

            CT_LB_LOAD_MAPPING_STATUS_HINT.Text = "Mapping 상태 계산 중...";
            UpdateLoadDbProgress(0, 88, "Scanning Routing/Reason mapping status...");
            await ScanAllMissingMappingsAsync(showProgressPopup: false);
            UpdateLoadDbProgress(100, 96, "Mapping status calculated.");
            UpdateLoadMappingStatusUi();
        }

        private void UpdateLoadMappingStatusUi()
        {
            if (_mappingScanResult == null)
            {
                ResetLoadMappingStatus();
                return;
            }

            CT_LB_LOAD_ROUTING_TABLE_COUNT.Text = _mappingScanResult.RoutingTableRows.ToString("N0");
            CT_LB_LOAD_REASON_TABLE_COUNT.Text = _mappingScanResult.ReasonTableRows.ToString("N0");
            CT_LB_LOAD_MISSING_ROUTING_COUNT.Text = _mappingScanResult.MissingRoutingRows.ToString("N0");
            CT_LB_LOAD_MISSING_REASON_COUNT.Text = _mappingScanResult.MissingReasonRows.ToString("N0");
            CT_LB_LOAD_MAPPING_STATUS_HINT.Text = "현재 DB 기준 자동 계산된 Mapping 상태입니다.";
            CT_LB_LOAD_MAPPING_STATUS_SUMMARY.Text =
                $"OriginalTable {_mappingScanResult.TotalOriginalRows:N0} rows 기준으로 " +
                $"Routing 미매핑 {_mappingScanResult.MissingRoutingRows:N0}건, " +
                $"Reason 미매핑 {_mappingScanResult.MissingReasonRows:N0}건입니다.";
        }

        private void UpdateLoadDbProgress(int taskProgress, int totalProgress, string message)
        {
            SetProgressPopupState("Load DB", message);
            UpdateProgressBars(taskProgress, totalProgress);
            LogLoadDbProgressIfNeeded(totalProgress, message);
        }

        private void ResetLoadDbProgressLogging()
        {
            _lastLoadDbLogBucket = -1;
            _lastLoadDbLogMessage = string.Empty;
        }

        private void InvalidateDatabaseSummaryCache()
        {
            _cachedSummaryDbPath = string.Empty;
            _cachedSummaryLastWriteTimeUtc = DateTime.MinValue;
            _cachedOriginalTableDateRange = string.Empty;
            _cachedDbLastUpdatedAt = string.Empty;
        }

        private void LogLoadDbProgressIfNeeded(int totalProgress, string message)
        {
            int safeTotal = Math.Clamp(totalProgress, 0, 100);
            int bucket = safeTotal / 5;
            bool messageChanged = !string.Equals(_lastLoadDbLogMessage, message, StringComparison.Ordinal);
            bool bucketChanged = bucket != _lastLoadDbLogBucket;

            if (!messageChanged && !bucketChanged)
            {
                return;
            }

            _lastLoadDbLogBucket = bucket;
            _lastLoadDbLogMessage = message;
            clLogger.LogImportant($"Load DB progress: {safeTotal}% - {message}");
        }

        private (int AppliedRoutingRows, int MissingRoutingRows, int AppliedReasonRows, int MissingReasonRows) CalculateMappingRowCounts(
            clSQLFileIO sql,
            IProgress<(int Percent, string Message)>? progress = null)
        {
            if (!sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ORG))
            {
                return (0, 0, 0, 0);
            }

            DataTable orgTable = sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ORG);
            DataTable routingTable = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                ? sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                : new DataTable();
            DataTable reasonTable = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.REASON)
                ? sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.REASON)
                : new DataTable();
            string routingModelColumnName = GetRoutingModelColumnName(routingTable);

            var routingKeys = new HashSet<(string ModelName, string ProcessCode, string ProcessName)>();
            foreach (DataRow row in routingTable.Rows)
            {
                routingKeys.Add((
                    NormalizeCompareValue(row[routingModelColumnName]?.ToString()),
                    NormalizeCompareValue(row["ProcessCode"]?.ToString()),
                    NormalizeCompareValue(row["ProcessName"]?.ToString(), normalize: true)));
            }

            var reasonKeys = new HashSet<(string ProcessName, string NgName)>();
            foreach (DataRow row in reasonTable.Rows)
            {
                reasonKeys.Add((
                    NormalizeCompareValue(row["processName"]?.ToString(), normalize: true),
                    NormalizeCompareValue(row["NgName"]?.ToString(), normalize: true)));
            }

            int appliedRoutingRows = 0;
            int missingRoutingRows = 0;
            int appliedReasonRows = 0;
            int missingReasonRows = 0;
            int total = Math.Max(orgTable.Rows.Count, 1);

            for (int i = 0; i < orgTable.Rows.Count; i++)
            {
                DataRow row = orgTable.Rows[i];
                string modelName = row.Field<string>(CONSTANT.MATERIALNAME.NEW) ?? string.Empty;
                string processCode = row.Field<string>(CONSTANT.PROCESSCODE.NEW) ?? string.Empty;
                string processName = row.Field<string>(CONSTANT.PROCESSNAME.NEW) ?? string.Empty;
                string ngName = row.Field<string>(CONSTANT.NGNAME.NEW) ?? string.Empty;

                bool hasRouting = routingKeys.Contains((
                    NormalizeCompareValue(modelName),
                    NormalizeCompareValue(processCode),
                    NormalizeCompareValue(processName, normalize: true)));
                bool hasReason = reasonKeys.Contains((
                    NormalizeCompareValue(processName, normalize: true),
                    NormalizeCompareValue(ngName, normalize: true)));

                if (hasRouting)
                {
                    appliedRoutingRows++;
                }
                else
                {
                    missingRoutingRows++;
                }

                if (hasReason)
                {
                    appliedReasonRows++;
                }
                else
                {
                    missingReasonRows++;
                }

                if (i == orgTable.Rows.Count - 1 || (i + 1) % LargeScanProgressLogInterval == 0)
                {
                    int percent = 75 + (int)(25d * (i + 1) / total);
                    progress?.Report((percent, $"Counting mapped rows... {i + 1:N0}/{orgTable.Rows.Count:N0}"));
                }
            }

            return (appliedRoutingRows, missingRoutingRows, appliedReasonRows, missingReasonRows);
        }

        private static string GetRoutingModelColumnName(DataTable routingTable)
        {
            if (routingTable.Columns.Contains("모델명"))
            {
                return "모델명";
            }

            if (routingTable.Columns.Contains(CONSTANT.MATERIALNAME.NEW))
            {
                return CONSTANT.MATERIALNAME.NEW;
            }

            throw new InvalidOperationException("RoutingTable does not contain a model name column.");
        }

        private void UpdateMappingSummaryText()
        {
            if (_mappingScanResult == null)
            {
                CT_LB_MAPPING_SUMMARY.Text = "Scan result is not available.";
                return;
            }

            int currentListRows = _currentMappingTableKind == MappingTableKind.Routing
                ? _mappingScanResult.MissingRoutingItems.Count
                : _mappingScanResult.MissingReasonItems.Count;

            CT_LB_MAPPING_SUMMARY.Text =
                $"Original rows        : {_mappingScanResult.TotalOriginalRows:N0}\n" +
                $"Routing mapped       : {_mappingScanResult.AppliedRoutingRows:N0}\n" +
                $"Routing unmapped     : {_mappingScanResult.MissingRoutingRows:N0}\n" +
                $"Reason mapped        : {_mappingScanResult.AppliedReasonRows:N0}\n" +
                $"Reason unmapped      : {_mappingScanResult.MissingReasonRows:N0}\n" +
                $"Current list rows    : {currentListRows:N0}";
        }

        private DataTable BuildCurrentMissingMappingTable()
        {
            IEnumerable<IReadOnlyDictionary<string, string>> items = _currentMappingTableKind == MappingTableKind.Routing
                ? _cachedMissingRoutingItems
                : _cachedMissingReasonItems;

            var itemList = items.ToList();
            var table = new DataTable();
            if (itemList.Count == 0)
            {
                return table;
            }

            foreach (string column in itemList.SelectMany(item => item.Keys).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                table.Columns.Add(column, typeof(string));
            }

            foreach (var item in itemList)
            {
                DataRow row = table.NewRow();
                foreach (DataColumn column in table.Columns)
                {
                    row[column.ColumnName] = item.TryGetValue(column.ColumnName, out string? value) ? value : string.Empty;
                }

                table.Rows.Add(row);
            }

            return table;
        }

        private static void ExportDataTableToTxt(DataTable table, string filePath)
        {
            using var writer = new StreamWriter(filePath, false, Encoding.UTF8);
            string[] headers = table.Columns.Cast<DataColumn>()
                .Select(column => column.ColumnName)
                .ToArray();
            writer.WriteLine(string.Join("\t", headers));

            foreach (DataRow row in table.Rows)
            {
                string[] values = table.Columns.Cast<DataColumn>()
                    .Select(column => row[column]?.ToString()?.Replace("\r", " ").Replace("\n", " ").Trim() ?? string.Empty)
                    .ToArray();
                writer.WriteLine(string.Join("\t", values));
            }
        }

        private void InsertMappingRowToDatabase(IReadOnlyList<string> columns, IReadOnlyList<string> values, string tableLabel)
        {
            if (columns.Count != values.Count)
            {
                throw new InvalidOperationException($"Column count does not match value count for {tableLabel}.");
            }

            string tableName = GetCurrentMappingTableName();
            using var sql = new clSQLFileIO(PathSelectedDB);
            if (!sql.IsTableExist(tableName))
            {
                throw new InvalidOperationException($"{tableLabel} table '{tableName}' does not exist in the current DB.");
            }

            var data = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < columns.Count; i++)
            {
                data[columns[i]] = NormalizeMappingValue(tableLabel, columns[i], values[i]);
            }

            if (TryUpdateExistingMappingRow(sql, tableName, data, tableLabel))
            {
                clLogger.Log($"{tableLabel} row already existed. Updated DB table '{tableName}'.");
                return;
            }

            sql.InsertData(tableName, data);
            clLogger.Log($"{tableLabel} row inserted into DB table '{tableName}'.");
        }

        private static string NormalizeMappingValue(string tableLabel, string columnName, string rawValue)
        {
            string sanitized = rawValue.Replace("\r", " ").Replace("\n", " ").Trim();

            if (tableLabel.Equals("Routing", StringComparison.OrdinalIgnoreCase))
            {
                if (columnName.Equals("ProcessName", StringComparison.OrdinalIgnoreCase) ||
                    columnName.Equals("ProcessType", StringComparison.OrdinalIgnoreCase))
                {
                    return CONSTANT.Normalize(sanitized);
                }
            }
            else if (tableLabel.Equals("Reason", StringComparison.OrdinalIgnoreCase))
            {
                if (columnName.Equals("processName", StringComparison.OrdinalIgnoreCase) ||
                    columnName.Equals("NgName", StringComparison.OrdinalIgnoreCase))
                {
                    return CONSTANT.Normalize(sanitized);
                }
            }

            return sanitized;
        }

        private IReadOnlyList<IReadOnlyDictionary<string, string>> GetMissingItemsForDialog(string tableLabel)
        {
            bool isRouting = tableLabel.Equals("Routing", StringComparison.OrdinalIgnoreCase);
            if (isRouting && _hasScannedMissingRoutingItems)
            {
                return _cachedMissingRoutingItems;
            }

            if (!isRouting && _hasScannedMissingReasonItems)
            {
                return _cachedMissingReasonItems;
            }

            var result = GetMissingItemsFromOrgTable();
            if (result != null)
            {
                if (isRouting)
                {
                    return result.Value.MissingRouting
                        .Select(item => (IReadOnlyDictionary<string, string>)new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["모델명"] = item.ModelName?.ToString() ?? string.Empty,
                            ["ProcessCode"] = item.ProcessCode?.ToString() ?? string.Empty,
                            ["ProcessName"] = item.ProcessName?.ToString() ?? string.Empty,
                            ["ProcessType"] = string.Empty
                        })
                        .ToList();
                }

                return result.Value.MissingReason
                    .Select(item => (IReadOnlyDictionary<string, string>)new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["processName"] = item.ProcessName?.ToString() ?? string.Empty,
                        ["NgName"] = item.NgName?.ToString() ?? string.Empty,
                        ["Reason"] = string.Empty
                    })
                    .ToList();
            }

            string missingTableName = isRouting
                ? CONSTANT.OPTION_TABLE_NAME.MISSING_ROUTING
                : CONSTANT.OPTION_TABLE_NAME.MISSING_REASON;

            using var sql = new clSQLFileIO(PathSelectedDB);
            if (!sql.IsTableExist(missingTableName))
            {
                return Array.Empty<IReadOnlyDictionary<string, string>>();
            }

            DataTable table = sql.LoadTable(missingTableName);
            return table.AsEnumerable()
                .Select(row => CreateMissingItemDialogValues(tableLabel, table, row))
                .Cast<IReadOnlyDictionary<string, string>>()
                .ToList();
        }

        private static IReadOnlyDictionary<string, string> CreateMissingItemDialogValues(string tableLabel, DataTable table, DataRow row)
        {
            var values = table.Columns.Cast<DataColumn>()
                .ToDictionary(
                    column => column.ColumnName,
                    column => row[column]?.ToString() ?? string.Empty,
                    StringComparer.OrdinalIgnoreCase);

            if (tableLabel.Equals("Routing", StringComparison.OrdinalIgnoreCase))
            {
                AddAliasIfMissing(values, "모델명", CONSTANT.MATERIALNAME.NEW);
                AddAliasIfMissing(values, "ProcessCode", CONSTANT.PROCESSCODE.NEW);
                AddAliasIfMissing(values, "ProcessName", CONSTANT.PROCESSNAME.NEW);
            }
            else
            {
                AddAliasIfMissing(values, "processName", CONSTANT.PROCESSNAME.NEW);
                AddAliasIfMissing(values, "NgName", CONSTANT.NGNAME.NEW);
            }

            return values;
        }

        private static void AddAliasIfMissing(IDictionary<string, string> values, string aliasColumn, string sourceColumn)
        {
            if (values.ContainsKey(aliasColumn))
            {
                return;
            }

            if (values.TryGetValue(sourceColumn, out string? value))
            {
                values[aliasColumn] = value;
            }
        }

        private static bool TryUpdateExistingMappingRow(clSQLFileIO sql, string tableName, Dictionary<string, object> data, string tableLabel)
        {
            string[] keyColumns = tableLabel.Equals("Routing", StringComparison.OrdinalIgnoreCase)
                ? new[] { "모델명", "ProcessCode", "ProcessName" }
                : new[] { "processName", "NgName" };

            if (keyColumns.Any(column => !data.ContainsKey(column)))
            {
                return false;
            }

            var conditions = new List<string>();
            foreach (string column in keyColumns)
            {
                string value = data[column]?.ToString()?.Replace("'", "''") ?? string.Empty;
                conditions.Add($"[{column}] = '{value}'");
            }

            string condition = string.Join(" AND ", conditions);
            using var existing = sql.Reader.Read(tableName, data
                .Where(kv => keyColumns.Contains(kv.Key))
                .Select(kv => new HashSet<(string ColumnName, string ColumnItem)>
                {
                    (kv.Key, kv.Value?.ToString() ?? string.Empty)
                })
                .ToList());

            if (existing.Rows.Count == 0)
            {
                return false;
            }

            sql.UpdateData(tableName, data, condition);
            return true;
        }

        private async Task ImportCurrentMappingFromFileAsync()
        {
            if (string.IsNullOrWhiteSpace(PathSelectedDB) || _currentMappingTableKind == MappingTableKind.None)
            {
                return;
            }

            bool isRouting = _currentMappingTableKind == MappingTableKind.Routing;
            string currentPath = GetCurrentMappingFilePath();
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                Title = isRouting ? "Select Routing File" : "Select Reason File",
                InitialDirectory = Path.GetDirectoryName(currentPath)
            };

            if (!ShowCommonDialog(openFileDialog))
            {
                return;
            }

            string selectedPath = openFileDialog.FileName;

            if (isRouting)
            {
                _processTypeCsvPath = selectedPath;
                await Task.Run(() =>
                {
                    var dataProcessor = new clDataProcessor(PathSelectedDB, selectedPath, _reasonCsvPath);
                    try
                    {
                        dataProcessor.UpdateProcessTypeTable(PathSelectedDB, selectedPath);
                    }
                    finally
                    {
                        dataProcessor.Dispose();
                    }
                });
            }
            else
            {
                _reasonCsvPath = selectedPath;
                await Task.Run(() =>
                {
                    var dataProcessor = new clDataProcessor(PathSelectedDB, _processTypeCsvPath, selectedPath);
                    try
                    {
                        dataProcessor.UpdateReasonTable(PathSelectedDB, selectedPath);
                    }
                    finally
                    {
                        dataProcessor.Dispose();
                    }
                });
            }

            InvalidateDebugGroupTableCache();
            await ScanAllMissingMappingsAsync(showProgressPopup: true);
            await RefreshMappingTableViewAsync();
        }

        private async void CT_BT_MAPPING_EXPORT_Click(object sender, RoutedEventArgs e)
        {
            if (_currentMappingTableKind == MappingTableKind.None)
            {
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                Title = _currentMappingTableKind == MappingTableKind.Routing ? "Export Unmapped Routing" : "Export Unmapped Reason",
                FileName = _currentMappingTableKind == MappingTableKind.Routing ? "UnmappedRouting.txt" : "UnmappedReason.txt"
            };

            if (!ShowCommonDialog(saveFileDialog))
            {
                return;
            }

            try
            {
                DataTable exportTable = _currentMappingTableData ?? BuildCurrentMissingMappingTable();
                await Task.Run(() => ExportDataTableToTxt(exportTable, saveFileDialog.FileName));
                ShowInfoMessage($"Export completed.\n{saveFileDialog.FileName}", "Export Txt");
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error exporting mapping txt");
                ShowErrorMessage($"Failed to export txt.\n{ex.Message}", "Error");
            }
        }

        private async void CT_BT_MAPPING_IMPORT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await ImportCurrentMappingFromFileAsync();
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error importing mapping file");
                ShowErrorMessage($"Failed to import mapping file.\n{ex.Message}", "Error");
            }
        }

        private void SetMappingButtonsEnabled(bool isEnabled)
        {
            CT_BT_MAPPING_EXPORT.IsEnabled = isEnabled;
            CT_BT_MAPPING_IMPORT.IsEnabled = isEnabled;
        }

        private void UpdateMappingSelectorState()
        {
            if (_currentMappingTableKind == MappingTableKind.Routing)
            {
                CT_BT_MAPPING_SHOW_ROUTING.Style = (Style)FindResource("AccentButton");
                CT_BT_MAPPING_SHOW_REASON.Style = (Style)FindResource("ModernButton");
            }
            else if (_currentMappingTableKind == MappingTableKind.Reason)
            {
                CT_BT_MAPPING_SHOW_ROUTING.Style = (Style)FindResource("ModernButton");
                CT_BT_MAPPING_SHOW_REASON.Style = (Style)FindResource("AccentButton");
            }
            else
            {
                CT_BT_MAPPING_SHOW_ROUTING.Style = (Style)FindResource("ModernButton");
                CT_BT_MAPPING_SHOW_REASON.Style = (Style)FindResource("ModernButton");
            }
        }

        private async void CT_BT_GETROUTING_Click(object sender, RoutedEventArgs e)
        {
            var credentials = FormSettingBMESWindow.EnsureLoadedInfo();
            if (string.IsNullOrWhiteSpace(credentials.LoginID) || string.IsNullOrWhiteSpace(credentials.Password))
            {
                MessageBox.Show("자격 증명이 설정되지 않았습니다. 먼저 BMES 설정을 완료해주세요.", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string id = credentials.LoginID;
            string password = credentials.Password;

            var sfd = new SaveFileDialog
            {
                Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                Title = "Save Routing Data As",
                FileName = "routing.txt"
            };

            if (!ShowCommonDialog(sfd))
                return;

            CT_BT_GETROUTING.IsEnabled = false;
            CT_BT_GETROUTING.Content = "Fetching...";

            try
            {
                var fetcher = new clFetchRoutingData(id, password);
                string[] columns = { "MAKTX", "VLSCH", "VLSCH_TX", "LGUBN_TX" };
                await fetcher.FetchAndSaveToTxtAsync(sfd.FileName, columns);
            }
            finally
            {
                CT_BT_GETROUTING.IsEnabled = true;
                CT_BT_GETROUTING.Content = "Get Routing";
            }
        }

        private void CT_CB_DEBUG_GROUP_DropDownOpened(object sender, EventArgs e)
        {
            RefreshDebugGroupDropdown();
        }

        private void CT_CB_DEBUG_GROUP_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingDebugCombo) return;
            RefreshDebugProcessDropdown();
        }

        private void CT_CB_DEBUG_PROCESS_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingDebugCombo) return;
            RefreshDebugNgDropdown();
        }

        private void CT_CB_DEBUG_NG_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingDebugCombo) return;
            RefreshDebugDateDropdown();
        }

        private void RefreshDebugGroupDropdown()
        {
            if (string.IsNullOrWhiteSpace(PathSelectedDB) || !File.Exists(PathSelectedDB))
            {
                InvalidateDebugGroupTableCache();
                return;
            }

            string previous = CT_CB_DEBUG_GROUP.Text?.Trim() ?? "";

            try
            {
                _isUpdatingDebugCombo = true;
                InvalidateDebugGroupTableCache();

                var groupMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var displays = new List<string>();

                using (var conn = new SQLiteConnection($"Data Source={PathSelectedDB};"))
                using (var cmd = new SQLiteCommand(
                    "SELECT name FROM sqlite_master " +
                    "WHERE type='table' AND name LIKE 'Group_%' AND name NOT LIKE '%__MODELS' " +
                    "ORDER BY name;", conn))
                {
                    conn.Open();
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string tableName = reader["name"]?.ToString()?.Trim() ?? "";
                            if (string.IsNullOrWhiteSpace(tableName))
                            {
                                continue;
                            }

                            string display = tableName.StartsWith("Group_", StringComparison.OrdinalIgnoreCase)
                                ? tableName.Substring("Group_".Length)
                                : tableName;

                            if (string.IsNullOrWhiteSpace(display))
                            {
                                continue;
                            }

                            if (!groupMap.ContainsKey(display))
                            {
                                groupMap[display] = tableName;
                                displays.Add(display);
                            }
                        }
                    }
                }

                _debugGroupDisplayToTable.Clear();
                foreach (var item in groupMap)
                {
                    _debugGroupDisplayToTable[item.Key] = item.Value;
                }

                CT_CB_DEBUG_GROUP.ItemsSource = displays;

                if (displays.Count == 0)
                {
                    CT_CB_DEBUG_GROUP.Text = "";
                    CT_CB_DEBUG_PROCESS.ItemsSource = null;
                    CT_CB_DEBUG_NG.ItemsSource = null;
                    CT_CB_DEBUG_DATE.ItemsSource = null;
                    InvalidateDebugGroupTableCache();
                    return;
                }

                if (!string.IsNullOrWhiteSpace(previous) && displays.Contains(previous))
                {
                    CT_CB_DEBUG_GROUP.SelectedItem = previous;
                }
                else if (CT_CB_DEBUG_GROUP.SelectedItem == null ||
                         !displays.Contains(CT_CB_DEBUG_GROUP.SelectedItem?.ToString() ?? ""))
                {
                    CT_CB_DEBUG_GROUP.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "RefreshDebugGroupDropdown");
            }
            finally
            {
                _isUpdatingDebugCombo = false;
            }

            var selectedGroupTable = LoadSelectedGroupTable();
            RefreshDebugProcessDropdown(selectedGroupTable);
        }

        private void RefreshDebugProcessDropdown(DataTable? selectedGroupTable = null)
        {
            string previous = CT_CB_DEBUG_PROCESS.Text?.Trim() ?? "";
            DataTable? data = selectedGroupTable ?? LoadSelectedGroupTable();

            try
            {
                _isUpdatingDebugCombo = true;

                var processes = new List<string>();

                if (data != null && data.Columns.Contains(CONSTANT.PROCESSNAME.NEW))
                {
                    processes = data.AsEnumerable()
                        .Select(r => r[CONSTANT.PROCESSNAME.NEW]?.ToString()?.Trim() ?? "")
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(v => v)
                        .ToList();
                }

                CT_CB_DEBUG_PROCESS.ItemsSource = processes;

                if (processes.Count == 0)
                {
                    CT_CB_DEBUG_PROCESS.Text = "";
                    CT_CB_DEBUG_NG.ItemsSource = null;
                    CT_CB_DEBUG_DATE.ItemsSource = null;
                    return;
                }

                if (!string.IsNullOrWhiteSpace(previous) && processes.Contains(previous))
                {
                    CT_CB_DEBUG_PROCESS.SelectedItem = previous;
                }
                else
                {
                    CT_CB_DEBUG_PROCESS.SelectedIndex = 0;
                }
            }
            finally
            {
                _isUpdatingDebugCombo = false;
            }

            RefreshDebugNgDropdown(data);
        }

        private void RefreshDebugNgDropdown(DataTable? selectedGroupTable = null)
        {
            string previous = CT_CB_DEBUG_NG.Text?.Trim() ?? "";
            string selectedProcess = CT_CB_DEBUG_PROCESS.Text?.Trim() ?? "";
            DataTable? data = selectedGroupTable ?? LoadSelectedGroupTable();

            try
            {
                _isUpdatingDebugCombo = true;

                var ngNames = new List<string>();

                if (data != null &&
                    data.Columns.Contains(CONSTANT.PROCESSNAME.NEW) &&
                    data.Columns.Contains(CONSTANT.NGNAME.NEW))
                {
                    string normalizedProcess = CONSTANT.Normalize(selectedProcess);
                    var rows = data.AsEnumerable();

                    if (!string.IsNullOrWhiteSpace(normalizedProcess))
                    {
                        rows = rows.Where(r =>
                            CONSTANT.Normalize(r.Field<string>(CONSTANT.PROCESSNAME.NEW) ?? "") == normalizedProcess);
                    }

                    ngNames = rows
                        .Select(r => r[CONSTANT.NGNAME.NEW]?.ToString()?.Trim() ?? "")
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(v => v)
                        .ToList();
                }

                CT_CB_DEBUG_NG.ItemsSource = ngNames;

                if (ngNames.Count == 0)
                {
                    CT_CB_DEBUG_NG.Text = "";
                    CT_CB_DEBUG_DATE.ItemsSource = null;
                    return;
                }

                if (!string.IsNullOrWhiteSpace(previous) && ngNames.Contains(previous))
                {
                    CT_CB_DEBUG_NG.SelectedItem = previous;
                }
                else
                {
                    CT_CB_DEBUG_NG.SelectedIndex = 0;
                }
            }
            finally
            {
                _isUpdatingDebugCombo = false;
            }

            RefreshDebugDateDropdown(data);
        }

        private void RefreshDebugDateDropdown(DataTable? selectedGroupTable = null)
        {
            string previous = CT_CB_DEBUG_DATE.Text?.Trim() ?? "";
            string selectedProcess = CT_CB_DEBUG_PROCESS.Text?.Trim() ?? "";
            string selectedNg = CT_CB_DEBUG_NG.Text?.Trim() ?? "";

            try
            {
                _isUpdatingDebugCombo = true;

                var data = selectedGroupTable ?? LoadSelectedGroupTable();
                var dates = new List<string>();

                if (data != null &&
                    data.Columns.Contains(CONSTANT.PROCESSNAME.NEW) &&
                    data.Columns.Contains(CONSTANT.NGNAME.NEW) &&
                    data.Columns.Contains(CONSTANT.PRODUCT_DATE.NEW))
                {
                    string normalizedProcess = CONSTANT.Normalize(selectedProcess);
                    string normalizedNg = CONSTANT.Normalize(selectedNg);

                    var rows = data.AsEnumerable().Where(r =>
                        (string.IsNullOrWhiteSpace(normalizedProcess) ||
                         CONSTANT.Normalize(r.Field<string>(CONSTANT.PROCESSNAME.NEW) ?? "") == normalizedProcess) &&
                        (string.IsNullOrWhiteSpace(normalizedNg) ||
                         CONSTANT.Normalize(r.Field<string>(CONSTANT.NGNAME.NEW) ?? "") == normalizedNg));

                    dates = rows
                        .Select(r => r[CONSTANT.PRODUCT_DATE.NEW]?.ToString()?.Trim() ?? "")
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderByDescending(v => ParseDateOrMin(v))
                        .ThenByDescending(v => v)
                        .ToList();
                }

                CT_CB_DEBUG_DATE.ItemsSource = dates;

                if (dates.Count == 0)
                {
                    if (string.IsNullOrWhiteSpace(previous))
                    {
                        CT_CB_DEBUG_DATE.Text = DateTime.Today.ToString("yyyy-MM-dd");
                    }
                    return;
                }

                if (!string.IsNullOrWhiteSpace(previous) && dates.Contains(previous))
                {
                    CT_CB_DEBUG_DATE.SelectedItem = previous;
                }
                else
                {
                    CT_CB_DEBUG_DATE.SelectedIndex = 0;
                }
            }
            finally
            {
                _isUpdatingDebugCombo = false;
            }
        }

        private DataTable? LoadSelectedGroupTable()
        {
            string selectedGroup = CT_CB_DEBUG_GROUP.Text?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(selectedGroup) || string.IsNullOrWhiteSpace(PathSelectedDB))
            {
                InvalidateDebugGroupTableCache();
                return null;
            }

            string tableName = ResolveGroupTableName(selectedGroup);
            if (string.IsNullOrWhiteSpace(tableName))
            {
                InvalidateDebugGroupTableCache();
                return null;
            }

            if (!string.IsNullOrWhiteSpace(_cachedDebugGroupTableName) &&
                string.Equals(_cachedDebugGroupTableName, tableName, StringComparison.OrdinalIgnoreCase) &&
                _cachedDebugGroupTable is not null)
            {
                return _cachedDebugGroupTable;
            }

            using (var sql = new clSQLFileIO(PathSelectedDB))
            {
                if (!sql.IsTableExist(tableName))
                {
                    InvalidateDebugGroupTableCache();
                    return null;
                }

                _cachedDebugGroupTable = sql.LoadTable(tableName);
                _cachedDebugGroupTableName = tableName;
                return _cachedDebugGroupTable;
            }
        }

        private string ResolveGroupTableName(string groupDisplayName)
        {
            if (string.IsNullOrWhiteSpace(groupDisplayName))
            {
                return "";
            }

            if (_debugGroupDisplayToTable.TryGetValue(groupDisplayName, out string tableName))
            {
                return tableName;
            }

            var groupInfo = new clModelGroupData
            {
                GroupName = groupDisplayName,
                ModelList = new List<string>()
            };

            return CONSTANT.GetGroupTableName(groupInfo);
        }

        private static DateTime ParseDateOrMin(string value)
        {
            return DateTime.TryParse(value, out DateTime dt) ? dt.Date : DateTime.MinValue;
        }

        private void CT_BT_DEBUG_RATE_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(PathSelectedDB))
            {
                MessageBox.Show("Please load BMES data first.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (CT_CB_DEBUG_GROUP.Items.Count == 0)
            {
                RefreshDebugGroupDropdown();
            }

            string processName = CT_CB_DEBUG_PROCESS.Text?.Trim() ?? "";
            string ngName = CT_CB_DEBUG_NG.Text?.Trim() ?? "";
            string productDate = CT_CB_DEBUG_DATE.Text?.Trim() ?? "";
            string groupName = CT_CB_DEBUG_GROUP.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(processName) ||
                string.IsNullOrWhiteSpace(ngName) ||
                string.IsNullOrWhiteSpace(productDate) ||
                string.IsNullOrWhiteSpace(groupName))
            {
                MessageBox.Show(
                    "ProcessName, NgName, Date, GroupName을 모두 입력해주세요.",
                    "Input Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            CT_BT_DEBUG_RATE.IsEnabled = false;
            CT_BT_DEBUG_RATE.Content = "Checking...";

            try
            {
                var groupTableMaker = new clGroupTableMaker(PathSelectedDB);
                try
                {
                    var debug = groupTableMaker.GetDebugRate(groupName, processName, ngName, productDate);
                    CT_TB_DEBUG_RESULT.Text = BuildDebugRateResultText(debug);
                }
                finally
                {
                    groupTableMaker.Dispose();
                }
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "CT_BT_DEBUG_RATE_Click");
                CT_TB_DEBUG_RESULT.Text = $"Error: {ex.Message}";
                MessageBox.Show($"Debug lookup failed:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                CT_BT_DEBUG_RATE.IsEnabled = true;
                CT_BT_DEBUG_RATE.Content = "Check";
            }
        }

        private static string BuildDebugRateResultText(clGroupDebugRateResult debug)
        {
            var sb = new StringBuilder();

            sb.AppendLine($"Table: {debug.GroupTableName}");
            sb.AppendLine($"Filter: ProcessName='{debug.ProcessName}', NgName='{debug.NGName}', Date='{debug.ProductDate}'");
            sb.AppendLine($"GroupRows={debug.MatchedRowCount}, Input={debug.InputQty:F4}, NG={debug.NGQty:F4}, NGRATE={debug.NGRate:F8} ({debug.NGRatePercent:F6}%)");

            if (!string.IsNullOrWhiteSpace(debug.SourceRowsMessage))
            {
                sb.AppendLine(debug.SourceRowsMessage);
            }

            sb.AppendLine($"SourceRows={debug.SourceRowCount}, SourceInput={debug.SourceInputQty:F4}, SourceNG={debug.SourceNGQty:F4}, SourceNGRATE={debug.SourceNGRate:F8}");

            if (debug.SourceRowCount == 0)
            {
                return sb.ToString().TrimEnd();
            }

            sb.AppendLine("----- Source Rows (from ProcTable) -----");
            int maxDisplay = 200;
            int printCount = Math.Min(debug.SourceRows.Count, maxDisplay);
            for (int i = 0; i < printCount; i++)
            {
                var row = debug.SourceRows[i];
                sb.AppendLine(
                    $"{i + 1}. Model={row.ModelName}, NGCode={row.NGCode}, Date={row.ProductDate}, Input={row.InputQty:F4}, NG={row.NGQty:F4}");
            }

            if (debug.SourceRows.Count > maxDisplay)
            {
                sb.AppendLine($"... ({debug.SourceRows.Count - maxDisplay} rows more)");
            }

            return sb.ToString().TrimEnd();
        }

        private void SaveUnmappedItems(string outputPath)
        {
            try
            {
                var result = GetMissingItemsFromOrgTable();
                if (result == null)
                {
                    return;
                }

                var missingReason = result.Value.MissingReason;
                var missingRouting = result.Value.MissingRouting;

                using (var sql = new clSQLFileIO(PathSelectedDB))
                {
                    SaveMissingItemsToDatabase(sql, missingReason, missingRouting);

                    if (missingReason.Count > 0)
                    {
                        string reasonPath = Path.Combine(outputPath, "MissingReason.txt");
                        var reasonLines = new List<string>();
                        reasonLines.Add("# Missing Reason Mappings");
                        reasonLines.Add("# Add these entries to Reason.txt file");
                        reasonLines.Add("# Format: processName\tNgName\tReason");
                        reasonLines.Add("");

                        foreach (var item in missingReason)
                        {
                            reasonLines.Add(
                                $"{NormalizeCompareValue(item.ProcessName?.ToString(), normalize: true)}\t" +
                                $"{NormalizeCompareValue(item.NgName?.ToString(), normalize: true)}\t");
                        }

                        File.WriteAllLines(reasonPath, reasonLines);
                        clLogger.Log($"  -> Saved MissingReason.txt with {missingReason.Count} unmapped items");
                    }
                    else
                    {
                        clLogger.Log($"  All Reason mappings found!");
                    }

                    if (missingRouting.Count > 0)
                    {
                        string routingPath = Path.Combine(outputPath, "MissingRouting.txt");
                        var routingLines = new List<string>();
                        routingLines.Add("# Missing Routing (ProcessType) Mappings");
                        routingLines.Add("# Add these entries to Routing.txt file");
                        routingLines.Add("# Format: 모델명\tProcessCode\tProcessName\tProcessType");
                        routingLines.Add("");

                        foreach (var item in missingRouting)
                        {
                            routingLines.Add(
                                $"{NormalizeCompareValue(item.ModelName?.ToString())}\t" +
                                $"{NormalizeCompareValue(item.ProcessCode?.ToString())}\t" +
                                $"{NormalizeCompareValue(item.ProcessName?.ToString(), normalize: true)}\t");
                        }

                        File.WriteAllLines(routingPath, routingLines);
                        clLogger.Log($"  -> Saved MissingRouting.txt with {missingRouting.Count} unmapped items");
                    }
                    else
                    {
                        clLogger.Log($"  All Routing mappings found!");
                    }

                    clLogger.Log($"Unmapped items check completed.");
                }
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error checking unmapped items");
            }
        }

        private void SaveMissingItemsToDatabaseOnly()
        {
            var result = GetMissingItemsFromOrgTable();
            if (result == null)
            {
                return;
            }

            using var sql = new clSQLFileIO(PathSelectedDB);
            SaveMissingItemsToDatabase(sql, result.Value.MissingReason, result.Value.MissingRouting);
            clLogger.Log("Missing mapping tables saved to DB.");
        }

        private (List<dynamic> MissingReason, List<dynamic> MissingRouting)? GetMissingItemsFromOrgTable()
        {
            if (string.IsNullOrEmpty(PathSelectedDB))
            {
                clLogger.LogWarning("Database path is not set. Cannot check unmapped items.");
                return null;
            }

            using var sql = new clSQLFileIO(PathSelectedDB);
            if (!sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ORG))
            {
                clLogger.LogWarning("OrginalTable does not exist. Cannot check unmapped items.");
                return null;
            }

            var orgTable = sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ORG);
            var routingTable = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                ? sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.ROUTING)
                : new DataTable();
            var reasonTable = sql.IsTableExist(CONSTANT.OPTION_TABLE_NAME.REASON)
                ? sql.LoadTable(CONSTANT.OPTION_TABLE_NAME.REASON)
                : new DataTable();

            var routingKeys = new HashSet<(string ModelName, string ProcessCode, string ProcessName)>();
            foreach (DataRow row in routingTable.Rows)
            {
                routingKeys.Add((
                    NormalizeCompareValue(row["모델명"]?.ToString()),
                    NormalizeCompareValue(row["ProcessCode"]?.ToString()),
                    NormalizeCompareValue(row["ProcessName"]?.ToString(), normalize: true)));
            }

            var reasonKeys = new HashSet<(string ProcessName, string NgName)>();
            foreach (DataRow row in reasonTable.Rows)
            {
                reasonKeys.Add((
                    NormalizeCompareValue(row["processName"]?.ToString(), normalize: true),
                    NormalizeCompareValue(row["NgName"]?.ToString(), normalize: true)));
            }

            var missingReason = orgTable.AsEnumerable()
                .Select(row => new
                {
                    ProcessName = row.Field<string>(CONSTANT.PROCESSNAME.NEW) ?? string.Empty,
                    NgName = row.Field<string>(CONSTANT.NGNAME.NEW) ?? string.Empty,
                    ProcessNameKey = NormalizeCompareValue(row.Field<string>(CONSTANT.PROCESSNAME.NEW), normalize: true),
                    NgNameKey = NormalizeCompareValue(row.Field<string>(CONSTANT.NGNAME.NEW), normalize: true)
                })
                .Where(item => !reasonKeys.Contains((item.ProcessNameKey, item.NgNameKey)))
                .Select(item => new
                {
                    item.ProcessName,
                    item.NgName
                })
                .DistinctBy(item => $"{NormalizeCompareValue(item.ProcessName, true)}|{NormalizeCompareValue(item.NgName, true)}")
                .OrderBy(x => x.ProcessName)
                .ThenBy(x => x.NgName)
                .Cast<dynamic>()
                .ToList();

            var missingRouting = orgTable.AsEnumerable()
                .Select(row => new
                {
                    ModelName = row.Field<string>(CONSTANT.MATERIALNAME.NEW) ?? string.Empty,
                    ProcessCode = row.Field<string>(CONSTANT.PROCESSCODE.NEW) ?? string.Empty,
                    ProcessName = row.Field<string>(CONSTANT.PROCESSNAME.NEW) ?? string.Empty,
                    ModelNameKey = NormalizeCompareValue(row.Field<string>(CONSTANT.MATERIALNAME.NEW)),
                    ProcessCodeKey = NormalizeCompareValue(row.Field<string>(CONSTANT.PROCESSCODE.NEW)),
                    ProcessNameKey = NormalizeCompareValue(row.Field<string>(CONSTANT.PROCESSNAME.NEW), normalize: true)
                })
                .Where(item => !routingKeys.Contains((item.ModelNameKey, item.ProcessCodeKey, item.ProcessNameKey)))
                .Select(item => new
                {
                    item.ModelName,
                    item.ProcessCode,
                    item.ProcessName
                })
                .DistinctBy(item => $"{NormalizeCompareValue(item.ModelName)}|{NormalizeCompareValue(item.ProcessCode)}|{NormalizeCompareValue(item.ProcessName, true)}")
                .OrderBy(x => x.ModelName)
                .ThenBy(x => x.ProcessCode)
                .ThenBy(x => x.ProcessName)
                .Cast<dynamic>()
                .ToList();

            return (missingReason, missingRouting);
        }

        private static string NormalizeCompareValue(string? value, bool normalize = false)
        {
            string text = value?.Trim() ?? string.Empty;
            return normalize ? CONSTANT.Normalize(text) : text;
        }

        private void SaveMissingItemsToDatabase(
            clSQLFileIO sql,
            IEnumerable<dynamic> missingReason,
            IEnumerable<dynamic> missingRouting)
        {
            var missingReasonTable = CreateMissingReasonTable(missingReason);
            RewriteTable(sql, CONSTANT.OPTION_TABLE_NAME.MISSING_REASON, missingReasonTable);
            clLogger.Log($"  -> Saved DB table {CONSTANT.OPTION_TABLE_NAME.MISSING_REASON} with {missingReasonTable.Rows.Count} rows");

            var missingRoutingTable = CreateMissingRoutingTable(missingRouting);
            RewriteTable(sql, CONSTANT.OPTION_TABLE_NAME.MISSING_ROUTING, missingRoutingTable);
            clLogger.Log($"  -> Saved DB table {CONSTANT.OPTION_TABLE_NAME.MISSING_ROUTING} with {missingRoutingTable.Rows.Count} rows");
        }

        private static void RewriteTable(clSQLFileIO sql, string tableName, DataTable table)
        {
            if (sql.IsTableExist(tableName))
            {
                sql.DropTable(tableName);
            }

            sql.Writer.Write(tableName, table);
        }

        private static DataTable CreateMissingReasonTable(IEnumerable<dynamic> missingReason)
        {
            var table = new DataTable(CONSTANT.OPTION_TABLE_NAME.MISSING_REASON);
            table.Columns.Add(CONSTANT.PROCESSNAME.NEW, typeof(string));
            table.Columns.Add(CONSTANT.NGNAME.NEW, typeof(string));

            foreach (var item in missingReason)
            {
                table.Rows.Add(
                    NormalizeCompareValue(item.ProcessName?.ToString(), normalize: true),
                    NormalizeCompareValue(item.NgName?.ToString(), normalize: true));
            }

            return table;
        }

        private static DataTable CreateMissingRoutingTable(IEnumerable<dynamic> missingRouting)
        {
            var table = new DataTable(CONSTANT.OPTION_TABLE_NAME.MISSING_ROUTING);
            table.Columns.Add(CONSTANT.MATERIALNAME.NEW, typeof(string));
            table.Columns.Add(CONSTANT.PROCESSCODE.NEW, typeof(string));
            table.Columns.Add(CONSTANT.PROCESSNAME.NEW, typeof(string));

            foreach (var item in missingRouting)
            {
                table.Rows.Add(
                    NormalizeCompareValue(item.ModelName?.ToString()),
                    NormalizeCompareValue(item.ProcessCode?.ToString()),
                    NormalizeCompareValue(item.ProcessName?.ToString(), normalize: true));
            }

            return table;
        }

        private string EnsureAiDirectory()
        {
            string directoryPath = Path.Combine(AppSettingsPathManager.GetModuleDirectory("DataMaker"), "AI");
            Directory.CreateDirectory(directoryPath);
            return directoryPath;
        }

        private async Task SelectAiTrainingFileAsync()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Text Files|*.txt;*.csv|All Files|*.*",
                Title = "Select AI Training Data"
            };

            if (!ShowCommonDialog(openFileDialog))
            {
                return;
            }

            _aiTrainingFilePath = openFileDialog.FileName;
            _aiTrainingRowCount = await CountNonEmptyLinesAsync(_aiTrainingFilePath);
            _aiStatusMessage = $"Training file selected: {Path.GetFileName(_aiTrainingFilePath)}";
            NotifyWebModuleSnapshotChanged();
        }

        private async Task SelectAiReasonTrainingFileAsync()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Text Files|*.txt;*.csv|All Files|*.*",
                Title = "Select AI Reason Training Data"
            };

            if (!ShowCommonDialog(openFileDialog))
            {
                return;
            }

            _aiReasonTrainingFilePath = openFileDialog.FileName;
            _aiReasonTrainingRowCount = await CountNonEmptyLinesAsync(_aiReasonTrainingFilePath);
            _aiReasonStatusMessage = $"Reason training file selected: {Path.GetFileName(_aiReasonTrainingFilePath)}";
            NotifyWebModuleSnapshotChanged();
        }

        private async Task SelectAiFeedbackFileAsync()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files|*.csv|Text Files|*.txt|All Files|*.*",
                Title = "Select AI Feedback CSV"
            };

            if (!ShowCommonDialog(openFileDialog))
            {
                return;
            }

            _aiFeedbackFilePath = openFileDialog.FileName;
            _aiFeedbackRowCount = await CountNonEmptyLinesAsync(_aiFeedbackFilePath);
            _aiStatusMessage = $"Feedback file selected: {Path.GetFileName(_aiFeedbackFilePath)}";
            NotifyWebModuleSnapshotChanged();
        }

        private async Task SelectAiReasonFeedbackFileAsync()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files|*.csv|Text Files|*.txt|All Files|*.*",
                Title = "Select AI Reason Feedback CSV"
            };

            if (!ShowCommonDialog(openFileDialog))
            {
                return;
            }

            _aiReasonFeedbackFilePath = openFileDialog.FileName;
            _aiReasonFeedbackRowCount = await CountNonEmptyLinesAsync(_aiReasonFeedbackFilePath);
            _aiReasonStatusMessage = $"Reason feedback file selected: {Path.GetFileName(_aiReasonFeedbackFilePath)}";
            NotifyWebModuleSnapshotChanged();
        }

        private async Task SelectAiInferenceFileAsync()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Text Files|*.txt;*.csv|All Files|*.*",
                Title = "Select AI Inference Data"
            };

            if (!ShowCommonDialog(openFileDialog))
            {
                return;
            }

            _aiInferenceFilePath = openFileDialog.FileName;
            _aiInferenceRowCount = await CountNonEmptyLinesAsync(_aiInferenceFilePath);
            _aiStatusMessage = $"Inference file selected: {Path.GetFileName(_aiInferenceFilePath)}";
            NotifyWebModuleSnapshotChanged();
        }

        private void TryAutoLoadAiModel()
        {
            try
            {
                string aiDirectory = EnsureAiDirectory();
                string modelZipPath = Path.Combine(aiDirectory, "model.zip");
                if (!File.Exists(modelZipPath))
                {
                    return;
                }

                if (TryLoadAiModelFromZip(modelZipPath))
                {
                    _aiStatusMessage = $"Model auto-loaded: {Path.GetFileName(modelZipPath)}";
                    _aiModelZipPath = modelZipPath;
                }
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Failed to auto-load AI process classifier model");
            }
        }

        private void TryAutoLoadAiReasonModel()
        {
            try
            {
                string aiDirectory = EnsureAiDirectory();
                string modelZipPath = Path.Combine(aiDirectory, "reason-model.zip");
                if (!File.Exists(modelZipPath))
                {
                    return;
                }

                if (TryLoadAiReasonModelFromZip(modelZipPath))
                {
                    _aiReasonStatusMessage = $"Reason model auto-loaded: {Path.GetFileName(modelZipPath)}";
                    _aiReasonModelZipPath = modelZipPath;
                }
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Failed to auto-load AI reason model");
            }
        }

        private bool TryLoadAiModelFromZip(string zipPath)
        {
            using ZipArchive archive = ZipFile.OpenRead(zipPath);
            ZipArchiveEntry? entry = archive.GetEntry("model.json");
            if (entry is null)
            {
                return false;
            }

            using Stream entryStream = entry.Open();
            AiModelPayload? payload = JsonSerializer.Deserialize<AiModelPayload>(entryStream);
            if (payload is null)
            {
                return false;
            }

            _aiLastTrainedAtUtc = payload.TrainedAtUtc;
            _aiTrainingRowCount = payload.TrainingRowCount;
            _aiLabelCounts = payload.LabelCounts ?? new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            _aiTokenLabelCounts = payload.TokenLabelCounts ?? new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
            _aiKeywordDictionary = payload.KeywordDictionary ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _aiExactLookup = payload.ExactLookup ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _aiNormalizedLookup = payload.NormalizedLookup ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _aiModelZipPath = zipPath;
            return _aiLabelCounts.Count > 0 && _aiTokenLabelCounts.Count > 0;
        }

        private bool TryLoadAiReasonModelFromZip(string zipPath)
        {
            using ZipArchive archive = ZipFile.OpenRead(zipPath);
            ZipArchiveEntry? entry = archive.GetEntry("model.json");
            if (entry is null)
            {
                return false;
            }

            using Stream entryStream = entry.Open();
            AiModelPayload? payload = JsonSerializer.Deserialize<AiModelPayload>(entryStream);
            if (payload is null)
            {
                return false;
            }

            _aiReasonLastTrainedAtUtc = payload.TrainedAtUtc;
            _aiReasonTrainingRowCount = payload.TrainingRowCount;
            _aiReasonLabelCounts = payload.LabelCounts ?? new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            _aiReasonTokenLabelCounts = payload.TokenLabelCounts ?? new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
            _aiReasonExactLookup = payload.ExactLookup ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _aiReasonNormalizedLookup = payload.NormalizedLookup ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _aiReasonModelZipPath = zipPath;
            return _aiReasonLabelCounts.Count > 0 && _aiReasonTokenLabelCounts.Count > 0;
        }

        private async Task TrainAiProcessClassifierAsync()
        {
            if (string.IsNullOrWhiteSpace(_aiTrainingFilePath) || !File.Exists(_aiTrainingFilePath))
            {
                ShowWarningMessage("Select a training file first.", "AI Process Classifier");
                return;
            }

            List<AiTrainingRecord> records = await LoadAiTrainingRecordsAsync(_aiTrainingFilePath);
            if (records.Count == 0)
            {
                ShowWarningMessage("The selected training file does not contain valid rows.", "AI Process Classifier");
                return;
            }
            await TrainAiProcessClassifierFromRecordsAsync(records, 0);
        }

        private async Task TrainAiReasonClassifierAsync()
        {
            if (string.IsNullOrWhiteSpace(_aiReasonTrainingFilePath) || !File.Exists(_aiReasonTrainingFilePath))
            {
                ShowWarningMessage("Select a reason training file first.", "AI Reason Classifier");
                return;
            }

            List<AiReasonTrainingRecord> records = await LoadAiReasonTrainingRecordsAsync(_aiReasonTrainingFilePath);
            if (records.Count == 0)
            {
                ShowWarningMessage("The selected reason training file does not contain valid rows.", "AI Reason Classifier");
                return;
            }

            Dictionary<string, int> labelCounts = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, Dictionary<string, int>> tokenLabelCounts = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> exactLookup = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> normalizedLookup = new(StringComparer.OrdinalIgnoreCase);

            foreach (AiReasonTrainingRecord record in records)
            {
                if (!labelCounts.TryAdd(record.Reason, 1))
                {
                    labelCounts[record.Reason]++;
                }

                foreach (string token in TokenizeForAi($"{record.ProcessName} {record.NgName}"))
                {
                    if (!tokenLabelCounts.TryGetValue(token, out Dictionary<string, int>? perLabel))
                    {
                        perLabel = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                        tokenLabelCounts[token] = perLabel;
                    }

                    if (!perLabel.TryAdd(record.Reason, 1))
                    {
                        perLabel[record.Reason]++;
                    }
                }

                string exactKey = BuildAiReasonExactKey(record.ProcessName, record.NgName);
                if (!exactLookup.ContainsKey(exactKey))
                {
                    exactLookup[exactKey] = record.Reason;
                }

                string normalizedKey = BuildAiReasonNormalizedKey(record.ProcessName, record.NgName);
                if (!normalizedLookup.ContainsKey(normalizedKey))
                {
                    normalizedLookup[normalizedKey] = record.Reason;
                }
            }

            _aiReasonLabelCounts = labelCounts;
            _aiReasonTokenLabelCounts = tokenLabelCounts;
            _aiReasonExactLookup = exactLookup;
            _aiReasonNormalizedLookup = normalizedLookup;
            _aiReasonTrainingRowCount = records.Count;
            _aiReasonLastTrainedAtUtc = DateTime.UtcNow;

            string aiDirectory = EnsureAiDirectory();
            string modelZipPath = Path.Combine(aiDirectory, "reason-model.zip");
            AiModelPayload payload = new()
            {
                TrainedAtUtc = _aiReasonLastTrainedAtUtc.Value,
                TrainingRowCount = _aiReasonTrainingRowCount,
                LabelCounts = _aiReasonLabelCounts,
                TokenLabelCounts = _aiReasonTokenLabelCounts,
                ExactLookup = _aiReasonExactLookup,
                NormalizedLookup = _aiReasonNormalizedLookup
            };

            await WriteAiModelZipAsync(modelZipPath, payload);
            _aiReasonModelZipPath = modelZipPath;
            _aiReasonStatusMessage = $"Reason model trained with {_aiReasonTrainingRowCount} rows.";
            NotifyWebModuleSnapshotChanged();
        }

        private async Task RetrainAiProcessClassifierWithFeedbackAsync()
        {
            if (string.IsNullOrWhiteSpace(_aiTrainingFilePath) || !File.Exists(_aiTrainingFilePath))
            {
                ShowWarningMessage("Select a base training file first.", "AI Process Classifier");
                return;
            }

            if (string.IsNullOrWhiteSpace(_aiFeedbackFilePath) || !File.Exists(_aiFeedbackFilePath))
            {
                ShowWarningMessage("Select a feedback CSV first.", "AI Process Classifier");
                return;
            }

            List<AiTrainingRecord> baseRecords = await LoadAiTrainingRecordsAsync(_aiTrainingFilePath);
            List<AiTrainingRecord> feedbackRecords = await LoadAiFeedbackTrainingRecordsAsync(_aiFeedbackFilePath);
            if (feedbackRecords.Count == 0)
            {
                ShowWarningMessage("The feedback file does not contain any corrected rows to learn from.", "AI Process Classifier");
                return;
            }

            await TrainAiProcessClassifierFromRecordsAsync(baseRecords.Concat(feedbackRecords).ToList(), feedbackRecords.Count);
        }

        private async Task RetrainAiReasonClassifierWithFeedbackAsync()
        {
            if (string.IsNullOrWhiteSpace(_aiReasonTrainingFilePath) || !File.Exists(_aiReasonTrainingFilePath))
            {
                ShowWarningMessage("Select a base reason training file first.", "AI Reason Classifier");
                return;
            }

            if (string.IsNullOrWhiteSpace(_aiReasonFeedbackFilePath) || !File.Exists(_aiReasonFeedbackFilePath))
            {
                ShowWarningMessage("Select a reason feedback CSV first.", "AI Reason Classifier");
                return;
            }

            List<AiReasonTrainingRecord> baseRecords = await LoadAiReasonTrainingRecordsAsync(_aiReasonTrainingFilePath);
            List<AiReasonTrainingRecord> feedbackRecords = await LoadAiReasonFeedbackTrainingRecordsAsync(_aiReasonFeedbackFilePath);
            if (feedbackRecords.Count == 0)
            {
                ShowWarningMessage("The reason feedback file does not contain any corrected rows to learn from.", "AI Reason Classifier");
                return;
            }

            await TrainAiReasonClassifierFromRecordsAsync(baseRecords.Concat(feedbackRecords).ToList(), feedbackRecords.Count);
        }

        private async Task TrainAiProcessClassifierFromRecordsAsync(List<AiTrainingRecord> records, int feedbackCount)
        {
            Dictionary<string, int> labelCounts = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, Dictionary<string, int>> tokenLabelCounts = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> exactLookup = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> normalizedLookup = new(StringComparer.OrdinalIgnoreCase);

            foreach (AiTrainingRecord record in records)
            {
                if (!labelCounts.TryAdd(record.Label, 1))
                {
                    labelCounts[record.Label]++;
                }

                foreach (string token in TokenizeForAi($"{record.ModelName} {record.CodeName} {record.ProcessName}"))
                {
                    if (!tokenLabelCounts.TryGetValue(token, out Dictionary<string, int>? perLabel))
                    {
                        perLabel = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                        tokenLabelCounts[token] = perLabel;
                    }

                    if (!perLabel.TryAdd(record.Label, 1))
                    {
                        perLabel[record.Label]++;
                    }
                }

                string exactKey = BuildAiExactKey(record.ModelName, record.CodeName, record.ProcessName);
                exactLookup[exactKey] = record.Label;

                string normalizedKey = BuildAiNormalizedKey(record.ModelName, record.CodeName, record.ProcessName);
                normalizedLookup[normalizedKey] = record.Label;
            }

            _aiLabelCounts = labelCounts;
            _aiTokenLabelCounts = tokenLabelCounts;
            _aiExactLookup = exactLookup;
            _aiNormalizedLookup = normalizedLookup;
            _aiTrainingRowCount = records.Count;
            _aiLastTrainedAtUtc = DateTime.UtcNow;

            string aiDirectory = EnsureAiDirectory();
            string modelZipPath = Path.Combine(aiDirectory, "model.zip");
            string dictionaryPath = Path.Combine(aiDirectory, "dictionary.csv");
            string rulesPath = Path.Combine(aiDirectory, "rules.json");

            AiModelPayload payload = new()
            {
                TrainedAtUtc = _aiLastTrainedAtUtc.Value,
                TrainingRowCount = _aiTrainingRowCount,
                LabelCounts = _aiLabelCounts,
                TokenLabelCounts = _aiTokenLabelCounts,
                KeywordDictionary = _aiKeywordDictionary,
                ExactLookup = _aiExactLookup,
                NormalizedLookup = _aiNormalizedLookup
            };

            await WriteAiModelZipAsync(modelZipPath, payload);
            await File.WriteAllLinesAsync(dictionaryPath, _aiKeywordDictionary.Select(item => $"{item.Key},{item.Value}"), Encoding.UTF8);
            await File.WriteAllTextAsync(
                rulesPath,
                JsonSerializer.Serialize(new
                {
                    rules = new[]
                    {
                        new { keyword = "VISUAL", label = "VISUAL" },
                        new { keyword = "TEST", label = "FUNCTION" },
                        new { keyword = "SUB", label = "SUB" },
                        new { keyword = "MAIN", label = "MAIN" }
                    }
                }, new JsonSerializerOptions { WriteIndented = true }),
                Encoding.UTF8);

            _aiModelZipPath = modelZipPath;
            _aiStatusMessage = feedbackCount > 0
                ? $"Model retrained with {records.Count} rows ({feedbackCount} feedback rows)."
                : $"Model trained with {records.Count} rows.";
            NotifyWebModuleSnapshotChanged();
        }

        private async Task TrainAiReasonClassifierFromRecordsAsync(List<AiReasonTrainingRecord> records, int feedbackCount)
        {
            Dictionary<string, int> labelCounts = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, Dictionary<string, int>> tokenLabelCounts = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> exactLookup = new(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> normalizedLookup = new(StringComparer.OrdinalIgnoreCase);

            foreach (AiReasonTrainingRecord record in records)
            {
                if (!labelCounts.TryAdd(record.Reason, 1))
                {
                    labelCounts[record.Reason]++;
                }

                foreach (string token in TokenizeForAi($"{record.ProcessName} {record.NgName}"))
                {
                    if (!tokenLabelCounts.TryGetValue(token, out Dictionary<string, int>? perLabel))
                    {
                        perLabel = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                        tokenLabelCounts[token] = perLabel;
                    }

                    if (!perLabel.TryAdd(record.Reason, 1))
                    {
                        perLabel[record.Reason]++;
                    }
                }

                string exactKey = BuildAiReasonExactKey(record.ProcessName, record.NgName);
                if (!exactLookup.ContainsKey(exactKey))
                {
                    exactLookup[exactKey] = record.Reason;
                }

                string normalizedKey = BuildAiReasonNormalizedKey(record.ProcessName, record.NgName);
                if (!normalizedLookup.ContainsKey(normalizedKey))
                {
                    normalizedLookup[normalizedKey] = record.Reason;
                }
            }

            _aiReasonLabelCounts = labelCounts;
            _aiReasonTokenLabelCounts = tokenLabelCounts;
            _aiReasonExactLookup = exactLookup;
            _aiReasonNormalizedLookup = normalizedLookup;
            _aiReasonTrainingRowCount = records.Count;
            _aiReasonLastTrainedAtUtc = DateTime.UtcNow;

            string aiDirectory = EnsureAiDirectory();
            string modelZipPath = Path.Combine(aiDirectory, "reason-model.zip");
            AiModelPayload payload = new()
            {
                TrainedAtUtc = _aiReasonLastTrainedAtUtc.Value,
                TrainingRowCount = _aiReasonTrainingRowCount,
                LabelCounts = _aiReasonLabelCounts,
                TokenLabelCounts = _aiReasonTokenLabelCounts
            };

            await WriteAiModelZipAsync(modelZipPath, payload);
            _aiReasonModelZipPath = modelZipPath;
            _aiReasonStatusMessage = feedbackCount > 0
                ? $"Reason model retrained with {records.Count} rows ({feedbackCount} feedback rows)."
                : $"Reason model trained with {records.Count} rows.";
            NotifyWebModuleSnapshotChanged();
        }

        private async Task RunAiInferenceAsync()
        {
            _aiStatusMessage = "Preparing inference from current DB missing-routing items...";
            NotifyWebModuleSnapshotChanged();

            if (_aiLabelCounts.Count == 0 || _aiTokenLabelCounts.Count == 0)
            {
                if (!TryLoadAiModelFromZip(Path.Combine(EnsureAiDirectory(), "model.zip")))
                {
                    _aiStatusMessage = "Inference stopped: no trained model was found.";
                    NotifyWebModuleSnapshotChanged();
                    ShowWarningMessage("Train the AI model first.", "AI Process Classifier");
                    return;
                }
            }

            IReadOnlyList<AiInferenceRecord> records = await BuildAiInferenceRecordsFromMissingRoutingAsync();
            if (records.Count == 0)
            {
                _aiStatusMessage = "Inference stopped: current DB has no missing Routing Mapping items.";
                NotifyWebModuleSnapshotChanged();
                ShowWarningMessage("No missing Routing Mapping items were found in the current DB.", "AI Process Classifier");
                return;
            }

            _aiPredictionResults = records
                .Select(PredictAiLabel)
                .ToList();
            _aiInferenceRowCount = records.Count;
            _aiOutputRowCount = _aiPredictionResults.Count;
            _aiStatusMessage = $"Inference completed from current DB missing-routing list: {_aiOutputRowCount} rows.";
            NotifyWebModuleSnapshotChanged();
            ShowAiInferencePreviewWindow();
        }

        private async Task<IReadOnlyList<AiInferenceRecord>> BuildAiInferenceRecordsFromMissingRoutingAsync()
        {
            if (string.IsNullOrWhiteSpace(PathSelectedDB) || !File.Exists(PathSelectedDB))
            {
                _aiStatusMessage = "Inference stopped: load DB first.";
                NotifyWebModuleSnapshotChanged();
                return Array.Empty<AiInferenceRecord>();
            }

            if (_mappingScanResult is null)
            {
                _aiStatusMessage = "Scanning current DB for missing Routing Mapping items...";
                NotifyWebModuleSnapshotChanged();
                await ScanAllMissingMappingsAsync(showProgressPopup: false);
            }

            IReadOnlyList<IReadOnlyDictionary<string, string>> items = _mappingScanResult?.MissingRoutingItems
                ?? _cachedMissingRoutingItems
                ?? Array.Empty<IReadOnlyDictionary<string, string>>();
            if (items.Count == 0)
            {
                return Array.Empty<AiInferenceRecord>();
            }

            List<AiInferenceRecord> records = new();
            foreach (IReadOnlyDictionary<string, string> item in items)
            {
                string modelName = TryGetMissingRoutingValue(item, CONSTANT.MATERIALNAME.NEW);
                string codeName = TryGetMissingRoutingValue(item, "ProcessCode");
                string processName = TryGetMissingRoutingValue(item, "ProcessName");
                if (string.IsNullOrWhiteSpace(modelName) ||
                    string.IsNullOrWhiteSpace(codeName) ||
                    string.IsNullOrWhiteSpace(processName))
                {
                    continue;
                }

                records.Add(new AiInferenceRecord
                {
                    ModelName = modelName,
                    CodeName = codeName,
                    ProcessName = processName,
                    NormalizedModelName = NormalizeCompareValue(modelName),
                    NormalizedCodeName = NormalizeCompareValue(codeName),
                    NormalizedProcessName = NormalizeCompareValue(processName, normalize: true)
                });
            }

            _aiInferenceFilePath = $"Current DB Missing Routing ({records.Count:N0} items)";
            return records;
        }

        private static string TryGetMissingRoutingValue(IReadOnlyDictionary<string, string> item, string key)
        {
            return item.TryGetValue(key, out string? value) ? value?.Trim() ?? string.Empty : string.Empty;
        }

        private void ShowAiInferencePreviewWindow()
        {
            DataTable table = BuildAiInferencePreviewTable();
            _aiInferencePreviewWindow = ShowAiPreviewWindow(
                $"AI Inference Preview ({table.Rows.Count:N0} rows)",
                table,
                _aiInferencePreviewWindow);
        }

        private DataTable BuildAiInferencePreviewTable()
        {
            var table = new DataTable("AiInferencePreview");
            table.Columns.Add("ModelName", typeof(string));
            table.Columns.Add("CodeName", typeof(string));
            table.Columns.Add("ProcessName", typeof(string));
            table.Columns.Add("NormalizedModelName", typeof(string));
            table.Columns.Add("NormalizedCodeName", typeof(string));
            table.Columns.Add("NormalizedProcessName", typeof(string));
            table.Columns.Add("PredictedLabel", typeof(string));
            table.Columns.Add("Confidence", typeof(string));
            table.Columns.Add("Source", typeof(string));

            foreach (AiPredictionResult result in _aiPredictionResults)
            {
                table.Rows.Add(
                    result.ModelName,
                    result.CodeName,
                    result.ProcessName,
                    NormalizeCompareValue(result.ModelName),
                    NormalizeCompareValue(result.CodeName),
                    NormalizeCompareValue(result.ProcessName, normalize: true),
                    result.PredictedLabel,
                    result.Confidence.ToString("0.00"),
                    result.Source);
            }

            return table;
        }

        private async Task ExportAiInferenceOutputAsync()
        {
            try
            {
                if (_aiPredictionResults.Count == 0)
                {
                    _aiStatusMessage = "Export stopped: there are no inference results yet.";
                    NotifyWebModuleSnapshotChanged();
                    ShowWarningMessage("Run inference before exporting output.", "AI Process Classifier");
                    return;
                }

                string exportDirectory = EnsureAiDirectory();
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                _aiOutputFilePath = Path.Combine(exportDirectory, $"ai-process-classification-output_{timestamp}.csv");
                string[] lines = new[]
                {
                    "ModelName,CodeName,ProcessName,PredictedLabel,Confidence,Source,CorrectLabel,IncludeForRetrain"
                }
                .Concat(_aiPredictionResults.Select(result =>
                    string.Join(",",
                        EscapeCsv(NormalizeCompareValue(result.ModelName)),
                        EscapeCsv(NormalizeCompareValue(result.CodeName)),
                        EscapeCsv(NormalizeCompareValue(result.ProcessName, normalize: true)),
                        EscapeCsv(result.PredictedLabel),
                        EscapeCsv(result.Confidence.ToString("0.00")),
                        EscapeCsv(result.Source),
                        string.Empty,
                        EscapeCsv("Y"))))
                .ToArray();
                await File.WriteAllLinesAsync(_aiOutputFilePath, lines, Encoding.UTF8);
                _aiStatusMessage = $"Output exported: {_aiOutputFilePath}";
                NotifyWebModuleSnapshotChanged();
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error exporting AI process classification output");
                _aiStatusMessage = $"Export failed: {ex.Message}";
                NotifyWebModuleSnapshotChanged();
                ShowWarningMessage($"Failed to export AI process classification output.\n{ex.Message}", "AI Process Classifier");
            }
        }

        private async Task RunAiReasonInferenceAsync()
        {
            _aiReasonStatusMessage = "Preparing reason inference from current DB missing-reason items...";
            NotifyWebModuleSnapshotChanged();

            if (_aiReasonLabelCounts.Count == 0 || _aiReasonTokenLabelCounts.Count == 0)
            {
                if (!TryLoadAiReasonModelFromZip(Path.Combine(EnsureAiDirectory(), "reason-model.zip")))
                {
                    _aiReasonStatusMessage = "Reason inference stopped: no trained reason model was found.";
                    NotifyWebModuleSnapshotChanged();
                    ShowWarningMessage("Train the reason model first.", "AI Reason Classifier");
                    return;
                }
            }

            IReadOnlyList<AiReasonInferenceRecord> records = await BuildAiReasonInferenceRecordsFromMissingReasonAsync();
            if (records.Count == 0)
            {
                _aiReasonStatusMessage = "Reason inference stopped: current DB has no missing Reason Mapping items.";
                NotifyWebModuleSnapshotChanged();
                ShowWarningMessage("No missing Reason Mapping items were found in the current DB.", "AI Reason Classifier");
                return;
            }

            _aiReasonPredictionResults = records.Select(PredictAiReason).ToList();
            _aiReasonInferenceRowCount = records.Count;
            _aiReasonOutputRowCount = _aiReasonPredictionResults.Count;
            _aiReasonStatusMessage = $"Reason inference completed from current DB missing-reason list: {_aiReasonOutputRowCount} rows.";
            NotifyWebModuleSnapshotChanged();
            ShowAiReasonInferencePreviewWindow();
        }

        private async Task<IReadOnlyList<AiReasonInferenceRecord>> BuildAiReasonInferenceRecordsFromMissingReasonAsync()
        {
            if (string.IsNullOrWhiteSpace(PathSelectedDB) || !File.Exists(PathSelectedDB))
            {
                _aiReasonStatusMessage = "Reason inference stopped: load DB first.";
                NotifyWebModuleSnapshotChanged();
                return Array.Empty<AiReasonInferenceRecord>();
            }

            if (_mappingScanResult is null)
            {
                _aiReasonStatusMessage = "Scanning current DB for missing Reason Mapping items...";
                NotifyWebModuleSnapshotChanged();
                await ScanAllMissingMappingsAsync(showProgressPopup: false);
            }

            IReadOnlyList<IReadOnlyDictionary<string, string>> items = _mappingScanResult?.MissingReasonItems
                ?? _cachedMissingReasonItems
                ?? Array.Empty<IReadOnlyDictionary<string, string>>();
            if (items.Count == 0)
            {
                return Array.Empty<AiReasonInferenceRecord>();
            }

            List<AiReasonInferenceRecord> records = new();
            foreach (IReadOnlyDictionary<string, string> item in items)
            {
                string processName = TryGetMissingRoutingValue(item, "processName");
                string ngName = TryGetMissingRoutingValue(item, "NgName");
                if (string.IsNullOrWhiteSpace(processName) || string.IsNullOrWhiteSpace(ngName))
                {
                    continue;
                }

                records.Add(new AiReasonInferenceRecord
                {
                    ProcessName = processName,
                    NgName = ngName,
                    NormalizedProcessName = NormalizeCompareValue(processName, normalize: true),
                    NormalizedNgName = NormalizeCompareValue(ngName, normalize: true)
                });
            }

            return records;
        }

        private AiReasonPredictionResult PredictAiReason(AiReasonInferenceRecord record)
        {
            string exactKey = BuildAiReasonExactKey(record.ProcessName, record.NgName);
            if (_aiReasonExactLookup.TryGetValue(exactKey, out string? exactReason))
            {
                return new AiReasonPredictionResult
                {
                    ProcessName = record.ProcessName,
                    NgName = record.NgName,
                    PredictedReason = exactReason,
                    Confidence = 1.00,
                    Source = "EXACT"
                };
            }

            string normalizedKey = $"{record.NormalizedProcessName}\t{record.NormalizedNgName}";
            if (_aiReasonNormalizedLookup.TryGetValue(normalizedKey, out string? normalizedReason))
            {
                return new AiReasonPredictionResult
                {
                    ProcessName = record.ProcessName,
                    NgName = record.NgName,
                    PredictedReason = normalizedReason,
                    Confidence = 0.95,
                    Source = "NORMALIZED"
                };
            }

            string[] tokens = TokenizeForAi($"{record.NormalizedProcessName} {record.NormalizedNgName}").ToArray();
            Dictionary<string, double> scores = new(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, int> labelEntry in _aiReasonLabelCounts)
            {
                scores[labelEntry.Key] = labelEntry.Value;
            }

            foreach (string token in tokens)
            {
                if (!_aiReasonTokenLabelCounts.TryGetValue(token, out Dictionary<string, int>? perLabel))
                {
                    continue;
                }

                foreach (KeyValuePair<string, int> labelEntry in perLabel)
                {
                    if (!scores.TryAdd(labelEntry.Key, labelEntry.Value * 2.0))
                    {
                        scores[labelEntry.Key] += labelEntry.Value * 2.0;
                    }
                }
            }

            if (scores.Count == 0)
            {
                return new AiReasonPredictionResult
                {
                    ProcessName = record.ProcessName,
                    NgName = record.NgName,
                    PredictedReason = string.Empty,
                    Confidence = 0.40,
                    Source = "ML"
                };
            }

            KeyValuePair<string, double> winner = scores.OrderByDescending(entry => entry.Value).First();
            double total = scores.Values.Sum();
            double confidence = total <= 0 ? 0.50 : Math.Max(0.50, Math.Min(0.99, winner.Value / total));

            return new AiReasonPredictionResult
            {
                ProcessName = record.ProcessName,
                NgName = record.NgName,
                PredictedReason = winner.Key,
                Confidence = Math.Max(0.55, confidence),
                Source = "ML"
            };
        }

        private static string BuildAiReasonExactKey(string processName, string ngName)
        {
            return $"{(processName ?? string.Empty).Trim()}\t{(ngName ?? string.Empty).Trim()}";
        }

        private static string BuildAiExactKey(string modelName, string codeName, string processName)
        {
            return $"{(modelName ?? string.Empty).Trim()}\t{(codeName ?? string.Empty).Trim()}\t{(processName ?? string.Empty).Trim()}";
        }

        private static string BuildAiNormalizedKey(string modelName, string codeName, string processName)
        {
            return $"{NormalizeCompareValue(modelName)}\t{NormalizeCompareValue(codeName)}\t{NormalizeCompareValue(processName, normalize: true)}";
        }

        private static string BuildAiReasonNormalizedKey(string processName, string ngName)
        {
            return $"{NormalizeCompareValue(processName, normalize: true)}\t{NormalizeCompareValue(ngName, normalize: true)}";
        }

        private void ShowAiReasonInferencePreviewWindow()
        {
            DataTable table = BuildAiReasonInferencePreviewTable();
            _aiReasonInferencePreviewWindow = ShowAiPreviewWindow(
                $"AI Reason Inference Preview ({table.Rows.Count:N0} rows)",
                table,
                _aiReasonInferencePreviewWindow);
        }

        private Window ShowAiPreviewWindow(string title, DataTable table, Window? existingWindow)
        {
            var dataGrid = new DataGrid
            {
                AutoGenerateColumns = true,
                IsReadOnly = true,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                CanUserReorderColumns = true,
                CanUserResizeColumns = true,
                CanUserSortColumns = true,
                Margin = new Thickness(12),
                HeadersVisibility = DataGridHeadersVisibility.Column,
                ItemsSource = table.DefaultView
            };
            var closeButton = new Button
            {
                Content = "Close",
                MinWidth = 96,
                Height = 34,
                Margin = new Thickness(12),
                HorizontalAlignment = HorizontalAlignment.Right
            };
            closeButton.Click += (_, _) =>
            {
                if (Window.GetWindow(closeButton) is Window previewWindow)
                {
                    previewWindow.Close();
                }
            };

            var root = new DockPanel();
            DockPanel.SetDock(closeButton, Dock.Bottom);
            root.Children.Add(closeButton);
            root.Children.Add(dataGrid);

            Window window = existingWindow ?? new Window
            {
                Title = title,
                Width = 1180,
                Height = 760,
                MinWidth = 860,
                MinHeight = 480,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ShowInTaskbar = false,
                WindowStyle = WindowStyle.SingleBorderWindow,
                ResizeMode = ResizeMode.CanResize,
                Content = root
            };

            if (existingWindow != null)
            {
                window.Title = title;
                window.Content = root;
            }

            dataGrid.PreviewKeyDown += (_, args) =>
            {
                if (args.Key == System.Windows.Input.Key.Escape &&
                    Window.GetWindow(dataGrid) is Window previewWindow)
                {
                    previewWindow.Close();
                    args.Handled = true;
                }
            };
            window.PreviewKeyDown += (_, args) =>
            {
                if (args.Key == System.Windows.Input.Key.Escape)
                {
                    window.Close();
                    args.Handled = true;
                }
            };

            Window? owner = ResolveDialogOwner(null);
            if (owner?.IsVisible == true)
            {
                window.Owner = owner;
                window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }

            if (!window.IsVisible)
            {
                window.Show();
            }
            else
            {
                if (window.WindowState == WindowState.Minimized)
                {
                    window.WindowState = WindowState.Normal;
                }

                window.Activate();
                window.Focus();
            }

            return window;
        }

        private DataTable BuildAiReasonInferencePreviewTable()
        {
            var table = new DataTable("AiReasonInferencePreview");
            table.Columns.Add("ProcessName", typeof(string));
            table.Columns.Add("NgName", typeof(string));
            table.Columns.Add("NormalizedProcessName", typeof(string));
            table.Columns.Add("NormalizedNgName", typeof(string));
            table.Columns.Add("PredictedReason", typeof(string));
            table.Columns.Add("Confidence", typeof(string));
            table.Columns.Add("Source", typeof(string));

            foreach (AiReasonPredictionResult result in _aiReasonPredictionResults)
            {
                table.Rows.Add(
                    result.ProcessName,
                    result.NgName,
                    NormalizeCompareValue(result.ProcessName, normalize: true),
                    NormalizeCompareValue(result.NgName, normalize: true),
                    result.PredictedReason,
                    result.Confidence.ToString("0.00"),
                    result.Source);
            }

            return table;
        }

        private async Task ExportAiReasonInferenceOutputAsync()
        {
            try
            {
                if (_aiReasonPredictionResults.Count == 0)
                {
                    _aiReasonStatusMessage = "Reason export stopped: there are no inference results yet.";
                    NotifyWebModuleSnapshotChanged();
                    ShowWarningMessage("Run reason inference before exporting output.", "AI Reason Classifier");
                    return;
                }

                string exportDirectory = EnsureAiDirectory();
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                _aiReasonOutputFilePath = Path.Combine(exportDirectory, $"ai-reason-classification-output_{timestamp}.csv");
                string[] lines = new[]
                {
                    "ProcessName,NgName,PredictedReason,Confidence,Source,CorrectReason,IncludeForRetrain"
                }
                .Concat(_aiReasonPredictionResults.Select(result =>
                    string.Join(",",
                        EscapeCsv(NormalizeCompareValue(result.ProcessName, normalize: true)),
                        EscapeCsv(NormalizeCompareValue(result.NgName, normalize: true)),
                        EscapeCsv(result.PredictedReason),
                        EscapeCsv(result.Confidence.ToString("0.00")),
                        EscapeCsv(result.Source),
                        string.Empty,
                        EscapeCsv("Y"))))
                .ToArray();
                await File.WriteAllLinesAsync(_aiReasonOutputFilePath, lines, Encoding.UTF8);
                _aiReasonStatusMessage = $"Reason output exported: {_aiReasonOutputFilePath}";
                NotifyWebModuleSnapshotChanged();
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error exporting AI reason classification output");
                _aiReasonStatusMessage = $"Reason export failed: {ex.Message}";
                NotifyWebModuleSnapshotChanged();
                ShowWarningMessage($"Failed to export AI reason classification output.\n{ex.Message}", "AI Reason Classifier");
            }
        }

        private static string EscapeCsv(string? value)
        {
            string text = value ?? string.Empty;
            if (text.Contains('"'))
            {
                text = text.Replace("\"", "\"\"");
            }

            if (text.Contains(',') || text.Contains('"') || text.Contains('\r') || text.Contains('\n'))
            {
                return $"\"{text}\"";
            }

            return text;
        }

        private static async Task<int> CountNonEmptyLinesAsync(string filePath)
        {
            int count = 0;
            foreach (string line in await File.ReadAllLinesAsync(filePath, Encoding.UTF8))
            {
                if (!string.IsNullOrWhiteSpace(line))
                {
                    count++;
                }
            }

            return count;
        }

        private static async Task<List<AiTrainingRecord>> LoadAiTrainingRecordsAsync(string filePath)
        {
            string[] lines = await File.ReadAllLinesAsync(filePath, Encoding.UTF8);
            List<AiTrainingRecord> records = new();

            foreach (string rawLine in lines)
            {
                if (TryParseAiTrainingRecord(rawLine, out AiTrainingRecord? record))
                {
                    records.Add(record!);
                }
            }

            return records;
        }

        private static async Task<List<AiReasonTrainingRecord>> LoadAiReasonTrainingRecordsAsync(string filePath)
        {
            string[] lines = await File.ReadAllLinesAsync(filePath, Encoding.UTF8);
            List<AiReasonTrainingRecord> records = new();

            foreach (string rawLine in lines)
            {
                if (TryParseAiReasonTrainingRecord(rawLine, out AiReasonTrainingRecord? record))
                {
                    records.Add(record!);
                }
            }

            return records;
        }

        private static async Task<List<AiInferenceRecord>> LoadAiInferenceRecordsAsync(string filePath)
        {
            string[] lines = await File.ReadAllLinesAsync(filePath, Encoding.UTF8);
            List<AiInferenceRecord> records = new();

            foreach (string rawLine in lines)
            {
                if (TryParseAiInferenceRecord(rawLine, out AiInferenceRecord? record))
                {
                    records.Add(record!);
                }
            }

            return records;
        }

        private static async Task<List<AiTrainingRecord>> LoadAiFeedbackTrainingRecordsAsync(string filePath)
        {
            string[] lines = await File.ReadAllLinesAsync(filePath, Encoding.UTF8);
            if (lines.Length == 0)
            {
                return new List<AiTrainingRecord>();
            }

            Dictionary<string, int> headers = ParseCsvLine(lines[0])
                .Select((header, index) => new { header, index })
                .ToDictionary(item => item.header.Trim(), item => item.index, StringComparer.OrdinalIgnoreCase);

            if (!headers.ContainsKey("ModelName") ||
                !headers.ContainsKey("CodeName") ||
                !headers.ContainsKey("ProcessName") ||
                !headers.ContainsKey("CorrectLabel"))
            {
                return new List<AiTrainingRecord>();
            }

            var records = new List<AiTrainingRecord>();
            foreach (string line in lines.Skip(1))
            {
                if (TryParseAiFeedbackTrainingRecord(line, headers, out AiTrainingRecord? record))
                {
                    records.Add(record!);
                }
            }

            return records;
        }

        private static async Task<List<AiReasonTrainingRecord>> LoadAiReasonFeedbackTrainingRecordsAsync(string filePath)
        {
            string[] lines = await File.ReadAllLinesAsync(filePath, Encoding.UTF8);
            if (lines.Length == 0)
            {
                return new List<AiReasonTrainingRecord>();
            }

            Dictionary<string, int> headers = ParseCsvLine(lines[0])
                .Select((header, index) => new { header, index })
                .ToDictionary(item => item.header.Trim(), item => item.index, StringComparer.OrdinalIgnoreCase);

            if (!headers.ContainsKey("ProcessName") ||
                !headers.ContainsKey("NgName") ||
                !headers.ContainsKey("CorrectReason"))
            {
                return new List<AiReasonTrainingRecord>();
            }

            var records = new List<AiReasonTrainingRecord>();
            foreach (string line in lines.Skip(1))
            {
                if (TryParseAiReasonFeedbackTrainingRecord(line, headers, out AiReasonTrainingRecord? record))
                {
                    records.Add(record!);
                }
            }

            return records;
        }

        private static bool TryParseAiTrainingRecord(string rawLine, out AiTrainingRecord? record)
        {
            record = null;
            string[] tokens = SplitAiLine(rawLine);
            if (tokens.Length < 4)
            {
                return false;
            }

            string label = NormalizeAiKeyword(tokens[^1]);
            string processName = string.Join(" ", tokens.Skip(2).Take(tokens.Length - 3)).Trim();
            if (string.IsNullOrWhiteSpace(processName))
            {
                return false;
            }

            record = new AiTrainingRecord
            {
                ModelName = tokens[0].Trim(),
                CodeName = tokens[1].Trim(),
                ProcessName = processName,
                Label = string.IsNullOrWhiteSpace(label) ? tokens[^1].Trim().ToUpperInvariant() : label
            };
            return true;
        }

        private static bool TryParseAiInferenceRecord(string rawLine, out AiInferenceRecord? record)
        {
            record = null;
            string[] tokens = SplitAiLine(rawLine);
            if (tokens.Length < 3)
            {
                return false;
            }

            string processName = string.Join(" ", tokens.Skip(2)).Trim();
            if (string.IsNullOrWhiteSpace(processName))
            {
                return false;
            }

            record = new AiInferenceRecord
            {
                ModelName = tokens[0].Trim(),
                CodeName = tokens[1].Trim(),
                ProcessName = processName,
                NormalizedModelName = NormalizeCompareValue(tokens[0].Trim()),
                NormalizedCodeName = NormalizeCompareValue(tokens[1].Trim()),
                NormalizedProcessName = NormalizeCompareValue(processName, normalize: true)
            };
            return true;
        }

        private static bool TryParseAiReasonTrainingRecord(string rawLine, out AiReasonTrainingRecord? record)
        {
            record = null;
            string line = rawLine?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(line))
            {
                return false;
            }

            string[] tabParts = line
                .Split('\t')
                .Select(part => part.Trim())
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .ToArray();

            string processName;
            string ngName;
            string reason;

            if (tabParts.Length >= 3)
            {
                processName = tabParts[0];
                ngName = tabParts[1];
                reason = string.Join(" ", tabParts.Skip(2)).Trim();
            }
            else
            {
                string[] tokens = SplitAiLine(line);
                if (tokens.Length < 3)
                {
                    return false;
                }

                processName = tokens[0].Trim();
                ngName = tokens[1].Trim();
                reason = string.Join(" ", tokens.Skip(2)).Trim();
            }

            if (string.IsNullOrWhiteSpace(processName) ||
                string.IsNullOrWhiteSpace(ngName) ||
                string.IsNullOrWhiteSpace(reason))
            {
                return false;
            }

            record = new AiReasonTrainingRecord
            {
                ProcessName = processName,
                NgName = ngName,
                Reason = reason
            };
            return true;
        }

        private static bool TryParseAiFeedbackTrainingRecord(string rawLine, IReadOnlyDictionary<string, int> headers, out AiTrainingRecord? record)
        {
            record = null;
            string[] values = ParseCsvLine(rawLine);
            if (values.Length == 0)
            {
                return false;
            }

            string includeValue = GetCsvValue(values, headers, "IncludeForRetrain");
            if (!ShouldIncludeFeedbackRow(includeValue))
            {
                return false;
            }

            string label = NormalizeAiKeyword(GetCsvValue(values, headers, "CorrectLabel"));

            string modelName = GetCsvValue(values, headers, "ModelName");
            string codeName = GetCsvValue(values, headers, "CodeName");
            string processName = GetCsvValue(values, headers, "ProcessName");
            if (string.IsNullOrWhiteSpace(modelName) || string.IsNullOrWhiteSpace(codeName) || string.IsNullOrWhiteSpace(processName))
            {
                return false;
            }

            record = new AiTrainingRecord
            {
                ModelName = NormalizeCompareValue(modelName),
                CodeName = NormalizeCompareValue(codeName),
                ProcessName = NormalizeCompareValue(processName, normalize: true),
                Label = label
            };
            return true;
        }

        private static bool TryParseAiReasonFeedbackTrainingRecord(string rawLine, IReadOnlyDictionary<string, int> headers, out AiReasonTrainingRecord? record)
        {
            record = null;
            string[] values = ParseCsvLine(rawLine);
            if (values.Length == 0)
            {
                return false;
            }

            string includeValue = GetCsvValue(values, headers, "IncludeForRetrain");
            if (!ShouldIncludeFeedbackRow(includeValue))
            {
                return false;
            }

            string reason = GetCsvValue(values, headers, "CorrectReason");
            string processName = GetCsvValue(values, headers, "ProcessName");
            string ngName = GetCsvValue(values, headers, "NgName");
            if (string.IsNullOrWhiteSpace(processName) || string.IsNullOrWhiteSpace(ngName))
            {
                return false;
            }

            record = new AiReasonTrainingRecord
            {
                ProcessName = NormalizeCompareValue(processName, normalize: true),
                NgName = NormalizeCompareValue(ngName, normalize: true),
                Reason = reason.Trim()
            };
            return true;
        }

        private static string GetCsvValue(string[] values, IReadOnlyDictionary<string, int> headers, string headerName)
        {
            return headers.TryGetValue(headerName, out int index) && index >= 0 && index < values.Length
                ? values[index].Trim()
                : string.Empty;
        }

        private static bool ShouldIncludeFeedbackRow(string includeValue)
        {
            string normalized = (includeValue ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return true;
            }

            return normalized.Equals("Y", StringComparison.OrdinalIgnoreCase) ||
                   normalized.Equals("YES", StringComparison.OrdinalIgnoreCase) ||
                   normalized.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
                   normalized.Equals("1", StringComparison.OrdinalIgnoreCase);
        }

        private static string[] ParseCsvLine(string line)
        {
            var values = new List<string>();
            var builder = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < (line ?? string.Empty).Length; i++)
            {
                char ch = line[i];
                if (ch == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        builder.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (ch == ',' && !inQuotes)
                {
                    values.Add(builder.ToString());
                    builder.Clear();
                }
                else
                {
                    builder.Append(ch);
                }
            }

            values.Add(builder.ToString());
            return values.ToArray();
        }

        private static string[] SplitAiLine(string rawLine)
        {
            return Regex.Split(rawLine?.Trim() ?? string.Empty, "\\s+")
                .Where(token => !string.IsNullOrWhiteSpace(token))
                .ToArray();
        }

        private static string NormalizeAiKeyword(string value)
        {
            string token = (value ?? string.Empty).Trim().ToUpperInvariant();
            return token switch
            {
                "VIS" => "VISUAL",
                "VSL" => "VISUAL",
                "FUNC" => "FUNCTION",
                _ => token
            };
        }

        private static string NormalizeAiText(string value)
        {
            string normalized = NormalizeCompareValue(value, normalize: true).ToUpperInvariant();
            normalized = normalized.Replace("/", " ")
                .Replace("_", " ")
                .Replace("(", " ")
                .Replace(")", " ");
            normalized = Regex.Replace(normalized, "[^A-Z0-9\\-\\s]", " ");
            normalized = Regex.Replace(normalized, "\\s+", " ").Trim();
            return normalized;
        }

        private static IEnumerable<string> TokenizeForAi(string value)
        {
            string normalized = NormalizeAiText(value);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return Array.Empty<string>();
            }

            return normalized
                .Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(NormalizeAiKeyword)
                .Where(token => token.Any(ch => !char.IsDigit(ch)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private AiPredictionResult PredictAiLabel(AiInferenceRecord record)
        {
            string exactKey = BuildAiExactKey(record.ModelName, record.CodeName, record.ProcessName);
            if (_aiExactLookup.TryGetValue(exactKey, out string? exactLabel))
            {
                return CreateAiPrediction(record, exactLabel, 1.00, "EXACT");
            }

            string normalizedKey = BuildAiNormalizedKey(record.NormalizedModelName, record.NormalizedCodeName, record.NormalizedProcessName);
            if (_aiNormalizedLookup.TryGetValue(normalizedKey, out string? normalizedLabel))
            {
                return CreateAiPrediction(record, normalizedLabel, 0.95, "NORMALIZED");
            }

            string[] tokens = TokenizeForAi($"{record.NormalizedModelName} {record.NormalizedCodeName} {record.NormalizedProcessName}").ToArray();
            string[] normalizedProcessTokens = TokenizeForAi(record.NormalizedProcessName).ToArray();

            if (normalizedProcessTokens.Contains("VISUAL", StringComparer.OrdinalIgnoreCase))
            {
                return CreateAiPrediction(record, "VISUAL", 0.99, "RULE");
            }

            if (normalizedProcessTokens.Contains("TEST", StringComparer.OrdinalIgnoreCase))
            {
                return CreateAiPrediction(record, "FUNCTION", 0.97, "RULE");
            }

            if (normalizedProcessTokens.Contains("SUB", StringComparer.OrdinalIgnoreCase))
            {
                return CreateAiPrediction(record, "SUB", 0.96, "RULE");
            }

            Dictionary<string, double> scores = new(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, int> labelEntry in _aiLabelCounts)
            {
                scores[labelEntry.Key] = labelEntry.Value;
            }

            foreach (string token in tokens)
            {
                if (!_aiTokenLabelCounts.TryGetValue(token, out Dictionary<string, int>? perLabel))
                {
                    continue;
                }

                foreach (KeyValuePair<string, int> labelEntry in perLabel)
                {
                    if (!scores.TryAdd(labelEntry.Key, labelEntry.Value * 2.0))
                    {
                        scores[labelEntry.Key] += labelEntry.Value * 2.0;
                    }
                }
            }

            if (scores.Count == 0)
            {
                return CreateAiPrediction(record, "MAIN", 0.50, "ML");
            }

            KeyValuePair<string, double> winner = scores
                .OrderByDescending(entry => entry.Value)
                .ThenBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
                .First();
            double total = scores.Values.Sum();
            double confidence = total <= 0 ? 0.50 : Math.Max(0.50, Math.Min(0.99, winner.Value / total));

            return CreateAiPrediction(record, winner.Key, confidence, "ML");
        }

        private static AiPredictionResult CreateAiPrediction(AiInferenceRecord record, string label, double confidence, string source)
        {
            return new AiPredictionResult
            {
                ModelName = record.ModelName,
                CodeName = record.CodeName,
                ProcessName = record.ProcessName,
                PredictedLabel = label,
                Confidence = confidence,
                Source = source
            };
        }

        private static async Task WriteAiModelZipAsync(string zipPath, AiModelPayload payload)
        {
            if (File.Exists(zipPath))
            {
                File.Delete(zipPath);
            }

            await using FileStream stream = new(zipPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
            using ZipArchive archive = new(stream, ZipArchiveMode.Create, leaveOpen: false);
            ZipArchiveEntry entry = archive.CreateEntry("model.json");
            await using Stream entryStream = entry.Open();
            await JsonSerializer.SerializeAsync(entryStream, payload, new JsonSerializerOptions { WriteIndented = true });
        }

        public object GetWebModuleSnapshot()
        {
            var selectedGroup = CT_LB_GROUPS.SelectedItem as LoadedGroupDisplay;
            var databaseSummary = GetDatabaseSummaryInfo();

            return new
            {
                hasDatabase = !string.IsNullOrWhiteSpace(PathSelectedDB),
                databaseName = string.IsNullOrWhiteSpace(PathSelectedDB) ? string.Empty : Path.GetFileName(PathSelectedDB),
                databasePath = PathSelectedDB,
                originalTableDateRange = databaseSummary.OriginalTableDateRange,
                lastUpdatedAt = databaseSummary.LastUpdatedAt,
                progressTitle = _currentProgressTitle,
                progressMessage = _currentProgressMessage,
                activeView = GetCurrentWebViewName(),
                currentTaskPercent = _currentTaskProgressPercent,
                totalTaskPercent = _currentTotalProgressPercent,
                isReportEnabled = CT_BT_GETREPORT.IsEnabled,
                loadedGroupCount = _loadedGroupDisplays.Count,
                selectedGroupName = selectedGroup?.GroupName ?? string.Empty,
                selectedSubGroupCount = selectedGroup?.Count ?? 0,
                startDate = CT_DATE_START.SelectedDate?.ToString("yyyy-MM-dd") ?? string.Empty,
                endDate = CT_DATE_END.SelectedDate?.ToString("yyyy-MM-dd") ?? string.Empty,
                loadedGroups = _loadedGroupDisplays
                    .Take(6)
                    .Select(group => new
                    {
                        name = group.GroupName,
                        count = group.Count
                    })
                    .ToArray(),
                debug = new
                {
                    group = CT_CB_DEBUG_GROUP.Text?.Trim() ?? string.Empty,
                    process = CT_CB_DEBUG_PROCESS.Text?.Trim() ?? string.Empty,
                    ng = CT_CB_DEBUG_NG.Text?.Trim() ?? string.Empty,
                    date = CT_CB_DEBUG_DATE.Text?.Trim() ?? string.Empty,
                    result = CT_TB_DEBUG_RESULT.Text ?? string.Empty,
                    availableGroups = GetComboItems(CT_CB_DEBUG_GROUP),
                    availableProcesses = GetComboItems(CT_CB_DEBUG_PROCESS),
                    availableNgs = GetComboItems(CT_CB_DEBUG_NG),
                    availableDates = GetComboItems(CT_CB_DEBUG_DATE)
                },
                mapping = _mappingScanResult is null
                    ? null
                    : new
                    {
                        totalOriginalRows = _mappingScanResult.TotalOriginalRows,
                        appliedRoutingRows = _mappingScanResult.AppliedRoutingRows,
                        missingRoutingRows = _mappingScanResult.MissingRoutingRows,
                        appliedReasonRows = _mappingScanResult.AppliedReasonRows,
                        missingReasonRows = _mappingScanResult.MissingReasonRows,
                        currentView = _currentMappingTableKind switch
                        {
                            MappingTableKind.Routing => "routing",
                            MappingTableKind.Reason => "reason",
                            _ => "none"
                        }
                    },
                aiClassifier = new
                {
                    trainingFilePath = _aiTrainingFilePath,
                    inferenceFilePath = _aiInferenceFilePath,
                    outputFilePath = _aiOutputFilePath,
                    feedbackFilePath = _aiFeedbackFilePath,
                    modelZipPath = _aiModelZipPath,
                    statusMessage = _aiStatusMessage,
                    trainingRowCount = _aiTrainingRowCount,
                    inferenceRowCount = _aiInferenceRowCount,
                    outputRowCount = _aiOutputRowCount,
                    feedbackRowCount = _aiFeedbackRowCount,
                    hasModel = _aiLabelCounts.Count > 0 && _aiTokenLabelCounts.Count > 0,
                    lastTrainedAt = _aiLastTrainedAtUtc?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") ?? string.Empty,
                    labels = _aiLabelCounts
                        .OrderByDescending(item => item.Value)
                        .Select(item => new
                        {
                            label = item.Key,
                            count = item.Value
                        })
                        .ToArray(),
                    recentResults = _aiPredictionResults
                        .Take(5)
                        .Select(result => new
                        {
                            modelName = result.ModelName,
                            codeName = result.CodeName,
                            processName = result.ProcessName,
                            predictedLabel = result.PredictedLabel,
                            confidence = result.Confidence.ToString("0.00"),
                            source = result.Source
                        })
                        .ToArray()
                },
                aiReasonClassifier = new
                {
                    trainingFilePath = _aiReasonTrainingFilePath,
                    outputFilePath = _aiReasonOutputFilePath,
                    feedbackFilePath = _aiReasonFeedbackFilePath,
                    modelZipPath = _aiReasonModelZipPath,
                    statusMessage = _aiReasonStatusMessage,
                    trainingRowCount = _aiReasonTrainingRowCount,
                    inferenceRowCount = _aiReasonInferenceRowCount,
                    outputRowCount = _aiReasonOutputRowCount,
                    feedbackRowCount = _aiReasonFeedbackRowCount,
                    hasModel = _aiReasonLabelCounts.Count > 0 && _aiReasonTokenLabelCounts.Count > 0,
                    lastTrainedAt = _aiReasonLastTrainedAtUtc?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") ?? string.Empty,
                    labels = _aiReasonLabelCounts
                        .OrderByDescending(item => item.Value)
                        .Select(item => new
                        {
                            label = item.Key,
                            count = item.Value
                        })
                        .ToArray(),
                    recentResults = _aiReasonPredictionResults
                        .Take(5)
                        .Select(result => new
                        {
                            processName = result.ProcessName,
                            ngName = result.NgName,
                            predictedReason = result.PredictedReason,
                            confidence = result.Confidence.ToString("0.00"),
                            source = result.Source
                        })
                        .ToArray()
                }
            };
        }

        public object UpdateWebModuleState(JsonElement payload)
        {
            if (payload.ValueKind != JsonValueKind.Object)
            {
                return GetWebModuleSnapshot();
            }

            if (TryGetString(payload, "startDate", out string? startDate) &&
                DateTime.TryParse(startDate, out DateTime parsedStart))
            {
                CT_DATE_START.SelectedDate = parsedStart.Date;
            }

            if (TryGetString(payload, "endDate", out string? endDate) &&
                DateTime.TryParse(endDate, out DateTime parsedEnd))
            {
                CT_DATE_END.SelectedDate = parsedEnd.Date;
            }

            bool shouldRefreshDebugGroups = false;

            if (TryGetString(payload, "debugGroup", out string? debugGroup))
            {
                shouldRefreshDebugGroups = true;
                RefreshDebugGroupDropdown();
                CT_CB_DEBUG_GROUP.Text = debugGroup ?? string.Empty;
            }

            if (TryGetString(payload, "debugProcess", out string? debugProcess))
            {
                if (shouldRefreshDebugGroups)
                {
                    RefreshDebugProcessDropdown();
                }
                CT_CB_DEBUG_PROCESS.Text = debugProcess ?? string.Empty;
                RefreshDebugNgDropdown();
            }

            if (TryGetString(payload, "debugNg", out string? debugNg))
            {
                CT_CB_DEBUG_NG.Text = debugNg ?? string.Empty;
                RefreshDebugDateDropdown();
            }

            if (TryGetString(payload, "debugDate", out string? debugDate))
            {
                CT_CB_DEBUG_DATE.Text = debugDate ?? string.Empty;
            }

            return GetWebModuleSnapshot();
        }

        public async Task<object> InvokeWebModuleActionAsync(string action)
        {
            switch (action)
            {
                case "show-get-data":
                    ShowMainContent();
                    CT_MAPPING_CONTENT.Visibility = Visibility.Collapsed;
                    break;
                case "load-db":
                    await LoadProcess();
                    break;
                case "load-group-json":
                    CT_BT_LOAD_GROUP_JSON_Click(this, new RoutedEventArgs());
                    break;
                case "scan-mapping":
                    await ShowMappingDashboardAsync(MappingTableKind.Routing, forceScan: true);
                    break;
                case "show-model-groups":
                    CT_BT_MODELGROUPS_Click(this, new RoutedEventArgs());
                    break;
                case "get-report":
                    await StartGetReportAsync();
                    break;
                case "show-settings":
                    CT_BT_SETTINGS_Click(this, new RoutedEventArgs());
                    break;
                case "bmes-settings":
                    CT_BT_BMESSETTING_Click(this, new RoutedEventArgs());
                    break;
                case "get-bmes-data":
                    CT_BT_GETBMESDATA_Click(this, new RoutedEventArgs());
                    break;
                case "get-routing":
                    CT_BT_GETROUTING_Click(this, new RoutedEventArgs());
                    break;
                case "check-debug-rate":
                    CT_BT_DEBUG_RATE_Click(this, new RoutedEventArgs());
                    break;
                case "mapping-export":
                    CT_BT_MAPPING_EXPORT_Click(this, new RoutedEventArgs());
                    break;
                case "mapping-import":
                    CT_BT_MAPPING_IMPORT_Click(this, new RoutedEventArgs());
                    break;
                case "clear-log":
                    CT_BT_CLEAR_LOG_Click(this, new RoutedEventArgs());
                    break;
                case "show-routing-mapping":
                    await ShowMappingDashboardAsync(MappingTableKind.Routing, forceScan: _mappingScanResult == null);
                    break;
                case "show-reason-mapping":
                    await ShowMappingDashboardAsync(MappingTableKind.Reason, forceScan: _mappingScanResult == null);
                    break;
                case "show-ai-process-classifier":
                    ShowMainContent();
                    _aiStatusMessage = "AI process classifier panel is ready.";
                    NotifyWebModuleSnapshotChanged();
                    break;
                case "ai-select-training":
                    await SelectAiTrainingFileAsync();
                    break;
                case "ai-train-model":
                    await TrainAiProcessClassifierAsync();
                    break;
                case "ai-run-inference":
                    await RunAiInferenceAsync();
                    break;
                case "ai-export-output":
                    await ExportAiInferenceOutputAsync();
                    break;
                case "ai-import-feedback":
                    await SelectAiFeedbackFileAsync();
                    break;
                case "ai-retrain-feedback":
                    await RetrainAiProcessClassifierWithFeedbackAsync();
                    break;
                case "ai-reason-select-training":
                    await SelectAiReasonTrainingFileAsync();
                    break;
                case "ai-reason-train-model":
                    await TrainAiReasonClassifierAsync();
                    break;
                case "ai-reason-run-inference":
                    await RunAiReasonInferenceAsync();
                    break;
                case "ai-reason-export-output":
                    await ExportAiReasonInferenceOutputAsync();
                    break;
                case "ai-reason-import-feedback":
                    await SelectAiReasonFeedbackFileAsync();
                    break;
                case "ai-reason-retrain-feedback":
                    await RetrainAiReasonClassifierWithFeedbackAsync();
                    break;
            }

            return GetWebModuleSnapshot();
        }

        private static string[] GetComboItems(ComboBox comboBox)
        {
            return comboBox.Items
                .Cast<object>()
                .Select(item => item?.ToString()?.Trim() ?? string.Empty)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private static bool TryGetString(JsonElement payload, string propertyName, out string? value)
        {
            value = null;
            if (!payload.TryGetProperty(propertyName, out JsonElement property))
            {
                return false;
            }

            value = property.ValueKind == JsonValueKind.String
                ? property.GetString()
                : property.ToString();
            return true;
        }

        private (string OriginalTableDateRange, string LastUpdatedAt) GetDatabaseSummaryInfo()
        {
            if (string.IsNullOrWhiteSpace(PathSelectedDB) || !File.Exists(PathSelectedDB))
            {
                return (string.Empty, string.Empty);
            }

            DateTime lastWriteTimeUtc = File.GetLastWriteTimeUtc(PathSelectedDB);

            if (string.Equals(_cachedSummaryDbPath, PathSelectedDB, StringComparison.OrdinalIgnoreCase)
                && _cachedSummaryLastWriteTimeUtc == lastWriteTimeUtc)
            {
                return (_cachedOriginalTableDateRange, _cachedDbLastUpdatedAt);
            }

            try
            {
                using var connection = new SQLiteConnection($"Data Source={PathSelectedDB};Version=3;");
                connection.Open();

                string productDateColumn = CONSTANT.PRODUCT_DATE.NEW;
                string originalTableName = CONSTANT.OPTION_TABLE_NAME.ORG;
                using var command = new SQLiteCommand(
                    $"SELECT MIN([{productDateColumn}]), MAX([{productDateColumn}]) FROM [{originalTableName}]",
                    connection);

                using var reader = command.ExecuteReader();
                if (!reader.Read())
                {
                    string metaUpdatedAt = TryReadOriginalTableUpdatedAt(connection) ?? string.Empty;
                    return (string.Empty, metaUpdatedAt);
                }

                string minDate = reader.IsDBNull(0) ? string.Empty : reader.GetValue(0)?.ToString()?.Trim() ?? string.Empty;
                string maxDate = reader.IsDBNull(1) ? string.Empty : reader.GetValue(1)?.ToString()?.Trim() ?? string.Empty;

                string dateRange = string.IsNullOrWhiteSpace(minDate) && string.IsNullOrWhiteSpace(maxDate)
                    ? string.Empty
                    : minDate == maxDate
                        ? minDate
                        : $"{minDate} ~ {maxDate}";

                string lastUpdatedAt = TryReadOriginalTableUpdatedAt(connection) ?? string.Empty;

                _cachedSummaryDbPath = PathSelectedDB;
                _cachedSummaryLastWriteTimeUtc = lastWriteTimeUtc;
                _cachedOriginalTableDateRange = dateRange;
                _cachedDbLastUpdatedAt = lastUpdatedAt;
                return (dateRange, lastUpdatedAt);
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Failed to load OriginalTable date range for WebView snapshot");
                return (string.Empty, string.Empty);
            }
        }

        private static string? TryReadOriginalTableUpdatedAt(SQLiteConnection connection)
        {
            using var command = new SQLiteCommand(
                $"SELECT [MetaValue] FROM [{OriginalTableMetaTableName}] WHERE [MetaKey] = @key LIMIT 1",
                connection);
            command.Parameters.AddWithValue("@key", OriginalTableUpdatedAtKey);

            object? result = command.ExecuteScalar();
            return result?.ToString()?.Trim();
        }

        private string GetCurrentWebViewName()
        {
            if (CT_MAPPING_CONTENT.Visibility == Visibility.Visible)
            {
                return _currentMappingTableKind switch
                {
                    MappingTableKind.Routing => "mapping-routing",
                    MappingTableKind.Reason => "mapping-reason",
                    _ => "mapping"
                };
            }

            if (CT_REPORT_SETUP_CONTENT.Visibility == Visibility.Visible)
            {
                return "report-setup";
            }

            if (CT_LOAD_CONTENT.Visibility == Visibility.Visible)
            {
                return "load";
            }

            if (CT_MAIN_CONTENT_SCROLL.Visibility == Visibility.Visible)
            {
                return "main";
            }

            return "home";
        }
    }
}
