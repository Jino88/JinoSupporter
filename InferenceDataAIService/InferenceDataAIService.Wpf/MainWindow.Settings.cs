using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Forms = System.Windows.Forms;
using WpfMessageBox = System.Windows.MessageBox;

namespace InferenceDataAIService.Wpf;

public partial class MainWindow
{
    private readonly ObservableCollection<PathSettingRow>
        _pathSettingRows = [];

    private void LoadPathSettingsUi() =>
        LoadPathSettingsUi(_pathSettings);

    private void LoadPathSettingsUi(
        AppPathSettings displayedSettings)
    {
        _pathSettingRows.Clear();
        AddPathSetting(
            "output_root",
            "Excel DB 통합 폴더",
            displayedSettings.OutputRootDirectory,
            PathSettingKind.Directory,
            "이 폴더 하나만 지정하면 DB, Excel 보관함, 처리 결과, 로그와 검수 파일을 하위 폴더로 자동 구성합니다.");
        AddPathSetting(
            "service",
            "서비스 폴더",
            displayedSettings.ServiceDirectory,
            PathSettingKind.Directory,
            "Python CLI와 스키마 파일이 있는 프로젝트 서비스 폴더");
        AddPathSetting(
            "python",
            "Python 실행 파일",
            displayedSettings.PythonExecutable,
            PathSettingKind.Executable,
            "python.exe 전체 경로 또는 PATH에서 찾을 명령 이름");
        AddPathSetting(
            "codex",
            "Codex 실행 파일",
            displayedSettings.CodexExecutable,
            PathSettingKind.Executable,
            "codex.cmd/exe 전체 경로 또는 PATH에서 찾을 명령 이름");
        PathSettingsItems.ItemsSource = _pathSettingRows;
        PathSettingsFileText.Text =
            AppPathSettingsStore.SettingsFilePath;
        PathSettingsState.Text =
            "Excel DB 관련 경로는 통합 폴더 아래에 자동 구성됩니다.";
    }

    private void AddPathSetting(
        string key,
        string label,
        string value,
        PathSettingKind kind,
        string description) =>
        _pathSettingRows.Add(new PathSettingRow(
            key,
            label,
            value,
            kind,
            description));

    private void BrowsePathSetting_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (sender is not FrameworkElement
            {
                DataContext: PathSettingRow row
            })
        {
            return;
        }

        if (row.Kind == PathSettingKind.Directory)
        {
            using var dialog = new Forms.FolderBrowserDialog
            {
                Description = $"{row.Label} 선택",
                UseDescriptionForTitle = true,
                ShowNewFolderButton = true,
                InitialDirectory = Directory.Exists(row.Value)
                    ? row.Value
                    : string.Empty,
            };
            if (dialog.ShowDialog() == Forms.DialogResult.OK)
            {
                row.Value = Path.GetFullPath(
                    dialog.SelectedPath);
                PathSettingsItems.Items.Refresh();
            }
            return;
        }

        using var fileDialog = new Forms.OpenFileDialog
        {
            Title = $"{row.Label} 선택",
            CheckFileExists = true,
            Multiselect = false,
            Filter = row.Kind == PathSettingKind.Executable
                ? "실행 파일 (*.exe;*.cmd;*.bat)|*.exe;*.cmd;*.bat|모든 파일 (*.*)|*.*"
                : "설정 파일 (*.sqlite;*.db;*.json)|*.sqlite;*.db;*.json|모든 파일 (*.*)|*.*",
        };
        if (Path.IsPathRooted(row.Value)
            && File.Exists(row.Value))
        {
            fileDialog.InitialDirectory =
                Path.GetDirectoryName(row.Value);
            fileDialog.FileName = Path.GetFileName(row.Value);
        }
        if (fileDialog.ShowDialog() == Forms.DialogResult.OK)
        {
            row.Value = Path.GetFullPath(
                fileDialog.FileName);
            PathSettingsItems.Items.Refresh();
        }
    }

    private void SavePathSettings_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_canonicalEvidenceBusy || _excelFolderSearchBusy)
        {
            WpfMessageBox.Show(
                "실행 중인 작업이 끝난 뒤 설정을 저장하세요.",
                "경로 설정",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        try
        {
            string Value(string key) =>
                _pathSettingRows.First(row => row.Key == key)
                    .Value.Trim();
            var defaults = AppPathSettings.CreateDefaults(
                Value("service"));
            var settings = new AppPathSettings
            {
                ServiceDirectory = Value("service"),
                PythonExecutable = Value("python"),
                CodexExecutable = Value("codex"),
                OutputRootDirectory = Value("output_root"),
            }.Normalize(defaults);
            settings.Validate();
            AppPathSettingsStore.Save(settings);
            ApplyPathSettings(settings);
            RecreatePathDependentClients();
            LoadPathSettingsUi();
            PathSettingsState.Text =
                "저장 완료 · DB와 모든 처리 산출물이 통합 폴더 기준으로 즉시 적용됩니다.";
            Log("경로 설정 저장·적용 완료");
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
            PathSettingsState.Text = "설정 저장 실패";
            WpfMessageBox.Show(
                exception.Message,
                "경로 설정",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void ResetPathSettings_Click(
        object sender,
        RoutedEventArgs e)
    {
        LoadPathSettingsUi(AppPathSettings.CreateDefaults(
            _discoveredServiceDirectory));
        PathSettingsState.Text =
            "기본값을 화면에 불러왔습니다. 저장해야 적용됩니다.";
    }

    private void OpenPathSettingsFolder_Click(
        object sender,
        RoutedEventArgs e)
    {
        var directory = Path.GetDirectoryName(
            AppPathSettingsStore.SettingsFilePath);
        if (string.IsNullOrWhiteSpace(directory))
            return;
        Directory.CreateDirectory(directory);
        Process.Start(new ProcessStartInfo
        {
            FileName = directory,
            UseShellExecute = true,
        });
    }

    private void RecreatePathDependentClients()
    {
        _canonicalEvidenceClient = new CanonicalEvidenceClient(
            _pathSettings);
        _workbookComparisonClient = new WorkbookComparisonClient(
            _pathSettings);
        CanonicalDbPathText.Text = _databasePath;
        ExcelArchivePathText.Text =
            _pathSettings.ExcelArchiveDirectory;
        ConfigureArchiveIngestSource(resetStages: true);
        RefreshIngestRetryAvailability();
        LoadLearnedResults();
    }
}
