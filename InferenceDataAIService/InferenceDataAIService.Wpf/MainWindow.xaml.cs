using System.Collections.ObjectModel;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Microsoft.Data.Sqlite;
using Forms = System.Windows.Forms;

namespace InferenceDataAIService.Wpf;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<ExcelItem> _items = [];
    private readonly string _serviceDirectory;
    private readonly string _databasePath;
    private readonly ConcurrentQueue<string> _pendingLogs = new();
    private readonly DispatcherTimer _logFlushTimer;

    public MainWindow()
    {
        InitializeComponent();
        _serviceDirectory = FindServiceDirectory();
        _databasePath = Path.Combine(_serviceDirectory, "outputs", "universal-grid", "InputDataFinish.sqlite");
        FilesGrid.ItemsSource = _items;
        _logFlushTimer = new DispatcherTimer(DispatcherPriority.Background) { Interval = TimeSpan.FromMilliseconds(120) };
        _logFlushTimer.Tick += (_, _) => FlushPendingLogs();
        _logFlushTimer.Start();
        Log($"서비스 폴더: {_serviceDirectory}");
        LoadStoredFiles();
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
        var files = await Task.Run(() => Directory.EnumerateFiles(dialog.SelectedPath, "*.*", SearchOption.AllDirectories)
            .Where(path => new[] { ".xlsx", ".xlsm", ".xls", ".xlsb" }.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase) && !Path.GetFileName(path).StartsWith("~$"))
            .ToList());
        AddFiles(files);
    }

    private void Refresh_Click(object sender, RoutedEventArgs e) => LoadStoredFiles();

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
                SELECT w.status, ar.overall_status, ar.overall_decision, ar.dashboard_html_path
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
            item.ResultHtmlPath = reader.IsDBNull(3) ? null : reader.GetString(3);
            item.Progress = reader.IsDBNull(1) ? "DB 적재 완료" : "분석 완료";
        }
        catch (Exception ex) { item.Progress = "상태 조회 오류"; Log(ex.Message); }
        return item;
    }

    private async void AnalyzeSelected_Click(object sender, RoutedEventArgs e)
    {
        var selected = FilesGrid.SelectedItems.Cast<ExcelItem>().ToList();
        if (selected.Count == 0) return;
        FilesGrid.IsEnabled = false;
        try
        {
            foreach (var item in selected)
            {
                item.Progress = "DB 적재 중"; FilesGrid.Items.Refresh();
                await RunPythonAsync("inference_data_ai_cli.py", "com-index", "--input", item.FullPath, "--dataset", "InputDataFinish", "--db", _databasePath, "--covered-cell-mode", "blank", "--verify-after-import", "--include-hidden");
                item.Progress = "AI 분석 중"; FilesGrid.Items.Refresh();
                await RunPythonAsync("inference_data_ai_analysis_runner.py", "--service-dir", _serviceDirectory, "--db", _databasePath, "--source", item.FullPath, "--dataset", "InputDataFinish", "--replace-auto-draft");
                ReplaceItem(item);
            }
            Log("선택한 Excel 분석이 완료되었습니다.");
        }
        catch (Exception ex) { Log($"분석 실패: {ex.Message}"); System.Windows.MessageBox.Show(ex.Message, "분석 실패", MessageBoxButton.OK, MessageBoxImage.Error); }
        finally { FilesGrid.IsEnabled = true; }
    }

    private async Task RunPythonAsync(string script, params string[] arguments)
    {
        var executable = Environment.GetEnvironmentVariable("INFERENCE_DATA_AI_PYTHON");
        if (string.IsNullOrWhiteSpace(executable)) executable = "python";
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
        info.ArgumentList.Add(Path.Combine(_serviceDirectory, script));
        foreach (var argument in arguments) info.ArgumentList.Add(argument);
        using var process = Process.Start(info) ?? throw new InvalidOperationException("Python 실행을 시작하지 못했습니다.");
        // Never synchronously invoke the UI for each Codex output line: it can
        // produce thousands of lines and starve repaint/input while analysing.
        process.OutputDataReceived += (_, e) => { if (e.Data is not null) QueueLog(e.Data); };
        process.ErrorDataReceived += (_, e) => { if (e.Data is not null) QueueLog(e.Data); };
        process.BeginOutputReadLine(); process.BeginErrorReadLine(); await process.WaitForExitAsync();
        if (process.ExitCode != 0) throw new InvalidOperationException($"{script} 실행 실패: {process.ExitCode}");
    }

    private void ReplaceItem(ExcelItem item)
    {
        var index = _items.IndexOf(item); _items[index] = ReadStatus(item.FullPath); FilesGrid.SelectedItem = _items[index];
    }

    private void FilesGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (FilesGrid.SelectedItem is ExcelItem item) ShowResult(item);
    }

    private void OpenResult_Click(object sender, RoutedEventArgs e)
    {
        if (FilesGrid.SelectedItem is ExcelItem item) ShowResult(item);
    }

    private void ShowResult(ExcelItem item)
    {
        ResultTitle.Text = $"분석 결과 — {item.FileName}";
        if (!string.IsNullOrWhiteSpace(item.ResultHtmlPath) && File.Exists(item.ResultHtmlPath))
        {
            // WPF WebBrowser can ignore a local file's UTF-8 declaration and fall
            // back to the system ANSI code page. Passing decoded Unicode prevents
            // Korean analysis text from being rendered as mojibake.
            ResultBrowser.NavigateToString(File.ReadAllText(item.ResultHtmlPath, Encoding.UTF8));
        }
        else ResultBrowser.NavigateToString("<html><body style='font-family:Segoe UI;padding:24px'><h2>아직 분석 결과가 없습니다.</h2><p>목록에서 우클릭한 뒤 ‘선택 파일 분석 시작’을 선택하세요.</p></body></html>");
    }

    private void Log(string text)
    {
        QueueLog(text);
    }

    private void QueueLog(string text) => _pendingLogs.Enqueue(text);

    private void FlushPendingLogs()
    {
        if (_pendingLogs.IsEmpty) return;
        var buffer = new StringBuilder();
        for (var count = 0; count < 150 && _pendingLogs.TryDequeue(out var line); count++)
            buffer.Append('[').Append(DateTime.Now.ToString("HH:mm:ss")).Append("] ").AppendLine(line);
        LogText.AppendText(buffer.ToString());
        LogText.ScrollToEnd();
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
}
