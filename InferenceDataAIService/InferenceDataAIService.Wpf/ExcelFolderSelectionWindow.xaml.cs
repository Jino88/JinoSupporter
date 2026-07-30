using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Forms = System.Windows.Forms;
using WpfMessageBox = System.Windows.MessageBox;

namespace InferenceDataAIService.Wpf;

public partial class ExcelFolderSelectionWindow : Window
{
    private readonly ObservableCollection<ExcelSearchFolderEntry>
        _folders = [];

    internal ExcelFolderSelectionWindow()
    {
        InitializeComponent();
        FoldersGrid.ItemsSource = _folders;
        foreach (var path in ExcelFolderHistoryStore.Load())
            _folders.Add(new ExcelSearchFolderEntry(path));
        UpdateSummary();
        Loaded += (_, _) => FolderPathEntryText.Focus();
    }

    internal IReadOnlyList<string> SelectedFolders { get; private set; } =
        [];

    private void CloseButton_Click(
        object sender,
        RoutedEventArgs e) =>
        Close();

    private void BrowseFolder_Click(
        object sender,
        RoutedEventArgs e)
    {
        using var folderDialog = new Forms.FolderBrowserDialog
        {
            Description =
                "Excel 파일을 검색할 폴더를 추가하세요. 여러 폴더는 이 버튼을 반복해서 눌러 추가할 수 있습니다.",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = false,
        };
        if (folderDialog.ShowDialog() != Forms.DialogResult.OK)
            return;

        AddFolder(folderDialog.SelectedPath);
    }

    private void AddTypedFolder_Click(
        object sender,
        RoutedEventArgs e) =>
        AddFolder(FolderPathEntryText.Text);

    private void FolderPathEntryText_KeyDown(
        object sender,
        System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
            return;

        AddFolder(FolderPathEntryText.Text);
        e.Handled = true;
    }

    private void AddFolder(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return;

        string normalized;
        try
        {
            normalized = Path.GetFullPath(value.Trim().Trim('"'));
        }
        catch (
            Exception exception
        ) when (
            exception is ArgumentException
            or NotSupportedException
            or PathTooLongException)
        {
            WpfMessageBox.Show(
                "올바른 폴더 경로를 입력하세요.",
                "여러 검색 폴더 지정",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var existing = _folders.FirstOrDefault(folder =>
            string.Equals(
                folder.Path,
                normalized,
                StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            FoldersGrid.SelectedItem = existing;
            FoldersGrid.ScrollIntoView(existing);
            FolderPathEntryText.SelectAll();
            return;
        }

        var entry = new ExcelSearchFolderEntry(normalized);
        _folders.Add(entry);
        FoldersGrid.SelectedItem = entry;
        FoldersGrid.ScrollIntoView(entry);
        FolderPathEntryText.Clear();
        UpdateSummary();
    }

    private void RemoveSelectedFolders_Click(
        object sender,
        RoutedEventArgs e)
    {
        var selected = FoldersGrid.SelectedItems
            .OfType<ExcelSearchFolderEntry>()
            .ToArray();
        foreach (var entry in selected)
            _folders.Remove(entry);
        UpdateSummary();
    }

    private void ClearFolders_Click(
        object sender,
        RoutedEventArgs e)
    {
        _folders.Clear();
        UpdateSummary();
    }

    private void Confirm_Click(
        object sender,
        RoutedEventArgs e)
    {
        if (_folders.Count == 0)
        {
            WpfMessageBox.Show(
                "검색할 폴더를 하나 이상 추가하세요.",
                "여러 검색 폴더 지정",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var unavailable = _folders
            .Where(folder => !folder.Exists)
            .Select(folder => folder.Path)
            .ToArray();
        if (unavailable.Length > 0)
        {
            WpfMessageBox.Show(
                "찾을 수 없는 폴더가 있습니다. 연결 상태를 확인하거나 목록에서 삭제하세요.\n\n"
                + string.Join(
                    Environment.NewLine,
                    unavailable.Take(5))
                + (unavailable.Length > 5
                    ? $"\n외 {unavailable.Length - 5:N0}개"
                    : string.Empty),
                "여러 검색 폴더 지정",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var selected = _folders
            .Select(folder => folder.Path)
            .ToArray();
        try
        {
            ExcelFolderHistoryStore.Save(selected);
        }
        catch (
            Exception exception
        ) when (
            exception is IOException
            or UnauthorizedAccessException
            or ArgumentException
            or NotSupportedException)
        {
            WpfMessageBox.Show(
                "검색 폴더 이력을 저장하지 못했습니다.\n\n"
                + exception.Message,
                "여러 검색 폴더 지정",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        SelectedFolders = selected;
        DialogResult = true;
    }

    private void UpdateSummary()
    {
        var available = _folders.Count(folder => folder.Exists);
        FolderHistorySummary.Text = _folders.Count == 0
            ? "저장된 이력 없음 · 폴더를 직접 추가하세요."
            : $"검색 폴더 {_folders.Count:N0}개 · 사용 가능 {available:N0}개"
              + (_folders.Count == available
                  ? string.Empty
                  : $" · 찾을 수 없음 {_folders.Count - available:N0}개");
    }
}

internal sealed class ExcelSearchFolderEntry
{
    internal ExcelSearchFolderEntry(string path)
    {
        Path = path;
        Exists = Directory.Exists(path);
    }

    public string Path { get; }
    public bool Exists { get; }
    public string Availability => Exists
        ? "사용 가능"
        : "찾을 수 없음";
}
