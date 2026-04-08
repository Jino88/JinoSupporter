using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using MessageBox = System.Windows.MessageBox;
using DragDropEffects = System.Windows.DragDropEffects;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using Color = System.Windows.Media.Color;

namespace GraphMaker
{
    public class RemovedLimitRow
    {
        public int OriginalIndex { get; set; }
        public Dictionary<string, string> Values { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    }

    public class ProcessTrendFileInfo : GraphFileInfoBase
    {
        public bool UseFirstColumnAsSampleId { get; set; }
        public int MaxSamples { get; set; } = 100;
        public int SavedPlotColorIndex { get; set; }
        public bool UseQuadraticRegression { get; set; }
        public List<ProcessAxisMapping> SavedAxisMappings { get; set; } = new();
        public Dictionary<int, string> SavedGroupNames { get; set; } = new();
        public Dictionary<string, double> UpperSpecLimits { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, double> LowerSpecLimits { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, double> SpecLimits { get; } = new(StringComparer.OrdinalIgnoreCase);
        public bool? UseSecondThirdRowsAsLimits { get; set; }
        public int? AssignedSpecRowIndex { get; set; }
        public int? AssignedUpperRowIndex { get; set; }
        public int? AssignedLowerRowIndex { get; set; }
        public RemovedLimitRow? RemovedSpecRow { get; set; }
        public RemovedLimitRow? RemovedUpperRow { get; set; }
        public RemovedLimitRow? RemovedLowerRow { get; set; }
    }

    public class ProcessAxisMapping
    {
        public int XAxisProcessIndex { get; set; } = -1;
        public List<int> YAxisProcessIndices { get; set; } = new();
        public int GroupId { get; set; } = 1;
    }

    public partial class ProcessFlowTrendView : GraphViewBase
    {
        private sealed class ProcessTrendSavedState
        {
            public List<string> RawColumns { get; set; } = new();
            public List<List<string>> RawRows { get; set; } = new();
            public List<string> FilePaths { get; set; } = new();
            public string Delimiter { get; set; } = "\t";
            public int HeaderRowNumber { get; set; } = 1;
            public bool UseFirstColumnAsSampleId { get; set; }
            public int MaxSamples { get; set; } = 100;
            public int SavedPlotColorIndex { get; set; }
            public bool UseQuadraticRegression { get; set; }
            public List<ProcessAxisMapping> SavedAxisMappings { get; set; } = new();
            public Dictionary<int, string> SavedGroupNames { get; set; } = new();
            public Dictionary<string, double> UpperSpecLimits { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, double> LowerSpecLimits { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, double> SpecLimits { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class RawTableData
        {
            public List<string> Columns { get; init; } = new();
            public List<List<string>> Rows { get; init; } = new();
        }

        private sealed class AxisEntry
        {
            public int XAxisProcessIndex { get; init; }
            public string Name { get; init; } = string.Empty;
            public Dictionary<int, int> YAxisGroups { get; set; } = new();
            public override string ToString()
            {
                if (YAxisGroups.Count == 0)
                {
                    return $"X: {Name}";
                }

                var yNames = YAxisGroups.Keys
                    .OrderBy(index => index)
                    .Select(index => _owner._processNames.TryGetValue(index, out string? value) ? value : "?");
                return $"X: {Name} -> Y: {string.Join(", ", yNames)}";
            }

            public required ProcessFlowTrendView _owner { get; init; }
        }

        private sealed class YAxisOption
        {
            public int Index { get; init; }
            public CheckBox CheckBox { get; init; } = null!;
        }

        private ProcessTrendFileInfo? _currentFile;
        private readonly List<string> _loadedFilePaths = new();
        private readonly List<string> _pendingSetupFilePaths = new();
        private PlotModel? _trendPlotModel;
        private List<PreviewColorChoice> _colorChoices = new();
        private List<ProcessAxisMapping> _axisMappings = new();
        private readonly Dictionary<int, string> _pairGroupNames = new();
        private readonly List<AxisEntry> _axisEntries = new();
        private readonly List<YAxisOption> _yAxisOptions = new();
        private readonly Dictionary<int, string> _processNames = new();
        private bool _isUpdatingAxisUi;
        private string _pendingSetupDelimiter = "\t";
        private int _pendingSetupHeaderRowNumber = 1;
        private System.Windows.Point _dataGridDragStartPoint;
        private static readonly JsonSerializerOptions SavedStateJsonOptions = new() { WriteIndented = true };
        public PlotModel? TrendPlotModel
        {
            get => _trendPlotModel;
            set
            {
                _trendPlotModel = value;
                OnPropertyChanged(nameof(TrendPlotModel));
            }
        }

        public ProcessFlowTrendView()
        {
            InitializeComponent();
            DataContext = this;
            PreviewGraphViewBase.WirePreviewGrid(DataPreviewGrid, files => LoadFiles(files), DeleteSelectedPreviewRow);
            InitializeColorOptions();
            TrendPlotModel = new PlotModel { Title = "Process Trend" };
            EnsureDefaultPairGroups();
            NotifyWebModuleSnapshotChanged();
        }


        private void InitializeColorOptions()
        {
            PreviewGraphViewBase.InitializeDefaultColorOptions(PlotColorComboBox, _colorChoices);
        }

        private Color GetSelectedPlotColor(ProcessTrendFileInfo fileInfo)
        {
            if (fileInfo.SavedPlotColorIndex >= 0 && fileInfo.SavedPlotColorIndex < _colorChoices.Count)
            {
                return _colorChoices[fileInfo.SavedPlotColorIndex].Color;
            }

            return Colors.Black;
        }

        private void EnsureDefaultPairGroups()
        {
            if (_pairGroupNames.Count > 0)
            {
                return;
            }

            _pairGroupNames[1] = "Group 1";
            _pairGroupNames[2] = "Group 2";
        }

        private OxyColor GetGroupOxyColor(int groupId, Color fallbackColor)
        {
            if (_colorChoices.Count == 0)
            {
                return OxyColor.FromArgb(fallbackColor.A, fallbackColor.R, fallbackColor.G, fallbackColor.B);
            }

            int idx = Math.Abs(groupId - 1) % _colorChoices.Count;
            var c = _colorChoices[idx].Color;
            return OxyColor.FromArgb(c.A, c.R, c.G, c.B);
        }

        private string GetPairGroupName(int groupId)
        {
            if (_pairGroupNames.TryGetValue(groupId, out string? name) && !string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            return $"Group {groupId}";
        }

        private OxyColor GetSeriesOxyColor(int seriesIndex, Color fallbackColor)
        {
            if (_colorChoices.Count == 0)
            {
                return OxyColor.FromArgb(fallbackColor.A, fallbackColor.R, fallbackColor.G, fallbackColor.B);
            }

            int idx = Math.Abs(seriesIndex) % _colorChoices.Count;
            var c = _colorChoices[idx].Color;
            return OxyColor.FromArgb(c.A, c.R, c.G, c.B);
        }

        private OxyColor GetColumnOxyColor(string columnName, Color fallbackColor)
        {
            if (_colorChoices.Count == 0)
            {
                return OxyColor.FromArgb(fallbackColor.A, fallbackColor.R, fallbackColor.G, fallbackColor.B);
            }

            int idx = Math.Abs(StringComparer.OrdinalIgnoreCase.GetHashCode(columnName)) % _colorChoices.Count;
            var c = _colorChoices[idx].Color;
            return OxyColor.FromArgb(c.A, c.R, c.G, c.B);
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select data file",
                Filter = "Text/CSV files (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*",
                Multiselect = true
            };

            if (dialog.ShowDialog() == true)
            {
                HandleWebDroppedFiles(dialog.FileNames);
            }
        }

        private void FileDropBorder_DragEnter(object sender, DragEventArgs e)
        {
            bool hasFiles = e.Data.GetDataPresent(DataFormats.FileDrop);
            e.Effects = hasFiles ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private void FileDropBorder_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            if (e.Data.GetData(DataFormats.FileDrop) is string[] filePaths && filePaths.Length > 0)
            {
                HandleWebDroppedFiles(filePaths);
            }

            e.Handled = true;
        }

        private void LoadFiles(IEnumerable<string> filePaths)
        {
            Window? owner = Application.Current?.MainWindow ?? Window.GetWindow(this);
            var setupWindow = new DailyDataTrendSetupWindow(
                filePaths,
                requireFirstColumnDate: false,
                convertWideToSingleY: false)
            {
                Owner = owner
            };
            if (setupWindow.ShowDialog() != true || setupWindow.ResultFileInfo?.FullData == null)
            {
                return;
            }

            ApplyLoadedFileInfo(setupWindow.ResultFileInfo);
        }

        public void HandleWebDroppedFiles(IReadOnlyList<string> filePaths)
        {
            string[] accepted = filePaths
                .Where(path =>
                {
                    string ext = Path.GetExtension(path);
                    return string.Equals(ext, ".txt", StringComparison.OrdinalIgnoreCase) ||
                           string.Equals(ext, ".csv", StringComparison.OrdinalIgnoreCase);
                })
                .ToArray();

            if (accepted.Length > 0)
            {
                System.Diagnostics.Debug.WriteLine($"[GraphDrop:ProcessTrend] Accepted files: {string.Join(" | ", accepted)}");
                _pendingSetupFilePaths.Clear();
                _pendingSetupFilePaths.AddRange(accepted);
                _pendingSetupDelimiter = "\t";
                _pendingSetupHeaderRowNumber = 1;
                StatusText.Text = $"Dropped {accepted.Length:N0} file(s). Review settings and apply.";
                NotifyWebModuleSnapshotChanged();
                ApplyPendingWebSetup();
            }
        }

        private void ApplyPendingWebSetup()
        {
            if (_pendingSetupFilePaths.Count == 0)
            {
                return;
            }

            try
            {
                System.Diagnostics.Debug.WriteLine($"[GraphDrop:ProcessTrend] ApplyPendingWebSetup delimiter={_pendingSetupDelimiter}, headerRow={_pendingSetupHeaderRowNumber}, files={string.Join(" | ", _pendingSetupFilePaths)}");
                ProcessTrendFileInfo fileInfo = DailyDataTrendSetupLoader.LoadProcessTrendFileInfo(
                    _pendingSetupFilePaths,
                    _pendingSetupDelimiter,
                    Math.Max(1, _pendingSetupHeaderRowNumber),
                    requireFirstColumnDate: false,
                    convertWideToSingleY: false);

                ApplyLoadedFileInfo(fileInfo);
                _pendingSetupFilePaths.Clear();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GraphDrop:ProcessTrend] ApplyPendingWebSetup failed: {ex}");
                StatusText.Text = ex.Message;
                NotifyWebModuleSnapshotChanged();
            }
        }

        private static string ParseWebDelimiter(string delimiter)
        {
            return delimiter switch
            {
                "comma" => ",",
                "space" => " ",
                _ => "\t"
            };
        }

        private void ReloadCurrentFilesWithDelimiter(string delimiter)
        {
            if (_currentFile == null || _loadedFilePaths.Count == 0)
            {
                return;
            }

            _currentFile.Delimiter = delimiter;
            _currentFile.HeaderRowNumber = Math.Max(1, _currentFile.HeaderRowNumber);
            _currentFile.FilePath = _loadedFilePaths[0];
            _currentFile.Name = _loadedFilePaths.Count > 1
                ? string.Join(" + ", _loadedFilePaths.Select(Path.GetFileName))
                : Path.GetFileName(_loadedFilePaths[0]);
            _currentFile.FullData = null;

            LoadFileData(_currentFile);
            for (int i = 1; i < _loadedFilePaths.Count; i++)
            {
                MergeAdditionalFile(_loadedFilePaths[i], delimiter, _currentFile.HeaderRowNumber);
            }
        }

        private void ApplyLoadedFileInfo(ProcessTrendFileInfo fileInfo)
        {
            fileInfo.UseSecondThirdRowsAsLimits = null;
            _currentFile = fileInfo;

            _loadedFilePaths.Clear();
            foreach (string path in fileInfo.FilePath.Split('|', StringSplitOptions.RemoveEmptyEntries))
            {
                if (!string.IsNullOrWhiteSpace(path))
                {
                    _loadedFilePaths.Add(path);
                }
            }

            if (_currentFile.FullData == null)
            {
                return;
            }

            PreviewGraphViewBase.ApplyPreviewSummary(
                CurrentFileNameText,
                RowCountText,
                ColumnCountText,
                StatusText,
                string.IsNullOrWhiteSpace(fileInfo.Name)
                    ? string.Join(" + ", _loadedFilePaths.Select(Path.GetFileName))
                    : fileInfo.Name,
                _currentFile.FullData.Rows.Count,
                _currentFile.FullData.Columns.Count,
                $"Loaded {_loadedFilePaths.Count:N0} file(s) into Data Preview.");
            RefreshDataPreviewGrid();
            UpdateLimitRowSummary();
            ApplyCurrentFileSettingsToUi();
            UpdateAxisSelectionOptions();
            NotifyWebModuleSnapshotChanged();
        }

        private void LoadFile(string filePath, string delimiter, int headerRowNumber)
        {
            try
            {
                _currentFile = new ProcessTrendFileInfo
                {
                    Name = Path.GetFileName(filePath),
                    FilePath = filePath,
                    Delimiter = delimiter,
                    HeaderRowNumber = headerRowNumber,
                    UseSecondThirdRowsAsLimits = null
                };

                _loadedFilePaths.Clear();
                _loadedFilePaths.Add(filePath);
                LoadFileData(_currentFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while loading file: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadFileData(ProcessTrendFileInfo fileInfo)
        {
            var lines = File.ReadAllLines(fileInfo.FilePath);
            if (lines.Length == 0)
            {
                MessageBox.Show("File is empty.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var headerIndex = fileInfo.HeaderRowNumber - 1;
            if (headerIndex < 0 || headerIndex >= lines.Length)
            {
                MessageBox.Show("Invalid header row number.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var headerTokens = GraphMakerTableHelper.SplitLine(lines[headerIndex], fileInfo.Delimiter);
            if (headerTokens.Length == 0)
            {
                MessageBox.Show("Cannot read header. Check the delimiter.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var uniqueHeaders = GraphMakerTableHelper.BuildUniqueHeaders(headerTokens);
            var table = new DataTable();
            foreach (var header in uniqueHeaders)
            {
                table.Columns.Add(header);
            }

            fileInfo.UpperSpecLimits.Clear();
            fileInfo.LowerSpecLimits.Clear();
            fileInfo.SpecLimits.Clear();
            fileInfo.AssignedSpecRowIndex = null;
            fileInfo.AssignedUpperRowIndex = null;
            fileInfo.AssignedLowerRowIndex = null;
            fileInfo.RemovedSpecRow = null;
            fileInfo.RemovedUpperRow = null;
            fileInfo.RemovedLowerRow = null;

            int dataStartIndex = headerIndex + 1;

            for (var i = dataStartIndex; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i]))
                {
                    continue;
                }

                var values = GraphMakerTableHelper.SplitLine(lines[i], fileInfo.Delimiter);
                if (values.Length == 0)
                {
                    continue;
                }

                var row = table.NewRow();
                for (var col = 0; col < Math.Min(values.Length, table.Columns.Count); col++)
                {
                    row[col] = values[col];
                }

                table.Rows.Add(row);
            }

            fileInfo.FullData = table;
            PreviewGraphViewBase.ApplyPreviewSummary(CurrentFileNameText, RowCountText, ColumnCountText, StatusText, fileInfo.Name, table.Rows.Count, table.Columns.Count,
                "File loaded. Additional files are appended below by matching column headers. Missing columns stay blank.");
            RefreshDataPreviewGrid();
            UpdateLimitRowSummary();
            ApplyCurrentFileSettingsToUi();
            UpdateAxisSelectionOptions();
        }

        private void MergeAdditionalFile(string filePath, string delimiter, int headerRowNumber)
        {
            if (_currentFile == null || _currentFile.FullData == null)
            {
                return;
            }

            var temp = new ProcessTrendFileInfo
            {
                Name = Path.GetFileName(filePath),
                FilePath = filePath,
                Delimiter = delimiter,
                HeaderRowNumber = headerRowNumber,
                UseFirstColumnAsSampleId = _currentFile.UseFirstColumnAsSampleId,
                MaxSamples = _currentFile.MaxSamples,
                SavedPlotColorIndex = _currentFile.SavedPlotColorIndex,
                UseQuadraticRegression = _currentFile.UseQuadraticRegression,
                UseSecondThirdRowsAsLimits = _currentFile.UseSecondThirdRowsAsLimits
            };

            LoadFileData(temp);
            if (temp.FullData == null)
            {
                return;
            }

            var merged = _currentFile.FullData;
            var extra = temp.FullData;
            foreach (DataColumn column in extra.Columns)
            {
                if (!merged.Columns.Contains(column.ColumnName))
                {
                    merged.Columns.Add(column.ColumnName);
                }
            }

            foreach (DataRow extraRow in extra.Rows)
            {
                DataRow newRow = merged.NewRow();
                foreach (DataColumn column in extra.Columns)
                {
                    newRow[column.ColumnName] = extraRow[column]?.ToString() ?? string.Empty;
                }
                merged.Rows.Add(newRow);
            }

            foreach (DataColumn column in extra.Columns)
            {
                string columnName = column.ColumnName;
                if (temp.UpperSpecLimits.TryGetValue(columnName, out double usl))
                {
                    _currentFile.UpperSpecLimits[columnName] = usl;
                }
                if (temp.LowerSpecLimits.TryGetValue(columnName, out double lsl))
                {
                    _currentFile.LowerSpecLimits[columnName] = lsl;
                }
                if (temp.SpecLimits.TryGetValue(columnName, out double spec))
                {
                    _currentFile.SpecLimits[columnName] = spec;
                }
            }

            _loadedFilePaths.Add(filePath);
            PreviewGraphViewBase.ApplyPreviewSummary(
                CurrentFileNameText,
                RowCountText,
                ColumnCountText,
                StatusText,
                string.Join(" + ", _loadedFilePaths.Select(Path.GetFileName)),
                merged.Rows.Count,
                merged.Columns.Count,
                $"Merged {_loadedFilePaths.Count} files into one Data Preview. Matching headers were combined and missing values left blank.");
            UpdateLimitRowSummary();
            RefreshDataPreviewGrid();
            UpdateAxisSelectionOptions();
        }

        private void RemoveFileButton_Click(object sender, RoutedEventArgs e)
        {
            _currentFile = null;
            _loadedFilePaths.Clear();
            DataPreviewGrid.ItemsSource = null;
            TrendPlotModel = new PlotModel { Title = "Process Trend" };
            PreviewGraphViewBase.ResetPreviewSummary(CurrentFileNameText, RowCountText, ColumnCountText, StatusText, "(No file)", "Drop file into preview grid or click Browse to load data.");
            SpecRowTextBlock.Text = "SPEC row: not assigned";
            UpperRowTextBlock.Text = "USL row: not assigned";
            LowerRowTextBlock.Text = "LSL row: not assigned";
            _axisMappings.Clear();
            UpdateAxisSelectionOptions();
            ApplyCurrentFileSettingsToUi();
        }

        private void DataPreviewGrid_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || _currentFile?.FullData == null)
            {
                _dataGridDragStartPoint = e.GetPosition(DataPreviewGrid);
                return;
            }

            System.Windows.Point currentPosition = e.GetPosition(DataPreviewGrid);
            if (Math.Abs(currentPosition.X - _dataGridDragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(currentPosition.Y - _dataGridDragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            if (DataPreviewGrid.SelectedItem is not DataRowView rowView)
            {
                return;
            }

            int rowIndex = _currentFile.FullData.Rows.IndexOf(rowView.Row);
            if (rowIndex < 0)
            {
                rowIndex = DataPreviewGrid.SelectedIndex;
            }
            if (rowIndex < 0)
            {
                return;
            }

            DragDrop.DoDragDrop(DataPreviewGrid, new DataObject(typeof(int), rowIndex), DragDropEffects.Move);
        }

        private void DeletePreviewRowMenuItem_Click(object sender, RoutedEventArgs e)
        {
            DeleteSelectedPreviewRow();
        }

        private void DeleteSelectedPreviewRow()
        {
            if (_currentFile?.FullData == null || DataPreviewGrid.SelectedItem is not DataRowView rowView)
            {
                return;
            }

            int rowIndex = _currentFile.FullData.Rows.IndexOf(rowView.Row);
            if (rowIndex < 0)
            {
                return;
            }

            _currentFile.FullData.Rows.RemoveAt(rowIndex);
            RowCountText.Text = _currentFile.FullData.Rows.Count.ToString("N0");
            RefreshDataPreviewGrid();
            UpdateAxisSelectionOptions();
            StatusText.Text = "Selected row deleted from Data Preview.";
        }

        private void LimitRowDropBorder_DragEnter(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(int)))
            {
                return;
            }

            e.Effects = DragDropEffects.Move;
            if (sender is Border border)
            {
                border.BorderBrush = new SolidColorBrush(Colors.DodgerBlue);
            }
        }

        private void LimitRowDropBorder_DragLeave(object sender, DragEventArgs e)
        {
            if (sender is Border border)
            {
                border.BorderBrush = new SolidColorBrush(Color.FromRgb(0xB8, 0xC7, 0xDB));
            }
        }

        private void LimitRowDropBorder_Drop(object sender, DragEventArgs e)
        {
            if (_currentFile?.FullData == null || sender is not Border border || !e.Data.GetDataPresent(typeof(int)))
            {
                return;
            }

            int rowIndex = (int)e.Data.GetData(typeof(int));
            border.BorderBrush = new SolidColorBrush(Color.FromRgb(0xB8, 0xC7, 0xDB));
            AssignLimitRow(border.Tag as string, rowIndex);
        }

        private void AssignLimitRow(string? target, int rowIndex)
        {
            if (_currentFile?.FullData == null || string.IsNullOrWhiteSpace(target))
            {
                return;
            }

            var table = _currentFile.FullData;
            if (rowIndex < 0 || rowIndex >= table.Rows.Count)
            {
                return;
            }

            DataRow row = table.Rows[rowIndex];
            var values = table.Columns.Cast<DataColumn>()
                .ToDictionary(column => column.ColumnName, column => row[column]?.ToString() ?? string.Empty, StringComparer.OrdinalIgnoreCase);

            RestoreRemovedLimitRow(target);

            switch (target)
            {
                case "SPEC":
                    _currentFile.SpecLimits.Clear();
                    ApplyLimitValues(_currentFile.SpecLimits, values);
                    _currentFile.AssignedSpecRowIndex = rowIndex;
                    _currentFile.RemovedSpecRow = new RemovedLimitRow { OriginalIndex = rowIndex, Values = new Dictionary<string, string>(values, StringComparer.OrdinalIgnoreCase) };
                    break;
                case "USL":
                    _currentFile.UpperSpecLimits.Clear();
                    ApplyLimitValues(_currentFile.UpperSpecLimits, values);
                    _currentFile.AssignedUpperRowIndex = rowIndex;
                    _currentFile.RemovedUpperRow = new RemovedLimitRow { OriginalIndex = rowIndex, Values = new Dictionary<string, string>(values, StringComparer.OrdinalIgnoreCase) };
                    break;
                case "LSL":
                    _currentFile.LowerSpecLimits.Clear();
                    ApplyLimitValues(_currentFile.LowerSpecLimits, values);
                    _currentFile.AssignedLowerRowIndex = rowIndex;
                    _currentFile.RemovedLowerRow = new RemovedLimitRow { OriginalIndex = rowIndex, Values = new Dictionary<string, string>(values, StringComparer.OrdinalIgnoreCase) };
                    break;
                default:
                    return;
            }

            table.Rows.Remove(row);
            RowCountText.Text = table.Rows.Count.ToString("N0");
            RefreshDataPreviewGrid();
            UpdateLimitRowSummary();
            UpdateAxisSelectionOptions();
            StatusText.Text = $"{target} row assigned from Data Preview.";
        }

        private void ClearLimitRowMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not MenuItem menuItem)
            {
                return;
            }

            ClearLimitRow(menuItem.CommandParameter as string);
        }

        private void EditLimitRowMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not MenuItem menuItem)
            {
                return;
            }

            EditLimitRowValues(menuItem.CommandParameter as string);
        }

        private void LimitRowBorder_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is not Border border)
            {
                return;
            }

            EditLimitRowValues(border.Tag as string);
            e.Handled = true;
        }

        private void ClearLimitRow(string? target)
        {
            if (_currentFile?.FullData == null || string.IsNullOrWhiteSpace(target))
            {
                return;
            }

            bool restored = RestoreRemovedLimitRow(target);
            switch (target)
            {
                case "SPEC":
                    _currentFile.SpecLimits.Clear();
                    _currentFile.AssignedSpecRowIndex = null;
                    _currentFile.RemovedSpecRow = null;
                    break;
                case "USL":
                    _currentFile.UpperSpecLimits.Clear();
                    _currentFile.AssignedUpperRowIndex = null;
                    _currentFile.RemovedUpperRow = null;
                    break;
                case "LSL":
                    _currentFile.LowerSpecLimits.Clear();
                    _currentFile.AssignedLowerRowIndex = null;
                    _currentFile.RemovedLowerRow = null;
                    break;
                default:
                    return;
            }

            if (restored)
            {
                RowCountText.Text = _currentFile.FullData.Rows.Count.ToString("N0");
                RefreshDataPreviewGrid();
                UpdateAxisSelectionOptions();
                StatusText.Text = $"{target} row cleared and restored to Data Preview.";
            }

            UpdateLimitRowSummary();
        }

        private bool RestoreRemovedLimitRow(string target)
        {
            if (_currentFile?.FullData == null)
            {
                return false;
            }

            RemovedLimitRow? removedRow = target switch
            {
                "SPEC" => _currentFile.RemovedSpecRow,
                "USL" => _currentFile.RemovedUpperRow,
                "LSL" => _currentFile.RemovedLowerRow,
                _ => null
            };

            if (removedRow == null)
            {
                return false;
            }

            DataTable table = _currentFile.FullData;
            DataRow restoredRow = table.NewRow();
            foreach (DataColumn column in table.Columns)
            {
                if (removedRow.Values.TryGetValue(column.ColumnName, out string? value))
                {
                    restoredRow[column.ColumnName] = value;
                }
            }

            int insertIndex = Math.Max(0, Math.Min(removedRow.OriginalIndex, table.Rows.Count));
            table.Rows.InsertAt(restoredRow, insertIndex);
            return true;
        }

        private static void ApplyLimitValues(Dictionary<string, double> target, IReadOnlyDictionary<string, string> values)
        {
            foreach (var pair in values)
            {
                if (GraphMakerParsingHelper.TryParseDouble(pair.Value, out double numeric))
                {
                    target[pair.Key] = numeric;
                }
            }
        }

        private void EditLimitRowValues(string? target)
        {
            if (_currentFile?.FullData == null || string.IsNullOrWhiteSpace(target))
            {
                return;
            }

            DataTable table = _currentFile.FullData;
            IReadOnlyDictionary<string, string> existingValues = GetExistingLimitValues(target);
            var window = new LimitValuesWindow(
                table.Columns.Cast<DataColumn>().Select(column => column.ColumnName),
                existingValues)
            {
                Owner = Window.GetWindow(this),
                Title = $"Edit {target} Values"
            };

            if (window.ShowDialog() != true)
            {
                return;
            }

            RestoreRemovedLimitRow(target);

            switch (target)
            {
                case "SPEC":
                    _currentFile.SpecLimits.Clear();
                    ApplyLimitValues(_currentFile.SpecLimits, window.Values);
                    _currentFile.AssignedSpecRowIndex = null;
                    _currentFile.RemovedSpecRow = null;
                    break;
                case "USL":
                    _currentFile.UpperSpecLimits.Clear();
                    ApplyLimitValues(_currentFile.UpperSpecLimits, window.Values);
                    _currentFile.AssignedUpperRowIndex = null;
                    _currentFile.RemovedUpperRow = null;
                    break;
                case "LSL":
                    _currentFile.LowerSpecLimits.Clear();
                    ApplyLimitValues(_currentFile.LowerSpecLimits, window.Values);
                    _currentFile.AssignedLowerRowIndex = null;
                    _currentFile.RemovedLowerRow = null;
                    break;
                default:
                    return;
            }

            RowCountText.Text = _currentFile.FullData.Rows.Count.ToString("N0");
            RefreshDataPreviewGrid();
            UpdateLimitRowSummary();
            UpdateAxisSelectionOptions();
            StatusText.Text = $"{target} values updated manually.";
        }

        private IReadOnlyDictionary<string, string> GetExistingLimitValues(string target)
        {
            Dictionary<string, double> source = target switch
            {
                "SPEC" => _currentFile?.SpecLimits ?? new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase),
                "USL" => _currentFile?.UpperSpecLimits ?? new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase),
                "LSL" => _currentFile?.LowerSpecLimits ?? new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase),
                _ => new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
            };

            return source.ToDictionary(pair => pair.Key, pair => pair.Value.ToString(CultureInfo.InvariantCulture), StringComparer.OrdinalIgnoreCase);
        }

        private void UpdateLimitRowSummary()
        {
            if (_currentFile == null)
            {
                return;
            }

            SpecRowTextBlock.Text = _currentFile.AssignedSpecRowIndex.HasValue
                ? $"SPEC row: #{_currentFile.AssignedSpecRowIndex.Value + 1}"
                : "SPEC row: not assigned";
            UpperRowTextBlock.Text = _currentFile.AssignedUpperRowIndex.HasValue
                ? $"USL row: #{_currentFile.AssignedUpperRowIndex.Value + 1}"
                : "USL row: not assigned";
            LowerRowTextBlock.Text = _currentFile.AssignedLowerRowIndex.HasValue
                ? $"LSL row: #{_currentFile.AssignedLowerRowIndex.Value + 1}"
                : "LSL row: not assigned";
        }

        private void ConfigureAxesButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile?.FullData == null)
            {
                MessageBox.Show("Please load a file first.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var processNames = GetProcessNames(_currentFile.FullData).ToList();
            if (processNames.Count < 2)
            {
                MessageBox.Show("At least two process columns are required.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (XAxisListBox.Items.Count > 0)
            {
                XAxisListBox.SelectedIndex = 0;
                XAxisListBox.ScrollIntoView(XAxisListBox.SelectedItem);
            }
        }

        private void ConfigurePairsButton_Click(object sender, RoutedEventArgs e)
        {
            ConfigureAxesButton_Click(sender, e);
        }

        private IEnumerable<string> GetProcessNames(DataTable data)
        {
            int startCol = _currentFile?.UseFirstColumnAsSampleId == true ? 1 : 0;
            for (int col = startCol; col < data.Columns.Count; col++)
            {
                yield return data.Columns[col].ColumnName;
            }
        }

        private void UpdateAxisSelectionSummary(IReadOnlyList<string> processNames)
        {
            PersistAxisSettings();
        }

        private void UpdateAxisSelectionOptions()
        {
            RebuildAxisSelectionUi();

            if (_currentFile?.FullData == null)
            {
                return;
            }

            if (_currentFile != null)
            {
                _pairGroupNames.Clear();
                if (_currentFile.SavedGroupNames.Count > 0)
                {
                    foreach (var group in _currentFile.SavedGroupNames.Where(g => g.Key > 0))
                    {
                        _pairGroupNames[group.Key] = string.IsNullOrWhiteSpace(group.Value) ? $"Group {group.Key}" : group.Value;
                    }
                }
                else
                {
                    EnsureDefaultPairGroups();
                }

                _axisMappings = _currentFile.SavedAxisMappings
                    .Select(mapping => new ProcessAxisMapping
                    {
                        XAxisProcessIndex = mapping.XAxisProcessIndex,
                        YAxisProcessIndices = mapping.YAxisProcessIndices.ToList(),
                        GroupId = mapping.GroupId
                    })
                    .ToList();

                RebuildAxisSelectionUi();
                UpdateAxisSelectionSummary(GetProcessNames(_currentFile.FullData).ToList());
            }
        }

        private void RefreshDataPreviewGrid()
        {
            if (_currentFile?.FullData == null)
            {
                DataPreviewGrid.ItemsSource = null;
                return;
            }

            DataPreviewGrid.ItemsSource = _currentFile.FullData.DefaultView;
        }

        private void RebuildAxisSelectionUi()
        {
            _isUpdatingAxisUi = true;
            try
            {
                _axisEntries.Clear();
                _yAxisOptions.Clear();
                _processNames.Clear();
                XAxisListBox.ItemsSource = null;
                YAxisItemsPanel.Children.Clear();
                SelectedYGroupListBox.Items.Clear();
                GroupManageComboBox.ItemsSource = null;
                GroupNameTextBox.Text = string.Empty;

                if (_currentFile?.FullData == null)
                {
                    return;
                }

                var processNames = GetProcessNames(_currentFile.FullData).ToList();
                for (int i = 0; i < processNames.Count; i++)
                {
                    _processNames[i] = processNames[i];
                }

                if (_pairGroupNames.Count == 0)
                {
                    EnsureDefaultPairGroups();
                }

                var savedByX = _axisMappings
                    .GroupBy(mapping => mapping.XAxisProcessIndex)
                    .ToDictionary(
                        group => group.Key,
                        group => group.SelectMany(mapping => mapping.YAxisProcessIndices.Select(y => new { Y = y, mapping.GroupId }))
                            .GroupBy(item => item.Y)
                            .ToDictionary(item => item.Key, item => item.Last().GroupId));

                for (int i = 0; i < processNames.Count; i++)
                {
                    var saved = savedByX.TryGetValue(i, out var existing) ? existing : null;
                    _axisEntries.Add(new AxisEntry
                    {
                        _owner = this,
                        XAxisProcessIndex = i,
                        Name = processNames[i],
                        YAxisGroups = saved ?? new Dictionary<int, int>()
                    });

                    var checkBox = new CheckBox
                    {
                        Content = processNames[i],
                        Margin = new Thickness(0, 0, 0, 6)
                    };
                    checkBox.Checked += YAxisCheckBox_Changed;
                    checkBox.Unchecked += YAxisCheckBox_Changed;
                    YAxisItemsPanel.Children.Add(checkBox);
                    _yAxisOptions.Add(new YAxisOption { Index = i, CheckBox = checkBox });
                }

                if (_axisMappings.Count == 0 && processNames.Count > 1 && _axisEntries.Count > 0)
                {
                    int defaultGroupId = _pairGroupNames.Keys.Min();
                    foreach (int index in Enumerable.Range(1, processNames.Count - 1))
                    {
                        _axisEntries[0].YAxisGroups[index] = defaultGroupId;
                    }
                    CaptureAxisMappingsFromUi();
                }

                RefreshGroupOptions(_pairGroupNames.Keys.Min());
                XAxisListBox.ItemsSource = _axisEntries;
                if (XAxisListBox.Items.Count > 0)
                {
                    XAxisListBox.SelectedIndex = 0;
                }
            }
            finally
            {
                _isUpdatingAxisUi = false;
            }

            RefreshSelectedYGroupList();
        }

        private void XAxisListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (XAxisListBox.SelectedItem is AxisEntry selected)
            {
                LoadAxisEntryToUi(selected);
            }
        }

        private void LoadAxisEntryToUi(AxisEntry? entry)
        {
            _isUpdatingAxisUi = true;
            try
            {
                foreach (var option in _yAxisOptions)
                {
                    option.CheckBox.IsChecked = entry != null &&
                                                option.Index != entry.XAxisProcessIndex &&
                                                entry.YAxisGroups.ContainsKey(option.Index);
                    option.CheckBox.IsEnabled = entry != null && option.Index != entry.XAxisProcessIndex;
                }
            }
            finally
            {
                _isUpdatingAxisUi = false;
            }

            RefreshSelectedYGroupList();
        }

        private void YAxisCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (_isUpdatingAxisUi || XAxisListBox.SelectedItem is not AxisEntry selected)
            {
                return;
            }

            int defaultGroupId = _pairGroupNames.Keys.Min();
            var preserved = _yAxisOptions
                .Where(option => option.CheckBox.IsChecked == true && option.Index != selected.XAxisProcessIndex)
                .Select(option => option.Index)
                .Distinct()
                .OrderBy(index => index)
                .ToDictionary(
                    index => index,
                    index => selected.YAxisGroups.TryGetValue(index, out int groupId) ? groupId : defaultGroupId);

            selected.YAxisGroups = preserved;
            CaptureAxisMappingsFromUi();
            RefreshSelectedYGroupList();
            XAxisListBox.Items.Refresh();
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var option in _yAxisOptions.Where(option => option.CheckBox.IsEnabled))
            {
                option.CheckBox.IsChecked = true;
            }
        }

        private void ClearAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var option in _yAxisOptions.Where(option => option.CheckBox.IsEnabled))
            {
                option.CheckBox.IsChecked = false;
            }
        }

        private void AddGroupButton_Click(object sender, RoutedEventArgs e)
        {
            int nextId = _pairGroupNames.Count == 0 ? 1 : _pairGroupNames.Keys.Max() + 1;
            _pairGroupNames[nextId] = $"Group {nextId}";
            RefreshGroupOptions(nextId);
            RefreshSelectedYGroupList();
            PersistAxisSettings();
        }

        private void RemoveGroupButton_Click(object sender, RoutedEventArgs e)
        {
            if (_pairGroupNames.Count <= 1 || GroupManageComboBox.SelectedValue is not int groupId || !_pairGroupNames.ContainsKey(groupId))
            {
                return;
            }

            int fallback = _pairGroupNames.Keys.First(key => key != groupId);
            _pairGroupNames.Remove(groupId);
            foreach (var entry in _axisEntries)
            {
                foreach (int yIndex in entry.YAxisGroups.Where(pair => pair.Value == groupId).Select(pair => pair.Key).ToList())
                {
                    entry.YAxisGroups[yIndex] = fallback;
                }
            }

            CaptureAxisMappingsFromUi();
            RefreshGroupOptions(fallback);
            RefreshSelectedYGroupList();
            XAxisListBox.Items.Refresh();
            PersistAxisSettings();
        }

        private void GroupManageComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingAxisUi || GroupManageComboBox.SelectedValue is not int groupId)
            {
                return;
            }

            GroupNameTextBox.Text = _pairGroupNames.TryGetValue(groupId, out string? name) ? name : string.Empty;
        }

        private void ApplyGroupNameButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupManageComboBox.SelectedValue is not int groupId || !_pairGroupNames.ContainsKey(groupId))
            {
                return;
            }

            string name = (GroupNameTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show("Group name cannot be empty.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _pairGroupNames[groupId] = name;
            RefreshGroupOptions(groupId);
            RefreshSelectedYGroupList();
            XAxisListBox.Items.Refresh();
            PersistAxisSettings();
        }

        private void RefreshGroupOptions(int selectedGroupId)
        {
            _isUpdatingAxisUi = true;
            try
            {
                GroupManageComboBox.ItemsSource = null;
                GroupManageComboBox.ItemsSource = _pairGroupNames.OrderBy(g => g.Key).ToList();
                GroupManageComboBox.SelectedValue = selectedGroupId;
                GroupNameTextBox.Text = _pairGroupNames.TryGetValue(selectedGroupId, out string? name) ? name : string.Empty;
            }
            finally
            {
                _isUpdatingAxisUi = false;
            }
        }

        private void RefreshSelectedYGroupList()
        {
            SelectedYGroupListBox.Items.Clear();

            foreach (var entry in _axisEntries.OrderBy(item => item.XAxisProcessIndex))
            {
                foreach (var pair in entry.YAxisGroups.OrderBy(pair => pair.Key))
                {
                    var rowGrid = new Grid
                    {
                        Margin = new Thickness(0, 0, 0, 2),
                        Height = 22
                    };
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(118) });

                    var textBlock = new TextBlock
                    {
                        Text = $"X: {entry.Name} -> Y: {_processNames[pair.Key]}",
                        VerticalAlignment = System.Windows.VerticalAlignment.Center,
                        Margin = new Thickness(0, 0, 6, 0),
                        FontSize = 11
                    };

                    var comboBox = new ComboBox
                    {
                        ItemsSource = _pairGroupNames.OrderBy(g => g.Key).ToList(),
                        DisplayMemberPath = "Value",
                        SelectedValuePath = "Key",
                        SelectedValue = pair.Value,
                        Height = 20,
                        MinHeight = 20,
                        Padding = new Thickness(4, 0, 4, 0),
                        FontSize = 11,
                        VerticalAlignment = System.Windows.VerticalAlignment.Center
                    };
                    int yIndex = pair.Key;
                    comboBox.SelectionChanged += (_, _) =>
                    {
                        if (_isUpdatingAxisUi || comboBox.SelectedValue is not int selectedGroupId)
                        {
                            return;
                        }

                        entry.YAxisGroups[yIndex] = selectedGroupId;
                        CaptureAxisMappingsFromUi();
                        RefreshSelectedYGroupList();
                        XAxisListBox.Items.Refresh();
                        PersistAxisSettings();
                    };

                    Grid.SetColumn(textBlock, 0);
                    Grid.SetColumn(comboBox, 1);
                    rowGrid.Children.Add(textBlock);
                    rowGrid.Children.Add(comboBox);
                    SelectedYGroupListBox.Items.Add(rowGrid);
                }
            }
        }

        private void CaptureAxisMappingsFromUi()
        {
            _axisMappings = _axisEntries
                .SelectMany(entry => entry.YAxisGroups
                    .GroupBy(pair => pair.Value)
                    .Select(group => new ProcessAxisMapping
                    {
                        XAxisProcessIndex = entry.XAxisProcessIndex,
                        YAxisProcessIndices = group.Select(pair => pair.Key).OrderBy(index => index).ToList(),
                        GroupId = group.Key
                    }))
                .Where(mapping => mapping.YAxisProcessIndices.Count > 0)
                .ToList();

            if (_currentFile?.FullData != null)
            {
                UpdateAxisSelectionSummary(GetProcessNames(_currentFile.FullData).ToList());
            }
        }

        private void PersistAxisSettings()
        {
            if (_currentFile == null)
            {
                return;
            }

            SyncCurrentFileSettingsFromUi();
            _currentFile.SavedAxisMappings = _axisMappings
                .Select(mapping => new ProcessAxisMapping
                {
                    XAxisProcessIndex = mapping.XAxisProcessIndex,
                    YAxisProcessIndices = mapping.YAxisProcessIndices.ToList(),
                    GroupId = mapping.GroupId
                })
                .ToList();
            _currentFile.SavedGroupNames = _pairGroupNames.ToDictionary(kv => kv.Key, kv => kv.Value);
        }

        private static RawTableData CaptureRawTableData(DataTable table)
        {
            var result = new RawTableData();
            foreach (DataColumn column in table.Columns)
            {
                result.Columns.Add(column.ColumnName);
            }

            foreach (DataRow row in table.Rows)
            {
                var values = new List<string>(table.Columns.Count);
                foreach (DataColumn column in table.Columns)
                {
                    values.Add(row[column]?.ToString() ?? string.Empty);
                }

                result.Rows.Add(values);
            }

            return result;
        }

        private static DataTable BuildTableFromRawData(IReadOnlyList<string> columns, IReadOnlyList<List<string>> rows)
        {
            var table = new DataTable();
            foreach (string columnName in columns)
            {
                table.Columns.Add(string.IsNullOrWhiteSpace(columnName) ? "Column" : columnName);
            }

            foreach (var sourceRow in rows)
            {
                DataRow row = table.NewRow();
                for (int i = 0; i < table.Columns.Count && i < sourceRow.Count; i++)
                {
                    row[i] = sourceRow[i] ?? string.Empty;
                }

                table.Rows.Add(row);
            }

            return table;
        }

        private void SaveStateButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile == null)
            {
                MessageBox.Show("Load a file first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            PersistAxisSettings();

            var dialog = new SaveFileDialog
            {
                Title = "Save Process Trend Settings",
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                DefaultExt = ".json",
                AddExtension = true,
                FileName = $"{Path.GetFileNameWithoutExtension(_currentFile.Name)}_processtrend.json"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            RawTableData rawTable = _currentFile.FullData != null
                ? CaptureRawTableData(_currentFile.FullData)
                : new RawTableData();

            var state = new ProcessTrendSavedState
            {
                RawColumns = rawTable.Columns,
                RawRows = rawTable.Rows,
                FilePaths = _loadedFilePaths.ToList(),
                Delimiter = _currentFile.Delimiter,
                HeaderRowNumber = _currentFile.HeaderRowNumber,
                UseFirstColumnAsSampleId = _currentFile.UseFirstColumnAsSampleId,
                MaxSamples = _currentFile.MaxSamples,
                SavedPlotColorIndex = _currentFile.SavedPlotColorIndex,
                UseQuadraticRegression = _currentFile.UseQuadraticRegression,
                SavedAxisMappings = _currentFile.SavedAxisMappings
                    .Select(mapping => new ProcessAxisMapping
                    {
                        XAxisProcessIndex = mapping.XAxisProcessIndex,
                        YAxisProcessIndices = mapping.YAxisProcessIndices.ToList(),
                        GroupId = mapping.GroupId
                    })
                    .ToList(),
                SavedGroupNames = new Dictionary<int, string>(_currentFile.SavedGroupNames),
                UpperSpecLimits = new Dictionary<string, double>(_currentFile.UpperSpecLimits, StringComparer.OrdinalIgnoreCase),
                LowerSpecLimits = new Dictionary<string, double>(_currentFile.LowerSpecLimits, StringComparer.OrdinalIgnoreCase),
                SpecLimits = new Dictionary<string, double>(_currentFile.SpecLimits, StringComparer.OrdinalIgnoreCase)
            };

            File.WriteAllText(dialog.FileName, JsonSerializer.Serialize(state, SavedStateJsonOptions), Encoding.UTF8);
            StatusText.Text = $"Saved settings to {Path.GetFileName(dialog.FileName)}.";
        }

        private void LoadStateButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Load Process Trend Settings",
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                Multiselect = false
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            ProcessTrendSavedState? state = JsonSerializer.Deserialize<ProcessTrendSavedState>(File.ReadAllText(dialog.FileName), SavedStateJsonOptions);
            if (state == null)
            {
                MessageBox.Show("Invalid settings file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ApplySavedState(state);
            StatusText.Text = $"Loaded settings from {Path.GetFileName(dialog.FileName)}.";
        }

        private void ApplySavedState(ProcessTrendSavedState state)
        {
            if ((_currentFile == null || _currentFile.FullData == null) && state.RawColumns.Count > 0)
            {
                var restoredTable = BuildTableFromRawData(state.RawColumns, state.RawRows);
                _currentFile = new ProcessTrendFileInfo
                {
                    Name = state.FilePaths.Count > 0
                        ? string.Join(" + ", state.FilePaths.Select(Path.GetFileName))
                        : "Saved Process Trend Data",
                    FilePath = state.FilePaths.FirstOrDefault() ?? string.Empty,
                    Delimiter = state.Delimiter,
                    HeaderRowNumber = state.HeaderRowNumber,
                    FullData = restoredTable,
                    UseSecondThirdRowsAsLimits = null
                };

                _loadedFilePaths.Clear();
                _loadedFilePaths.AddRange(state.FilePaths.Where(path => !string.IsNullOrWhiteSpace(path)));
                RefreshDataPreviewGrid();
                PreviewGraphViewBase.ApplyPreviewSummary(
                    CurrentFileNameText,
                    RowCountText,
                    ColumnCountText,
                    StatusText,
                    _currentFile.Name,
                    restoredTable.Rows.Count,
                    restoredTable.Columns.Count,
                    "Loaded saved raw data from JSON.");
            }

            if ((_currentFile == null || _currentFile.FullData == null) && state.FilePaths.Count > 0)
            {
                var existingFiles = state.FilePaths.Where(File.Exists).ToList();
                if (existingFiles.Count == 0)
                {
                    MessageBox.Show("Saved data files were not found.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                LoadFile(existingFiles[0], state.Delimiter, state.HeaderRowNumber);
                for (int i = 1; i < existingFiles.Count; i++)
                {
                    MergeAdditionalFile(existingFiles[i], state.Delimiter, state.HeaderRowNumber);
                }
            }

            if (_currentFile == null)
            {
                return;
            }

            _currentFile.Delimiter = state.Delimiter;
            _currentFile.HeaderRowNumber = state.HeaderRowNumber;
            _currentFile.UseFirstColumnAsSampleId = state.UseFirstColumnAsSampleId;
            _currentFile.MaxSamples = state.MaxSamples;
            _currentFile.SavedPlotColorIndex = state.SavedPlotColorIndex;
            _currentFile.UseQuadraticRegression = state.UseQuadraticRegression;
            _currentFile.SavedAxisMappings = state.SavedAxisMappings
                .Select(mapping => new ProcessAxisMapping
                {
                    XAxisProcessIndex = mapping.XAxisProcessIndex,
                    YAxisProcessIndices = mapping.YAxisProcessIndices.ToList(),
                    GroupId = mapping.GroupId
                })
                .ToList();
            _currentFile.SavedGroupNames = new Dictionary<int, string>(state.SavedGroupNames);

            _currentFile.UpperSpecLimits.Clear();
            foreach (var pair in state.UpperSpecLimits)
            {
                _currentFile.UpperSpecLimits[pair.Key] = pair.Value;
            }

            _currentFile.LowerSpecLimits.Clear();
            foreach (var pair in state.LowerSpecLimits)
            {
                _currentFile.LowerSpecLimits[pair.Key] = pair.Value;
            }

            _currentFile.SpecLimits.Clear();
            foreach (var pair in state.SpecLimits)
            {
                _currentFile.SpecLimits[pair.Key] = pair.Value;
            }

            _currentFile.AssignedSpecRowIndex = null;
            _currentFile.AssignedUpperRowIndex = null;
            _currentFile.AssignedLowerRowIndex = null;
            _currentFile.RemovedSpecRow = null;
            _currentFile.RemovedUpperRow = null;
            _currentFile.RemovedLowerRow = null;

            ApplyCurrentFileSettingsToUi();
            UpdateLimitRowSummary();
            UpdateAxisSelectionOptions();
        }

        private void SyncCurrentFileSettingsFromUi()
        {
            if (_currentFile == null)
            {
                return;
            }

            _currentFile.UseFirstColumnAsSampleId = UseFirstColumnAsSampleIdCheckBox.IsChecked == true;
            _currentFile.UseQuadraticRegression = QuadraticDegreeRadio.IsChecked == true;
            _currentFile.SavedPlotColorIndex = PlotColorComboBox.SelectedIndex < 0 ? 0 : PlotColorComboBox.SelectedIndex;
            _currentFile.MaxSamples = int.TryParse(MaxSamplesTextBox.Text, out int maxSamples) && maxSamples > 0 ? maxSamples : 100;
        }

        private void ApplyCurrentFileSettingsToUi()
        {
            if (_currentFile == null)
            {
                UseFirstColumnAsSampleIdCheckBox.IsChecked = false;
                MaxSamplesTextBox.Text = "100";
                PlotColorComboBox.SelectedIndex = _colorChoices.Count > 0 ? 0 : -1;
                LinearDegreeRadio.IsChecked = true;
                QuadraticDegreeRadio.IsChecked = false;
                return;
            }

            UseFirstColumnAsSampleIdCheckBox.IsChecked = _currentFile.UseFirstColumnAsSampleId;
            MaxSamplesTextBox.Text = (_currentFile.MaxSamples > 0 ? _currentFile.MaxSamples : 100).ToString(CultureInfo.InvariantCulture);
            PlotColorComboBox.SelectedIndex = _currentFile.SavedPlotColorIndex >= 0 && _currentFile.SavedPlotColorIndex < _colorChoices.Count
                ? _currentFile.SavedPlotColorIndex
                : 0;
            LinearDegreeRadio.IsChecked = !_currentFile.UseQuadraticRegression;
            QuadraticDegreeRadio.IsChecked = _currentFile.UseQuadraticRegression;
        }

        private void GenerateTrendButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile?.FullData == null)
            {
                MessageBox.Show("Please load a file first.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            PersistAxisSettings();
            int maxSamples = _currentFile.MaxSamples > 0 ? _currentFile.MaxSamples : 100;

            var data = _currentFile.FullData;
            var useFirstColAsSampleId = _currentFile.UseFirstColumnAsSampleId;
            int startCol = useFirstColAsSampleId ? 1 : 0;

            if (data.Columns.Count <= startCol)
            {
                MessageBox.Show("No process columns found. Check data format.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int processCount = data.Columns.Count - startCol;
            if (processCount < 2)
            {
                MessageBox.Show("No process pairs to compare. At least two process columns are required.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (_axisMappings.Count == 0)
            {
                MessageBox.Show("Configure axis mappings first.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int rowsToPlot = Math.Min(data.Rows.Count, maxSamples);
            var selectedColor = GetSelectedPlotColor(_currentFile);
            bool useQuadratic = _currentFile.UseQuadraticRegression;
            var pairResults = new List<ProcessPairPlotResult>();
            int totalPoints = 0;

            var groupedMappings = _axisMappings
                .Where(mapping =>
                    mapping.XAxisProcessIndex >= 0 &&
                    mapping.XAxisProcessIndex < processCount &&
                    mapping.YAxisProcessIndices.Any(index => index >= 0 && index < processCount && index != mapping.XAxisProcessIndex))
                .Select(mapping => new ProcessAxisMapping
                {
                    XAxisProcessIndex = mapping.XAxisProcessIndex,
                    YAxisProcessIndices = mapping.YAxisProcessIndices
                        .Where(index => index >= 0 && index < processCount && index != mapping.XAxisProcessIndex)
                        .Distinct()
                        .ToList(),
                    GroupId = mapping.GroupId
                })
                .Where(mapping => mapping.YAxisProcessIndices.Count > 0)
                .GroupBy(mapping => mapping.GroupId)
                .OrderBy(g => g.Key)
                .ToList();

            foreach (var group in groupedMappings)
            {
                int groupId = group.Key;
                string groupName = GetPairGroupName(groupId);

                var pairModel = new PlotModel
                {
                    Title = groupName
                };
                pairModel.Axes.Add(new LinearAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = "X",
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot
                });
                pairModel.Axes.Add(new LinearAxis
                {
                    Position = AxisPosition.Left,
                    Title = "Y",
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot
                });

                var detailSb = new StringBuilder();
                detailSb.AppendLine($"Group: {groupName}");
                int groupPointCount = 0;
                var candidates = new List<ProcessTrendComputationCandidate>();

                foreach (var mapping in group)
                {
                        string leftName = data.Columns[startCol + mapping.XAxisProcessIndex].ColumnName;
                    foreach (int right in mapping.YAxisProcessIndices)
                    {
                        string rightName = data.Columns[startCol + right].ColumnName;
                        string pairTitle = $"{leftName} -> {rightName}";
                        var paired = CollectPairedValues(data, startCol, rowsToPlot, mapping.XAxisProcessIndex, right);
                        var seriesColor = GetColumnOxyColor(rightName, selectedColor);

                        var scatter = new ScatterSeries
                        {
                            Title = pairTitle,
                            MarkerType = MarkerType.Circle,
                            MarkerSize = 3,
                            MarkerFill = seriesColor,
                            RenderInLegend = true
                        };
                        foreach (var p in paired)
                        {
                            scatter.Points.Add(new ScatterPoint(p.X, p.Y));
                        }
                        pairModel.Series.Add(scatter);
                        AddLimitLines(pairModel, leftName, rightName, paired);
                        candidates.Add(new ProcessTrendComputationCandidate
                        {
                            PairTitle = pairTitle,
                            PlotModel = BuildSinglePairModel(pairTitle, leftName, rightName, paired, seriesColor, useQuadratic),
                            XAxisTitle = leftName,
                            YAxisTitle = rightName,
                            RawPoints = paired.ToList()
                        });

                        groupPointCount += paired.Count;
                        detailSb.AppendLine($"- {pairTitle} | n={paired.Count}");

                        if (paired.Count < 2)
                        {
                            detailSb.AppendLine("  Formula: insufficient data");
                            continue;
                        }

                        double minX = paired.Min(p => p.X);
                        double maxX = paired.Max(p => p.X);
                        if (Math.Abs(maxX - minX) <= 1e-12)
                        {
                            detailSb.AppendLine("  Formula: x range too small");
                            continue;
                        }

                        if (!useQuadratic && TryCalculateTrendLine(paired, out double slope, out double intercept, out double r2))
                        {
                            var trend = new LineSeries
                            {
                                Title = $"{pairTitle} Regression",
                                StrokeThickness = 2,
                                Color = seriesColor,
                                RenderInLegend = false
                            };
                            trend.Points.Add(new DataPoint(minX, slope * minX + intercept));
                            trend.Points.Add(new DataPoint(maxX, slope * maxX + intercept));
                            pairModel.Series.Add(trend);
                            detailSb.AppendLine($"  Formula: y = {slope:F4}x + {intercept:F4} | R^2 = {r2:F4}");
                        }
                        else if (useQuadratic &&
                            TryCalculateQuadraticTrendLine(paired, out double a, out double b, out double c, out double r2Q))
                        {
                            var trend = new LineSeries
                            {
                                Title = $"{pairTitle} Regression",
                                StrokeThickness = 2,
                                Color = seriesColor,
                                RenderInLegend = false
                            };
                            int segments = 60;
                            for (int i = 0; i <= segments; i++)
                            {
                                double x = minX + (maxX - minX) * i / segments;
                                double y = a * x * x + b * x + c;
                                trend.Points.Add(new DataPoint(x, y));
                            }
                            pairModel.Series.Add(trend);
                            detailSb.AppendLine($"  Formula: y = {a:F4}x^2 + {b:F4}x + {c:F4} | R^2 = {r2Q:F4}");
                        }
                        else
                        {
                            detailSb.AppendLine("  Formula: regression unavailable");
                        }
                    }
                }

                pairResults.Add(new ProcessPairPlotResult
                {
                    PairTitle = groupName,
                    PlotModel = pairModel,
                    DetailText = detailSb.ToString().TrimEnd(),
                    XAxisTitle = "X",
                    YAxisTitle = "Y",
                    RawPoints = candidates.Count > 0 ? candidates[0].RawPoints : new List<DataPoint>(),
                    ComputationCandidates = candidates
                });
                totalPoints += groupPointCount;
            }

            if (pairResults.Count == 0)
            {
                MessageBox.Show("No valid axis mappings were configured.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var resultWindow = new ProcessFlowTrendResultWindow(pairResults);
            Window? ownerWindow = Window.GetWindow(this);
            if (ownerWindow is not null && ownerWindow.IsVisible)
            {
                resultWindow.Owner = ownerWindow;
            }
            resultWindow.Show();

            StatusText.Text = $"Displayed {pairResults.Count:N0} group graphs (points {totalPoints:N0}).";
        }


        private PlotModel BuildSinglePairModel(
            string pairTitle,
            string leftName,
            string rightName,
            IReadOnlyList<DataPoint> paired,
            OxyColor seriesColor,
            bool useQuadratic)
        {
            var model = new PlotModel { Title = pairTitle };
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = leftName,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = rightName,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            var scatter = new ScatterSeries
            {
                Title = $"{pairTitle}_avg",
                MarkerType = MarkerType.Circle,
                MarkerSize = 3,
                MarkerFill = seriesColor
            };
            foreach (var p in paired)
            {
                scatter.Points.Add(new ScatterPoint(p.X, p.Y));
            }
            model.Series.Add(scatter);
            AddLimitLines(model, leftName, rightName, paired);

            if (paired.Count >= 2)
            {
                double minX = paired.Min(p => p.X);
                double maxX = paired.Max(p => p.X);
                if (Math.Abs(maxX - minX) > 1e-12)
                {
                    if (!useQuadratic && TryCalculateTrendLine(paired, out double slope, out double intercept, out _))
                    {
                        var trend = new LineSeries
                        {
                            Title = $"{pairTitle} Regression",
                            StrokeThickness = 2,
                            Color = seriesColor
                        };
                        trend.Points.Add(new DataPoint(minX, slope * minX + intercept));
                        trend.Points.Add(new DataPoint(maxX, slope * maxX + intercept));
                        model.Series.Add(trend);
                    }
                    else if (useQuadratic &&
                        TryCalculateQuadraticTrendLine(paired, out double a, out double b, out double c, out _))
                    {
                        var trend = new LineSeries
                        {
                            Title = $"{pairTitle} Regression",
                            StrokeThickness = 2,
                            Color = seriesColor
                        };
                        int segments = 60;
                        for (int i = 0; i <= segments; i++)
                        {
                            double x = minX + (maxX - minX) * i / segments;
                            double y = a * x * x + b * x + c;
                            trend.Points.Add(new DataPoint(x, y));
                        }
                        model.Series.Add(trend);
                    }
                }
            }

            return model;
        }

        private void AddLimitLines(PlotModel model, string xColumn, string yColumn, IReadOnlyList<DataPoint> paired)
        {
            if (_currentFile == null || paired.Count == 0)
            {
                return;
            }

            double minX = paired.Min(p => p.X);
            double maxX = paired.Max(p => p.X);
            double minY = paired.Min(p => p.Y);
            double maxY = paired.Max(p => p.Y);
            var fallbackColor = GetSelectedPlotColor(_currentFile);
            var xColumnColor = GetColumnOxyColor(xColumn, fallbackColor);
            var yColumnColor = GetColumnOxyColor(yColumn, fallbackColor);

            if (_currentFile.UpperSpecLimits.TryGetValue(xColumn, out double xUsl))
            {
                model.Series.Add(new LineSeries
                {
                    Title = $"{xColumn} USL",
                    Color = xColumnColor,
                    LineStyle = LineStyle.Dash,
                    StrokeThickness = 1.8,
                    RenderInLegend = false,
                    Points = { new DataPoint(xUsl, minY), new DataPoint(xUsl, maxY) }
                });
            }

            if (_currentFile.LowerSpecLimits.TryGetValue(xColumn, out double xLsl))
            {
                model.Series.Add(new LineSeries
                {
                    Title = $"{xColumn} LSL",
                    Color = xColumnColor,
                    LineStyle = LineStyle.Dash,
                    StrokeThickness = 1.8,
                    RenderInLegend = false,
                    Points = { new DataPoint(xLsl, minY), new DataPoint(xLsl, maxY) }
                });
            }

            if (_currentFile.UpperSpecLimits.TryGetValue(yColumn, out double yUsl))
            {
                model.Series.Add(new LineSeries
                {
                    Title = $"{yColumn} USL",
                    Color = yColumnColor,
                    LineStyle = LineStyle.Dash,
                    StrokeThickness = 1.8,
                    RenderInLegend = false,
                    Points = { new DataPoint(minX, yUsl), new DataPoint(maxX, yUsl) }
                });
            }

            if (_currentFile.LowerSpecLimits.TryGetValue(yColumn, out double yLsl))
            {
                model.Series.Add(new LineSeries
                {
                    Title = $"{yColumn} LSL",
                    Color = yColumnColor,
                    LineStyle = LineStyle.Dash,
                    StrokeThickness = 1.8,
                    RenderInLegend = false,
                    Points = { new DataPoint(minX, yLsl), new DataPoint(maxX, yLsl) }
                });
            }
        }

        private void OpenFileSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile == null)
            {
                MessageBox.Show("Load a file first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            LoadFileData(_currentFile);
        }

        private static bool TryCalculateTrendLine(
            IReadOnlyList<DataPoint> points,
            out double slope,
            out double intercept,
            out double rSquared)
        {
            slope = 0;
            intercept = 0;
            rSquared = 0;

            if (points.Count < 2)
            {
                return false;
            }

            double n = points.Count;
            double sumX = points.Sum(p => p.X);
            double sumY = points.Sum(p => p.Y);
            double sumXY = points.Sum(p => p.X * p.Y);
            double sumXX = points.Sum(p => p.X * p.X);
            double denominator = n * sumXX - sumX * sumX;

            if (Math.Abs(denominator) < 1e-12)
            {
                return false;
            }

            slope = (n * sumXY - sumX * sumY) / denominator;
            intercept = (sumY - slope * sumX) / n;

            double meanY = sumY / n;
            double ssTot = points.Sum(p => Math.Pow(p.Y - meanY, 2));
            double slopeLocal = slope;
            double interceptLocal = intercept;
            double ssRes = points.Sum(p =>
            {
                double predicted = slopeLocal * p.X + interceptLocal;
                return Math.Pow(p.Y - predicted, 2);
            });

            rSquared = ssTot <= 1e-12 ? 1.0 : 1.0 - (ssRes / ssTot);
            return true;
        }

        private static bool TryCalculateQuadraticTrendLine(
            IReadOnlyList<DataPoint> points,
            out double a,
            out double b,
            out double c,
            out double rSquared)
        {
            a = 0;
            b = 0;
            c = 0;
            rSquared = 0;

            if (points.Count < 3)
            {
                return false;
            }

            double n = points.Count;
            double sx = points.Sum(p => p.X);
            double sx2 = points.Sum(p => p.X * p.X);
            double sx3 = points.Sum(p => p.X * p.X * p.X);
            double sx4 = points.Sum(p => p.X * p.X * p.X * p.X);
            double sy = points.Sum(p => p.Y);
            double sxy = points.Sum(p => p.X * p.Y);
            double sx2y = points.Sum(p => p.X * p.X * p.Y);

            // Normal equation:
            // [sx4 sx3 sx2][a]   [sx2y]
            // [sx3 sx2 sx ][b] = [sxy ]
            // [sx2 sx  n  ][c]   [sy  ]
            if (!Solve3x3(
                sx4, sx3, sx2,
                sx3, sx2, sx,
                sx2, sx, n,
                sx2y, sxy, sy,
                out a, out b, out c))
            {
                return false;
            }

            double meanY = sy / n;
            double ssTot = points.Sum(p => Math.Pow(p.Y - meanY, 2));
            double aa = a;
            double bb = b;
            double cc = c;
            double ssRes = points.Sum(p =>
            {
                double predicted = aa * p.X * p.X + bb * p.X + cc;
                return Math.Pow(p.Y - predicted, 2);
            });

            rSquared = ssTot <= 1e-12 ? 1.0 : 1.0 - (ssRes / ssTot);
            return true;
        }

        private static bool Solve3x3(
            double a11, double a12, double a13,
            double a21, double a22, double a23,
            double a31, double a32, double a33,
            double b1, double b2, double b3,
            out double x1, out double x2, out double x3)
        {
            x1 = x2 = x3 = 0;

            double detA =
                a11 * (a22 * a33 - a23 * a32) -
                a12 * (a21 * a33 - a23 * a31) +
                a13 * (a21 * a32 - a22 * a31);

            if (Math.Abs(detA) < 1e-12)
            {
                return false;
            }

            double detX1 =
                b1 * (a22 * a33 - a23 * a32) -
                a12 * (b2 * a33 - a23 * b3) +
                a13 * (b2 * a32 - a22 * b3);

            double detX2 =
                a11 * (b2 * a33 - a23 * b3) -
                b1 * (a21 * a33 - a23 * a31) +
                a13 * (a21 * b3 - b2 * a31);

            double detX3 =
                a11 * (a22 * b3 - b2 * a32) -
                a12 * (a21 * b3 - b2 * a31) +
                b1 * (a21 * a32 - a22 * a31);

            x1 = detX1 / detA;
            x2 = detX2 / detA;
            x3 = detX3 / detA;
            return true;
        }

        private static string BuildTrendLog(
            DataTable data,
            int startCol,
            int rowsToPlot,
            int plottedSamples,
            int plottedPoints,
            Dictionary<int, List<double>> processValues,
            string formulaText,
            string r2Text,
            bool useQuadratic,
            ISet<string> selectedPairKeys)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Rows Used: {rowsToPlot:N0}");
            sb.AppendLine($"Samples Plotted: {plottedSamples:N0}");
            sb.AppendLine($"Scatter Points: {plottedPoints:N0}");
            sb.AppendLine($"Trend Line: {formulaText}");
            sb.AppendLine(r2Text);
            sb.AppendLine();
            sb.AppendLine("[Pairwise Regression]");

            int processCount = data.Columns.Count - startCol;
            for (int left = 0; left < processCount; left++)
            {
                for (int right = left + 1; right < processCount; right++)
                {
                    string pairKey = $"{left}-{right}";
                    int totalPairs = processCount * (processCount - 1) / 2;
                    bool useAllPairs = selectedPairKeys.Count == 0 || selectedPairKeys.Count == totalPairs;
                    if (!useAllPairs && !selectedPairKeys.Contains(pairKey))
                    {
                        continue;
                    }

                    string leftName = data.Columns[startCol + left].ColumnName;
                    string rightName = data.Columns[startCol + right].ColumnName;

                    var paired = CollectPairedValues(data, startCol, rowsToPlot, left, right);
                    if (paired.Count < 2)
                    {
                        sb.AppendLine($"{leftName} -> {rightName}: insufficient data");
                        continue;
                    }

                    if (useQuadratic &&
                        TryCalculateQuadraticTrendLine(paired, out double a, out double b, out double c, out double pairR2Q))
                    {
                        sb.AppendLine($"{leftName} -> {rightName}: y = {a:F4}x^2 + {b:F4}x + {c:F4}, R^2 = {pairR2Q:F4}, n={paired.Count}");
                    }
                    else if (!useQuadratic &&
                        TryCalculateTrendLine(paired, out double slope, out double intercept, out double pairR2))
                    {
                        sb.AppendLine($"{leftName} -> {rightName}: y = {slope:F4}x + {intercept:F4}, R^2 = {pairR2:F4}, n={paired.Count}");
                    }
                    else
                    {
                        sb.AppendLine($"{leftName} -> {rightName}: regression unavailable");
                    }
                }
            }

            sb.AppendLine();
            sb.AppendLine("[Process Transition Delta (Mean)]");

            for (int processIndex = 0; processIndex < data.Columns.Count - startCol - 1; processIndex++)
            {
                if (!processValues.TryGetValue(processIndex, out var leftValues) || leftValues.Count == 0)
                {
                    continue;
                }

                if (!processValues.TryGetValue(processIndex + 1, out var rightValues) || rightValues.Count == 0)
                {
                    continue;
                }

                double leftMean = leftValues.Average();
                double rightMean = rightValues.Average();
                double delta = rightMean - leftMean;

                string leftName = data.Columns[startCol + processIndex].ColumnName;
                string rightName = data.Columns[startCol + processIndex + 1].ColumnName;
                sb.AppendLine($"{leftName} -> {rightName}: mean {leftMean:F4} -> {rightMean:F4} (Delta {delta:+0.0000;-0.0000;0.0000})");
            }

            return sb.ToString();
        }

        private static List<DataPoint> CollectPairedValues(
            DataTable data,
            int startCol,
            int rowsToPlot,
            int leftProcessIndex,
            int rightProcessIndex)
        {
            var pairs = new List<DataPoint>();

            int leftCol = startCol + leftProcessIndex;
            int rightCol = startCol + rightProcessIndex;
            int rowLimit = Math.Min(rowsToPlot, data.Rows.Count);

            for (int rowIndex = 0; rowIndex < rowLimit; rowIndex++)
            {
                var row = data.Rows[rowIndex];
                if (GraphMakerParsingHelper.TryParseDouble(row[leftCol]?.ToString(), out double x) &&
                    GraphMakerParsingHelper.TryParseDouble(row[rightCol]?.ToString(), out double y))
                {
                    pairs.Add(new DataPoint(x, y));
                }
            }

            return pairs;
        }

        public object GetWebModuleSnapshot()
        {
            return new
            {
                moduleType = "GraphMakerProcessTrend",
                fileName = CurrentFileNameText.Text ?? "(No file)",
                rowCount = RowCountText.Text ?? "0",
                columnCount = ColumnCountText.Text ?? "0",
                status = StatusText.Text ?? string.Empty,
                pendingSetup = _pendingSetupFilePaths.Count == 0 ? null : new
                {
                    files = _pendingSetupFilePaths.Select(Path.GetFileName).ToArray(),
                    filePaths = _pendingSetupFilePaths.ToArray(),
                    delimiter = _pendingSetupDelimiter == "\t" ? "tab" : _pendingSetupDelimiter == "," ? "comma" : "space",
                    headerRow = _pendingSetupHeaderRowNumber.ToString(CultureInfo.InvariantCulture)
                },
                delimiter = _currentFile == null ? "tab" : _currentFile.Delimiter == "\t" ? "tab" : _currentFile.Delimiter == "," ? "comma" : "space",
                maxSamples = MaxSamplesTextBox.Text ?? "100",
                plotColor = PlotColorComboBox.SelectedItem?.ToString() ?? string.Empty,
                useQuadratic = QuadraticDegreeRadio.IsChecked == true,
                useFirstColumnAsSampleId = UseFirstColumnAsSampleIdCheckBox.IsChecked == true,
                specRow = SpecRowTextBlock.Text ?? string.Empty,
                upperRow = UpperRowTextBlock.Text ?? string.Empty,
                lowerRow = LowerRowTextBlock.Text ?? string.Empty,
                xAxisItems = XAxisListBox.Items.Cast<object>().Select(item => item?.ToString() ?? string.Empty).ToArray(),
                selectedXAxis = XAxisListBox.SelectedItem?.ToString() ?? string.Empty,
                selectedGroupId = GroupManageComboBox.SelectedValue is int selectedGroupId ? selectedGroupId : _pairGroupNames.Keys.DefaultIfEmpty(1).First(),
                selectedGroupName = GroupNameTextBox.Text ?? string.Empty,
                groupOptions = _pairGroupNames.OrderBy(pair => pair.Key).Select(pair => new { id = pair.Key, name = pair.Value }).ToArray(),
                axisGroupAssignments = _axisEntries.SelectMany(entry => entry.YAxisGroups.OrderBy(pair => pair.Key).Select(pair => new
                {
                    xAxis = entry.Name,
                    yAxis = _processNames.TryGetValue(pair.Key, out string? yName) ? yName : pair.Key.ToString(),
                    groupId = pair.Value
                })).ToArray(),
                previewColumns = BuildPreviewColumns(_currentFile?.FullData),
                previewRows = BuildPreviewRows(_currentFile?.FullData),
                yAxisItems = YAxisItemsPanel.Children.OfType<CheckBox>().Select(cb => new
                {
                    label = cb.Content?.ToString() ?? string.Empty,
                    isChecked = cb.IsChecked == true
                }).ToArray()
            };
        }

        public object UpdateWebModuleState(JsonElement payload)
        {
            if (payload.TryGetProperty("maxSamples", out JsonElement maxSamplesElement))
            {
                MaxSamplesTextBox.Text = maxSamplesElement.GetString() ?? "100";
            }

            if (payload.TryGetProperty("pendingDelimiter", out JsonElement pendingDelimiterElement))
            {
                string pendingDelimiter = pendingDelimiterElement.GetString() ?? "tab";
                _pendingSetupDelimiter = ParseWebDelimiter(pendingDelimiter);
            }

            if (payload.TryGetProperty("delimiter", out JsonElement delimiterElement))
            {
                string webDelimiter = delimiterElement.GetString() ?? "tab";
                string parsedDelimiter = ParseWebDelimiter(webDelimiter);
                if (_currentFile == null || !string.Equals(_currentFile.Delimiter, parsedDelimiter, StringComparison.Ordinal))
                {
                    ReloadCurrentFilesWithDelimiter(parsedDelimiter);
                }
            }

            if (payload.TryGetProperty("pendingHeaderRow", out JsonElement pendingHeaderRowElement) &&
                int.TryParse(pendingHeaderRowElement.GetString(), out int pendingHeaderRowNumber) &&
                pendingHeaderRowNumber > 0)
            {
                _pendingSetupHeaderRowNumber = pendingHeaderRowNumber;
            }

            if (payload.TryGetProperty("useQuadratic", out JsonElement quadraticElement))
            {
                bool useQuadratic = quadraticElement.GetBoolean();
                QuadraticDegreeRadio.IsChecked = useQuadratic;
                LinearDegreeRadio.IsChecked = !useQuadratic;
            }

            if (payload.TryGetProperty("useFirstColumnAsSampleId", out JsonElement sampleIdElement))
            {
                UseFirstColumnAsSampleIdCheckBox.IsChecked = sampleIdElement.GetBoolean();
            }

            if (payload.TryGetProperty("selectedXAxis", out JsonElement selectedXAxisElement))
            {
                string selectedXAxis = selectedXAxisElement.GetString() ?? string.Empty;
                AxisEntry? match = _axisEntries.FirstOrDefault(entry => string.Equals(entry.ToString(), selectedXAxis, StringComparison.Ordinal));
                if (match is not null)
                {
                    XAxisListBox.SelectedItem = match;
                    XAxisListBox.ScrollIntoView(match);
                }
            }

            if (payload.TryGetProperty("selectedGroupId", out JsonElement selectedGroupIdElement) &&
                selectedGroupIdElement.TryGetInt32(out int selectedGroupId) &&
                _pairGroupNames.ContainsKey(selectedGroupId))
            {
                RefreshGroupOptions(selectedGroupId);
            }

            if (payload.TryGetProperty("selectedGroupName", out JsonElement selectedGroupNameElement))
            {
                GroupNameTextBox.Text = selectedGroupNameElement.GetString() ?? string.Empty;
            }

            if (payload.TryGetProperty("axisGroupAssignments", out JsonElement axisGroupAssignmentsElement) &&
                axisGroupAssignmentsElement.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement assignmentElement in axisGroupAssignmentsElement.EnumerateArray())
                {
                    string xAxis = assignmentElement.TryGetProperty("xAxis", out JsonElement xAxisElement) ? xAxisElement.GetString() ?? string.Empty : string.Empty;
                    string yAxis = assignmentElement.TryGetProperty("yAxis", out JsonElement yAxisElement) ? yAxisElement.GetString() ?? string.Empty : string.Empty;
                    if (!assignmentElement.TryGetProperty("groupId", out JsonElement groupIdElement) || !groupIdElement.TryGetInt32(out int groupId))
                    {
                        continue;
                    }

                    AxisEntry? entry = _axisEntries.FirstOrDefault(item => string.Equals(item.Name, xAxis, StringComparison.Ordinal));
                    if (entry is null || !_pairGroupNames.ContainsKey(groupId))
                    {
                        continue;
                    }

                    int yIndex = _processNames.FirstOrDefault(pair => string.Equals(pair.Value, yAxis, StringComparison.Ordinal)).Key;
                    if (!_processNames.ContainsKey(yIndex))
                    {
                        continue;
                    }

                    entry.YAxisGroups[yIndex] = groupId;
                }

                CaptureAxisMappingsFromUi();
                RefreshSelectedYGroupList();
                XAxisListBox.Items.Refresh();
                PersistAxisSettings();
            }

            if (payload.TryGetProperty("selectedYAxis", out JsonElement selectedYAxisElement) &&
                selectedYAxisElement.ValueKind == JsonValueKind.Array)
            {
                HashSet<string> selectedLabels = selectedYAxisElement.EnumerateArray()
                    .Select(item => item.GetString() ?? string.Empty)
                    .ToHashSet(StringComparer.Ordinal);

                foreach (YAxisOption option in _yAxisOptions)
                {
                    string label = option.CheckBox.Content?.ToString() ?? string.Empty;
                    option.CheckBox.IsChecked = selectedLabels.Contains(label);
                }
            }

            NotifyWebModuleSnapshotChanged();
            return GetWebModuleSnapshot();
        }

        public object InvokeWebModuleAction(string action)
        {
            switch (action)
            {
                case "browse-file":
                    BrowseButton_Click(this, new RoutedEventArgs());
                    break;
                case "remove-file":
                    RemoveFileButton_Click(this, new RoutedEventArgs());
                    break;
                case "select-all-y":
                    SelectAllButton_Click(this, new RoutedEventArgs());
                    break;
                case "clear-all-y":
                    ClearAllButton_Click(this, new RoutedEventArgs());
                    break;
                case "add-group":
                    AddGroupButton_Click(this, new RoutedEventArgs());
                    break;
                case "remove-group":
                    RemoveGroupButton_Click(this, new RoutedEventArgs());
                    break;
                case "apply-group-name":
                    ApplyGroupNameButton_Click(this, new RoutedEventArgs());
                    break;
                case "configure-pairs":
                    ConfigurePairsButton_Click(this, new RoutedEventArgs());
                    break;
                case "save-state":
                    SaveStateButton_Click(this, new RoutedEventArgs());
                    break;
                case "load-state":
                    LoadStateButton_Click(this, new RoutedEventArgs());
                    break;
                case "generate-trend":
                    GenerateTrendButton_Click(this, new RoutedEventArgs());
                    break;
                case "apply-inline-setup":
                    ApplyPendingWebSetup();
                    break;
                case "cancel-inline-setup":
                    _pendingSetupFilePaths.Clear();
                    NotifyWebModuleSnapshotChanged();
                    break;
            }

            return GetWebModuleSnapshot();
        }

        private static string[] BuildPreviewColumns(DataTable? table)
        {
            return table?.Columns.Cast<DataColumn>().Take(24).Select(column => column.ColumnName ?? string.Empty).ToArray()
                ?? Array.Empty<string>();
        }

        private static string[][] BuildPreviewRows(DataTable? table)
        {
            if (table == null)
            {
                return Array.Empty<string[]>();
            }

            DataColumn[] columns = table.Columns.Cast<DataColumn>().Take(24).ToArray();
            return table.Rows.Cast<DataRow>()
                .Take(40)
                .Select(row => columns.Select(column => row[column]?.ToString() ?? string.Empty).ToArray())
                .ToArray();
        }

    }
}
