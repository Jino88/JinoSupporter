using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using JinoSupporter.Controls;
using Microsoft.Win32;
using OxyPlot;
using UserControl = System.Windows.Controls.UserControl;
using CheckBox = System.Windows.Controls.CheckBox;
using Color = System.Windows.Media.Color;
using DragEventArgs = System.Windows.DragEventArgs;
using MessageBox = System.Windows.MessageBox;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Border = System.Windows.Controls.Border;

namespace GraphMaker
{
    public class SavedGroupSetting
    {
        public int GroupId { get; set; }
        public string Name { get; set; } = string.Empty;
        public int ColorIndex { get; set; }
        public int LineStyleIndex { get; set; }
        public int RefColumnIndex { get; set; } = -1;
        public int UpperLimitColumnIndex { get; set; } = -1;
        public int LowerLimitColumnIndex { get; set; } = -1;
        public int RefColorIndex { get; set; } = 0;
        public int UpperColorIndex { get; set; } = 0;
        public int LowerColorIndex { get; set; } = 0;
        public int RefLineStyleIndex { get; set; } = 1;
        public int UpperLineStyleIndex { get; set; } = 1;
        public int LowerLineStyleIndex { get; set; } = 1;
    }

    public class FileInfo_Custom : GraphFileInfoBase
    {
        public List<string> FirstColumn { get; set; }

        public int SavedXAxisIndex { get; set; } = -1;
        public List<int> SavedYAxisIndices { get; set; } = new List<int>();
        public int SavedSpecColumnIndex { get; set; } = -1;
        public int SavedUpperLimitColumnIndex { get; set; } = -1;
        public int SavedLowerLimitColumnIndex { get; set; } = -1;
        public int SavedSpecColorIndex { get; set; } = 0;
        public int SavedUpperColorIndex { get; set; } = 0;
        public int SavedLowerColorIndex { get; set; } = 0;
        public int SavedSpecLineStyleIndex { get; set; } = 1;
        public int SavedUpperLineStyleIndex { get; set; } = 1;
        public int SavedLowerLineStyleIndex { get; set; } = 1;
        public int SavedYAxisColorIndex { get; set; } = 0;
        public bool SavedXAxisLogScale { get; set; } = true;
        public Dictionary<int, int> SavedColumnGroups { get; set; } = new Dictionary<int, int>();
        public List<SavedGroupSetting> SavedGroups { get; set; } = new List<SavedGroupSetting>
        {
            new SavedGroupSetting { GroupId = 1, Name = "Group 1", ColorIndex = 2, LineStyleIndex = 0 },
            new SavedGroupSetting { GroupId = 2, Name = "Group 2", ColorIndex = 1, LineStyleIndex = 0 }
        };
    }

    public partial class ScatterPlotView : GraphViewBase
    {
        private sealed class ScatterPlotReportState
        {
            public List<ScatterPlotFileState> Files { get; set; } = new();
            public string? SelectedFilePath { get; set; }
        }

        private sealed class ScatterPlotFileState
        {
            public string Name { get; set; } = string.Empty;
            public string FilePath { get; set; } = string.Empty;
            public string Delimiter { get; set; } = "\t";
            public int HeaderRowNumber { get; set; } = 1;
            public RawTableData RawData { get; set; } = new();
            public int SavedXAxisIndex { get; set; } = -1;
            public List<int> SavedYAxisIndices { get; set; } = new();
            public int SavedSpecColumnIndex { get; set; } = -1;
            public int SavedUpperLimitColumnIndex { get; set; } = -1;
            public int SavedLowerLimitColumnIndex { get; set; } = -1;
            public int SavedSpecColorIndex { get; set; }
            public int SavedUpperColorIndex { get; set; }
            public int SavedLowerColorIndex { get; set; }
            public int SavedSpecLineStyleIndex { get; set; } = 1;
            public int SavedUpperLineStyleIndex { get; set; } = 1;
            public int SavedLowerLineStyleIndex { get; set; } = 1;
            public int SavedYAxisColorIndex { get; set; }
            public bool SavedXAxisLogScale { get; set; } = true;
            public Dictionary<int, int> SavedColumnGroups { get; set; } = new();
            public List<SavedGroupSetting> SavedGroups { get; set; } = new();
        }

        private sealed class ColumnItem
        {
            public int Index { get; set; }
            public string Header { get; set; } = string.Empty;
            public override string ToString() => Header;
        }

        private sealed class GroupRow
        {
            public int GroupId { get; set; }
            public TextBox NameTextBox { get; set; } = null!;
            public ComboBox ColorComboBox { get; set; } = null!;
            public ComboBox LineStyleComboBox { get; set; } = null!;
            public ListBox GroupColumnsListBox { get; set; } = null!;
        }

        private ObservableCollection<FileInfo_Custom> _loadedFiles = new ObservableCollection<FileInfo_Custom>();
        private FileInfo_Custom _currentFile;
        private List<PreviewColorChoice> _colorOptions;
        private readonly Dictionary<int, int> _columnGroups = new Dictionary<int, int>();
        private readonly List<GroupRow> _groupRows = new List<GroupRow>();
        private int _nextGroupId = 1;
        private readonly List<string> _lineStyleOptions = new List<string> { "Solid", "Dash", "Dot", "DashDot" };
        private System.Windows.Point _unassignedDragStartPoint;

        private static readonly OxyColor[] ColorPalette = new[]
        {
            OxyColors.Blue, OxyColors.Red, OxyColors.Green, OxyColors.Orange,
            OxyColors.Purple, OxyColors.Brown, OxyColors.Cyan, OxyColors.Magenta,
            OxyColors.Gold, OxyColors.Navy, OxyColors.Teal, OxyColors.Maroon,
            OxyColors.Olive, OxyColors.Lime, OxyColors.Pink, OxyColors.Indigo
        };

        public ScatterPlotView()
        {
            InitializeComponent();
            DataContext = this;
            FileListBox.ItemsSource = _loadedFiles;
            _loadedFiles.CollectionChanged += (_, _) => NotifyWebModuleSnapshotChanged();
            InitializeColorOptions();
            NotifyWebModuleSnapshotChanged();
        }

        private void InitializeColorOptions()
        {
            _colorOptions = new List<PreviewColorChoice>();
            PreviewGraphViewBase.InitializeDefaultColorOptions(SpecColorComboBox, _colorOptions);
            PreviewGraphViewBase.BindColorComboBox(UpperColorComboBox, _colorOptions);
            PreviewGraphViewBase.BindColorComboBox(LowerColorComboBox, _colorOptions);

            SpecLineStyleComboBox.ItemsSource = _lineStyleOptions;
            SpecLineStyleComboBox.SelectedIndex = 1;
            UpperLimitLineStyleComboBox.ItemsSource = _lineStyleOptions;
            UpperLimitLineStyleComboBox.SelectedIndex = 1;
            LowerLimitLineStyleComboBox.ItemsSource = _lineStyleOptions;
            LowerLimitLineStyleComboBox.SelectedIndex = 1;

            PreviewGraphViewBase.BindColorComboBox(YAxisColorComboBox, _colorOptions);
        }

        private void FileDropBox_FilesSelected(object sender, FilesSelectedEventArgs e)
        {
            foreach (var path in e.FilePaths)
            {
                LoadFile(path);
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select data files",
                Filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string fileName in openFileDialog.FileNames)
                {
                    LoadFile(fileName);
                }
            }
        }

        private void LoadFile(string filePath)
        {
            try
            {
                // ì´ë¯¸ ë¡œë“œëœ íŒŒì¼ì¸ì§€ í™•ì¸
                if (_loadedFiles.Any(f => f.FilePath == filePath))
                {
                    MessageBox.Show($"{Path.GetFileName(filePath)} is already loaded.",
                        "Duplicate", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var fileInfo = new FileInfo_Custom
                {
                    Name = Path.GetFileName(filePath),
                    FilePath = filePath
                };
                EnsureSavedGroups(fileInfo);

                _loadedFiles.Add(fileInfo);
                LoadFileData(fileInfo);

                // ì²« ë²ˆì§¸ íŒŒì¼ì´ë©´ ìžë™ ì„ íƒ
                if (_loadedFiles.Count == 1)
                {
                    FileListBox.SelectedIndex = 0;
                }

                NotifyWebModuleSnapshotChanged();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while loading file: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void HandleWebDroppedFiles(IReadOnlyList<string> filePaths)
        {
            foreach (string file in filePaths)
            {
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".txt" || ext == ".csv")
                {
                    LoadFile(file);
                }
            }
        }

        private void FileListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // ì´ì „ íŒŒì¼ì˜ ì„¤ì • ì €ìž¥
            if (_currentFile != null)
            {
                SaveCurrentFileSettings();
            }

            if (FileListBox.SelectedItem is FileInfo_Custom selectedFile)
            {
                _currentFile = selectedFile;
                LoadFileData(selectedFile);
            }

            NotifyWebModuleSnapshotChanged();
        }

        private void SaveCurrentFileSettings()
        {
            if (_currentFile == null) return;

            // Xì¶• ì €ìž¥
            _currentFile.SavedXAxisIndex = XAxisComboBox.SelectedIndex;

            _currentFile.SavedYAxisIndices = _columnGroups.Keys.OrderBy(x => x).ToList();

            // SPEC/Upper/Lower Limit ì €ìž¥
            _currentFile.SavedSpecColumnIndex = GetSelectedLimitIndex(RefLimitListBox);
            _currentFile.SavedUpperLimitColumnIndex = GetSelectedLimitIndex(UpperLimitListBox);
            _currentFile.SavedLowerLimitColumnIndex = GetSelectedLimitIndex(LowerLimitListBox);

            // ìƒ‰ìƒ ì €ìž¥
            _currentFile.SavedSpecColorIndex = SpecColorComboBox.SelectedIndex;
            _currentFile.SavedUpperColorIndex = UpperColorComboBox.SelectedIndex;
            _currentFile.SavedLowerColorIndex = LowerColorComboBox.SelectedIndex;
            _currentFile.SavedSpecLineStyleIndex = SpecLineStyleComboBox.SelectedIndex;
            _currentFile.SavedUpperLineStyleIndex = UpperLimitLineStyleComboBox.SelectedIndex;
            _currentFile.SavedLowerLineStyleIndex = LowerLimitLineStyleComboBox.SelectedIndex;
            _currentFile.SavedXAxisLogScale = XAxisLogScaleCheckBox.IsChecked == true;
            _currentFile.SavedColumnGroups = new Dictionary<int, int>(_columnGroups);
            _currentFile.SavedGroups = _groupRows
                .Select(g => new SavedGroupSetting
                {
                    GroupId = g.GroupId,
                    Name = string.IsNullOrWhiteSpace(g.NameTextBox.Text) ? $"Group {g.GroupId}" : g.NameTextBox.Text.Trim(),
                    ColorIndex = g.ColorComboBox.SelectedIndex < 0 ? 0 : g.ColorComboBox.SelectedIndex,
                    LineStyleIndex = g.LineStyleComboBox.SelectedIndex < 0 ? 0 : g.LineStyleComboBox.SelectedIndex
                })
                .OrderBy(g => g.GroupId)
                .ToList();
        }

        private void LoadFileData(FileInfo_Custom fileInfo)
        {
            try
            {
                if (fileInfo.FullData != null && fileInfo.HeaderRow != null && fileInfo.HeaderRow.Count > 0)
                {
                    UpdateComboBoxes(fileInfo);
                    return;
                }

                // Always re-parse from source file to avoid stale cached parsing results.
                var lines = File.ReadAllLines(fileInfo.FilePath);
                if (lines.Length == 0)
                {
                    MessageBox.Show("File is empty.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int headerRowIndex = fileInfo.HeaderRowNumber - 1;
                if (headerRowIndex >= lines.Length)
                {
                    MessageBox.Show("Header row number exceeds total rows.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Detect delimiter from header line first to avoid broken parsing when file is tab-delimited.
                string effectiveDelimiter = DetectDelimiter(lines[headerRowIndex], fileInfo.Delimiter);
                fileInfo.Delimiter = effectiveDelimiter;

                var headerRow = GraphMakerTableHelper.SplitLine(lines[headerRowIndex], effectiveDelimiter);
                fileInfo.HeaderRow = headerRow.ToList();

                // DataTable ìƒì„± (ê°€ìƒí™”ë¥¼ ìœ„í•´ ëª¨ë“  ë°ì´í„° ë¡œë“œ)
                var dataTable = new DataTable();
                for (int i = 0; i < headerRow.Length; i++)
                {
                    dataTable.Columns.Add(headerRow[i]);
                }

                // ì²« ë²ˆì§¸ ì—´ ê°’ ì €ìž¥
                fileInfo.FirstColumn = new List<string>();

                // ë°ì´í„° í–‰ íŒŒì‹±
                for (int i = headerRowIndex + 1; i < lines.Length; i++)
                {
                    var values = ParseDataRow(lines[i], effectiveDelimiter, headerRow.Length);
                    if (values.Length > 0)
                    {
                        fileInfo.FirstColumn.Add(values[0]);

                        var row = dataTable.NewRow();
                        for (int j = 0; j < Math.Min(values.Length, headerRow.Length); j++)
                        {
                            row[j] = values[j];
                        }
                        dataTable.Rows.Add(row);
                    }
                }

                fileInfo.FullData = dataTable;

                // ì½¤ë³´ë°•ìŠ¤ ì—…ë°ì´íŠ¸
                UpdateComboBoxes(fileInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while loading data: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateComboBoxes(FileInfo_Custom fileInfo)
        {
            if (fileInfo.HeaderRow == null || fileInfo.HeaderRow.Count == 0)
                return;

            // Xì¶• ì½¤ë³´ë°•ìŠ¤
            XAxisComboBox.Items.Clear();
            XAxisComboBox.Items.Add("(None)");

            foreach (var header in fileInfo.HeaderRow)
            {
                XAxisComboBox.Items.Add(header);
            }

            _columnGroups.Clear();
            foreach (var kv in fileInfo.SavedColumnGroups)
            {
                _columnGroups[kv.Key] = kv.Value;
            }
            BuildGroupRowsFromSavedGroups(fileInfo);

            RefreshLimitLists(fileInfo);
            RefreshGroupLists(fileInfo);

            // ì €ìž¥ëœ ì„¤ì •ì´ ìžˆìœ¼ë©´ ë³µì›, ì—†ìœ¼ë©´ ê¸°ë³¸ê°’ ì„¤ì •
            if (fileInfo.SavedXAxisIndex >= 0 && fileInfo.SavedXAxisIndex < XAxisComboBox.Items.Count)
            {
                // ì €ìž¥ëœ ì„¤ì • ë³µì›
                XAxisComboBox.SelectedIndex = fileInfo.SavedXAxisIndex;

                // ìƒ‰ìƒ ë³µì›
                if (fileInfo.SavedSpecColorIndex >= 0 && fileInfo.SavedSpecColorIndex < SpecColorComboBox.Items.Count)
                    SpecColorComboBox.SelectedIndex = fileInfo.SavedSpecColorIndex;
                if (fileInfo.SavedUpperColorIndex >= 0 && fileInfo.SavedUpperColorIndex < UpperColorComboBox.Items.Count)
                    UpperColorComboBox.SelectedIndex = fileInfo.SavedUpperColorIndex;
                if (fileInfo.SavedLowerColorIndex >= 0 && fileInfo.SavedLowerColorIndex < LowerColorComboBox.Items.Count)
                    LowerColorComboBox.SelectedIndex = fileInfo.SavedLowerColorIndex;
                if (fileInfo.SavedSpecLineStyleIndex >= 0 && fileInfo.SavedSpecLineStyleIndex < SpecLineStyleComboBox.Items.Count)
                    SpecLineStyleComboBox.SelectedIndex = fileInfo.SavedSpecLineStyleIndex;
                if (fileInfo.SavedUpperLineStyleIndex >= 0 && fileInfo.SavedUpperLineStyleIndex < UpperLimitLineStyleComboBox.Items.Count)
                    UpperLimitLineStyleComboBox.SelectedIndex = fileInfo.SavedUpperLineStyleIndex;
                if (fileInfo.SavedLowerLineStyleIndex >= 0 && fileInfo.SavedLowerLineStyleIndex < LowerLimitLineStyleComboBox.Items.Count)
                    LowerLimitLineStyleComboBox.SelectedIndex = fileInfo.SavedLowerLineStyleIndex;
                XAxisLogScaleCheckBox.IsChecked = fileInfo.SavedXAxisLogScale;
            }
            else
            {
                // ê¸°ë³¸ ì„ íƒ (ì²˜ìŒ ë¡œë“œí•œ íŒŒì¼)
                if (XAxisComboBox.Items.Count > 1)
                {
                    XAxisComboBox.SelectedIndex = 1; // ì²« ë²ˆì§¸ ì—´ ì„ íƒ

                    fileInfo.SavedXAxisIndex = 1;
                }

                XAxisLogScaleCheckBox.IsChecked = true;
                fileInfo.SavedXAxisLogScale = true;

                // Auto-assign all non-X columns to Group 1 so users can generate immediately
                if (fileInfo.HeaderRow != null && _groupRows.Count > 0)
                {
                    int xColIndex = 0; // XAxisComboBox index 1 → actual column 0
                    int group1Id = _groupRows[0].GroupId;
                    for (int i = 0; i < fileInfo.HeaderRow.Count; i++)
                    {
                        if (i != xColIndex)
                        {
                            _columnGroups[i] = group1Id;
                        }
                    }

                    RefreshGroupLists(fileInfo);
                }
            }

            NotifyWebModuleSnapshotChanged();
        }

        private void DelimiterChanged(object sender, RoutedEventArgs e)
        {
            if (_currentFile == null) return;

            if (TabDelimiterRadio.IsChecked == true)
                _currentFile.Delimiter = "\t";
            else if (CommaDelimiterRadio.IsChecked == true)
                _currentFile.Delimiter = ",";
            else if (SpaceDelimiterRadio.IsChecked == true)
                _currentFile.Delimiter = " ";

            _currentFile.FullData = null;
            _currentFile.HeaderRow = null;
            _currentFile.FirstColumn = null;
            LoadFileData(_currentFile);
            NotifyWebModuleSnapshotChanged();
        }

        private static string DetectDelimiter(string headerLine, string currentDelimiter)
        {
            if (headerLine.Contains('\t'))
            {
                return "\t";
            }

            if (headerLine.Contains(","))
            {
                return ",";
            }

            return currentDelimiter;
        }

        private static string[] ParseDataRow(string line, string delimiter, int expectedColumns)
        {
            var values = GraphMakerTableHelper.SplitLine(line, delimiter);
            if (values.Length >= expectedColumns)
            {
                return values;
            }

            var fallbackValues = Regex.Split(line.Trim(), @"\s+");
            if (fallbackValues.Length >= expectedColumns)
            {
                return fallbackValues;
            }

            return values;
        }

        private void HeaderRowChanged(object sender, TextChangedEventArgs e)
        {
            if (_currentFile == null) return;

            if (int.TryParse(HeaderRowTextBox.Text, out int headerRow) && headerRow > 0)
            {
                _currentFile.HeaderRowNumber = headerRow;
                _currentFile.FullData = null;
                _currentFile.HeaderRow = null;
                _currentFile.FirstColumn = null;
                LoadFileData(_currentFile);
            }

            NotifyWebModuleSnapshotChanged();
        }

        private void XAxisComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Xì¶• ì—´ì´ ì„ íƒë˜ë©´ Yì¶•ì—ì„œ ìžë™ìœ¼ë¡œ ì œì™¸
            UpdateYAxisSelection();
            NotifyWebModuleSnapshotChanged();
        }

        private void XAxisModeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // X-Mode removed from UI.
        }

        private void AddGroupButton_Click(object sender, RoutedEventArgs e)
        {
            int groupId = _nextGroupId++;
            AddGroupRow(groupId, $"Group {groupId}", 2, 0);
            if (_currentFile != null)
            {
                RefreshGroupLists(_currentFile);
            }
        }

        private void UpdateYAxisSelection()
        {
            if (YAxisListBox.Items.Count == 0) return;

            // Xì¶• ì¸ë±ìŠ¤ (0ì€ "(ì„ íƒ ì•ˆ í•¨)"ì´ë¯€ë¡œ ì‹¤ì œ ì—´ ì¸ë±ìŠ¤ëŠ” -1)
            int xAxisColumnIndex = XAxisComboBox.SelectedIndex - 1;

            // SPEC/Upper/Lower Limit ì¸ë±ìŠ¤
            int specColumnIndex = GetSelectedLimitIndex(RefLimitListBox) - 1;
            int upperColumnIndex = GetSelectedLimitIndex(UpperLimitListBox) - 1;
            int lowerColumnIndex = GetSelectedLimitIndex(LowerLimitListBox) - 1;

            // ì œì™¸í•  ì—´ ì¸ë±ìŠ¤ ìˆ˜ì§‘
            var excludedIndices = new HashSet<int>();
            if (xAxisColumnIndex >= 0) excludedIndices.Add(xAxisColumnIndex);
            if (specColumnIndex >= 0) excludedIndices.Add(specColumnIndex);
            if (upperColumnIndex >= 0) excludedIndices.Add(upperColumnIndex);
            if (lowerColumnIndex >= 0) excludedIndices.Add(lowerColumnIndex);

            // Yì¶• ListBox ì—…ë°ì´íŠ¸
            for (int i = 0; i < YAxisListBox.Items.Count; i++)
            {
                var item = YAxisListBox.Items[i] as ListBoxItem;
                if (item != null && excludedIndices.Contains(i))
                {
                    // ì œì™¸í•  ì—´ì´ë©´ ì„ íƒ í•´ì œ
                    item.IsSelected = false;
                }
            }
        }

        private void RemoveFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (FileListBox.SelectedItem is FileInfo_Custom selectedFile)
            {
                // í˜„ìž¬ ì„ íƒëœ íŒŒì¼ì´ ì œê±°ë  íŒŒì¼ì´ë©´ ì„¤ì • ì €ìž¥ ë¶ˆí•„ìš”
                if (selectedFile == _currentFile)
                {
                    _currentFile = null;
                }

                _loadedFiles.Remove(selectedFile);

                if (_loadedFiles.Count > 0)
                {
                    FileListBox.SelectedIndex = 0;
                }
                else
                {
                    _currentFile = null;
                }
            }

            NotifyWebModuleSnapshotChanged();
        }

        private void SaveReportButton_Click(object sender, RoutedEventArgs e)
        {
            if (_loadedFiles.Count == 0)
            {
                MessageBox.Show("Load data first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            SaveCurrentFileSettings();

            var state = new ScatterPlotReportState
            {
                SelectedFilePath = _currentFile?.FilePath,
                Files = _loadedFiles.Select(file => new ScatterPlotFileState
                {
                    Name = file.Name,
                    FilePath = file.FilePath,
                    Delimiter = file.Delimiter,
                    HeaderRowNumber = file.HeaderRowNumber,
                    RawData = GraphReportStorageHelper.CaptureRawTableData(file.FullData),
                    SavedXAxisIndex = file.SavedXAxisIndex,
                    SavedYAxisIndices = file.SavedYAxisIndices.ToList(),
                    SavedSpecColumnIndex = file.SavedSpecColumnIndex,
                    SavedUpperLimitColumnIndex = file.SavedUpperLimitColumnIndex,
                    SavedLowerLimitColumnIndex = file.SavedLowerLimitColumnIndex,
                    SavedSpecColorIndex = file.SavedSpecColorIndex,
                    SavedUpperColorIndex = file.SavedUpperColorIndex,
                    SavedLowerColorIndex = file.SavedLowerColorIndex,
                    SavedSpecLineStyleIndex = file.SavedSpecLineStyleIndex,
                    SavedUpperLineStyleIndex = file.SavedUpperLineStyleIndex,
                    SavedLowerLineStyleIndex = file.SavedLowerLineStyleIndex,
                    SavedYAxisColorIndex = file.SavedYAxisColorIndex,
                    SavedXAxisLogScale = file.SavedXAxisLogScale,
                    SavedColumnGroups = new Dictionary<int, int>(file.SavedColumnGroups),
                    SavedGroups = file.SavedGroups.Select(group => new SavedGroupSetting
                    {
                        GroupId = group.GroupId,
                        Name = group.Name,
                        ColorIndex = group.ColorIndex,
                        LineStyleIndex = group.LineStyleIndex,
                        RefColumnIndex = group.RefColumnIndex,
                        UpperLimitColumnIndex = group.UpperLimitColumnIndex,
                        LowerLimitColumnIndex = group.LowerLimitColumnIndex,
                        RefColorIndex = group.RefColorIndex,
                        UpperColorIndex = group.UpperColorIndex,
                        LowerColorIndex = group.LowerColorIndex,
                        RefLineStyleIndex = group.RefLineStyleIndex,
                        UpperLineStyleIndex = group.UpperLineStyleIndex,
                        LowerLineStyleIndex = group.LowerLineStyleIndex
                    }).ToList()
                }).ToList()
            };

            GraphReportFileDialogHelper.SaveState("Save Graph Report", "scatterplot.graphreport.json", state);
        }

        private void LoadReportButton_Click(object sender, RoutedEventArgs e)
        {
            ScatterPlotReportState? state = GraphReportFileDialogHelper.LoadState<ScatterPlotReportState>("Load Graph Report");
            if (state == null || state.Files.Count == 0)
            {
                MessageBox.Show("Invalid report file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _loadedFiles.Clear();
            _currentFile = null;

            foreach (ScatterPlotFileState fileState in state.Files)
            {
                var table = GraphReportStorageHelper.BuildTableFromRawData(fileState.RawData);
                var fileInfo = new FileInfo_Custom
                {
                    Name = fileState.Name,
                    FilePath = fileState.FilePath,
                    Delimiter = string.IsNullOrWhiteSpace(fileState.Delimiter) ? "\t" : fileState.Delimiter,
                    HeaderRowNumber = Math.Max(1, fileState.HeaderRowNumber),
                    FullData = table,
                    HeaderRow = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToList(),
                    FirstColumn = table.Rows.Cast<DataRow>()
                        .Select(row => table.Columns.Count > 0 ? row[0]?.ToString() ?? string.Empty : string.Empty)
                        .ToList(),
                    SavedXAxisIndex = fileState.SavedXAxisIndex,
                    SavedYAxisIndices = fileState.SavedYAxisIndices ?? new List<int>(),
                    SavedSpecColumnIndex = fileState.SavedSpecColumnIndex,
                    SavedUpperLimitColumnIndex = fileState.SavedUpperLimitColumnIndex,
                    SavedLowerLimitColumnIndex = fileState.SavedLowerLimitColumnIndex,
                    SavedSpecColorIndex = fileState.SavedSpecColorIndex,
                    SavedUpperColorIndex = fileState.SavedUpperColorIndex,
                    SavedLowerColorIndex = fileState.SavedLowerColorIndex,
                    SavedSpecLineStyleIndex = fileState.SavedSpecLineStyleIndex,
                    SavedUpperLineStyleIndex = fileState.SavedUpperLineStyleIndex,
                    SavedLowerLineStyleIndex = fileState.SavedLowerLineStyleIndex,
                    SavedYAxisColorIndex = fileState.SavedYAxisColorIndex,
                    SavedXAxisLogScale = fileState.SavedXAxisLogScale,
                    SavedColumnGroups = fileState.SavedColumnGroups ?? new Dictionary<int, int>(),
                    SavedGroups = fileState.SavedGroups ?? new List<SavedGroupSetting>()
                };

                EnsureSavedGroups(fileInfo);
                _loadedFiles.Add(fileInfo);
            }

            FileInfo_Custom? selectedFile = _loadedFiles.FirstOrDefault(file =>
                string.Equals(file.FilePath, state.SelectedFilePath, StringComparison.OrdinalIgnoreCase))
                ?? _loadedFiles.FirstOrDefault();

            if (selectedFile != null)
            {
                FileListBox.SelectedItem = selectedFile;
            }

            NotifyWebModuleSnapshotChanged();
        }

        public object GetWebModuleSnapshot()
        {
            string delimiter = TabDelimiterRadio.IsChecked == true
                ? "tab"
                : CommaDelimiterRadio.IsChecked == true
                    ? "comma"
                    : "space";

            List<string> yAxisCandidates = new();
            if (_currentFile?.HeaderRow is not null)
            {
                int xAxisColumnIndex = XAxisComboBox.SelectedIndex - 1;
                int specColumnIndex = GetSelectedLimitIndex(RefLimitListBox) - 1;
                int upperColumnIndex = GetSelectedLimitIndex(UpperLimitListBox) - 1;
                int lowerColumnIndex = GetSelectedLimitIndex(LowerLimitListBox) - 1;

                var excludedIndices = new HashSet<int>();
                if (xAxisColumnIndex >= 0) excludedIndices.Add(xAxisColumnIndex);
                if (specColumnIndex >= 0) excludedIndices.Add(specColumnIndex);
                if (upperColumnIndex >= 0) excludedIndices.Add(upperColumnIndex);
                if (lowerColumnIndex >= 0) excludedIndices.Add(lowerColumnIndex);

                for (int i = 0; i < _currentFile.HeaderRow.Count; i++)
                {
                    if (!excludedIndices.Contains(i))
                    {
                        yAxisCandidates.Add(_currentFile.HeaderRow[i] ?? string.Empty);
                    }
                }
            }

            return new
            {
                moduleType = "GraphMakerScatterPlot",
                fileCount = _loadedFiles.Count,
                selectedFilePath = _currentFile?.FilePath ?? string.Empty,
                selectedFileName = _currentFile?.Name ?? string.Empty,
                delimiter,
                headerRow = HeaderRowTextBox.Text ?? "1",
                xAxisIndex = XAxisComboBox.SelectedIndex,
                xAxisLabel = XAxisComboBox.SelectedItem?.ToString() ?? "(None)",
                xAxisLogScale = XAxisLogScaleCheckBox.IsChecked == true,
                xAxisOptions = XAxisComboBox.Items.Cast<object>().Select((item, index) => new
                {
                    index,
                    label = item?.ToString() ?? string.Empty
                }).ToArray(),
                limitOptions = _currentFile?.HeaderRow?.Select(header => header ?? string.Empty).ToArray() ?? Array.Empty<string>(),
                loadedFiles = _loadedFiles.Select(file => new
                {
                    name = file.Name,
                    filePath = file.FilePath,
                    rowCount = file.FullData?.Rows.Count ?? 0,
                    columnCount = file.HeaderRow?.Count ?? 0,
                    isSelected = string.Equals(file.FilePath, _currentFile?.FilePath, StringComparison.OrdinalIgnoreCase)
                }).ToArray(),
                yAxisCandidates = yAxisCandidates.ToArray(),
                unassignedColumns = UnassignedColumnsListBox.Items.Cast<object>().Select(item => item?.ToString() ?? string.Empty).Take(50).ToArray(),
                refLimit = RefLimitListBox.Items.Cast<object>().Select(item => item?.ToString() ?? string.Empty).FirstOrDefault() ?? string.Empty,
                upperLimit = UpperLimitListBox.Items.Cast<object>().Select(item => item?.ToString() ?? string.Empty).FirstOrDefault() ?? string.Empty,
                lowerLimit = LowerLimitListBox.Items.Cast<object>().Select(item => item?.ToString() ?? string.Empty).FirstOrDefault() ?? string.Empty,
                previewColumns = BuildPreviewColumns(_currentFile?.FullData),
                previewRows = BuildPreviewRows(_currentFile?.FullData),
                groups = _groupRows.Select(row => new
                {
                    groupId = row.GroupId,
                    name = string.IsNullOrWhiteSpace(row.NameTextBox.Text) ? $"Group {row.GroupId}" : row.NameTextBox.Text.Trim(),
                    columns = row.GroupColumnsListBox.Items.Cast<object>().Select(item => item?.ToString() ?? string.Empty).ToArray()
                }).ToArray()
            };
        }

        public object UpdateWebModuleState(JsonElement payload)
        {
            if (payload.TryGetProperty("selectedFilePath", out JsonElement selectedFilePathElement))
            {
                string? selectedFilePath = selectedFilePathElement.GetString();
                FileInfo_Custom? selectedFile = _loadedFiles.FirstOrDefault(file =>
                    string.Equals(file.FilePath, selectedFilePath, StringComparison.OrdinalIgnoreCase));
                if (selectedFile is not null)
                {
                    FileListBox.SelectedItem = selectedFile;
                }
            }

            if (payload.TryGetProperty("delimiter", out JsonElement delimiterElement))
            {
                string delimiter = delimiterElement.GetString() ?? "tab";
                TabDelimiterRadio.IsChecked = delimiter == "tab";
                CommaDelimiterRadio.IsChecked = delimiter == "comma";
                SpaceDelimiterRadio.IsChecked = delimiter == "space";
                if (_currentFile != null)
                {
                    DelimiterChanged(this, new RoutedEventArgs());
                }
            }

            if (payload.TryGetProperty("headerRow", out JsonElement headerRowElement))
            {
                HeaderRowTextBox.Text = headerRowElement.GetString() ?? "1";
            }

            if (payload.TryGetProperty("xAxisIndex", out JsonElement xAxisIndexElement) && xAxisIndexElement.TryGetInt32(out int xAxisIndex))
            {
                if (xAxisIndex >= 0 && xAxisIndex < XAxisComboBox.Items.Count)
                {
                    XAxisComboBox.SelectedIndex = xAxisIndex;
                }
            }

            if (payload.TryGetProperty("xAxisLogScale", out JsonElement logScaleElement))
            {
                XAxisLogScaleCheckBox.IsChecked = logScaleElement.GetBoolean();
            }

            if (_currentFile?.HeaderRow is not null)
            {
                if (payload.TryGetProperty("refLimit", out JsonElement refLimitElement))
                {
                    ApplyLimitSelection(RefLimitListBox, _currentFile.HeaderRow, refLimitElement.GetString());
                }

                if (payload.TryGetProperty("upperLimit", out JsonElement upperLimitElement))
                {
                    ApplyLimitSelection(UpperLimitListBox, _currentFile.HeaderRow, upperLimitElement.GetString());
                }

                if (payload.TryGetProperty("lowerLimit", out JsonElement lowerLimitElement))
                {
                    ApplyLimitSelection(LowerLimitListBox, _currentFile.HeaderRow, lowerLimitElement.GetString());
                }

                RebuildLimitUnassignedList(_currentFile);
                UpdateYAxisSelection();
            }

            if (_currentFile?.HeaderRow is not null &&
                payload.TryGetProperty("groups", out JsonElement groupsElement) &&
                groupsElement.ValueKind == JsonValueKind.Array)
            {
                Dictionary<int, List<string>> groupColumnsById = new();

                foreach (JsonElement groupElement in groupsElement.EnumerateArray())
                {
                    if (!groupElement.TryGetProperty("groupId", out JsonElement groupIdElement) ||
                        !groupIdElement.TryGetInt32(out int groupId))
                    {
                        continue;
                    }

                    if (groupElement.TryGetProperty("name", out JsonElement nameElement))
                    {
                        GroupRow? row = _groupRows.FirstOrDefault(item => item.GroupId == groupId);
                        if (row is not null)
                        {
                            row.NameTextBox.Text = string.IsNullOrWhiteSpace(nameElement.GetString())
                                ? $"Group {groupId}"
                                : nameElement.GetString()!;
                        }
                    }

                    if (groupElement.TryGetProperty("columns", out JsonElement columnsElement) &&
                        columnsElement.ValueKind == JsonValueKind.Array)
                    {
                        groupColumnsById[groupId] = columnsElement.EnumerateArray()
                            .Select(item => item.GetString() ?? string.Empty)
                            .Where(name => !string.IsNullOrWhiteSpace(name))
                            .Distinct(StringComparer.Ordinal)
                            .ToList();
                    }
                }

                _columnGroups.Clear();
                foreach ((int groupId, List<string> names) in groupColumnsById)
                {
                    foreach (string name in names)
                    {
                        int columnIndex = _currentFile.HeaderRow.FindIndex(header => string.Equals(header, name, StringComparison.Ordinal));
                        if (columnIndex >= 0)
                        {
                            _columnGroups[columnIndex] = groupId;
                        }
                    }
                }

                RefreshGroupLists(_currentFile);
            }

            NotifyWebModuleSnapshotChanged();
            return GetWebModuleSnapshot();
        }

        public object InvokeWebModuleAction(string action)
        {
            switch (action)
            {
                case "browse-files":
                    BrowseButton_Click(this, new RoutedEventArgs());
                    break;
                case "remove-selected-file":
                    RemoveFileButton_Click(this, new RoutedEventArgs());
                    break;
                case "save-report":
                    SaveReportButton_Click(this, new RoutedEventArgs());
                    break;
                case "load-report":
                    LoadReportButton_Click(this, new RoutedEventArgs());
                    break;
                case "add-group":
                    AddGroupRow(_nextGroupId++, $"Group {_nextGroupId - 1}", 0, 0);
                    if (_currentFile is not null)
                    {
                        RefreshGroupLists(_currentFile);
                    }
                    break;
                case "generate-graph":
                    GenerateGraphButton_Click(this, new RoutedEventArgs());
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

        private static void ApplyLimitSelection(ListBox target, List<string> headers, string? headerName)
        {
            target.Items.Clear();

            if (string.IsNullOrWhiteSpace(headerName))
            {
                return;
            }

            int index = headers.FindIndex(header => string.Equals(header, headerName, StringComparison.Ordinal));
            if (index >= 0)
            {
                target.Items.Add(new ColumnItem { Index = index, Header = headers[index] });
            }
        }

        private static int GetSelectedLimitIndex(ListBox limitListBox)
        {
            if (limitListBox.Items.Count == 0)
            {
                return 0;
            }

            return limitListBox.Items[0] is ColumnItem item ? item.Index + 1 : 0;
        }

        private void RefreshLimitLists(FileInfo_Custom fileInfo)
        {
            SetLimitListItem(RefLimitListBox, fileInfo.HeaderRow, fileInfo.SavedSpecColumnIndex - 1);
            SetLimitListItem(UpperLimitListBox, fileInfo.HeaderRow, fileInfo.SavedUpperLimitColumnIndex - 1);
            SetLimitListItem(LowerLimitListBox, fileInfo.HeaderRow, fileInfo.SavedLowerLimitColumnIndex - 1);
            RebuildLimitUnassignedList(fileInfo);
            UpdateYAxisSelection();
        }

        private void SetLimitListItem(ListBox target, List<string>? headers, int index)
        {
            target.Items.Clear();
            if (headers == null)
            {
                return;
            }

            if (index >= 0 && index < headers.Count)
            {
                target.Items.Add(new ColumnItem { Index = index, Header = headers[index] });
            }
        }

        private void RebuildLimitUnassignedList(FileInfo_Custom fileInfo)
        {
            if (fileInfo.HeaderRow == null)
            {
                return;
            }

            var used = new HashSet<int>();
            if (RefLimitListBox.Items.Count > 0 && RefLimitListBox.Items[0] is ColumnItem refItem) used.Add(refItem.Index);
            if (UpperLimitListBox.Items.Count > 0 && UpperLimitListBox.Items[0] is ColumnItem upperItem) used.Add(upperItem.Index);
            if (LowerLimitListBox.Items.Count > 0 && LowerLimitListBox.Items[0] is ColumnItem lowerItem) used.Add(lowerItem.Index);

            UnassignedColumnsListBox.Items.Clear();
            for (int i = 0; i < fileInfo.HeaderRow.Count; i++)
            {
                if (!used.Contains(i) && !_columnGroups.ContainsKey(i))
                {
                    UnassignedColumnsListBox.Items.Add(new ColumnItem { Index = i, Header = fileInfo.HeaderRow[i] });
                }
            }
        }

        private void MoveSelectedToLimit(ListBox target)
        {
            if (_currentFile == null)
            {
                return;
            }

            var selected = UnassignedColumnsListBox.SelectedItems.OfType<ColumnItem>().FirstOrDefault();
            if (selected == null)
            {
                return;
            }

            target.Items.Clear();
            target.Items.Add(new ColumnItem { Index = selected.Index, Header = selected.Header });
            RebuildLimitUnassignedList(_currentFile);
            UpdateYAxisSelection();
        }

        private void RemoveFromLimit(ListBox target)
        {
            if (_currentFile == null)
            {
                return;
            }

            target.Items.Clear();
            RebuildLimitUnassignedList(_currentFile);
            UpdateYAxisSelection();
        }

        private void MoveToRefButton_Click(object sender, RoutedEventArgs e) => MoveSelectedToLimit(RefLimitListBox);
        private void RemoveFromRefButton_Click(object sender, RoutedEventArgs e) => RemoveFromLimit(RefLimitListBox);
        private void MoveToUpperButton_Click(object sender, RoutedEventArgs e) => MoveSelectedToLimit(UpperLimitListBox);
        private void RemoveFromUpperButton_Click(object sender, RoutedEventArgs e) => RemoveFromLimit(UpperLimitListBox);
        private void MoveToLowerButton_Click(object sender, RoutedEventArgs e) => MoveSelectedToLimit(LowerLimitListBox);
        private void RemoveFromLowerButton_Click(object sender, RoutedEventArgs e) => RemoveFromLimit(LowerLimitListBox);

        private void EnsureSavedGroups(FileInfo_Custom fileInfo)
        {
            fileInfo.SavedGroups ??= new List<SavedGroupSetting>();
            if (fileInfo.SavedGroups.Count == 0)
            {
                fileInfo.SavedGroups.Add(new SavedGroupSetting { GroupId = 1, Name = "Group 1", ColorIndex = 2, LineStyleIndex = 0 });
                fileInfo.SavedGroups.Add(new SavedGroupSetting { GroupId = 2, Name = "Group 2", ColorIndex = 1, LineStyleIndex = 0 });
            }
        }

        private void BuildGroupRowsFromSavedGroups(FileInfo_Custom fileInfo)
        {
            EnsureSavedGroups(fileInfo);
            GroupRowsHost.Children.Clear();
            _groupRows.Clear();
            _nextGroupId = fileInfo.SavedGroups.Max(g => g.GroupId) + 1;
            foreach (var saved in fileInfo.SavedGroups.OrderBy(g => g.GroupId))
            {
                AddGroupRow(saved.GroupId, saved.Name, saved.ColorIndex, saved.LineStyleIndex);
            }
        }

        private void AddGroupRow(int groupId, string name, int colorIndex, int lineStyleIndex)
        {
            var rowGrid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var buttonPanel = new StackPanel { Margin = new Thickness(0, 36, 8, 0), VerticalAlignment = System.Windows.VerticalAlignment.Top };
            var addButton = new Button { Content = ">", Width = 20, Height = 20, Padding = new Thickness(0), Tag = groupId };
            addButton.Click += AddToGroupButton_Click;
            var removeButton = new Button { Content = "<", Width = 20, Height = 20, Padding = new Thickness(0), Margin = new Thickness(0, 4, 0, 0), Tag = groupId };
            removeButton.Click += RemoveFromGroupButton_Click;
            buttonPanel.Children.Add(addButton);
            buttonPanel.Children.Add(removeButton);
            Grid.SetColumn(buttonPanel, 0);
            rowGrid.Children.Add(buttonPanel);

            var groupBorder = new Border { BorderBrush = System.Windows.Media.Brushes.LightGray, BorderThickness = new Thickness(1), Padding = new Thickness(8) };
            var groupGrid = new Grid();
            groupGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            groupGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(170) });
            groupGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            groupGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
            groupGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });

            var nameTextBox = new TextBox { Text = string.IsNullOrWhiteSpace(name) ? $"Group {groupId}" : name, Margin = new Thickness(0, 0, 6, 6) };
            Grid.SetRow(nameTextBox, 0);
            Grid.SetColumn(nameTextBox, 0);
            groupGrid.Children.Add(nameTextBox);

            var colorCombo = new ComboBox { Margin = new Thickness(0, 0, 6, 6), ItemsSource = _colorOptions, DisplayMemberPath = "Name", SelectedIndex = Math.Max(0, Math.Min(colorIndex, _colorOptions.Count - 1)) };
            Grid.SetRow(colorCombo, 0);
            Grid.SetColumn(colorCombo, 1);
            groupGrid.Children.Add(colorCombo);

            var lineStyleCombo = new ComboBox { Margin = new Thickness(0, 0, 0, 6), ItemsSource = _lineStyleOptions, SelectedIndex = Math.Max(0, Math.Min(lineStyleIndex, _lineStyleOptions.Count - 1)) };
            Grid.SetRow(lineStyleCombo, 0);
            Grid.SetColumn(lineStyleCombo, 2);
            groupGrid.Children.Add(lineStyleCombo);

            var columnsList = new ListBox { SelectionMode = System.Windows.Controls.SelectionMode.Extended };
            columnsList.AllowDrop = true;
            columnsList.DragOver += GroupColumnsListBox_DragOver;
            columnsList.Drop += GroupColumnsListBox_Drop;
            Grid.SetRow(columnsList, 1);
            Grid.SetColumn(columnsList, 0);
            Grid.SetColumnSpan(columnsList, 3);
            groupGrid.Children.Add(columnsList);

            groupBorder.Child = groupGrid;
            Grid.SetColumn(groupBorder, 1);
            rowGrid.Children.Add(groupBorder);

            GroupRowsHost.Children.Add(rowGrid);
            _groupRows.Add(new GroupRow { GroupId = groupId, NameTextBox = nameTextBox, ColorComboBox = colorCombo, LineStyleComboBox = lineStyleCombo, GroupColumnsListBox = columnsList });
        }

        private void RefreshGroupLists(FileInfo_Custom fileInfo)
        {
            if (fileInfo.HeaderRow == null) return;
            var validGroupIds = _groupRows.Select(g => g.GroupId).ToHashSet();
            var invalid = _columnGroups.Where(kv => !validGroupIds.Contains(kv.Value)).Select(kv => kv.Key).ToList();
            foreach (var idx in invalid) _columnGroups.Remove(idx);

            foreach (var row in _groupRows)
            {
                row.GroupColumnsListBox.Items.Clear();
                for (int i = 0; i < fileInfo.HeaderRow.Count; i++)
                {
                    if (_columnGroups.TryGetValue(i, out int gid) && gid == row.GroupId)
                    {
                        row.GroupColumnsListBox.Items.Add(new ColumnItem { Index = i, Header = fileInfo.HeaderRow[i] });
                    }
                }
            }
            RebuildLimitUnassignedList(fileInfo);
        }

        private void AddToGroupButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile == null || sender is not Button b || b.Tag is not int gid) return;
            foreach (var selected in UnassignedColumnsListBox.SelectedItems.OfType<ColumnItem>())
            {
                _columnGroups[selected.Index] = gid;
            }
            RefreshGroupLists(_currentFile);
        }

        private void RemoveFromGroupButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile == null || sender is not Button b || b.Tag is not int gid) return;
            var row = _groupRows.FirstOrDefault(r => r.GroupId == gid);
            if (row == null) return;
            foreach (var selected in row.GroupColumnsListBox.SelectedItems.OfType<ColumnItem>().ToList())
            {
                _columnGroups.Remove(selected.Index);
            }
            RefreshGroupLists(_currentFile);
        }

        private void UnassignedColumnsListBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _unassignedDragStartPoint = e.GetPosition(UnassignedColumnsListBox);
        }

        private void UnassignedColumnsListBox_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed)
            {
                return;
            }

            var currentPos = e.GetPosition(UnassignedColumnsListBox);
            if (Math.Abs(currentPos.X - _unassignedDragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(currentPos.Y - _unassignedDragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            var selectedIndices = UnassignedColumnsListBox.SelectedItems
                .OfType<ColumnItem>()
                .Select(c => c.Index)
                .Distinct()
                .ToList();

            if (selectedIndices.Count == 0)
            {
                return;
            }

            var data = new DataObject(typeof(List<int>), selectedIndices);
            DragDrop.DoDragDrop(UnassignedColumnsListBox, data, DragDropEffects.Move);
        }

        private void GroupColumnsListBox_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(List<int>)))
            {
                e.Effects = DragDropEffects.Move;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = true;
        }

        private void GroupColumnsListBox_Drop(object sender, DragEventArgs e)
        {
            if (_currentFile == null || sender is not ListBox targetListBox)
            {
                return;
            }

            if (!e.Data.GetDataPresent(typeof(List<int>)))
            {
                return;
            }

            var row = _groupRows.FirstOrDefault(r => ReferenceEquals(r.GroupColumnsListBox, targetListBox));
            if (row == null)
            {
                return;
            }

            if (e.Data.GetData(typeof(List<int>)) is not List<int> indices || indices.Count == 0)
            {
                return;
            }

            foreach (int index in indices)
            {
                _columnGroups[index] = row.GroupId;
            }

            RefreshGroupLists(_currentFile);
            e.Handled = true;
        }

        private void GenerateGraphButton_Click(object sender, RoutedEventArgs e)
        {
            if (_loadedFiles.Count == 0)
            {
                MessageBox.Show("Please load a file first.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                // í˜„ìž¬ íŒŒì¼ì˜ ì„¤ì • ì €ìž¥
                SaveCurrentFileSettings();

                // ê° íŒŒì¼ì˜ ì„¤ì • ìˆ˜ì§‘
                var graphDataList = new List<GraphData>();

                foreach (var file in _loadedFiles)
                {
                    // Xì¶• í™•ì¸ (ì €ìž¥ëœ ì„¤ì • ì‚¬ìš©)
                    int xIndex = -1;
                    if (file.SavedXAxisIndex > 0)
                    {
                        xIndex = file.SavedXAxisIndex - 1;
                    }

                    // Yì¶• í™•ì¸ (ì €ìž¥ëœ ì„¤ì • ì‚¬ìš©)
                    var yIndices = new List<int>(file.SavedYAxisIndices);

                    if (yIndices.Count == 0)
                    {
                        MessageBox.Show($"{file.Name}: Please select at least one Y axis.", "Settings Error",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // SPEC/Upper/Lower Limit ì—´ ì¸ë±ìŠ¤ (ì €ìž¥ëœ ì„¤ì • ì‚¬ìš©)
                    int? specIndex = null;
                    int? upperIndex = null;
                    int? lowerIndex = null;

                    if (file.SavedSpecColumnIndex > 0)
                        specIndex = file.SavedSpecColumnIndex - 1;
                    if (file.SavedUpperLimitColumnIndex > 0)
                        upperIndex = file.SavedUpperLimitColumnIndex - 1;
                    if (file.SavedLowerLimitColumnIndex > 0)
                        lowerIndex = file.SavedLowerLimitColumnIndex - 1;

                    // ìƒ‰ìƒ ê°€ì ¸ì˜¤ê¸° (ì €ìž¥ëœ ì„¤ì • ì‚¬ìš©)
                    Color specColor = Colors.Black;
                    Color upperColor = Colors.Black;
                    Color lowerColor = Colors.Black;
                    Color yAxisColor = Colors.Black;

                    if (file.SavedSpecColorIndex >= 0 && file.SavedSpecColorIndex < _colorOptions.Count)
                        specColor = _colorOptions[file.SavedSpecColorIndex].Color;
                    if (file.SavedUpperColorIndex >= 0 && file.SavedUpperColorIndex < _colorOptions.Count)
                        upperColor = _colorOptions[file.SavedUpperColorIndex].Color;
                    if (file.SavedLowerColorIndex >= 0 && file.SavedLowerColorIndex < _colorOptions.Count)
                        lowerColor = _colorOptions[file.SavedLowerColorIndex].Color;
                    if (file.SavedYAxisColorIndex >= 0 && file.SavedYAxisColorIndex < _colorOptions.Count)
                        yAxisColor = _colorOptions[file.SavedYAxisColorIndex].Color;

                    var groupNameMap = new Dictionary<int, string>();
                    var groupColorMap = new Dictionary<int, Color>();
                    var groupLineStyleMap = new Dictionary<int, LineStyle>();
                    var groupRefColumnMap = new Dictionary<int, int>();
                    var groupUpperColumnMap = new Dictionary<int, int>();
                    var groupLowerColumnMap = new Dictionary<int, int>();
                    var groupRefColorMap = new Dictionary<int, Color>();
                    var groupUpperColorMap = new Dictionary<int, Color>();
                    var groupLowerColorMap = new Dictionary<int, Color>();
                    var groupRefLineStyleMap = new Dictionary<int, LineStyle>();
                    var groupUpperLineStyleMap = new Dictionary<int, LineStyle>();
                    var groupLowerLineStyleMap = new Dictionary<int, LineStyle>();
                    foreach (var group in file.SavedGroups)
                    {
                        if (group.GroupId <= 0)
                        {
                            continue;
                        }

                        groupNameMap[group.GroupId] = string.IsNullOrWhiteSpace(group.Name)
                            ? $"Group {group.GroupId}"
                            : group.Name;
                        if (group.ColorIndex >= 0 && group.ColorIndex < _colorOptions.Count)
                        {
                            groupColorMap[group.GroupId] = _colorOptions[group.ColorIndex].Color;
                        }
                        else
                        {
                            groupColorMap[group.GroupId] = Colors.Blue;
                        }

                        groupLineStyleMap[group.GroupId] = GetLineStyleFromIndex(group.LineStyleIndex);
                        if (group.RefColumnIndex >= 0)
                        {
                            groupRefColumnMap[group.GroupId] = group.RefColumnIndex;
                            groupRefColorMap[group.GroupId] = group.RefColorIndex >= 0 && group.RefColorIndex < _colorOptions.Count
                                ? _colorOptions[group.RefColorIndex].Color
                                : Colors.Black;
                            groupRefLineStyleMap[group.GroupId] = GetLineStyleFromIndex(group.RefLineStyleIndex);
                        }

                        if (group.UpperLimitColumnIndex >= 0)
                        {
                            groupUpperColumnMap[group.GroupId] = group.UpperLimitColumnIndex;
                            groupUpperColorMap[group.GroupId] = group.UpperColorIndex >= 0 && group.UpperColorIndex < _colorOptions.Count
                                ? _colorOptions[group.UpperColorIndex].Color
                                : Colors.Black;
                            groupUpperLineStyleMap[group.GroupId] = GetLineStyleFromIndex(group.UpperLineStyleIndex);
                        }

                        if (group.LowerLimitColumnIndex >= 0)
                        {
                            groupLowerColumnMap[group.GroupId] = group.LowerLimitColumnIndex;
                            groupLowerColorMap[group.GroupId] = group.LowerColorIndex >= 0 && group.LowerColorIndex < _colorOptions.Count
                                ? _colorOptions[group.LowerColorIndex].Color
                                : Colors.Black;
                            groupLowerLineStyleMap[group.GroupId] = GetLineStyleFromIndex(group.LowerLineStyleIndex);
                        }
                    }

                    var graphData = new GraphData
                    {
                        FileName = file.Name,
                        XColumnIndex = xIndex,
                        YColumnIndices = yIndices,
                        SpecColumnIndex = specIndex,
                        UpperLimitColumnIndex = upperIndex,
                        LowerLimitColumnIndex = lowerIndex,
                        Data = file.FullData,
                        HeaderRow = file.HeaderRow,
                        ColumnGroupMap = new Dictionary<int, int>(file.SavedColumnGroups),
                        SpecColor = specColor,
                        UpperColor = upperColor,
                        LowerColor = lowerColor,
                        YAxisColor = yAxisColor,
                        SpecLineStyle = GetLineStyleFromIndex(file.SavedSpecLineStyleIndex),
                        UpperLineStyle = GetLineStyleFromIndex(file.SavedUpperLineStyleIndex),
                        LowerLineStyle = GetLineStyleFromIndex(file.SavedLowerLineStyleIndex),
                        GroupNames = groupNameMap,
                        GroupColors = groupColorMap,
                        GroupLineStyles = groupLineStyleMap,
                        GroupRefColumnIndices = groupRefColumnMap,
                        GroupUpperLimitColumnIndices = groupUpperColumnMap,
                        GroupLowerLimitColumnIndices = groupLowerColumnMap,
                        GroupRefColors = groupRefColorMap,
                        GroupUpperColors = groupUpperColorMap,
                        GroupLowerColors = groupLowerColorMap,
                        GroupRefLineStyles = groupRefLineStyleMap,
                        GroupUpperLineStyles = groupUpperLineStyleMap,
                        GroupLowerLineStyles = groupLowerLineStyleMap
                    };

                    graphDataList.Add(graphData);
                }

                bool useLogXAxis = _currentFile?.SavedXAxisLogScale == true;

                if (graphDataList.Count > 0)
                {
                    var graphViewerWindow = new GraphViewerWindow(graphDataList, useLogXAxis)
                    {
                        WindowState = WindowState.Maximized,
                        Title = "SPL Scatter Graph"
                    };

                    Window? ownerWindow = Window.GetWindow(this);
                    if (ownerWindow is not null && ownerWindow.IsVisible)
                    {
                        graphViewerWindow.Owner = ownerWindow;
                    }

                    graphViewerWindow.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while creating graph:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static LineStyle GetLineStyleFromIndex(int index)
        {
            return index switch
            {
                1 => LineStyle.Dash,
                2 => LineStyle.Dot,
                3 => LineStyle.DashDot,
                _ => LineStyle.Solid
            };
        }

        private static List<SavedGroupSetting> CloneSavedGroups(List<SavedGroupSetting>? groups)
        {
            if (groups == null || groups.Count == 0)
            {
                return new List<SavedGroupSetting>
                {
                    new SavedGroupSetting { GroupId = 1, Name = "Group 1", ColorIndex = 2, LineStyleIndex = 0 },
                    new SavedGroupSetting { GroupId = 2, Name = "Group 2", ColorIndex = 1, LineStyleIndex = 0 }
                };
            }

            return groups
                .Select(g => new SavedGroupSetting
                {
                    GroupId = g.GroupId,
                    Name = g.Name,
                    ColorIndex = g.ColorIndex,
                    LineStyleIndex = g.LineStyleIndex,
                    RefColumnIndex = g.RefColumnIndex,
                    UpperLimitColumnIndex = g.UpperLimitColumnIndex,
                    LowerLimitColumnIndex = g.LowerLimitColumnIndex,
                    RefColorIndex = g.RefColorIndex,
                    UpperColorIndex = g.UpperColorIndex,
                    LowerColorIndex = g.LowerColorIndex,
                    RefLineStyleIndex = g.RefLineStyleIndex,
                    UpperLineStyleIndex = g.UpperLineStyleIndex,
                    LowerLineStyleIndex = g.LowerLineStyleIndex
                })
                .ToList();
        }

    }
}

