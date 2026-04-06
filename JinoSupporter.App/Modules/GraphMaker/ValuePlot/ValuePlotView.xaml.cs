using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Annotations;
using OxyPlot.Series;
using UserControl = System.Windows.Controls.UserControl;
using CheckBox = System.Windows.Controls.CheckBox;
using Color = System.Windows.Media.Color;
using DragEventArgs = System.Windows.DragEventArgs;
using MessageBox = System.Windows.MessageBox;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace GraphMaker
{
    public class FileInfo_DailySampling
    {
        public string Name { get; set; }
        public string FilePath { get; set; }
        public DataTable FullData { get; set; }
        public List<string> Dates { get; set; }
        public List<string> SampleNumbers { get; set; }
        public string Delimiter { get; set; } = "\t";
        public int HeaderRowNumber { get; set; } = 1;

        // ?? ??
        public int SavedDataColorIndex { get; set; } = 0;
        public int SavedSpecColorIndex { get; set; } = 0;
        public int SavedUpperColorIndex { get; set; } = 0;
        public int SavedLowerColorIndex { get; set; } = 0;
        public string SavedSpecValue { get; set; } = "";
        public string SavedUpperValue { get; set; } = "";
        public string SavedLowerValue { get; set; } = "";
        public Dictionary<string, ColumnLimitSetting> SavedColumnLimits { get; set; } =
            new Dictionary<string, ColumnLimitSetting>(StringComparer.Ordinal);
        public bool IsXAxisDate { get; set; } = true; // true: ??, false: ??
        public XAxisValueType? ForcedXAxisType { get; set; }
        public XAxisMode SavedXAxisMode { get; set; } = XAxisMode.Date;
        public NoXAxisDisplayMode SavedNoXAxisDisplayMode { get; set; } = NoXAxisDisplayMode.SampleOrderOnXAxis;
        public ValuePlotDisplayMode SavedDisplayMode { get; set; } = ValuePlotDisplayMode.Combined;
    }

    public partial class ValuePlotView : UserControl, INotifyPropertyChanged
    {
        private sealed class ValuePlotReportState
        {
            public List<ValuePlotFileState> Files { get; set; } = new();
            public string? SelectedFilePath { get; set; }
        }

        private sealed class ValuePlotFileState
        {
            public string Name { get; set; } = string.Empty;
            public string FilePath { get; set; } = string.Empty;
            public string Delimiter { get; set; } = "\t";
            public int HeaderRowNumber { get; set; } = 1;
            public RawTableData RawData { get; set; } = new();
            public int SavedDataColorIndex { get; set; }
            public int SavedSpecColorIndex { get; set; }
            public int SavedUpperColorIndex { get; set; }
            public int SavedLowerColorIndex { get; set; }
            public string SavedSpecValue { get; set; } = string.Empty;
            public string SavedUpperValue { get; set; } = string.Empty;
            public string SavedLowerValue { get; set; } = string.Empty;
            public bool IsXAxisDate { get; set; } = true;
        }

        private ObservableCollection<FileInfo_DailySampling> _loadedFiles = new ObservableCollection<FileInfo_DailySampling>();
        private FileInfo_DailySampling _currentFile;
        private List<PreviewColorChoice> _colorOptions;

        private PlotModel _scatterPlotModel;
        private PlotModel _cpkPlotModel;

        public PlotModel ScatterPlotModel
        {
            get => _scatterPlotModel;
            set { _scatterPlotModel = value; OnPropertyChanged(nameof(ScatterPlotModel)); }
        }

        public PlotModel CpkPlotModel
        {
            get => _cpkPlotModel;
            set { _cpkPlotModel = value; OnPropertyChanged(nameof(CpkPlotModel)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public ValuePlotView()
        {
            InitializeComponent();
            DataContext = this;
            FileListBox.ItemsSource = _loadedFiles;
            InitializeColorOptions();
        }

        private void InitializeColorOptions()
        {
            _colorOptions = new List<PreviewColorChoice>();
            PreviewGraphViewBase.InitializeDefaultColorOptions(DataColorComboBox, _colorOptions);
            PreviewGraphViewBase.BindColorComboBox(SpecColorComboBox, _colorOptions);
            PreviewGraphViewBase.BindColorComboBox(UpperColorComboBox, _colorOptions);
            PreviewGraphViewBase.BindColorComboBox(LowerColorComboBox, _colorOptions);
        }

        #region Drag and Drop

        private void DropZone_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (ext == ".txt" || ext == ".csv")
                    {
                        LoadFile(file);
                    }
                }
            }
        }

        private void DropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                SetDropHintForeground(sender, Colors.Blue);
            }
        }

        private void DropZone_DragLeave(object sender, DragEventArgs e)
        {
            SetDropHintForeground(sender, Color.FromRgb(102, 102, 102));
        }

        private void DropZone_DragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private static void SetDropHintForeground(object sender, Color color)
        {
            if (sender is Border border &&
                border.Child is Panel panel &&
                panel.Children.OfType<TextBlock>().FirstOrDefault() is TextBlock textBlock)
            {
                textBlock.Foreground = new SolidColorBrush(color);
            }
        }

        #endregion

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
                // ?? ??? ???? ??
                if (_loadedFiles.Any(f => f.FilePath == filePath))
                {
                    MessageBox.Show($"{Path.GetFileName(filePath)} is already loaded.",
                        "Duplicate", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var fileInfo = new FileInfo_DailySampling
                {
                    Name = Path.GetFileName(filePath),
                    FilePath = filePath
                };

                _loadedFiles.Add(fileInfo);

                // ? ?? ???? ?? ??
                if (_loadedFiles.Count == 1)
                {
                    FileListBox.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while loading file: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void FileListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // ?? ??? ?? ??
            if (_currentFile != null)
            {
                SaveCurrentFileSettings();
            }

            if (FileListBox.SelectedItem is FileInfo_DailySampling selectedFile)
            {
                _currentFile = selectedFile;
                LoadFileData(selectedFile);
            }
        }

        private void SaveCurrentFileSettings()
        {
            if (_currentFile == null) return;

            // ?? ??
            _currentFile.SavedDataColorIndex = DataColorComboBox.SelectedIndex;
            _currentFile.SavedSpecColorIndex = SpecColorComboBox.SelectedIndex;
            _currentFile.SavedUpperColorIndex = UpperColorComboBox.SelectedIndex;
            _currentFile.SavedLowerColorIndex = LowerColorComboBox.SelectedIndex;

            // SPEC/Limit ? ??
            _currentFile.SavedSpecValue = SpecValueTextBox.Text;
            _currentFile.SavedUpperValue = UpperLimitValueTextBox.Text;
            _currentFile.SavedLowerValue = LowerLimitValueTextBox.Text;

            // X? ?? ??
            _currentFile.IsXAxisDate = XAxisDateRadio.IsChecked == true;
        }

        private void LoadFileData(FileInfo_DailySampling fileInfo)
        {
            try
            {
                // ?? ??? ???? ???? ?? ?? ?? UI? ????
                if (fileInfo.FullData != null && fileInfo.Dates != null && fileInfo.SampleNumbers != null)
                {
                    // UI ????? ??
                    CurrentFileNameText.Text = fileInfo.Name;
                    RowCountText.Text = fileInfo.FullData.Rows.Count.ToString("N0");
                    ColumnCountText.Text = fileInfo.FullData.Columns.Count.ToString();
                    DataPreviewGrid.ItemsSource = fileInfo.FullData.DefaultView;

                    // ??? ?? ??
                    RestoreSettings(fileInfo);
                    return;
                }

                // ?? ???? ??: ?? ??
                var lines = File.ReadAllLines(fileInfo.FilePath);
                if (lines.Length < 2)
                {
                    MessageBox.Show("Invalid file format. At least 2 rows are required.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int headerRowIndex = fileInfo.HeaderRowNumber - 1;
                if (headerRowIndex >= lines.Length)
                {
                    MessageBox.Show("Header row number exceeds total rows.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // ? ?: ?? ?? (? ?? "??" ?? ??)
                var headerRow = lines[headerRowIndex].Split(new[] { fileInfo.Delimiter }, StringSplitOptions.None);
                fileInfo.SampleNumbers = new List<string>();

                // DataTable ??
                var dataTable = new DataTable();
                dataTable.Columns.Add("Date");

                for (int i = 1; i < headerRow.Length; i++)
                {
                    fileInfo.SampleNumbers.Add(headerRow[i]);
                    dataTable.Columns.Add(headerRow[i]);
                }

                // ?? ? ??? ? ??
                fileInfo.Dates = new List<string>();

                for (int i = headerRowIndex + 1; i < lines.Length; i++)
                {
                    var values = lines[i].Split(new[] { fileInfo.Delimiter }, StringSplitOptions.None);
                    if (values.Length > 0 && !string.IsNullOrWhiteSpace(values[0]))
                    {
                        fileInfo.Dates.Add(values[0]);

                        var row = dataTable.NewRow();
                        row[0] = values[0]; // ??
                        for (int j = 1; j < Math.Min(values.Length, headerRow.Length); j++)
                        {
                            row[j] = values[j];
                        }
                        dataTable.Rows.Add(row);
                    }
                }

                fileInfo.FullData = dataTable;

                // UI ????
                CurrentFileNameText.Text = fileInfo.Name;
                RowCountText.Text = dataTable.Rows.Count.ToString("N0");
                ColumnCountText.Text = dataTable.Columns.Count.ToString();
                DataPreviewGrid.ItemsSource = dataTable.DefaultView;

                // ??? ?? ?? (??? ???)
                RestoreSettings(fileInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while loading data: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RestoreSettings(FileInfo_DailySampling fileInfo)
        {
            // ?? ??
            if (fileInfo.SavedDataColorIndex >= 0 && fileInfo.SavedDataColorIndex < DataColorComboBox.Items.Count)
                DataColorComboBox.SelectedIndex = fileInfo.SavedDataColorIndex;
            if (fileInfo.SavedSpecColorIndex >= 0 && fileInfo.SavedSpecColorIndex < SpecColorComboBox.Items.Count)
                SpecColorComboBox.SelectedIndex = fileInfo.SavedSpecColorIndex;
            if (fileInfo.SavedUpperColorIndex >= 0 && fileInfo.SavedUpperColorIndex < UpperColorComboBox.Items.Count)
                UpperColorComboBox.SelectedIndex = fileInfo.SavedUpperColorIndex;
            if (fileInfo.SavedLowerColorIndex >= 0 && fileInfo.SavedLowerColorIndex < LowerColorComboBox.Items.Count)
                LowerColorComboBox.SelectedIndex = fileInfo.SavedLowerColorIndex;

            // SPEC/Limit ? ??
            SpecValueTextBox.Text = fileInfo.SavedSpecValue;
            UpperLimitValueTextBox.Text = fileInfo.SavedUpperValue;
            LowerLimitValueTextBox.Text = fileInfo.SavedLowerValue;

            // X? ?? ??
            XAxisDateRadio.IsChecked = fileInfo.IsXAxisDate;
            XAxisSequenceRadio.IsChecked = !fileInfo.IsXAxisDate;
        }

        private DateTime? ParseDate(string dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr))
                return null;

            // ??? ?? ?? ??
            var formats = new[]
            {
                "yyyy-MM-dd",
                "yyyy/MM/dd",
                "MM-dd",
                "MM/dd",
                "M-d",
                "M/d"
            };

            foreach (var format in formats)
            {
                if (DateTime.TryParseExact(dateStr, format, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out DateTime result))
                {
                    // MM-dd ?? MM/dd ??? ?? ?? ?? ??
                    if (format.StartsWith("M"))
                    {
                        result = new DateTime(DateTime.Now.Year, result.Month, result.Day);
                    }
                    return result;
                }
            }

            // ?? ?? ??
            if (DateTime.TryParse(dateStr, out DateTime parsedDate))
            {
                return parsedDate;
            }

            return null;
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

            // ??? ?? ??
            _currentFile.FullData = null;
            _currentFile.Dates = null;
            _currentFile.SampleNumbers = null;
            LoadFileData(_currentFile);
        }

        private void HeaderRowChanged(object sender, TextChangedEventArgs e)
        {
            if (_currentFile == null) return;

            if (int.TryParse(HeaderRowTextBox.Text, out int headerRow) && headerRow > 0)
            {
                _currentFile.HeaderRowNumber = headerRow;
                // ??? ?? ??
                _currentFile.FullData = null;
                _currentFile.Dates = null;
                _currentFile.SampleNumbers = null;
                LoadFileData(_currentFile);
            }
        }

        private void RemoveFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (FileListBox.SelectedItem is FileInfo_DailySampling selectedFile)
            {
                // ?? ??? ??? ??? ???? ?? ?? ???
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
                    DataPreviewGrid.ItemsSource = null;
                    CurrentFileNameText.Text = "(No file selected)";
                    RowCountText.Text = "0";
                    ColumnCountText.Text = "0";
                }
            }
        }

        private void SaveReportButton_Click(object sender, RoutedEventArgs e)
        {
            if (_loadedFiles.Count == 0)
            {
                MessageBox.Show("Load data first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            SaveCurrentFileSettings();

            var state = new ValuePlotReportState
            {
                SelectedFilePath = _currentFile?.FilePath,
                Files = _loadedFiles.Select(file => new ValuePlotFileState
                {
                    Name = file.Name,
                    FilePath = file.FilePath,
                    Delimiter = file.Delimiter,
                    HeaderRowNumber = file.HeaderRowNumber,
                    RawData = GraphReportStorageHelper.CaptureRawTableData(file.FullData),
                    SavedDataColorIndex = file.SavedDataColorIndex,
                    SavedSpecColorIndex = file.SavedSpecColorIndex,
                    SavedUpperColorIndex = file.SavedUpperColorIndex,
                    SavedLowerColorIndex = file.SavedLowerColorIndex,
                    SavedSpecValue = file.SavedSpecValue ?? string.Empty,
                    SavedUpperValue = file.SavedUpperValue ?? string.Empty,
                    SavedLowerValue = file.SavedLowerValue ?? string.Empty,
                    IsXAxisDate = file.IsXAxisDate
                }).ToList()
            };

            GraphReportFileDialogHelper.SaveState("Save Graph Report", "valueplot.graphreport.json", state);
        }

        private void LoadReportButton_Click(object sender, RoutedEventArgs e)
        {
            ValuePlotReportState? state = GraphReportFileDialogHelper.LoadState<ValuePlotReportState>("Load Graph Report");
            if (state == null || state.Files.Count == 0)
            {
                MessageBox.Show("Invalid report file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _loadedFiles.Clear();
            _currentFile = null;

            foreach (ValuePlotFileState fileState in state.Files)
            {
                var table = GraphReportStorageHelper.BuildTableFromRawData(fileState.RawData);
                var fileInfo = new FileInfo_DailySampling
                {
                    Name = fileState.Name,
                    FilePath = fileState.FilePath,
                    Delimiter = string.IsNullOrWhiteSpace(fileState.Delimiter) ? "\t" : fileState.Delimiter,
                    HeaderRowNumber = Math.Max(1, fileState.HeaderRowNumber),
                    FullData = table,
                    Dates = table.Rows.Cast<DataRow>().Select(row => row[0]?.ToString() ?? string.Empty).ToList(),
                    SampleNumbers = table.Columns.Cast<DataColumn>().Skip(1).Select(column => column.ColumnName).ToList(),
                    SavedDataColorIndex = fileState.SavedDataColorIndex,
                    SavedSpecColorIndex = fileState.SavedSpecColorIndex,
                    SavedUpperColorIndex = fileState.SavedUpperColorIndex,
                    SavedLowerColorIndex = fileState.SavedLowerColorIndex,
                    SavedSpecValue = fileState.SavedSpecValue,
                    SavedUpperValue = fileState.SavedUpperValue,
                    SavedLowerValue = fileState.SavedLowerValue,
                    IsXAxisDate = fileState.IsXAxisDate
                };

                _loadedFiles.Add(fileInfo);
            }

            FileInfo_DailySampling? selectedFile = _loadedFiles.FirstOrDefault(file =>
                string.Equals(file.FilePath, state.SelectedFilePath, StringComparison.OrdinalIgnoreCase))
                ?? _loadedFiles.FirstOrDefault();

            if (selectedFile != null)
            {
                FileListBox.SelectedItem = selectedFile;
            }
        }

        private void OpenFileSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile == null)
            {
                MessageBox.Show("Select a file first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            LoadFileData(_currentFile);
        }

        private void GenerateGraphButton_Click(object sender, RoutedEventArgs e)
        {
            if (_loadedFiles.Count == 0)
            {
                MessageBox.Show("Please load data files first.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                // ?? ??? ?? ??
                SaveCurrentFileSettings();

                // ? ??? ?? ??
                var graphDataList = new List<DailySamplingGraphData>();

                foreach (var file in _loadedFiles)
                {
                    // ??? ??? ??
                    if (file.FullData == null || file.Dates == null || file.SampleNumbers == null)
                    {
                        MessageBox.Show($"{file.Name}: Required data is missing.", "Settings Error",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // SPEC/Limit ? (?? ??)
                    double? specValue = null;
                    double? upperValue = null;
                    double? lowerValue = null;

                    if (!string.IsNullOrWhiteSpace(file.SavedSpecValue) &&
                        double.TryParse(file.SavedSpecValue, out double spec))
                        specValue = spec;
                    if (!string.IsNullOrWhiteSpace(file.SavedUpperValue) &&
                        double.TryParse(file.SavedUpperValue, out double upper))
                        upperValue = upper;
                    if (!string.IsNullOrWhiteSpace(file.SavedLowerValue) &&
                        double.TryParse(file.SavedLowerValue, out double lower))
                        lowerValue = lower;

                    // ?? ????
                    Color dataColor = Colors.Black;
                    Color specColor = Colors.Black;
                    Color upperColor = Colors.Black;
                    Color lowerColor = Colors.Black;

                    if (file.SavedDataColorIndex >= 0 && file.SavedDataColorIndex < _colorOptions.Count)
                        dataColor = _colorOptions[file.SavedDataColorIndex].Color;
                    if (file.SavedSpecColorIndex >= 0 && file.SavedSpecColorIndex < _colorOptions.Count)
                        specColor = _colorOptions[file.SavedSpecColorIndex].Color;
                    if (file.SavedUpperColorIndex >= 0 && file.SavedUpperColorIndex < _colorOptions.Count)
                        upperColor = _colorOptions[file.SavedUpperColorIndex].Color;
                    if (file.SavedLowerColorIndex >= 0 && file.SavedLowerColorIndex < _colorOptions.Count)
                        lowerColor = _colorOptions[file.SavedLowerColorIndex].Color;

                    var graphData = new DailySamplingGraphData
                    {
                        FileName = file.Name,
                        Dates = file.Dates,
                        SampleNumbers = file.SampleNumbers,
                        Data = file.FullData,
                        SpecValue = specValue,
                        UpperLimitValue = upperValue,
                        LowerLimitValue = lowerValue,
                        DataColor = dataColor,
                        SpecColor = specColor,
                        UpperColor = upperColor,
                        LowerColor = lowerColor,
                        IsXAxisDate = file.IsXAxisDate
                    };

                    graphDataList.Add(graphData);
                }

                // ??? ?? (? ? ?????? ??)
                var graphViewerWindow = new DailySamplingGraphViewerWindow(graphDataList, this)
                {
                    Owner = Window.GetWindow(this),
                    WindowState = WindowState.Maximized
                };
                graphViewerWindow.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while creating graph:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GenerateCharts(List<DailySamplingGraphData> graphDataList)
        {
            try
            {
                ScatterPlotModel = CreateScatterPlot(graphDataList);
                CpkPlotModel = CreateCpkPlot(graphDataList);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while creating graph:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // DailySamplingGraphViewerWindow? CreateScatterPlot ??? ??? ??
        private PlotModel CreateScatterPlot(List<DailySamplingGraphData> graphDataList)
        {
            var model = new PlotModel { Title = "Value Plot" };

            // X? ?? ?? (? ?? ?? ??)
            bool useDate = graphDataList.Count > 0 && graphDataList[0].IsXAxisDate;

            // ?? NG Rate ??? ??
            int totalNgPoints = 0;
            int totalPoints = 0;

            // X? ??
            if (useDate)
            {
                // X?: ?? (DateTimeAxis)
                var dateAxis = new DateTimeAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = "Date",
                    StringFormat = "MM-dd",
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot,
                    IntervalType = DateTimeIntervalType.Days
                };
                model.Axes.Add(dateAxis);
            }
            else
            {
                // X?: ?? (LinearAxis)
                var sequenceAxis = new LinearAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = "Sequence",
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot
                };
                model.Axes.Add(sequenceAxis);
            }

            // Y?: ?? ?
            var valueAxis = new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "?",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            };
            model.Axes.Add(valueAxis);

            // ? ??? ?? ??? ??
            foreach (var graphData in graphDataList)
            {
                var scatterSeries = new ScatterSeries
                {
                    Title = "Data",
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 3,
                    MarkerFill = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                                                    graphData.DataColor.G, graphData.DataColor.B)
                };

                if (useDate)
                {
                    // ??? ??? ?? ???? (?? ?? ??)
                    var dateStats = new Dictionary<DateTime, List<double>>();
                    var dateNgCounts = new Dictionary<DateTime, int>();  // ??? NG ??? ?
                    var dateTotalCounts = new Dictionary<DateTime, int>();  // ??? ? ??? ?

                    // ? ?? ??? ???? ?? ??
                    for (int rowIndex = 0; rowIndex < graphData.Dates.Count; rowIndex++)
                    {
                        var dateStr = graphData.Dates[rowIndex];
                        var parsedDate = ParseDateString(dateStr);

                        if (parsedDate.HasValue)
                        {
                            var dateValue = DateTimeAxis.ToDouble(parsedDate.Value);

                            // ???? ? ??? (?? ??? ?? ?? ???? ???)
                            if (!dateStats.ContainsKey(parsedDate.Value))
                            {
                                dateStats[parsedDate.Value] = new List<double>();
                                dateNgCounts[parsedDate.Value] = 0;
                                dateTotalCounts[parsedDate.Value] = 0;
                            }

                            // ?? ??? ?? ?? ? (2??? ???)
                            var dataRow = graphData.Data.Rows[rowIndex];
                            for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
                            {
                                var valueStr = dataRow[colIndex]?.ToString();
                                if (!string.IsNullOrWhiteSpace(valueStr) &&
                                    double.TryParse(valueStr, out double value))
                                {
                                    scatterSeries.Points.Add(new ScatterPoint(dateValue, value));
                                    dateStats[parsedDate.Value].Add(value);

                                    // NG ?? (Upper/Lower Limit? ?? ??)
                                    if (graphData.UpperLimitValue.HasValue && graphData.LowerLimitValue.HasValue)
                                    {
                                        dateTotalCounts[parsedDate.Value]++;
                                        totalPoints++;

                                        if (value > graphData.UpperLimitValue.Value || value < graphData.LowerLimitValue.Value)
                                        {
                                            dateNgCounts[parsedDate.Value]++;
                                            totalNgPoints++;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    model.Series.Add(scatterSeries);

                    // Avg ? ?? (?? ???? ??, ??)
                    var avgSeries = new LineSeries
                    {
                        Title = "Avg",
                        Color = OxyColors.Green,
                        StrokeThickness = 2
                    };

                    // ??? ?? ?? ? ??? ?? ??
                    foreach (var kvp in dateStats.OrderBy(x => x.Key))
                    {
                        var date = kvp.Key;
                        var values = kvp.Value;

                        if (values.Count > 0)
                        {
                            var dateValue = DateTimeAxis.ToDouble(date);
                            var max = values.Max();
                            var min = values.Min();
                            var avg = values.Average();

                            // Avg ?? ??? ??
                            avgSeries.Points.Add(new DataPoint(dateValue, avg));

                            // Max ??? ??
                            var maxAnnotation = new TextAnnotation
                            {
                                Text = $"Max: {max:F2}",
                                TextPosition = new DataPoint(dateValue, max),
                                TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                TextVerticalAlignment = OxyPlot.VerticalAlignment.Bottom,
                                Stroke = OxyColors.Transparent,
                                FontSize = 9,
                                TextColor = OxyColors.Red
                            };
                            model.Annotations.Add(maxAnnotation);

                            // Min ??? ??
                            var minAnnotation = new TextAnnotation
                            {
                                Text = $"Min: {min:F2}",
                                TextPosition = new DataPoint(dateValue, min),
                                TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                TextVerticalAlignment = OxyPlot.VerticalAlignment.Top,
                                Stroke = OxyColors.Transparent,
                                FontSize = 9,
                                TextColor = OxyColors.Blue
                            };
                            model.Annotations.Add(minAnnotation);

                            // Avg ??? ??
                            var avgAnnotation = new TextAnnotation
                            {
                                Text = $"Avg: {avg:F2}",
                                TextPosition = new DataPoint(dateValue, avg),
                                TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                                Stroke = OxyColors.Transparent,
                                FontSize = 9,
                                TextColor = OxyColors.Green
                            };
                            model.Annotations.Add(avgAnnotation);

                            // ??? NG Rate ?? ?? (Limit? ?? ??)
                            if (graphData.UpperLimitValue.HasValue && graphData.LowerLimitValue.HasValue &&
                                dateTotalCounts.ContainsKey(date) && dateTotalCounts[date] > 0)
                            {
                                int ngCount = dateNgCounts[date];
                                int totalCount = dateTotalCounts[date];
                                double ngRate = (double)ngCount / totalCount * 100.0;

                                // Y? ?? ??? NG Rate ??
                                var yPosition = (max + min) / 2.0;
                                var ngRateAnnotation = new TextAnnotation
                                {
                                    Text = $"NG: {ngCount}/{totalCount} ({ngRate:F1}%)",
                                    TextPosition = new DataPoint(dateValue, yPosition),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Right,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    FontWeight = OxyPlot.FontWeights.Bold,
                                    TextColor = ngCount > 0 ? OxyColors.Red : OxyColors.DarkGreen
                                };
                                model.Annotations.Add(ngRateAnnotation);
                            }
                        }
                    }

                    model.Series.Add(avgSeries);
                }
                else
                {
                    // ??? ??? ?? ????
                    var sequenceStats = new Dictionary<int, List<double>>();
                    var sequenceNgCounts = new Dictionary<int, int>();  // ??? NG ??? ?
                    var sequenceTotalCounts = new Dictionary<int, int>();  // ??? ? ??? ?

                    // ? ?? ???? ??? ??
                    for (int rowIndex = 0; rowIndex < graphData.Dates.Count; rowIndex++)
                    {
                        int sequence = rowIndex + 1; // 1?? ???? ??

                        // ???? ? ???
                        if (!sequenceStats.ContainsKey(sequence))
                        {
                            sequenceStats[sequence] = new List<double>();
                            sequenceNgCounts[sequence] = 0;
                            sequenceTotalCounts[sequence] = 0;
                        }

                        // ?? ??? ?? ?? ? (2??? ???)
                        var dataRow = graphData.Data.Rows[rowIndex];
                        for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
                        {
                            var valueStr = dataRow[colIndex]?.ToString();
                            if (!string.IsNullOrWhiteSpace(valueStr) &&
                                double.TryParse(valueStr, out double value))
                            {
                                scatterSeries.Points.Add(new ScatterPoint(sequence, value));
                                sequenceStats[sequence].Add(value);

                                // NG ?? (Upper/Lower Limit? ?? ??)
                                if (graphData.UpperLimitValue.HasValue && graphData.LowerLimitValue.HasValue)
                                {
                                    sequenceTotalCounts[sequence]++;
                                    totalPoints++;

                                    if (value > graphData.UpperLimitValue.Value || value < graphData.LowerLimitValue.Value)
                                    {
                                        sequenceNgCounts[sequence]++;
                                        totalNgPoints++;
                                    }
                                }
                            }
                        }
                    }

                    model.Series.Add(scatterSeries);

                    // Avg ? ??
                    var avgSeries = new LineSeries
                    {
                        Title = "Avg",
                        Color = OxyColors.Green,
                        StrokeThickness = 2
                    };

                    // ??? ?? ?? ? ??? ?? ??
                    foreach (var kvp in sequenceStats.OrderBy(x => x.Key))
                    {
                        var sequence = kvp.Key;
                        var values = kvp.Value;

                        if (values.Count > 0)
                        {
                            var max = values.Max();
                            var min = values.Min();
                            var avg = values.Average();

                            // Avg ?? ??? ??
                            avgSeries.Points.Add(new DataPoint(sequence, avg));

                            // Max ??? ??
                            var maxAnnotation = new TextAnnotation
                            {
                                Text = $"Max: {max:F2}",
                                TextPosition = new DataPoint(sequence, max),
                                TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                TextVerticalAlignment = OxyPlot.VerticalAlignment.Bottom,
                                Stroke = OxyColors.Transparent,
                                FontSize = 9,
                                TextColor = OxyColors.Red
                            };
                            model.Annotations.Add(maxAnnotation);

                            // Min ??? ??
                            var minAnnotation = new TextAnnotation
                            {
                                Text = $"Min: {min:F2}",
                                TextPosition = new DataPoint(sequence, min),
                                TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                TextVerticalAlignment = OxyPlot.VerticalAlignment.Top,
                                Stroke = OxyColors.Transparent,
                                FontSize = 9,
                                TextColor = OxyColors.Blue
                            };
                            model.Annotations.Add(minAnnotation);

                            // Avg ??? ??
                            var avgAnnotation = new TextAnnotation
                            {
                                Text = $"Avg: {avg:F2}",
                                TextPosition = new DataPoint(sequence, avg),
                                TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                                Stroke = OxyColors.Transparent,
                                FontSize = 9,
                                TextColor = OxyColors.Green
                            };
                            model.Annotations.Add(avgAnnotation);

                            // ??? NG Rate ?? ?? (Limit? ?? ??)
                            if (graphData.UpperLimitValue.HasValue && graphData.LowerLimitValue.HasValue &&
                                sequenceTotalCounts.ContainsKey(sequence) && sequenceTotalCounts[sequence] > 0)
                            {
                                int ngCount = sequenceNgCounts[sequence];
                                int totalCount = sequenceTotalCounts[sequence];
                                double ngRate = (double)ngCount / totalCount * 100.0;

                                // Y? ?? ??? NG Rate ??
                                var yPosition = (max + min) / 2.0;
                                var ngRateAnnotation = new TextAnnotation
                                {
                                    Text = $"NG: {ngCount}/{totalCount} ({ngRate:F1}%)",
                                    TextPosition = new DataPoint(sequence, yPosition),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Right,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    FontWeight = OxyPlot.FontWeights.Bold,
                                    TextColor = ngCount > 0 ? OxyColors.Red : OxyColors.DarkGreen
                                };
                                model.Annotations.Add(ngRateAnnotation);
                            }
                        }
                    }

                    model.Series.Add(avgSeries);
                }

                // SPEC/Limit ? ?? (?? ??)
                AddLimitLines(model, graphData, useDate);
            }

            // ?? NG Rate? ??? ??
            if (totalPoints > 0)
            {
                double totalNgRate = (double)totalNgPoints / totalPoints * 100.0;
                model.Title = $"Value Plot - Total NG: {totalNgPoints}/{totalPoints} (NG Rate: {totalNgRate:F1}%)";
            }

            model.IsLegendVisible = true;
            return model;
        }

        private void AddLimitLines(PlotModel model, DailySamplingGraphData graphData, bool useDate)
        {
            // ?? ??? ?? ?? ??? ?? SPEC/Limit ? ?? ??? DailySamplingGraphViewerWindow? ??
            // ???? ??? ?? (??? ??)
        }

        private PlotModel CreateCpkPlot(List<DailySamplingGraphData> graphDataList)
        {
            // DailySamplingGraphViewerWindow? CreateCpkPlot ?? ??
            var model = new PlotModel { Title = "Value CPK" };

            // X? ?? ?? (? ?? ?? ??)
            bool useDate = graphDataList.Count > 0 && graphDataList[0].IsXAxisDate;

            // X? ??
            if (useDate)
            {
                // X?: ?? (DateTimeAxis)
                var dateAxis = new DateTimeAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = "Date",
                    StringFormat = "MM-dd",
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot,
                    IntervalType = DateTimeIntervalType.Days
                };
                model.Axes.Add(dateAxis);
            }
            else
            {
                // X?: ?? (LinearAxis)
                var sequenceAxis = new LinearAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = "Sequence",
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot
                };
                model.Axes.Add(sequenceAxis);
            }

            // Y?: CPK ?
            var cpkAxis = new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "CPK",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            };
            model.Axes.Add(cpkAxis);

            // ? ??? ?? CPK ??? ??
            foreach (var graphData in graphDataList)
            {
                // SPEC, Upper, Lower ?? ?? ??? CPK ?? ??
                if (!graphData.SpecValue.HasValue ||
                    !graphData.UpperLimitValue.HasValue ||
                    !graphData.LowerLimitValue.HasValue)
                {
                    continue;
                }

                var cpkSeries = new LineSeries
                {
                    Title = "CPK",
                    Color = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                                              graphData.DataColor.G, graphData.DataColor.B),
                    StrokeThickness = 2,
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 4,
                    MarkerFill = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                                                    graphData.DataColor.G, graphData.DataColor.B)
                };

                if (useDate)
                {
                    // ???? ??? ??? (?? ?? ??)
                    var dateValues = new Dictionary<DateTime, List<double>>();

                    for (int rowIndex = 0; rowIndex < graphData.Dates.Count; rowIndex++)
                    {
                        var dateStr = graphData.Dates[rowIndex];
                        var parsedDate = ParseDateString(dateStr);

                        if (parsedDate.HasValue)
                        {
                            if (!dateValues.ContainsKey(parsedDate.Value))
                            {
                                dateValues[parsedDate.Value] = new List<double>();
                            }

                            // ?? ??? ?? ?? ? ??
                            var dataRow = graphData.Data.Rows[rowIndex];
                            for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
                            {
                                var valueStr = dataRow[colIndex]?.ToString();
                                if (!string.IsNullOrWhiteSpace(valueStr) &&
                                    double.TryParse(valueStr, out double value))
                                {
                                    dateValues[parsedDate.Value].Add(value);
                                }
                            }
                        }
                    }

                    // ??? CPK ??
                    foreach (var kvp in dateValues.OrderBy(x => x.Key))
                    {
                        var date = kvp.Key;
                        var values = kvp.Value;

                        if (values.Count > 0)
                        {
                            var dateValue = DateTimeAxis.ToDouble(date);
                            var mean = values.Average();
                            var variance = values.Sum(v => Math.Pow(v - mean, 2)) / values.Count;
                            var stdDev = Math.Sqrt(variance);

                            if (stdDev > 0)
                            {
                                var usl = graphData.UpperLimitValue.Value;
                                var lsl = graphData.LowerLimitValue.Value;

                                var cpkUpper = (usl - mean) / (3 * stdDev);
                                var cpkLower = (mean - lsl) / (3 * stdDev);
                                var cpk = Math.Min(cpkUpper, cpkLower);

                                cpkSeries.Points.Add(new DataPoint(dateValue, cpk));

                                // CPK ? ??? ?? ??
                                var cpkAnnotation = new TextAnnotation
                                {
                                    Text = $"{cpk:F2}",
                                    TextPosition = new DataPoint(dateValue, cpk),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Center,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Bottom,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    TextColor = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                                                                  graphData.DataColor.G, graphData.DataColor.B)
                                };
                                model.Annotations.Add(cpkAnnotation);
                            }
                        }
                    }
                }
                else
                {
                    // ???? ??? ???
                    var sequenceValues = new Dictionary<int, List<double>>();

                    for (int rowIndex = 0; rowIndex < graphData.Dates.Count; rowIndex++)
                    {
                        int sequence = rowIndex + 1;

                        if (!sequenceValues.ContainsKey(sequence))
                        {
                            sequenceValues[sequence] = new List<double>();
                        }

                        // ?? ??? ?? ?? ? ??
                        var dataRow = graphData.Data.Rows[rowIndex];
                        for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
                        {
                            var valueStr = dataRow[colIndex]?.ToString();
                            if (!string.IsNullOrWhiteSpace(valueStr) &&
                                double.TryParse(valueStr, out double value))
                            {
                                sequenceValues[sequence].Add(value);
                            }
                        }
                    }

                    // ??? CPK ??
                    foreach (var kvp in sequenceValues.OrderBy(x => x.Key))
                    {
                        var sequence = kvp.Key;
                        var values = kvp.Value;

                        if (values.Count > 0)
                        {
                            var mean = values.Average();
                            var variance = values.Sum(v => Math.Pow(v - mean, 2)) / values.Count;
                            var stdDev = Math.Sqrt(variance);

                            if (stdDev > 0)
                            {
                                var usl = graphData.UpperLimitValue.Value;
                                var lsl = graphData.LowerLimitValue.Value;

                                var cpkUpper = (usl - mean) / (3 * stdDev);
                                var cpkLower = (mean - lsl) / (3 * stdDev);
                                var cpk = Math.Min(cpkUpper, cpkLower);

                                cpkSeries.Points.Add(new DataPoint(sequence, cpk));

                                // CPK ? ??? ?? ??
                                var cpkAnnotation = new TextAnnotation
                                {
                                    Text = $"{cpk:F2}",
                                    TextPosition = new DataPoint(sequence, cpk),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Center,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Bottom,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    TextColor = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                                                                  graphData.DataColor.G, graphData.DataColor.B)
                                };
                                model.Annotations.Add(cpkAnnotation);
                            }
                        }
                    }
                }

                model.Series.Add(cpkSeries);
            }

            model.IsLegendVisible = true;
            return model;
        }

        // ?? ?? ?? ??? (public?? ???? ?? ??)
        public DateTime? ParseDateString(string dateStr)
        {
            return ParseDate(dateStr);
        }
    }

    // ??? ??? ??? ??? ???
    public class DailySamplingGraphData
    {
        public string FileName { get; set; }
        public List<string> Dates { get; set; }
        public List<string> SampleNumbers { get; set; }
        public DataTable Data { get; set; }
        public double? SpecValue { get; set; }
        public double? UpperLimitValue { get; set; }
        public double? LowerLimitValue { get; set; }
        public Dictionary<string, ColumnLimitSetting> ColumnLimits { get; set; } =
            new Dictionary<string, ColumnLimitSetting>(StringComparer.Ordinal);
        public Color DataColor { get; set; } = Colors.Black;
        public Color SpecColor { get; set; } = Colors.Black;
        public Color UpperColor { get; set; } = Colors.Black;
        public Color LowerColor { get; set; } = Colors.Black;
        public bool IsXAxisDate { get; set; } = true; // true: ??, false: ??
        public XAxisMode XAxisMode { get; set; } = XAxisMode.Date;
        public NoXAxisDisplayMode NoXAxisDisplayMode { get; set; } = NoXAxisDisplayMode.SampleOrderOnXAxis;
        public ValuePlotDisplayMode DisplayMode { get; set; } = ValuePlotDisplayMode.Combined;
    }
}
