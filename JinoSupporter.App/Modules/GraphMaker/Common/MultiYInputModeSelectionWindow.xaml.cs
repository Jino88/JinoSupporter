using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;

namespace GraphMaker
{
    public enum MultiYInputMode
    {
        HeaderMultiY,
        DateSingleY,
        DateNoHeaderMultiY
    }

    public partial class MultiYInputModeSelectionWindow : Window
    {
        private sealed class ModeAnalysisResult
        {
            public bool IsSupported { get; init; }
            public string Delimiter { get; init; } = "\t";
            public string Description { get; init; } = string.Empty;
        }

        private static readonly string[] CandidateDelimiters = { "\t", ",", " " };

        private readonly string? _previewFilePath;
        private readonly List<string> _previewLines = new();
        private readonly Dictionary<MultiYInputMode, ModeAnalysisResult> _modeAnalysis = new();

        public MultiYInputMode SelectedMode { get; private set; } = MultiYInputMode.HeaderMultiY;
        public string SelectedDelimiter { get; private set; } = "\t";
        public bool SelectedHasHeader { get; private set; } = true;
        public int SelectedHeaderRowNumber { get; private set; } = 1;

        public MultiYInputModeSelectionWindow(string? previewFilePath = null)
        {
            InitializeComponent();
            _previewFilePath = previewFilePath;
            LoadPreviewLines();
            Loaded += MultiYInputModeSelectionWindow_Loaded;
        }

        private void MultiYInputModeSelectionWindow_Loaded(object sender, RoutedEventArgs e)
        {
            RefreshAll();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (!TryGetSelectedMode(out MultiYInputMode selectedMode))
            {
                MessageBox.Show("No compatible mode was detected for this file.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            SelectedMode = selectedMode;
            SelectedDelimiter = GetEffectiveDelimiter(selectedMode);
            SelectedHasHeader = selectedMode != MultiYInputMode.DateNoHeaderMultiY && HasHeaderCheckBox.IsChecked == true;
            SelectedHeaderRowNumber = GetHeaderRowNumber();
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void ModeRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            UpdateHeaderSettingsAvailability();
            RefreshPreviewOnly();
        }

        private void PreviewSettingChanged(object sender, RoutedEventArgs e)
        {
            RefreshAll();
        }

        private void PreviewSettingChanged(object sender, TextChangedEventArgs e)
        {
            RefreshAll();
        }

        private void RefreshAll()
        {
            if (!IsLoaded || PreviewPlotView == null || PreviewDescriptionTextBlock == null)
            {
                return;
            }

            AnalyzeAvailableModes();
            ApplyModeAvailability();
            UpdateHeaderSettingsAvailability();
            RefreshPreviewOnly();
        }

        private void RefreshPreviewOnly()
        {
            if (PreviewPlotView == null || PreviewDescriptionTextBlock == null || PreviewSummaryTextBlock == null || PreviewDataGrid == null)
            {
                return;
            }

            if (!TryGetSelectedMode(out MultiYInputMode selectedMode))
            {
                BuildEmptyPreview(new PlotModel());
                return;
            }

            UpdatePreviewGrid(selectedMode);

            var model = new PlotModel();
            if (_previewLines.Count == 0)
            {
                BuildEmptyPreview(model);
            }
            else if (selectedMode == MultiYInputMode.HeaderMultiY)
            {
                BuildHeaderMultiYPreview(model);
            }
            else if (selectedMode == MultiYInputMode.DateSingleY)
            {
                BuildDateSingleYPreview(model);
            }
            else
            {
                BuildDateNoHeaderMultiYPreview(model);
            }

            PreviewPlotView.Model = model;
        }

        private void LoadPreviewLines()
        {
            _previewLines.Clear();
            if (string.IsNullOrWhiteSpace(_previewFilePath) || !File.Exists(_previewFilePath))
            {
                return;
            }

            foreach (string line in File.ReadLines(_previewFilePath).Take(100))
            {
                if (!string.IsNullOrWhiteSpace(line))
                {
                    _previewLines.Add(line);
                }
            }
        }

        private void BuildEmptyPreview(PlotModel model)
        {
            model.Title = "Preview unavailable";
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, IsAxisVisible = false });
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, IsAxisVisible = false });
            PreviewSummaryTextBlock.Text = "No preview rows were loaded from the selected file.";
            PreviewDescriptionTextBlock.Text = "No readable preview data was found. Select read options and a mode to continue.";
            PreviewDataGrid.ItemsSource = null;
            PreviewPlotView.Model = model;
        }

        private void BuildHeaderMultiYPreview(PlotModel model)
        {
            model.Title = "SingleX(No) - Multi Y Preview";
            List<string[]> rows = ParseRowsForMode(MultiYInputMode.HeaderMultiY);
            int headerIndex = Math.Max(0, GetHeaderRowNumber() - 1);
            if (rows.Count <= headerIndex + 1)
            {
                PreviewDescriptionTextBlock.Text = "Preview needs one header row and at least one data row.";
                return;
            }

            List<string> headers = GraphMakerTableHelper.BuildUniqueHeaders(rows[headerIndex]);
            var xAxis = new CategoryAxis { Position = AxisPosition.Bottom, Title = "Columns" };
            var yAxis = new LinearAxis { Position = AxisPosition.Left, Title = "Value" };
            model.Axes.Add(xAxis);
            model.Axes.Add(yAxis);

            int plottedColumns = 0;
            for (int columnIndex = 0; columnIndex < headers.Count && plottedColumns < 6; columnIndex++)
            {
                var series = new ScatterSeries
                {
                    Title = headers[columnIndex],
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 2.5
                };

                foreach (string[] row in rows.Skip(headerIndex + 1).Take(40))
                {
                    if (row.Length <= columnIndex)
                    {
                        continue;
                    }

                    if (GraphMakerParsingHelper.TryParseDouble(row[columnIndex], out double value))
                    {
                        series.Points.Add(new ScatterPoint(plottedColumns, value));
                    }
                }

                if (series.Points.Count == 0)
                {
                    continue;
                }

                xAxis.Labels.Add(headers[columnIndex]);
                model.Series.Add(series);
                plottedColumns++;
            }

            PreviewDescriptionTextBlock.Text = plottedColumns == 0
                ? "No numeric Y columns were found for the current delimiter/header setting."
                : $"{Path.GetFileName(_previewFilePath)} preview: header-based multi-column distribution.";
        }

        private void BuildDateSingleYPreview(PlotModel model)
        {
            model.Title = "SingleX(Date) - SingleY Preview";
            DataTable previewTable = BuildCollapsedDateValueTable(MultiYInputMode.DateSingleY, 60);
            if (previewTable.Rows.Count == 0)
            {
                PreviewDescriptionTextBlock.Text = "No valid Date/Value rows were found for the current setting.";
                return;
            }

            var xAxis = new CategoryAxis { Position = AxisPosition.Bottom, Title = "Date" };
            var yAxis = new LinearAxis { Position = AxisPosition.Left, Title = "Value" };
            var sampleSeries = new ScatterSeries
            {
                Title = "Sample",
                MarkerType = MarkerType.Circle,
                MarkerSize = 2.5,
                MarkerFill = OxyColors.Black
            };
            var minSeries = new ScatterSeries
            {
                Title = "Min",
                MarkerType = MarkerType.Circle,
                MarkerSize = 3,
                MarkerFill = OxyColor.FromRgb(52, 120, 246)
            };
            var maxSeries = new ScatterSeries
            {
                Title = "Max",
                MarkerType = MarkerType.Circle,
                MarkerSize = 3,
                MarkerFill = OxyColor.FromRgb(238, 62, 62)
            };
            var avgSeries = new LineSeries
            {
                Title = "Avg",
                Color = OxyColor.FromRgb(46, 160, 67),
                StrokeThickness = 2,
                MarkerType = MarkerType.Circle,
                MarkerSize = 3
            };
            model.Axes.Add(xAxis);
            model.Axes.Add(yAxis);

            var grouped = previewTable.Rows.Cast<DataRow>()
                .Select(row => new
                {
                    DateText = row["Date"]?.ToString()?.Trim() ?? string.Empty,
                    ValueText = row["Value"]?.ToString()?.Trim() ?? string.Empty
                })
                .Where(item => GraphMakerParsingHelper.TryParseDate(item.DateText, out _) &&
                               GraphMakerParsingHelper.TryParseDouble(item.ValueText, out _))
                .GroupBy(item => item.DateText)
                .Take(20)
                .ToList();

            int pointIndex = 0;
            foreach (var group in grouped)
            {
                xAxis.Labels.Add(group.Key);
                List<double> values = group
                    .Select(item => GraphMakerParsingHelper.TryParseDouble(item.ValueText, out double value) ? (double?)value : null)
                    .Where(value => value.HasValue)
                    .Select(value => value!.Value)
                    .ToList();

                foreach (double value in values)
                {
                    sampleSeries.Points.Add(new ScatterPoint(pointIndex, value));
                }

                minSeries.Points.Add(new ScatterPoint(pointIndex, values.Min()));
                maxSeries.Points.Add(new ScatterPoint(pointIndex, values.Max()));
                avgSeries.Points.Add(new DataPoint(pointIndex, values.Average()));
                pointIndex++;
            }

            model.Series.Add(sampleSeries);
            model.Series.Add(minSeries);
            model.Series.Add(maxSeries);
            model.Series.Add(avgSeries);
            PreviewDescriptionTextBlock.Text = $"{Path.GetFileName(_previewFilePath)} preview: selected value columns are combined, then summarized by Date like the actual SingleY result.";
        }

        private void BuildDateNoHeaderMultiYPreview(PlotModel model)
        {
            model.Title = "SingleX(Date) - Multi Y (No Header) Preview";
            List<string[]> rows = ParseRowsForMode(MultiYInputMode.DateNoHeaderMultiY);
            if (rows.Count == 0)
            {
                PreviewDescriptionTextBlock.Text = "Preview needs at least one readable data row.";
                return;
            }

            var xAxis = new CategoryAxis { Position = AxisPosition.Bottom, Title = "Date" };
            var yAxis = new LinearAxis { Position = AxisPosition.Left, Title = "Value" };
            var series = new LineSeries { Title = "Combined Value", MarkerType = MarkerType.Circle, StrokeThickness = 2 };
            model.Axes.Add(xAxis);
            model.Axes.Add(yAxis);

            int pointIndex = 0;
            foreach (string[] row in rows.Take(30))
            {
                if (row.Length < 2 || !GraphMakerParsingHelper.TryParseDate(row[0], out _))
                {
                    continue;
                }

                for (int columnIndex = 1; columnIndex < row.Length && pointIndex < 40; columnIndex++)
                {
                    if (!GraphMakerParsingHelper.TryParseDouble(row[columnIndex], out double value))
                    {
                        continue;
                    }

                    xAxis.Labels.Add(row[0]);
                    series.Points.Add(new DataPoint(pointIndex, value));
                    pointIndex++;
                }
            }

            model.Series.Add(series);
            PreviewDescriptionTextBlock.Text = $"{Path.GetFileName(_previewFilePath)} preview: wide no-header data combined into one Date/Value stream.";
        }

        private void UpdatePreviewGrid(MultiYInputMode mode)
        {
            DataTable table = BuildPreviewTable(mode);
            PreviewDataGrid.ItemsSource = table.DefaultView;

            string fileName = string.IsNullOrWhiteSpace(_previewFilePath) ? "(No file)" : Path.GetFileName(_previewFilePath);
            string delimiterLabel = GetDelimiterLabel(GetEffectiveDelimiter(mode));
            string headerText = mode == MultiYInputMode.DateNoHeaderMultiY
                ? "No header"
                : HasHeaderCheckBox.IsChecked == true
                    ? $"Header row {GetHeaderRowNumber()}"
                    : "No header";
            PreviewSummaryTextBlock.Text = $"{fileName} | Delimiter: {delimiterLabel} | {headerText} | Grid rows: {table.Rows.Count:N0}";
        }

        private DataTable BuildPreviewTable(MultiYInputMode mode)
        {
            List<string[]> rows = ParseRowsForMode(mode);
            var table = new DataTable();
            if (rows.Count == 0)
            {
                return table;
            }

            if (mode == MultiYInputMode.DateNoHeaderMultiY)
            {
                int maxColumnCount = rows.Max(row => row.Length);
                for (int columnIndex = 0; columnIndex < maxColumnCount; columnIndex++)
                {
                    table.Columns.Add(columnIndex == 0 ? "Date" : $"Value{columnIndex}");
                }

                foreach (string[] row in rows.Take(30))
                {
                    DataRow dataRow = table.NewRow();
                    for (int columnIndex = 0; columnIndex < Math.Min(maxColumnCount, row.Length); columnIndex++)
                    {
                        dataRow[columnIndex] = row[columnIndex];
                    }

                    table.Rows.Add(dataRow);
                }

                return table;
            }

            if (mode == MultiYInputMode.DateSingleY)
            {
                return BuildCollapsedDateValueTable(mode, 30);
            }

            int headerIndex = Math.Max(0, GetHeaderRowNumber() - 1);
            if (rows.Count <= headerIndex)
            {
                return table;
            }

            List<string> headers = GraphMakerTableHelper.BuildUniqueHeaders(rows[headerIndex]);
            foreach (string header in headers)
            {
                table.Columns.Add(string.IsNullOrWhiteSpace(header) ? "Column" : header);
            }

            foreach (string[] row in rows.Skip(headerIndex + 1).Take(30))
            {
                DataRow dataRow = table.NewRow();
                for (int columnIndex = 0; columnIndex < Math.Min(headers.Count, row.Length); columnIndex++)
                {
                    dataRow[columnIndex] = row[columnIndex];
                }

                table.Rows.Add(dataRow);
            }

            return table;
        }

        private DataTable BuildCollapsedDateValueTable(MultiYInputMode mode, int maxRows)
        {
            List<string[]> rows = ParseRowsForMode(mode);
            var table = new DataTable();
            table.Columns.Add("Date");
            table.Columns.Add("Value");

            int headerIndex = Math.Max(0, GetHeaderRowNumber() - 1);
            if (rows.Count <= headerIndex + 1)
            {
                return table;
            }

            int addedRows = 0;
            foreach (string[] row in rows.Skip(headerIndex + 1))
            {
                if (row.Length < 2)
                {
                    continue;
                }

                string dateText = row[0]?.Trim() ?? string.Empty;
                if (!GraphMakerParsingHelper.TryParseDate(dateText, out _))
                {
                    continue;
                }

                for (int columnIndex = 1; columnIndex < row.Length; columnIndex++)
                {
                    string valueText = row[columnIndex]?.Trim() ?? string.Empty;
                    if (!GraphMakerParsingHelper.TryParseDouble(valueText, out _))
                    {
                        continue;
                    }

                    DataRow dataRow = table.NewRow();
                    dataRow["Date"] = dateText;
                    dataRow["Value"] = valueText;
                    table.Rows.Add(dataRow);
                    addedRows++;

                    if (addedRows >= maxRows)
                    {
                        return table;
                    }
                }
            }

            return table;
        }

        private List<string[]> ParseDelimitedRows(string delimiter)
        {
            return _previewLines
                .Select(line => GraphMakerTableHelper.SplitLine(line, delimiter).Select(token => token.Trim()).ToArray())
                .Where(tokens => tokens.Length > 0)
                .ToList();
        }

        private List<string[]> ParseRowsForMode(MultiYInputMode mode)
        {
            return ParseDelimitedRows(GetEffectiveDelimiter(mode));
        }

        private void AnalyzeAvailableModes()
        {
            _modeAnalysis.Clear();
            _modeAnalysis[MultiYInputMode.HeaderMultiY] = AnalyzeHeaderMultiY();
            _modeAnalysis[MultiYInputMode.DateSingleY] = AnalyzeDateSingleY();
            _modeAnalysis[MultiYInputMode.DateNoHeaderMultiY] = AnalyzeDateNoHeaderMultiY();
        }

        private void ApplyModeAvailability()
        {
            ApplyModeAvailability(HeaderModeRadioButton, HeaderModeDescriptionTextBlock, MultiYInputMode.HeaderMultiY,
                "Use when the file has a normal header row and multiple Y columns.");
            ApplyModeAvailability(DateSingleYModeRadioButton, DateSingleYModeDescriptionTextBlock, MultiYInputMode.DateSingleY,
                "Use when the first column is Date and one Y column should be shown as a trend graph.");
            ApplyModeAvailability(DateNoHeaderModeRadioButton, DateNoHeaderModeDescriptionTextBlock, MultiYInputMode.DateNoHeaderMultiY,
                "Use when the first column is Date and the file does not have a normal header row.");

            if (TryGetFirstSupportedMode(out MultiYInputMode firstSupportedMode))
            {
                if (!IsModeCheckedAndEnabled())
                {
                    SetSelectedMode(firstSupportedMode);
                }
            }
            else
            {
                HeaderModeRadioButton.IsChecked = false;
                DateSingleYModeRadioButton.IsChecked = false;
                DateNoHeaderModeRadioButton.IsChecked = false;
            }
        }

        private void ApplyModeAvailability(RadioButton radioButton, TextBlock descriptionTextBlock, MultiYInputMode mode, string baseDescription)
        {
            ModeAnalysisResult result = _modeAnalysis[mode];
            radioButton.IsEnabled = result.IsSupported;
            descriptionTextBlock.Text = result.IsSupported
                ? $"{baseDescription} Delimiter: {GetDelimiterLabel(result.Delimiter)}."
                : $"{baseDescription} Not available: {result.Description}";
            descriptionTextBlock.Opacity = result.IsSupported ? 1.0 : 0.65;
        }

        private ModeAnalysisResult AnalyzeHeaderMultiY()
        {
            if (HasHeaderCheckBox.IsChecked != true)
            {
                return new ModeAnalysisResult { IsSupported = false, Description = "Header mode requires header rows." };
            }

            foreach (string delimiter in CandidateDelimiters)
            {
                if (!IsDelimiterAllowed(delimiter))
                {
                    continue;
                }

                List<string[]> rows = ParseDelimitedRows(delimiter);
                int headerIndex = Math.Max(0, GetHeaderRowNumber() - 1);
                if (rows.Count <= headerIndex + 1 || rows[headerIndex].Length < 2)
                {
                    continue;
                }

                bool hasNumericColumn = false;
                for (int columnIndex = 0; columnIndex < rows[headerIndex].Length; columnIndex++)
                {
                    foreach (string[] row in rows.Skip(headerIndex + 1))
                    {
                        if (row.Length > columnIndex && GraphMakerParsingHelper.TryParseDouble(row[columnIndex], out _))
                        {
                            hasNumericColumn = true;
                            break;
                        }
                    }

                    if (hasNumericColumn)
                    {
                        break;
                    }
                }

                if (hasNumericColumn)
                {
                    return new ModeAnalysisResult
                    {
                        IsSupported = true,
                        Delimiter = delimiter,
                        Description = "Header row and numeric columns detected."
                    };
                }
            }

            return new ModeAnalysisResult { IsSupported = false, Description = "No header-based numeric columns detected." };
        }

        private ModeAnalysisResult AnalyzeDateSingleY()
        {
            if (HasHeaderCheckBox.IsChecked != true)
            {
                return new ModeAnalysisResult { IsSupported = false, Description = "This mode requires header rows." };
            }

            foreach (string delimiter in CandidateDelimiters)
            {
                if (!IsDelimiterAllowed(delimiter))
                {
                    continue;
                }

                List<string[]> rows = ParseDelimitedRows(delimiter);
                int headerIndex = Math.Max(0, GetHeaderRowNumber() - 1);
                if (rows.Count <= headerIndex + 1)
                {
                    continue;
                }

                int yColumnIndex = FindFirstNumericColumn(rows, headerIndex + 1, 1);
                if (yColumnIndex < 0)
                {
                    continue;
                }

                int validRows = 0;
                foreach (string[] row in rows.Skip(headerIndex + 1))
                {
                    if (row.Length > yColumnIndex &&
                        GraphMakerParsingHelper.TryParseDate(row[0], out _) &&
                        GraphMakerParsingHelper.TryParseDouble(row[yColumnIndex], out _))
                    {
                        validRows++;
                    }
                }

                if (validRows >= 3)
                {
                    return new ModeAnalysisResult
                    {
                        IsSupported = true,
                        Delimiter = delimiter,
                        Description = "Date first column and one numeric Y column detected."
                    };
                }
            }

            return new ModeAnalysisResult { IsSupported = false, Description = "No Date + numeric Y pattern detected." };
        }

        private ModeAnalysisResult AnalyzeDateNoHeaderMultiY()
        {
            if (HasHeaderCheckBox.IsChecked == true)
            {
                return new ModeAnalysisResult { IsSupported = false, Description = "Disable header row to use no-header mode." };
            }

            foreach (string delimiter in CandidateDelimiters)
            {
                if (!IsDelimiterAllowed(delimiter))
                {
                    continue;
                }

                List<string[]> rows = ParseDelimitedRows(delimiter);
                int validRows = 0;
                foreach (string[] row in rows)
                {
                    if (row.Length < 2 || !GraphMakerParsingHelper.TryParseDate(row[0], out _))
                    {
                        continue;
                    }

                    if (row.Skip(1).Any(value => GraphMakerParsingHelper.TryParseDouble(value, out _)))
                    {
                        validRows++;
                    }
                }

                if (validRows >= 3)
                {
                    return new ModeAnalysisResult
                    {
                        IsSupported = true,
                        Delimiter = delimiter,
                        Description = "Date/value wide rows detected."
                    };
                }
            }

            return new ModeAnalysisResult { IsSupported = false, Description = "No-header Date/value pattern not detected." };
        }

        private bool TryGetFirstSupportedMode(out MultiYInputMode mode)
        {
            foreach (MultiYInputMode candidate in new[]
                     {
                         MultiYInputMode.HeaderMultiY,
                         MultiYInputMode.DateSingleY,
                         MultiYInputMode.DateNoHeaderMultiY
                     })
            {
                if (_modeAnalysis.TryGetValue(candidate, out ModeAnalysisResult? result) && result.IsSupported)
                {
                    mode = candidate;
                    return true;
                }
            }

            mode = MultiYInputMode.HeaderMultiY;
            return false;
        }

        private bool TryGetSelectedMode(out MultiYInputMode mode)
        {
            if (HeaderModeRadioButton.IsChecked == true && HeaderModeRadioButton.IsEnabled)
            {
                mode = MultiYInputMode.HeaderMultiY;
                return true;
            }

            if (DateSingleYModeRadioButton.IsChecked == true && DateSingleYModeRadioButton.IsEnabled)
            {
                mode = MultiYInputMode.DateSingleY;
                return true;
            }

            if (DateNoHeaderModeRadioButton.IsChecked == true && DateNoHeaderModeRadioButton.IsEnabled)
            {
                mode = MultiYInputMode.DateNoHeaderMultiY;
                return true;
            }

            mode = MultiYInputMode.HeaderMultiY;
            return false;
        }

        private bool IsModeCheckedAndEnabled()
        {
            return (HeaderModeRadioButton.IsChecked == true && HeaderModeRadioButton.IsEnabled) ||
                   (DateSingleYModeRadioButton.IsChecked == true && DateSingleYModeRadioButton.IsEnabled) ||
                   (DateNoHeaderModeRadioButton.IsChecked == true && DateNoHeaderModeRadioButton.IsEnabled);
        }

        private void SetSelectedMode(MultiYInputMode mode)
        {
            HeaderModeRadioButton.IsChecked = mode == MultiYInputMode.HeaderMultiY;
            DateSingleYModeRadioButton.IsChecked = mode == MultiYInputMode.DateSingleY;
            DateNoHeaderModeRadioButton.IsChecked = mode == MultiYInputMode.DateNoHeaderMultiY;
        }

        private void UpdateHeaderSettingsAvailability()
        {
            if (HasHeaderCheckBox == null || HeaderRowTextBox == null ||
                DateNoHeaderModeRadioButton == null)
            {
                return;
            }

            bool isNoHeaderMode = DateNoHeaderModeRadioButton.IsChecked == true;
            HasHeaderCheckBox.IsEnabled = !isNoHeaderMode;
            HeaderRowTextBox.IsEnabled = !isNoHeaderMode && HasHeaderCheckBox.IsChecked == true;
        }

        private string GetEffectiveDelimiter(MultiYInputMode mode)
        {
            string selected = GetSelectedDelimiterPreference();
            if (selected != "AUTO")
            {
                return selected;
            }

            return _modeAnalysis.TryGetValue(mode, out ModeAnalysisResult? result) ? result.Delimiter : "\t";
        }

        private string GetSelectedDelimiterPreference()
        {
            if (DelimiterComboBox.SelectedItem is ComboBoxItem item && item.Tag is string tag)
            {
                return tag;
            }

            return "AUTO";
        }

        private bool IsDelimiterAllowed(string delimiter)
        {
            string selected = GetSelectedDelimiterPreference();
            return selected == "AUTO" || string.Equals(selected, delimiter, StringComparison.Ordinal);
        }

        private int GetHeaderRowNumber()
        {
            return int.TryParse(HeaderRowTextBox.Text, out int headerRow) && headerRow > 0 ? headerRow : 1;
        }

        private static string GetDelimiterLabel(string delimiter)
        {
            return delimiter switch
            {
                "\t" => "Tab",
                "," => "Comma",
                " " => "Space",
                _ => delimiter
            };
        }

        private static int FindFirstNumericColumn(IReadOnlyList<string[]> rows, int startRowIndex, int startColumnIndex)
        {
            int maxColumns = rows.Max(row => row.Length);
            for (int columnIndex = startColumnIndex; columnIndex < maxColumns; columnIndex++)
            {
                for (int rowIndex = startRowIndex; rowIndex < rows.Count; rowIndex++)
                {
                    if (rows[rowIndex].Length > columnIndex &&
                        GraphMakerParsingHelper.TryParseDouble(rows[rowIndex][columnIndex], out _))
                    {
                        return columnIndex;
                    }
                }
            }

            return -1;
        }
    }
}
