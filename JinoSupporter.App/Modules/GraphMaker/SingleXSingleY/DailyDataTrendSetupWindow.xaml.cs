using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;

namespace GraphMaker;

public partial class DailyDataTrendSetupWindow : Window
{
    private sealed class ColumnRenameItem
    {
        public string OriginalName { get; init; } = string.Empty;
        public string NewName { get; set; } = string.Empty;
    }

    private readonly ObservableCollection<string> _filePaths = new();
    private readonly ObservableCollection<ColumnRenameItem> _columnItems = new();
    private readonly bool _requireFirstColumnDate;
    private readonly bool _convertWideToSingleY;
    private DataTable? _loadedTable;

    public ProcessTrendFileInfo? ResultFileInfo { get; private set; }

    public DailyDataTrendSetupWindow(
        IEnumerable<string>? initialFiles = null,
        bool requireFirstColumnDate = true,
        bool convertWideToSingleY = true)
    {
        InitializeComponent();
        _requireFirstColumnDate = requireFirstColumnDate;
        _convertWideToSingleY = convertWideToSingleY;
        InitializeDelimiterOptions();
        FileListBox.ItemsSource = _filePaths;
        ColumnRenameDataGrid.ItemsSource = _columnItems;

        if (initialFiles != null)
        {
            foreach (string path in initialFiles.Where(path => !string.IsNullOrWhiteSpace(path)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                _filePaths.Add(path);
            }
        }

        Loaded += (_, _) => RefreshPreviewIfPossible();
    }

    private void InitializeDelimiterOptions()
    {
        DelimiterComboBox.Items.Add(new System.Windows.Controls.ComboBoxItem { Content = "Tab", Tag = "\t" });
        DelimiterComboBox.Items.Add(new System.Windows.Controls.ComboBoxItem { Content = "Comma (,)", Tag = "," });
        DelimiterComboBox.Items.Add(new System.Windows.Controls.ComboBoxItem { Content = "Semicolon (;)", Tag = ";" });
        DelimiterComboBox.Items.Add(new System.Windows.Controls.ComboBoxItem { Content = "Space", Tag = " " });
        DelimiterComboBox.SelectedIndex = 0;
    }

    private void BrowseFilesButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Select data files",
            Filter = "Text/CSV files (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*",
            Multiselect = true
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        foreach (string path in dialog.FileNames)
        {
            if (_filePaths.Any(existingPath => string.Equals(existingPath, path, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            _filePaths.Add(path);
        }

        StatusTextBlock.Text = $"{_filePaths.Count:N0} file(s) selected.";
        RefreshPreviewIfPossible();
    }

    private void RemoveFileButton_Click(object sender, RoutedEventArgs e)
    {
        if (FileListBox.SelectedItem is not string selectedPath)
        {
            return;
        }

        _filePaths.Remove(selectedPath);
        StatusTextBlock.Text = $"{_filePaths.Count:N0} file(s) selected.";
        RefreshPreviewIfPossible();
    }

    private void ClearFilesButton_Click(object sender, RoutedEventArgs e)
    {
        _filePaths.Clear();
        _columnItems.Clear();
        _loadedTable = null;
        PreviewDataGrid.ItemsSource = null;
        StatusTextBlock.Text = "Cleared selected files.";
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_loadedTable == null)
            {
                LoadPreviewTable();
            }

            if (_loadedTable == null)
            {
                return;
            }

            ApplyColumnRenames(_loadedTable);
            ResultFileInfo = new ProcessTrendFileInfo
            {
                Name = string.Join(" + ", _filePaths.Select(Path.GetFileName)),
                FilePath = string.Join("|", _filePaths),
                Delimiter = GetSelectedDelimiter(),
                HeaderRowNumber = GetHeaderRowNumber(),
                FullData = _loadedTable
            };

            DialogResult = true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Setup", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    private void DelimiterComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        RefreshPreviewIfPossible();
    }

    private void HeaderRowTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        RefreshPreviewIfPossible();
    }

    private void LoadPreviewTable()
    {
        if (_filePaths.Count == 0)
        {
            throw new InvalidOperationException("Select at least one file.");
        }

        int headerRowNumber = GetHeaderRowNumber();
        string delimiter = GetSelectedDelimiter();

        DataTable mergedTable = LoadSingleTable(_filePaths[0], delimiter, headerRowNumber);
        for (int i = 1; i < _filePaths.Count; i++)
        {
            DataTable extraTable = LoadSingleTable(_filePaths[i], delimiter, headerRowNumber);
            MergeTable(mergedTable, extraTable);
        }

        if (_requireFirstColumnDate)
        {
            EnsureFirstColumnIsDate(mergedTable);
        }

        if (_convertWideToSingleY)
        {
            mergedTable = ConvertWideToSingleYTable(mergedTable);
        }

        _loadedTable = mergedTable;
        RefreshColumnItems(mergedTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName));
        PreviewDataGrid.ItemsSource = mergedTable.DefaultView;
        StatusTextBlock.Text = $"Loaded {_filePaths.Count:N0} file(s), {mergedTable.Columns.Count:N0} columns, {mergedTable.Rows.Count:N0} rows.";
    }

    private void RefreshPreviewIfPossible()
    {
        if (!IsLoaded)
        {
            return;
        }

        if (_filePaths.Count == 0)
        {
            _loadedTable = null;
            _columnItems.Clear();
            PreviewDataGrid.ItemsSource = null;
            StatusTextBlock.Text = "Select files to preview.";
            return;
        }

        try
        {
            LoadPreviewTable();
        }
        catch (Exception ex)
        {
            _loadedTable = null;
            _columnItems.Clear();
            PreviewDataGrid.ItemsSource = null;
            StatusTextBlock.Text = ex.Message;
        }
    }

    private void RefreshColumnItems(IEnumerable<string> columnNames)
    {
        Dictionary<string, string> previousNames = _columnItems.ToDictionary(item => item.OriginalName, item => item.NewName, StringComparer.OrdinalIgnoreCase);
        _columnItems.Clear();
        foreach (string columnName in columnNames)
        {
            _columnItems.Add(new ColumnRenameItem
            {
                OriginalName = columnName,
                NewName = previousNames.TryGetValue(columnName, out string? existingName) ? existingName : columnName
            });
        }
    }

    private void ApplyColumnRenames(DataTable table)
    {
        HashSet<string> usedNames = new(StringComparer.OrdinalIgnoreCase);
        foreach (ColumnRenameItem item in _columnItems)
        {
            string newName = string.IsNullOrWhiteSpace(item.NewName) ? item.OriginalName : item.NewName.Trim();
            if (!usedNames.Add(newName))
            {
                throw new InvalidOperationException($"Duplicate column name '{newName}' is not allowed.");
            }
        }

        foreach (ColumnRenameItem item in _columnItems)
        {
            string newName = string.IsNullOrWhiteSpace(item.NewName) ? item.OriginalName : item.NewName.Trim();
            if (string.Equals(item.OriginalName, newName, StringComparison.Ordinal))
            {
                continue;
            }

            table.Columns[item.OriginalName]!.ColumnName = newName;
        }
    }

    private static DataTable LoadSingleTable(string filePath, string delimiter, int headerRowNumber)
    {
        string[] lines = File.ReadAllLines(filePath);
        if (lines.Length == 0)
        {
            throw new InvalidOperationException($"File is empty: {Path.GetFileName(filePath)}");
        }

        int headerIndex = headerRowNumber - 1;
        if (headerIndex < 0 || headerIndex >= lines.Length)
        {
            throw new InvalidOperationException($"Invalid header row number for file: {Path.GetFileName(filePath)}");
        }

        string[] headerTokens = GraphMakerTableHelper.SplitLine(lines[headerIndex], delimiter);
        if (headerTokens.Length == 0)
        {
            throw new InvalidOperationException($"Cannot read header from file: {Path.GetFileName(filePath)}");
        }

        DataTable table = new();
        foreach (string header in GraphMakerTableHelper.BuildUniqueHeaders(headerTokens))
        {
            table.Columns.Add(header);
        }

        for (int i = headerIndex + 1; i < lines.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i]))
            {
                continue;
            }

            string[] values = GraphMakerTableHelper.SplitLine(lines[i], delimiter);
            if (values.Length == 0)
            {
                continue;
            }

            DataRow row = table.NewRow();
            for (int col = 0; col < Math.Min(values.Length, table.Columns.Count); col++)
            {
                row[col] = values[col];
            }

            table.Rows.Add(row);
        }

        return table;
    }

    private static void MergeTable(DataTable merged, DataTable extra)
    {
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
    }

    private string GetSelectedDelimiter()
    {
        if (DelimiterComboBox.SelectedItem is not System.Windows.Controls.ComboBoxItem selectedItem ||
            selectedItem.Tag is not string delimiter)
        {
            return "\t";
        }

        return delimiter;
    }

    private int GetHeaderRowNumber()
    {
        if (!int.TryParse(HeaderRowTextBox.Text.Trim(), out int headerRowNumber) || headerRowNumber < 1)
        {
            throw new InvalidOperationException("Header Row must be an integer greater than or equal to 1.");
        }

        return headerRowNumber;
    }

    private static void EnsureFirstColumnIsDate(DataTable table)
    {
        if (table.Columns.Count == 0)
        {
            throw new InvalidOperationException("Daily Data Trend requires at least one column.");
        }

        string firstColumnName = table.Columns[0].ColumnName;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            string text = table.Rows[rowIndex][0]?.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out _) ||
                DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out _))
            {
                continue;
            }

            throw new InvalidOperationException(
                $"Daily Data Trend requires the first column to contain date values. Invalid value '{text}' was found at row {rowIndex + 1:N0} in column '{firstColumnName}'.");
        }
    }

    private static DataTable ConvertWideToSingleYTable(DataTable wideTable)
    {
        if (wideTable.Columns.Count < 2)
        {
            throw new InvalidOperationException("SingleX(Date) - SingleY requires at least one date column and one value column.");
        }

        DataTable longTable = new();
        longTable.Columns.Add("Date");
        longTable.Columns.Add("Value");

        foreach (DataRow sourceRow in wideTable.Rows)
        {
            string dateText = sourceRow[0]?.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(dateText))
            {
                continue;
            }

            for (int columnIndex = 1; columnIndex < wideTable.Columns.Count; columnIndex++)
            {
                string valueText = sourceRow[columnIndex]?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(valueText))
                {
                    continue;
                }

                DataRow newRow = longTable.NewRow();
                newRow["Date"] = dateText;
                newRow["Value"] = valueText;
                longTable.Rows.Add(newRow);
            }
        }

        return longTable;
    }
}
