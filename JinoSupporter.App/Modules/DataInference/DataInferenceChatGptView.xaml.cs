using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using WorkbenchHost.Infrastructure;

namespace JinoSupporter.App.Modules.DataInference;

public partial class DataInferenceChatGptView : UserControl
{
    public event Action? WebModuleSnapshotChanged;

    private static readonly HttpClient HttpClient = new() { Timeout = TimeSpan.FromSeconds(120) };
    private static readonly JsonSerializerOptions JsonOpts = new() { PropertyNameCaseInsensitive = true };

    // GPT-4o mini pricing placeholder (per 1M tokens, USD)
    private const double PriceInputPer1M  = 0.15;
    private const double PriceOutputPer1M = 0.60;

    private readonly DataInferenceRepository _repository = new();
    private readonly List<PendingTable> _pendingTables = [];
    private readonly List<PendingImage> _pendingImages = [];
    private bool _isBusy;
    private int _sessionInputTokens;
    private int _sessionOutputTokens;

    public DataInferenceChatGptView()
    {
        InitializeComponent();
        DbPathTextBlock.Text = DataInferenceRepository.DatabasePath;
        RefreshDbInfo();
        RenderTablesPanel();
        IsVisibleChanged += (_, e) =>
        {
            if ((bool)e.NewValue)
                DbPathTextBlock.Text = DataInferenceRepository.DatabasePath;
        };
    }

    // -- Shell interface --

    public object GetWebModuleSnapshot() => new
    {
        moduleType = "DataInferenceChatGpt",
        hasData = !string.IsNullOrWhiteSpace(DataInputTextBox.Text),
        isBusy = _isBusy,
        pendingTables = _pendingTables.Count,
        pendingImages = _pendingImages.Count,
        totalRecords = _repository.GetTotalRowCount(),
        status = StatusTextBlock.Text
    };

    public object InvokeWebModuleAction(string action)
    {
        switch (action)
        {
            case "analyze": _ = AnalyzeAsync(); break;
            case "clear":   ClearAll();          break;
        }

        return GetWebModuleSnapshot();
    }

    // -- Ctrl+V: if TextBox focused => text paste handled by TextBox; else => image paste --

    protected override void OnPreviewKeyDown(KeyEventArgs e)
    {
        if (e.Key == Key.V && Keyboard.Modifiers == ModifierKeys.Control
            && Clipboard.ContainsImage()
            && !DataInputTextBox.IsKeyboardFocused)
        {
            PasteImageFromClipboard();
            e.Handled = true;
        }

        base.OnPreviewKeyDown(e);
    }

    // -- Event handlers --

    private async void AnalyzeButton_Click(object sender, RoutedEventArgs e) => await AnalyzeAsync();

    private void SaveDbButton_Click(object sender, RoutedEventArgs e) => SavePendingToDb();

    private void ClearButton_Click(object sender, RoutedEventArgs e) => ClearAll();

    private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (e.OriginalSource is TabControl && MainTabControl.SelectedItem == DbDataTab)
            RefreshDbTab();
    }

    private void DatasetDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (DatasetDataGrid.SelectedItem is not DatasetItem ds)
        {
            TablesDataGrid.ItemsSource = null;
            RowDetailDataGrid.ItemsSource = null;
            RowDetailDataGrid.Columns.Clear();
            DbImageBorder.Visibility = Visibility.Collapsed;
            return;
        }

        LoadDbTables(ds.Name);
        RefreshDbImageStrip(ds.Name);
    }

    private void RefreshDbButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshDbTab();
        RefreshDbInfo();
    }

    private void DeleteDbRowButton_Click(object sender, RoutedEventArgs e)
    {
        if (TablesDataGrid.SelectedItem is not DataTableInfo table)
        {
            SetStatus("Select a table to delete.");
            return;
        }

        if (MessageBox.Show(
                $"Delete the selected table?\nTable: {table.TableName}",
                "Confirm Delete",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            return;

        _repository.DeleteTable(table.Id);
        RefreshDbTab();
        RefreshDbInfo();
        SetStatus("Table deleted.");
    }

    // -- Core logic --

    private async Task AnalyzeAsync()
    {
        if (_isBusy) return;

        string rawData = NormalizeMergedCellsFillDown(DataInputTextBox.Text).Trim();

        if (string.IsNullOrWhiteSpace(rawData))
        {
            SetStatus("Paste table data into the left input area first.");
            return;
        }

        string? apiKey = WorkbenchSettingsStore.TryGetOpenAiApiKey();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            SetStatus("OpenAI API key is missing. Go to Settings > OpenAI API Key.");
            return;
        }

        await RunBusyAsync(async ct =>
        {
            int tableNo = _pendingTables.Count + 1;
            SetStatus($"Organizing table {tableNo} via ChatGPT...");

            (List<ColumnDef> columns, List<Dictionary<string, string>> rows, int inTok, int outTok)
                = await CallOpenAiAsync(apiKey, rawData, ct);

            if (rows.Count == 0)
            {
                SetStatus("No data extracted. Check the input format.");
                return;
            }

            AddPendingTable(tableNo, columns, rows);

            _sessionInputTokens  += inTok;
            _sessionOutputTokens += outTok;
            int total = inTok + outTok;
            UpdateTokenDisplay();
            SetStatus($"Table {tableNo} done - {rows.Count:N0} rows  |  " +
                      $"Tokens: in {inTok:N0} + out {outTok:N0} = {total:N0}");
        });
    }

    private void AddPendingTable(int tableNo, List<ColumnDef> columns, List<Dictionary<string, string>> rows)
    {
        var table = new PendingTable
        {
            TableIndex = tableNo,
            TableName  = $"Table {tableNo}",
            Columns    = columns,
            Rows       = rows
        };

        _pendingTables.Add(table);
        DataInputTextBox.Clear();
        RenderTablesPanel();
        UpdateSaveButton();
    }

    // Parse tab-separated text: first row = headers, rest = data rows
    private static (List<ColumnDef> Columns, List<Dictionary<string, string>> Rows)
        ParseTabSeparated(string text)
    {
        string[] lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length < 2) return ([], []);

        // Build columns from first line — deduplicate field names
        string[] headers = lines[0].Split('\t');
        var seenFields = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var columns = new List<ColumnDef>();
        for (int i = 0; i < headers.Length; i++)
        {
            string label = headers[i].Trim();
            string baseField = string.IsNullOrWhiteSpace(label)
                ? $"col{i}"
                : System.Text.RegularExpressions.Regex.Replace(label, @"[^a-zA-Z0-9가-힣]", "_");

            string field;
            if (!seenFields.ContainsKey(baseField))
            {
                seenFields[baseField] = 1;
                field = baseField;
            }
            else
            {
                seenFields[baseField]++;
                field = $"{baseField}_{seenFields[baseField]}";
            }

            columns.Add(new ColumnDef { Field = field, Label = label });
        }

        // Build rows from remaining lines
        var rows = new List<Dictionary<string, string>>();
        for (int li = 1; li < lines.Length; li++)
        {
            string line = lines[li].TrimEnd('\r');
            if (string.IsNullOrWhiteSpace(line)) continue;

            string[] cells = line.Split('\t');
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int ci = 0; ci < columns.Count; ci++)
                dict[columns[ci].Field] = ci < cells.Length ? cells[ci].Trim() : string.Empty;
            rows.Add(dict);
        }

        return (columns, rows);
    }

    private static string NormalizeMergedCellsFillDown(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        string[] rawLines = text.Replace("\r\n", "\n").Split('\n');
        List<string[]> rows = [];
        int maxColumns = 0;

        foreach (string rawLine in rawLines)
        {
            string line = rawLine.TrimEnd('\r');
            if (string.IsNullOrEmpty(line))
            {
                rows.Add([]);
                continue;
            }

            string[] cells = line.Split('\t');
            rows.Add(cells);
            if (cells.Length > maxColumns)
                maxColumns = cells.Length;
        }

        string[] previous = new string[maxColumns];

        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            string[] row = rows[rowIndex];
            if (row.Length == 0)
                continue;

            if (row.Length < maxColumns)
                Array.Resize(ref row, maxColumns);

            for (int columnIndex = 0; columnIndex < maxColumns; columnIndex++)
            {
                string current = row[columnIndex]?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(current) && !string.IsNullOrWhiteSpace(previous[columnIndex]))
                    row[columnIndex] = previous[columnIndex];
                else
                    row[columnIndex] = current;
            }

            for (int columnIndex = 0; columnIndex < maxColumns; columnIndex++)
            {
                string current = row[columnIndex]?.Trim() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(current))
                    previous[columnIndex] = current;
            }

            rows[rowIndex] = row;
        }

        return string.Join("\n", rows.Select(row =>
        {
            if (row.Length == 0)
                return string.Empty;

            int lastUsed = row.Length - 1;
            while (lastUsed >= 0 && string.IsNullOrEmpty(row[lastUsed]))
                lastUsed--;

            return lastUsed >= 0
                ? string.Join('\t', row.Take(lastUsed + 1))
                : string.Empty;
        }));
    }

    private void SavePendingToDb()
    {
        string datasetName = DatasetNameTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(datasetName))
        {
            SetStatus("Enter the dataset name (Excel filename) first.");
            DatasetNameTextBox.Focus();
            return;
        }

        if (_pendingTables.Count == 0 && _pendingImages.Count == 0)
        {
            SetStatus("Nothing to save.");
            return;
        }

        try
        {
            int savedRec = 0;
            foreach (PendingTable table in _pendingTables)
            {
                if (table.Rows.Count == 0) continue;
                _repository.SaveTable(datasetName, table.TableName, table.Columns, table.Rows);
                savedRec += table.Rows.Count;
            }

            foreach (PendingImage pi in _pendingImages)
                _repository.SaveImage(datasetName, pi.FileName, pi.Data);

            int savedImg = _pendingImages.Count;
            _pendingTables.Clear();
            _pendingImages.Clear();
            RenderTablesPanel();
            RefreshPendingImageStrip();
            UpdateSaveButton();
            RefreshDbInfo();
            if (MainTabControl.SelectedItem == DbDataTab) RefreshDbTab();

            SetStatus($"Saved '{datasetName}' - {savedRec:N0} records, {savedImg:N0} images.");
        }
        catch (Exception ex)
        {
            SetStatus($"DB save error: {ex.Message}");
        }
    }

    private void ClearAll()
    {
        DataInputTextBox.Clear();
        _pendingTables.Clear();
        _pendingImages.Clear();
        RenderTablesPanel();
        RefreshPendingImageStrip();
        UpdateSaveButton();
        SetStatus("Cleared.");
    }

    // -- Tables panel rendering --

    private void RenderTablesPanel()
    {
        TablesPanel.Children.Clear();

        if (_pendingTables.Count == 0)
        {
            TablesPanel.Children.Add(new Border
            {
                Padding = new Thickness(30, 50, 30, 50),
                Background = Brushes.White,
                CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(Color.FromRgb(211, 220, 234)),
                BorderThickness = new Thickness(1),
                Child = new StackPanel { HorizontalAlignment = HorizontalAlignment.Center, Children =
                {
                    new TextBlock { Text = "No tables yet",
                        FontSize = 14, FontWeight = FontWeights.SemiBold,
                        Foreground = new SolidColorBrush(Color.FromRgb(34, 48, 74)),
                        HorizontalAlignment = HorizontalAlignment.Center },
                    new TextBlock { Text = "Paste table data on the left and click Organize.",
                        FontSize = 11, Foreground = new SolidColorBrush(Color.FromRgb(97, 112, 132)),
                        HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 6, 0, 0) }
                }}
            });
            ResultGridHeaderText.Text = "Result";
            return;
        }

        foreach (PendingTable table in _pendingTables)
            TablesPanel.Children.Add(BuildTableCard(table));

        int totalRows = _pendingTables.Sum(t => t.Rows.Count);
        ResultGridHeaderText.Text = $"{_pendingTables.Count} table(s) / {totalRows:N0} rows total";
    }

    private Border BuildTableCard(PendingTable table)
    {
        DockPanel header = new() { Margin = new Thickness(0, 0, 0, 8) };

        Button deleteBtn = new()
        {
            Content = "Delete", FontSize = 11, Padding = new Thickness(10, 3, 10, 3),
            Background = new SolidColorBrush(Color.FromRgb(239, 68, 68)),
            Foreground = Brushes.White, BorderThickness = new Thickness(0),
            Cursor = Cursors.Hand
        };
        deleteBtn.Click += (_, _) => { _pendingTables.Remove(table); RenderTablesPanel(); UpdateSaveButton(); };
        DockPanel.SetDock(deleteBtn, Dock.Right);
        header.Children.Add(deleteBtn);

        TextBlock countLabel = new()
        {
            Text = $" - {table.Rows.Count:N0} rows",
            FontSize = 12, Foreground = new SolidColorBrush(Color.FromRgb(97, 112, 132)),
            VerticalAlignment = VerticalAlignment.Center
        };
        DockPanel.SetDock(countLabel, Dock.Right);
        header.Children.Add(countLabel);

        Border nameBox = new()
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(211, 220, 234)),
            BorderThickness = new Thickness(0, 0, 0, 1),
            VerticalAlignment = VerticalAlignment.Center
        };
        TextBox nameTb = new()
        {
            Text = table.TableName, FontSize = 13, FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(34, 48, 74)),
            BorderThickness = new Thickness(0), Background = Brushes.Transparent,
            Padding = new Thickness(0, 0, 4, 0), MinWidth = 80
        };
        nameTb.TextChanged += (_, _) => table.TableName = nameTb.Text;
        nameBox.Child = nameTb;
        header.Children.Add(nameBox);

        DataGrid dg = BuildDynamicDataGrid(table.Columns, table.Rows, maxHeight: 300);

        StackPanel content = new();
        content.Children.Add(header);
        content.Children.Add(new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(211, 220, 234)),
            BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6),
            Child = dg
        });

        return new Border
        {
            Margin = new Thickness(0, 0, 0, 12),
            Padding = new Thickness(12),
            Background = Brushes.White,
            CornerRadius = new CornerRadius(8),
            BorderBrush = new SolidColorBrush(Color.FromRgb(211, 220, 234)),
            BorderThickness = new Thickness(1),
            Child = content
        };
    }

    private DataGrid BuildDynamicDataGrid(List<ColumnDef> columns,
                                          List<Dictionary<string, string>> rows,
                                          double maxHeight = double.NaN)
    {
        var dt = new System.Data.DataTable();
        foreach (ColumnDef col in columns)
            dt.Columns.Add(col.Field, typeof(string));

        foreach (Dictionary<string, string> row in rows)
        {
            DataRow dr = dt.NewRow();
            foreach (ColumnDef col in columns)
                dr[col.Field] = row.TryGetValue(col.Field, out string? v) ? v : string.Empty;
            dt.Rows.Add(dr);
        }

        var dg = new DataGrid
        {
            AutoGenerateColumns = false,
            CanUserAddRows = false, CanUserDeleteRows = false,
            SelectionMode = DataGridSelectionMode.Single,
            GridLinesVisibility = DataGridGridLinesVisibility.Horizontal,
            HorizontalGridLinesBrush = new SolidColorBrush(Color.FromRgb(238, 242, 247)),
            AlternatingRowBackground = new SolidColorBrush(Color.FromRgb(248, 250, 253)),
            RowBackground = Brushes.White,
            BorderThickness = new Thickness(0),
            FontSize = 12, RowHeight = 28,
            ItemsSource = dt.DefaultView
        };

        if (!double.IsNaN(maxHeight)) dg.MaxHeight = maxHeight;
        if (TryFindResource("DgHeaderStyle") is Style hs) dg.ColumnHeaderStyle = hs;
        if (TryFindResource("DgCellStyle")   is Style cs) dg.CellStyle = cs;

        foreach (ColumnDef col in columns)
        {
            dg.Columns.Add(new DataGridTextColumn
            {
                Header  = string.IsNullOrWhiteSpace(col.Label) ? col.Field : col.Label,
                Binding = new Binding(col.Field),
                Width   = DataGridLength.Auto
            });
        }

        return dg;
    }

    private void UpdateSaveButton()
    {
        SaveDbButton.IsEnabled = _pendingTables.Count > 0 || _pendingImages.Count > 0;
    }

    private void UpdateTokenDisplay()
    {
        double cost = (_sessionInputTokens / 1_000_000.0 * PriceInputPer1M)
                    + (_sessionOutputTokens / 1_000_000.0 * PriceOutputPer1M);
        TokenCostTextBlock.Text = $"${cost:F4}";
        TokenUsageTextBlock.Text = $"{_sessionInputTokens + _sessionOutputTokens:N0} tokens";
    }

    // -- DB tab --

    private void RefreshDbTab()
    {
        LoadDbDatasets();
    }

    private void LoadDbDatasets()
    {
        List<(string Name, int TableCount, int ImageCount)> summary = _repository.GetDatasetSummary();
        var items = summary
            .Select(s => new DatasetItem { Name = s.Name, TableCount = s.TableCount, ImageCount = s.ImageCount })
            .ToList();
        DatasetDataGrid.ItemsSource = items;
        TablesDataGrid.ItemsSource = null;
        RowDetailDataGrid.ItemsSource = null;
        RowDetailDataGrid.Columns.Clear();
        DbImageBorder.Visibility = Visibility.Collapsed;
        DbStatusTextBlock.Text = $"{items.Count} dataset(s)";
    }

    private void LoadDbTables(string datasetName)
    {
        List<DataTableInfo> tables = _repository.GetTables(datasetName);
        TablesDataGrid.ItemsSource = tables;
        RowDetailDataGrid.ItemsSource = null;
        RowDetailDataGrid.Columns.Clear();
        DbStatusTextBlock.Text = $"{tables.Count:N0} table(s) in '{datasetName}'";
    }

    private void RefreshDbInfo() =>
        DbCountTextBlock.Text = $"{_repository.GetTotalRowCount():N0} records";

    private void RefreshDbImageStrip(string datasetName)
    {
        DbImagePanel.Children.Clear();
        List<DatasetImageRow> images = _repository.GetImages(datasetName);
        DbImageBorder.Visibility = images.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        foreach (DatasetImageRow img in images) DbImagePanel.Children.Add(BuildDbImageThumbnail(img));
    }

    // -- TextArea / Image handlers --

    private void TextAreaBorder_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        DataInputTextBox.Focus();
        e.Handled = true;
    }

    private void TablesDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (TablesDataGrid.SelectedItem is not DataTableInfo table)
        {
            RowDetailDataGrid.ItemsSource = null;
            RowDetailDataGrid.Columns.Clear();
            return;
        }

        List<(long Id, Dictionary<string, string> Data)> rows = _repository.GetTableRows(table.Id);
        BuildRowDetailGrid(table, rows);
        DbStatusTextBlock.Text = $"{table.TableName} - {rows.Count:N0} rows";
    }

    private void BuildRowDetailGrid(DataTableInfo table, List<(long Id, Dictionary<string, string> Data)> rows)
    {
        RowDetailDataGrid.Columns.Clear();

        List<Dictionary<string, string>> normalizedRows = rows.Select(row => row.Data).ToList();
        foreach (ColumnDef column in table.Columns)
        {
            RowDetailDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header  = string.IsNullOrWhiteSpace(column.Label) ? column.Field : column.Label,
                Binding = new Binding($"[{column.Field}]")
            });
        }

        RowDetailDataGrid.ItemsSource = normalizedRows;
    }

    private void PasteImageFromClipboard()
    {
        BitmapSource? bmp = Clipboard.GetImage();
        if (bmp is null) return;

        byte[] data = BitmapSourceToBytes(bmp);
        string fileName = $"capture_{DateTime.Now:yyyyMMdd_HHmmss}.png";
        _pendingImages.Add(new PendingImage { FileName = fileName, Data = data });
        RefreshPendingImageStrip();
        UpdateSaveButton();
        SetStatus($"Image attached: {fileName}  ({_pendingImages.Count} pending)");
    }

    private void RefreshPendingImageStrip()
    {
        PendingImagePanel.Children.Clear();
        foreach (PendingImage pi in _pendingImages) PendingImagePanel.Children.Add(BuildPendingThumbnail(pi));
    }

    private Border BuildPendingThumbnail(PendingImage pi)
    {
        BitmapImage bmp = BytesToBitmap(pi.Data);
        Grid grid = new() { Width = 64, Height = 64 };
        Image img = new() { Source = bmp, Stretch = Stretch.UniformToFill, Width = 64, Height = 64 };
        RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality);

        Button btn = new()
        {
            Content = "x", Width = 18, Height = 18, FontSize = 10,
            Padding = new Thickness(0),
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Top,
            Background = new SolidColorBrush(Color.FromArgb(210, 220, 50, 50)),
            Foreground = Brushes.White, BorderThickness = new Thickness(0), Cursor = Cursors.Hand
        };
        btn.Click += (_, _) => { _pendingImages.Remove(pi); RefreshPendingImageStrip(); UpdateSaveButton(); };

        grid.Children.Add(img);
        grid.Children.Add(btn);

        return new Border
        {
            Margin = new Thickness(0, 0, 6, 0), CornerRadius = new CornerRadius(4),
            BorderBrush = new SolidColorBrush(Color.FromRgb(211, 220, 234)),
            BorderThickness = new Thickness(1), Child = grid,
            Cursor = Cursors.Hand, ToolTip = pi.FileName
        };
    }

    private Border BuildDbImageThumbnail(DatasetImageRow row)
    {
        BitmapImage bmp = BytesToBitmap(row.ImageData);
        Grid grid = new() { Width = 80, Height = 80 };
        Image img = new() { Source = bmp, Stretch = Stretch.UniformToFill, Width = 80, Height = 80 };
        RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.HighQuality);
        grid.Children.Add(img);

        Border outer = new()
        {
            Margin = new Thickness(0, 0, 6, 0), CornerRadius = new CornerRadius(4),
            BorderBrush = new SolidColorBrush(Color.FromRgb(211, 220, 234)),
            BorderThickness = new Thickness(1), Child = grid,
            Cursor = Cursors.Hand, ToolTip = row.FileName
        };
        outer.MouseLeftButtonUp += (_, _) => OpenImageViewer(bmp, row.FileName);
        return outer;
    }

    private static void OpenImageViewer(BitmapImage bmp, string title)
    {
        new Window
        {
            Title = title,
            Width  = Math.Min(bmp.PixelWidth  + 40, SystemParameters.PrimaryScreenWidth  * 0.9),
            Height = Math.Min(bmp.PixelHeight + 60, SystemParameters.PrimaryScreenHeight * 0.9),
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            Content = new ScrollViewer
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
                Content = new Image { Source = bmp, Stretch = Stretch.None }
            }
        }.Show();
    }

    // -- OpenAI API --

    private static async Task<(List<ColumnDef> Columns, List<Dictionary<string, string>> Rows,
                                int InputTokens, int OutputTokens)>
        CallOpenAiAsync(string apiKey, string rawData, CancellationToken ct)
    {
        const string prompt =
            "The following is tab-separated text copied from Excel.\n\n" +
            "You must preserve the original data exactly.\n" +
            "Do not infer, summarize, normalize, translate, calculate, or replace values.\n" +
            "Copy visible cell text exactly as it appears after fill-down for merged cells.\n\n" +
            "Rules:\n" +
            "1. If a cell is blank because Excel merged cells vertically, fill it with the value from the same column in the row above.\n" +
            "2. Detect header rows, including grouped multi-row headers.\n" +
            "3. Build output columns by combining grouped headers with subheaders when needed.\n" +
            "4. Preserve row order exactly.\n" +
            "5. Do not change any numeric value.\n" +
            "6. Do not turn counts into percentages or percentages into counts.\n" +
            "7. Do not merge two rows into one or split one row into two.\n" +
            "8. Exclude only pure decorative/header rows and blank rows. Keep TOTAL rows if they exist in the source data area.\n" +
            "9. Output JSON only.\n\n" +
            "Return format (JSON object only):\n" +
            "{\n" +
            "  \"columns\": [{\"field\": \"camelCaseIdentifier\", \"label\": \"OriginalHeaderName\"}],\n" +
            "  \"rows\": [{\"camelCaseIdentifier\": \"value\", ...}]\n" +
            "}\n\n" +
            "Additional requirements:\n" +
            "- field must be English camelCase\n" +
            "- label must keep the original header meaning as-is\n" +
            "- every cell value in rows must be returned as text exactly as seen\n" +
            "- if parent+child headers exist, label them clearly, for example \"NG AUDIOBUS / SPL\"\n\n" +
            "Data:\n";

        string limited = rawData.Length > 20000 ? rawData[..20000] + "\n...(truncated)" : rawData;

        var body = new
        {
            model = "gpt-4o-mini",
            max_tokens = 8192,
            temperature = 0.0,
            response_format = new { type = "json_object" },
            messages = new[]
            {
                new { role = "system", content = "You extract spreadsheet tables into strict JSON." },
                new { role = "user", content = prompt + "\n" + limited }
            }
        };

        using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/chat/completions");
        req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
        
        req.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

        using HttpResponseMessage resp = await HttpClient.SendAsync(req, ct);
        resp.EnsureSuccessStatusCode();

        string json = await resp.Content.ReadAsStringAsync(ct);
        using JsonDocument doc = JsonDocument.Parse(json);
        JsonElement root = doc.RootElement;

        string text = string.Empty;
        if (root.TryGetProperty("choices", out JsonElement choices) && choices.GetArrayLength() > 0)
        {
            JsonElement message = choices[0].GetProperty("message");
            if (message.TryGetProperty("content", out JsonElement content))
                text = content.GetString() ?? string.Empty;
        }

        int inTok = 0, outTok = 0;
        if (root.TryGetProperty("usage", out JsonElement usage))
        {
            if (usage.TryGetProperty("prompt_tokens", out JsonElement i) && i.ValueKind == JsonValueKind.Number) inTok = i.GetInt32();
            if (usage.TryGetProperty("completion_tokens", out JsonElement o) && o.ValueKind == JsonValueKind.Number) outTok = o.GetInt32();
        }

        System.Diagnostics.Debug.WriteLine("=== OpenAI Response ===");
        System.Diagnostics.Debug.WriteLine(text);
        System.Diagnostics.Debug.WriteLine($"=== tokens: in={inTok} out={outTok} ===");

        var (cols, rows) = ParseOpenAiResponse(text);
        return (cols, rows, inTok, outTok);
    }

    private static (List<ColumnDef> Columns, List<Dictionary<string, string>> Rows)
        ParseOpenAiResponse(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return ([], []);
        string cleaned = text.Trim();

        // Strip markdown code fences (```json ... ``` or ``` ... ```)
        if (cleaned.StartsWith("```", StringComparison.Ordinal))
        {
            int nl = cleaned.IndexOf('\n');
            if (nl >= 0) cleaned = cleaned[(nl + 1)..];
            int closing = cleaned.LastIndexOf("```", StringComparison.Ordinal);
            if (closing > 0) cleaned = cleaned[..closing].Trim();
        }

        // If there is prose before the JSON object, extract the first {...} block
        int braceOpen = cleaned.IndexOf('{');
        int braceClose = cleaned.LastIndexOf('}');
        if (braceOpen > 0 && braceClose > braceOpen)
            cleaned = cleaned[braceOpen..(braceClose + 1)];

        try
        {
            using JsonDocument doc = JsonDocument.Parse(cleaned);
            JsonElement root = doc.RootElement;

            List<ColumnDef> cols = [];
            if (root.TryGetProperty("columns", out JsonElement colsEl))
                cols = JsonSerializer.Deserialize<List<ColumnDef>>(colsEl.GetRawText(), JsonOpts) ?? [];

            List<Dictionary<string, string>> rows = [];
            if (root.TryGetProperty("rows", out JsonElement rowsEl))
            {
                foreach (JsonElement rowEl in rowsEl.EnumerateArray())
                {
                    var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (JsonProperty prop in rowEl.EnumerateObject())
                        dict[prop.Name] = prop.Value.ValueKind == JsonValueKind.String
                            ? (prop.Value.GetString() ?? string.Empty)
                            : prop.Value.ToString();
                    rows.Add(dict);
                }
            }

            return (cols, rows);
        }
        catch { return ([], []); }
    }

    // -- Helpers --

    private async Task RunBusyAsync(Func<CancellationToken, Task> operation)
    {
        _isBusy = true; IsEnabled = false;
        try { await operation(CancellationToken.None); }
        catch (Exception ex) { SetStatus($"Error: {ex.Message}"); }
        finally { _isBusy = false; IsEnabled = true; NotifyWebModuleSnapshotChanged(); }
    }

    private void SetStatus(string message) => StatusTextBlock.Text = message;
    private void NotifyWebModuleSnapshotChanged() => WebModuleSnapshotChanged?.Invoke();

    private static BitmapImage BytesToBitmap(byte[] data)
    {
        using var ms = new MemoryStream(data);
        var bmp = new BitmapImage();
        bmp.BeginInit(); bmp.CacheOption = BitmapCacheOption.OnLoad; bmp.StreamSource = ms; bmp.EndInit();
        bmp.Freeze(); return bmp;
    }

    private static byte[] BitmapSourceToBytes(BitmapSource bmp)
    {
        using var ms = new MemoryStream();
        var enc = new PngBitmapEncoder();
        enc.Frames.Add(BitmapFrame.Create(bmp)); enc.Save(ms); return ms.ToArray();
    }

    // -- Private model classes --

    private sealed class PendingTable
    {
        public int    TableIndex { get; set; }
        public string TableName  { get; set; } = string.Empty;
        public List<ColumnDef>                  Columns { get; set; } = [];
        public List<Dictionary<string, string>> Rows    { get; set; } = [];
    }

    private sealed class PendingImage
    {
        public string FileName { get; set; } = string.Empty;
        public byte[] Data     { get; set; } = [];
    }

    private sealed class DatasetItem
    {
        public string Name       { get; init; } = string.Empty;
        public int    TableCount { get; init; }
        public int    ImageCount { get; init; }
    }
}
