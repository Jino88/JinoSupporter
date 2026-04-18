using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Controls.Primitives;
using System.Windows.Media.Imaging;
using WorkbenchHost.Infrastructure;

namespace JinoSupporter.App.Modules.DataInference;

public partial class DataInferenceView : UserControl
{
    public event Action? WebModuleSnapshotChanged;

    private static readonly HttpClient HttpClient = new() { Timeout = TimeSpan.FromSeconds(120) };
    private static readonly JsonSerializerOptions JsonOpts = new() { PropertyNameCaseInsensitive = true };

    // Haiku 4.5 pricing (per 1M tokens, USD)
    private const double PriceInputPer1M  = 0.80;
    private const double PriceOutputPer1M = 4.00;

    private readonly DataInferenceRepository _repository = new();
    private readonly List<PendingTable> _pendingTables = [];
    private readonly List<PendingImage> _pendingImages = [];
    private List<string> _pendingTags = [];
    private bool _tagsExtracted;
    private System.Data.DataTable? _dbEditTable;
    private DataTableInfo? _dbEditTableInfo;
    private bool _isBusy;
    private int _sessionInputTokens;
    private int _sessionOutputTokens;

    public DataInferenceView()
    {
        InitializeComponent();
        RefreshView();
        IsVisibleChanged += (_, e) =>
        {
            if ((bool)e.NewValue)
                DbPathTextBlock.Text = DataInferenceRepository.DatabasePath;
        };
        DataInputTextBox.TextChanged += (_, _) => UpdateOrganizeButton();
        UpdateOrganizeButton();
    }

    private void RefreshView()
    {
        DbPathTextBlock.Text = DataInferenceRepository.DatabasePath;
        RefreshDbInfo();
        RenderTablesPanel();
        if (!string.IsNullOrWhiteSpace(WorkbenchSettingsStore.GetAnthropicSessionKey()))
            _ = RefreshCreditBalanceAsync();
    }

    // -- Shell interface --

    public object GetWebModuleSnapshot() => new
    {
        moduleType = "DataInference",
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
            && !DataInputTextBox.IsKeyboardFocused
            && HasClipboardImage())
        {
            PasteImageFromClipboard();
            e.Handled = true;
        }

        base.OnPreviewKeyDown(e);
    }

    /// <summary>
    /// Returns true when the clipboard contains any image-like format,
    /// including the DIB / EMF formats that Excel uses when copying a cell range.
    /// </summary>
    private static bool HasClipboardImage()
    {
        if (Clipboard.ContainsImage()) return true;
        var data = Clipboard.GetDataObject();
        if (data is null) return false;
        return data.GetDataPresent("PNG")
            || data.GetDataPresent(DataFormats.Dib)
            || data.GetDataPresent(DataFormats.EnhancedMetafile)
            || data.GetDataPresent(DataFormats.Bitmap);
    }

    // -- Event handlers --

    private void TokenCostBadge_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        bool shiftHeld = (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) != 0;
        if (shiftHeld)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(
                "https://platform.claude.com/settings/billing") { UseShellExecute = true });
            return;
        }
        _ = RefreshCreditBalanceAsync();
    }

    private void TokenCostBadge_RightClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        var loginWin = new AnthropicLoginWindow { Owner = Window.GetWindow(this) };
        loginWin.ShowDialog();

        if (string.IsNullOrWhiteSpace(loginWin.SessionKey)) return;

        WorkbenchSettingsStore.SaveAnthropicSessionKey(loginWin.SessionKey);
        CreditBalanceTextBlock.Text = "Session key saved, fetching balance...";
        _ = RefreshCreditBalanceAsync();
    }

    private async Task RefreshCreditBalanceAsync()
    {
        string cookieHeader = WorkbenchSettingsStore.GetAnthropicSessionKey();
        if (string.IsNullOrWhiteSpace(cookieHeader))
        {
            CreditBalanceTextBlock.Text = "Balance: Session key required (right-click)";
            return;
        }

        CreditBalanceTextBlock.Text = "Fetching balance...";
        try
        {
            // platform.claude.com/settings/billing is a Next.js SPA → rendered via WebView2, text extracted by JS
            string? balance = await FetchBillingBalanceViaWebViewAsync(cookieHeader);

            if (balance != null)
            {
                CreditBalanceTextBlock.Text = $"Balance: {balance}";
                string raw = balance.Replace("US$", "").Replace("$", "").Replace(",", "").Trim();
                CreditBalanceTextBlock.Foreground = new SolidColorBrush(
                    double.TryParse(raw, out double v) && v < 1.0
                        ? Color.FromRgb(220, 38, 38)
                        : Color.FromRgb(5, 150, 105));
            }
            else
            {
                CreditBalanceTextBlock.Text = "Balance: Session expired (right-click → re-login)";
                CreditBalanceTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38));
            }
        }
        catch
        {
            CreditBalanceTextBlock.Text = "Balance: Fetch failed (right-click → re-login)";
        }
    }

    private Task<string?> FetchBillingBalanceViaWebViewAsync(string cookieHeader)
        => BillingHelper.FetchBalanceAsync(cookieHeader, Dispatcher);

    private async void AnalyzeButton_Click(object sender, RoutedEventArgs e) => await AnalyzeAsync();

    private void IncludeImagesCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        => UpdateOrganizeButton();

    private async void ExtractTagsButton_Click(object sender, RoutedEventArgs e) => await ExtractTagsAsync();

    private void GlossaryButton_Click(object sender, RoutedEventArgs e)
        => new GlossaryWindow(_repository) { Owner = Window.GetWindow(this) }.ShowDialog();

    private async void ReTagDbButton_Click(object sender, RoutedEventArgs e) => await ReTagDbDatasetAsync();

    private async void ReExtractGlossaryDbButton_Click(object sender, RoutedEventArgs e) => await ReExtractGlossaryFromDbAsync();

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
            DbImageBorder.Visibility  = Visibility.Collapsed;
            DbTagsMemoGrid.Visibility = Visibility.Collapsed;
            DbMemoTextBox.Clear();
            ReTagDbButton.IsEnabled = false;
            return;
        }

        LoadDbTables(ds.Name);
        RefreshDbImageStrip(ds.Name);
        RefreshDbTagStrip(ds.Name);
        DbMemoTextBox.Text           = _repository.GetMemo(ds.Name);
        DbTagsMemoGrid.Visibility    = Visibility.Visible;
        DbMemoAreaBorder.Visibility  = Visibility.Visible;
        DbTagsAreaBorder.Visibility  = Visibility.Visible;
        ReTagDbButton.IsEnabled = true;
        ReExtractGlossaryDbButton.IsEnabled = true;
    }

    private void RefreshDbButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshDbTab();
        RefreshDbInfo();
    }

    private void DeleteDbRowButton_Click(object sender, RoutedEventArgs e) =>
        DeleteSelectedTable();

    private void TablesDataGrid_PreviewMouseRightButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        // Select the row at the clicked position
        var row = ItemsControl.ContainerFromElement(TablesDataGrid, e.OriginalSource as DependencyObject) as DataGridRow;
        if (row == null) return;
        TablesDataGrid.SelectedItem = row.Item;

        var menu = new ContextMenu();
        var deleteItem = new MenuItem { Header = "Delete Row" };
        deleteItem.Click += (_, _) => DeleteSelectedTable();
        menu.Items.Add(deleteItem);
        menu.IsOpen = true;
        e.Handled = true;
    }

    private void RowDetailDataGrid_PreviewMouseRightButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        var row = ItemsControl.ContainerFromElement(RowDetailDataGrid, e.OriginalSource as DependencyObject) as DataGridRow;
        if (row == null) return;
        RowDetailDataGrid.SelectedItem = row.Item;

        var menu = new ContextMenu();
        var deleteItem = new MenuItem { Header = "Delete Row" };
        deleteItem.Click += (_, _) => DeleteSelectedDetailRow();
        menu.Items.Add(deleteItem);
        menu.IsOpen = true;
        e.Handled = true;
    }

    private void DeleteSelectedDetailRow()
    {
        if (_dbEditTable is null || _dbEditTableInfo is null) return;
        if (RowDetailDataGrid.SelectedItem is not System.Data.DataRowView rowView) return;

        if (MessageBox.Show(
                "Delete the selected row?",
                "Delete Row",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            return;

        _dbEditTable.Rows.Remove(rowView.Row);

        // Apply to DB immediately
        _repository.UpdateTableRows(_dbEditTableInfo.Id, _dbEditTableInfo.Columns, _dbEditTable);
        SetStatus($"Row deleted — {_dbEditTable.Rows.Count:N0} rows remaining");
    }

    private void DeleteSelectedTable()
    {
        if (TablesDataGrid.SelectedItem is not EditableTableInfo table)
        {
            SetStatus("Select a table to delete.");
            return;
        }

        if (MessageBox.Show(
                $"Delete the selected table?\nTable: {table.TableName}",
                "Delete Table",
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

        string rawData = DataInputTextBox.Text.Trim();
        bool includeImages = IncludeImagesCheckBox.IsChecked == true;
        bool hasImages     = _pendingImages.Count > 0;

        if (string.IsNullOrWhiteSpace(rawData) && !(includeImages && hasImages))
        {
            SetStatus("Paste table data into the left input area first, or attach images with Ctrl+V.");
            return;
        }

        string? apiKey = WorkbenchSettingsStore.TryGetClaudeApiKey();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            SetStatus("Claude API key is missing. Go to Settings > Claude API Key.");
            return;
        }

        List<PendingImage> imagesToSend = (includeImages && hasImages) ? [.. _pendingImages] : [];

        await RunBusyAsync(async ct =>
        {
                string glossary = _repository.GetGlossaryText();

                // If multiple images: each image → its own table
                // If text only, or text + images: one combined call
                bool perImage = imagesToSend.Count > 1 && string.IsNullOrWhiteSpace(rawData);

                var calls = perImage
                    ? imagesToSend.Select(img => (Text: string.Empty, Images: new List<PendingImage> { img })).ToList()
                    : new List<(string Text, List<PendingImage> Images)> { (rawData, imagesToSend) };

                int totalIn = 0, totalOut = 0;
                foreach (var (callText, callImages) in calls)
                {
                    int tableNo = _pendingTables.Count + 1;
                    string modeLabel = callImages.Count > 0
                        ? (perImage ? $"AI image {tableNo}/{calls.Count}" : "AI (image+text)")
                        : "AI";
                    SetStatus($"Organizing table {tableNo} via {modeLabel}...");

                    (List<ColumnDef> columns, List<Dictionary<string, string>> rows, int inTok, int outTok)
                        = await CallClaudeAsync(apiKey, callText, glossary, callImages, ct);

                    if (rows.Count == 0)
                    {
                        SetStatus($"Table {tableNo}: No data extracted.");
                        continue;
                    }

                    AddPendingTable(tableNo, columns, rows);
                    totalIn  += inTok;
                    totalOut += outTok;
                    _sessionInputTokens  += inTok;
                    _sessionOutputTokens += outTok;
                    UpdateTokenDisplay();
                    SetStatus($"Table {tableNo} done - {rows.Count:N0} rows. Extracting terms...");

                    (List<(string Term, string Desc)> terms, int gInTok, int gOutTok)
                        = await ExtractGlossaryTermsAsync(apiKey, columns, rows, ct);

                    _sessionInputTokens  += gInTok;
                    _sessionOutputTokens += gOutTok;
                    totalIn  += gInTok;
                    totalOut += gOutTok;
                    UpdateTokenDisplay();

                    if (terms.Count > 0)
                        MergeGlossaryToDb(terms);
                }

                int made = calls.Count;
                SetStatus($"{made} table(s) done | tokens {totalIn + totalOut:N0}");
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
                List<Dictionary<string, string>> rowsToSave;
                if (table.SourceTable is not null)
                {
                    rowsToSave = [];
                    foreach (System.Data.DataRow dr in table.SourceTable.Rows)
                    {
                        if (dr.RowState == System.Data.DataRowState.Deleted) continue;
                        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (ColumnDef col in table.Columns)
                            dict[col.Field] = dr[col.Field]?.ToString() ?? string.Empty;
                        rowsToSave.Add(dict);
                    }
                }
                else
                {
                    rowsToSave = table.Rows;
                }

                if (rowsToSave.Count == 0) continue;
                _repository.SaveTable(datasetName, table.TableName, table.Columns, rowsToSave);
                savedRec += rowsToSave.Count;
            }

            foreach (PendingImage pi in _pendingImages)
                _repository.SaveImage(datasetName, pi.FileName, pi.Data);

            int savedTags = _pendingTags.Count;
            if (savedTags > 0)
                _repository.SaveTags(datasetName, _pendingTags);

            string memo = DatasetMemoTextBox.Text.Trim();
            if (!string.IsNullOrEmpty(memo))
                _repository.SaveMemo(datasetName, memo);

            int savedImg = _pendingImages.Count;
            _pendingTables.Clear();
            _pendingImages.Clear();
            _pendingTags.Clear();
            _tagsExtracted = false;
            RenderTagBadges([]);
            RenderTablesPanel();
            RefreshPendingImageStrip();
            UpdateSaveButton();
            RefreshDbInfo();
            if (MainTabControl.SelectedItem == DbDataTab) RefreshDbTab();

            // Clear input fields after successful save
            DatasetNameTextBox.Clear();
            DatasetMemoTextBox.Clear();

            SetStatus($"Saved '{datasetName}' - {savedRec:N0} records, {savedImg:N0} images, {savedTags} tags.");
        }
        catch (Exception ex)
        {
            SetStatus($"DB save error: {ex.Message}");
        }
    }

    private void ClearAll()
    {
        DataInputTextBox.Clear();
        DatasetMemoTextBox.Clear();
        _pendingTables.Clear();
        _pendingImages.Clear();
        _pendingTags.Clear();
        _tagsExtracted = false;
        RenderTagBadges([]);
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

        DataGrid dg = BuildDynamicDataGrid(table.Columns, table.Rows, maxHeight: 300, owner: table);

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
                                          double maxHeight = double.NaN,
                                          PendingTable? owner = null)
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

        if (owner is not null) owner.SourceTable = dt;

        var dg = new DataGrid
        {
            AutoGenerateColumns  = false,
            CanUserAddRows       = false,
            CanUserDeleteRows    = owner is not null,  // row deletion allowed only for pending cards
            IsReadOnly           = owner is null,
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

    // -- Tag extraction (pending data) --

    private async Task ExtractTagsAsync()
    {
        string datasetName = DatasetNameTextBox.Text.Trim();
        string memo        = DatasetMemoTextBox.Text.Trim();

        // Extraction is possible even without a table if DatasetName or Memo is provided
        if (_pendingTables.Count == 0 && string.IsNullOrWhiteSpace(datasetName) && string.IsNullOrWhiteSpace(memo))
        {
            SetStatus("Enter a Dataset Name or Memo first.");
            DatasetNameTextBox.Focus();
            return;
        }

        if (string.IsNullOrWhiteSpace(datasetName))
        {
            SetStatus("Enter the file title (Dataset Name) first.");
            DatasetNameTextBox.Focus();
            return;
        }

        string? apiKey = WorkbenchSettingsStore.TryGetClaudeApiKey();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            SetStatus("Claude API key is missing. Please enter it in Settings.");
            return;
        }

        await RunBusyAsync(async ct =>
        {
            SetStatus("Extracting tags...");
            string prompt = BuildTagExtractionPrompt(datasetName, _pendingTables, memo);
            (List<string> tags, int inTok, int outTok) = await CallClaudeForTagsAsync(apiKey, prompt, ct);

            if (tags.Count == 0)
            {
                SetStatus("Tag extraction failed. Please try again.");
                return;
            }

            _pendingTags = tags;
            _tagsExtracted = true;
            _sessionInputTokens  += inTok;
            _sessionOutputTokens += outTok;
            UpdateTokenDisplay();
            RenderTagBadges(tags);
            UpdateSaveButton();
            SetStatus($"{tags.Count} tag(s) extracted. Click Save to DB to store them.");
        });
    }

    // -- Tag re-assignment (existing DB dataset) --

    private async Task ReTagDbDatasetAsync()
    {
        if (DatasetDataGrid.SelectedItem is not DatasetItem ds) return;

        string? apiKey = WorkbenchSettingsStore.TryGetClaudeApiKey();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            SetStatus("Claude API key is missing. Please enter it in Settings.");
            return;
        }

        await RunBusyAsync(async ct =>
        {
            SetStatus($"Re-tagging '{ds.Name}'...");

            // Load tables from DB and build virtual PendingTable list for prompt
            List<DataTableInfo> dbTables = _repository.GetTables(ds.Name);
            var virtualTables = new List<PendingTable>();
            foreach (DataTableInfo tbl in dbTables)
            {
                List<(long Id, Dictionary<string, string> Data)> dbRows = _repository.GetTableRows(tbl.Id);
                virtualTables.Add(new PendingTable
                {
                    TableIndex = (int)tbl.Id,
                    TableName  = tbl.TableName,
                    Columns    = tbl.Columns,
                    Rows       = dbRows.Select(r => r.Data).ToList()
                });
            }

            string memo = _repository.GetMemo(ds.Name);
            string prompt = BuildTagExtractionPrompt(ds.Name, virtualTables, memo);
            (List<string> tags, int inTok, int outTok) = await CallClaudeForTagsAsync(apiKey, prompt, ct);

            if (tags.Count == 0)
            {
                SetStatus("Tag extraction failed. Please try again.");
                return;
            }

            _repository.SaveTags(ds.Name, tags);
            _sessionInputTokens  += inTok;
            _sessionOutputTokens += outTok;
            UpdateTokenDisplay();
            RefreshDbTagStrip(ds.Name);
            SetStatus($"'{ds.Name}' re-tagged with {tags.Count} tag(s).");
        });
    }

    private string BuildTagExtractionPrompt(string datasetName, List<PendingTable> tables, string memo = "")
    {
        List<string> existingTags = _repository.GetAllDistinctTags();

        var sb = new StringBuilder();
        sb.AppendLine("Based on the information below, extract tags that describe what this data was reviewing.");
        sb.AppendLine();
        sb.AppendLine($"[Dataset Name] {datasetName}");

        if (!string.IsNullOrWhiteSpace(memo))
            sb.AppendLine($"[Memo / Purpose] {memo}");

        if (tables.Count > 0)
        {
            sb.AppendLine("[Column name list]");
            foreach (PendingTable table in tables)
                sb.AppendLine("  " + string.Join(", ", table.Columns.Select(c => c.Label)));
        }

        sb.AppendLine();

        if (existingTags.Count > 0)
        {
            sb.AppendLine("=== Existing tags (reuse the same name if meaning is the same) ===");
            sb.AppendLine(string.Join(", ", existingTags));
            sb.AppendLine();
        }

        sb.AppendLine("Respond with a JSON array only (no other text):");
        sb.AppendLine("[\"tag1\", \"tag2\", ...]\n");
        sb.AppendLine("Tag guidelines:");
        sb.AppendLine("- Focus on content inferred from dataset name, memo, and column names (review target, process, model, period, inspection item, etc.)");
        sb.AppendLine("- If an existing tag has the same or similar meaning, reuse that exact name");
        sb.AppendLine("- Add a new tag only for content not covered by existing tags");
        sb.AppendLine("- Short, clear words/phrases in Korean, at least 5, no duplicates");

        return sb.ToString();
    }

    private static async Task<(List<string> Tags, int InputTokens, int OutputTokens)>
        CallClaudeForTagsAsync(string apiKey, string prompt, CancellationToken ct)
    {
        string limited = prompt.Length > 12000 ? prompt[..12000] + "\n...(truncated)" : prompt;

        var body = new
        {
            model = "claude-haiku-4-5-20251001",
            max_tokens = 1024,
            messages = new[] { new { role = "user", content = limited } }
        };

        using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
        req.Headers.Add("x-api-key", apiKey);
        req.Headers.Add("anthropic-version", "2023-06-01");
        req.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

        using HttpResponseMessage resp = await HttpClient.SendAsync(req, ct);
        resp.EnsureSuccessStatusCode();

        string json = await resp.Content.ReadAsStringAsync(ct);
        using JsonDocument doc = JsonDocument.Parse(json);
        JsonElement root = doc.RootElement;

        string text = string.Empty;
        if (root.TryGetProperty("content", out JsonElement content) && content.GetArrayLength() > 0)
            text = content[0].GetProperty("text").GetString() ?? string.Empty;

        int inTok = 0, outTok = 0;
        if (root.TryGetProperty("usage", out JsonElement usage))
        {
            if (usage.TryGetProperty("input_tokens",  out JsonElement i) && i.ValueKind == JsonValueKind.Number) inTok  = i.GetInt32();
            if (usage.TryGetProperty("output_tokens", out JsonElement o) && o.ValueKind == JsonValueKind.Number) outTok = o.GetInt32();
        }

        System.Diagnostics.Debug.WriteLine($"=== Tag Response ===\n{text}\n=== tokens: in={inTok} out={outTok} ===");

        text = text.Trim();
        if (text.StartsWith("```", StringComparison.Ordinal))
        {
            int nl = text.IndexOf('\n');
            if (nl >= 0) text = text[(nl + 1)..];
            int closing = text.LastIndexOf("```", StringComparison.Ordinal);
            if (closing > 0) text = text[..closing].Trim();
        }

        int arrOpen  = text.IndexOf('[');
        int arrClose = text.LastIndexOf(']');
        if (arrOpen >= 0 && arrClose > arrOpen)
            text = text[arrOpen..(arrClose + 1)];

        try { return (JsonSerializer.Deserialize<List<string>>(text) ?? [], inTok, outTok); }
        catch { return ([], inTok, outTok); }
    }

    private void RenderTagBadges(List<string> tags)
    {
        TagsWrapPanel.Children.Clear();
        foreach (string tag in tags)
        {
            TagsWrapPanel.Children.Add(new Border
            {
                Margin          = new Thickness(0, 0, 4, 4),
                Padding         = new Thickness(8, 3, 8, 3),
                CornerRadius    = new CornerRadius(12),
                Background      = new SolidColorBrush(Color.FromRgb(219, 234, 254)),
                BorderBrush     = new SolidColorBrush(Color.FromRgb(147, 197, 253)),
                BorderThickness = new Thickness(1),
                Child           = new TextBlock
                {
                    Text       = tag,
                    FontSize   = 10,
                    Foreground = new SolidColorBrush(Color.FromRgb(29, 78, 216))
                }
            });
        }
        TagsAreaBorder.Visibility = tags.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void RefreshDbTagStrip(string datasetName)
    {
        List<string> tags = _repository.GetTags(datasetName);
        DbTagsWrapPanel.Children.Clear();
        foreach (string tag in tags)
        {
            DbTagsWrapPanel.Children.Add(new Border
            {
                Margin          = new Thickness(0, 0, 4, 4),
                Padding         = new Thickness(8, 3, 8, 3),
                CornerRadius    = new CornerRadius(12),
                Background      = new SolidColorBrush(Color.FromRgb(219, 234, 254)),
                BorderBrush     = new SolidColorBrush(Color.FromRgb(147, 197, 253)),
                BorderThickness = new Thickness(1),
                Child           = new TextBlock
                {
                    Text       = tag,
                    FontSize   = 10,
                    Foreground = new SolidColorBrush(Color.FromRgb(29, 78, 216))
                }
            });
        }
        DbTagsMemoGrid.Visibility    = tags.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        DbTagsAreaBorder.Visibility  = tags.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void UpdateSaveButton()
    {
        bool hasData   = _pendingTables.Count > 0 || _pendingImages.Count > 0;
        bool hasTables = _pendingTables.Count > 0;
        ExtractTagsButton.IsEnabled = hasTables;
        SaveDbButton.IsEnabled = hasData && (!hasTables || _tagsExtracted);
        UpdateOrganizeButton();
    }

    private void UpdateOrganizeButton()
    {
        bool imageOnlyEnabled = IncludeImagesCheckBox.IsChecked == true && _pendingImages.Count > 0;
        bool hasText          = !string.IsNullOrWhiteSpace(DataInputTextBox.Text);
        OrganizeButton.IsEnabled = hasText || imageOnlyEnabled;
    }

    // -- Glossary extraction --

    private async Task ReExtractGlossaryFromDbAsync()
    {
        if (DatasetDataGrid.SelectedItem is not DatasetItem ds) return;

        string? apiKey = WorkbenchSettingsStore.TryGetClaudeApiKey();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            SetStatus("Claude API key is missing.");
            return;
        }

        await RunBusyAsync(async ct =>
        {
            SetStatus($"Extracting terms from '{ds.Name}'...");

            List<DataTableInfo> dbTables = _repository.GetTables(ds.Name);
            var allColumns = new List<ColumnDef>();
            var sampleRows = new List<Dictionary<string, string>>();

            foreach (DataTableInfo tbl in dbTables)
            {
                allColumns.AddRange(tbl.Columns);
                var dbRows = _repository.GetTableRows(tbl.Id);
                sampleRows.AddRange(dbRows.Take(5).Select(r => r.Data));
            }

            (List<(string Term, string Desc)> terms, int inTok, int outTok)
                = await ExtractGlossaryTermsAsync(apiKey, allColumns, sampleRows, ct);

            _sessionInputTokens  += inTok;
            _sessionOutputTokens += outTok;
            UpdateTokenDisplay();

            if (terms.Count > 0)
                MergeGlossaryToDb(terms);

            SetStatus($"{terms.Count} term(s) extracted.");
        });
    }

    private static async Task<(List<(string Term, string Desc)> Terms, int InputTokens, int OutputTokens)>
        ExtractGlossaryTermsAsync(string apiKey, List<ColumnDef> columns, List<Dictionary<string, string>> rows, CancellationToken ct)
    {
        var sb = new StringBuilder();
        sb.AppendLine("From the following data, extract proper nouns, abbreviations, and domain-specific terms that AI may not fully understand.");
        sb.AppendLine("Exclude common words; target only specialized terms related to work/products/processes/inspections.\n");
        sb.AppendLine("Respond with a JSON array only (no other text):");
        sb.AppendLine("[{\"term\": \"term\", \"description\": \"brief explanation\"}, ...]\n");
        sb.AppendLine("Column names: " + string.Join(", ", columns.Select(c => c.Label)));
        sb.AppendLine("Sample values (up to 5 rows):");
        int sampleCount = Math.Min(rows.Count, 5);
        for (int i = 0; i < sampleCount; i++)
            sb.AppendLine(string.Join("\t", columns.Select(c => rows[i].TryGetValue(c.Field, out string? v) ? v : string.Empty)));

        var body = new
        {
            model = "claude-haiku-4-5-20251001",
            max_tokens = 1024,
            messages = new[] { new { role = "user", content = sb.ToString() } }
        };

        using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
        req.Headers.Add("x-api-key", apiKey);
        req.Headers.Add("anthropic-version", "2023-06-01");
        req.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

        using HttpResponseMessage resp = await HttpClient.SendAsync(req, ct);
        resp.EnsureSuccessStatusCode();

        string json = await resp.Content.ReadAsStringAsync(ct);
        using JsonDocument doc = JsonDocument.Parse(json);
        JsonElement root = doc.RootElement;

        string text = string.Empty;
        if (root.TryGetProperty("content", out JsonElement content) && content.GetArrayLength() > 0)
            text = content[0].GetProperty("text").GetString() ?? string.Empty;

        int inTok = 0, outTok = 0;
        if (root.TryGetProperty("usage", out JsonElement usage))
        {
            if (usage.TryGetProperty("input_tokens",  out JsonElement i) && i.ValueKind == JsonValueKind.Number) inTok  = i.GetInt32();
            if (usage.TryGetProperty("output_tokens", out JsonElement o) && o.ValueKind == JsonValueKind.Number) outTok = o.GetInt32();
        }

        System.Diagnostics.Debug.WriteLine($"=== Glossary Response ===\n{text}\n=== tokens: in={inTok} out={outTok} ===");

        text = text.Trim();
        if (text.StartsWith("```", StringComparison.Ordinal))
        {
            int nl = text.IndexOf('\n');
            if (nl >= 0) text = text[(nl + 1)..];
            int closing = text.LastIndexOf("```", StringComparison.Ordinal);
            if (closing > 0) text = text[..closing].Trim();
        }

        int arrOpen  = text.IndexOf('[');
        int arrClose = text.LastIndexOf(']');
        if (arrOpen >= 0 && arrClose > arrOpen)
            text = text[arrOpen..(arrClose + 1)];

        try
        {
            using JsonDocument arr = JsonDocument.Parse(text);
            var terms = new List<(string, string)>();
            foreach (JsonElement el in arr.RootElement.EnumerateArray())
            {
                string term = el.TryGetProperty("term", out JsonElement t) ? t.GetString() ?? string.Empty : string.Empty;
                string desc = el.TryGetProperty("description", out JsonElement d) ? d.GetString() ?? string.Empty : string.Empty;
                if (!string.IsNullOrWhiteSpace(term))
                    terms.Add((term, desc));
            }
            return (terms, inTok, outTok);
        }
        catch { return ([], inTok, outTok); }
    }

    private void MergeGlossaryToDb(List<(string Term, string Desc)> newTerms)
    {
        var existing = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string line in _repository.GetGlossaryText().Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            int colon = line.IndexOf(':', StringComparison.Ordinal);
            if (colon <= 0) continue;
            string t = line[..colon].Trim();
            if (!string.IsNullOrWhiteSpace(t))
                existing[t] = line[(colon + 1)..].Trim();
        }

        foreach ((string term, string desc) in newTerms)
            if (!existing.ContainsKey(term))
                existing[term] = desc;

        string merged = string.Join("\n",
            existing.OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                    .Select(kv => $"{kv.Key}: {kv.Value}"));
        _repository.MergeGlossary(merged);
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
        DbImageBorder.Visibility   = Visibility.Collapsed;
        DbTagsMemoGrid.Visibility  = Visibility.Collapsed;
        DbMemoTextBox.Clear();
        SaveDbRowsButton.IsEnabled              = false;
        ReTagDbButton.IsEnabled                 = false;
        ReExtractGlossaryDbButton.IsEnabled     = false;
        DbStatusTextBlock.Text = $"{items.Count} dataset(s)";
    }

    private void LoadDbTables(string datasetName)
    {
        List<DataTableInfo> tables = _repository.GetTables(datasetName);
        TablesDataGrid.ItemsSource = tables.Select(t => new EditableTableInfo(t)).ToList();
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
        if (TablesDataGrid.SelectedItem is not EditableTableInfo table)
        {
            RowDetailDataGrid.ItemsSource = null;
            RowDetailDataGrid.Columns.Clear();
            _dbEditTable     = null;
            _dbEditTableInfo = null;
            SaveDbRowsButton.IsEnabled = false;
            return;
        }

        List<DataTableInfo> fresh = _repository.GetTables(table.DatasetName);
        DataTableInfo? freshTable = fresh.FirstOrDefault(t => t.Id == table.Id);
        if (freshTable is null) return;

        List<(long Id, Dictionary<string, string> Data)> rows = _repository.GetTableRows(table.Id);
        BuildRowDetailGrid(freshTable, rows);
        DbStatusTextBlock.Text = $"{table.TableName} - {rows.Count:N0} rows";
        SaveDbRowsButton.IsEnabled = true;
    }

    private void TablesDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (e.EditAction != DataGridEditAction.Commit) return;
        if (e.Row.Item is not EditableTableInfo item) return;
        if (e.EditingElement is not TextBox tb) return;

        string newName = tb.Text.Trim();
        if (string.IsNullOrEmpty(newName) || newName == item.TableName) return;

        _repository.RenameTable(item.Id, newName);
        item.TableName = newName;
        SetStatus($"Table renamed to '{newName}'");
    }

    private void BuildRowDetailGrid(DataTableInfo table, List<(long Id, Dictionary<string, string> Data)> rows)
    {
        var dt = new System.Data.DataTable();
        foreach (ColumnDef col in table.Columns)
            dt.Columns.Add(col.Field, typeof(string));

        foreach ((_, Dictionary<string, string> data) in rows)
        {
            System.Data.DataRow dr = dt.NewRow();
            foreach (ColumnDef col in table.Columns)
                dr[col.Field] = data.TryGetValue(col.Field, out string? v) ? v : string.Empty;
            dt.Rows.Add(dr);
        }

        _dbEditTable     = dt;
        _dbEditTableInfo = table;

        RowDetailDataGrid.Columns.Clear();
        RowDetailDataGrid.IsReadOnly = false;
        foreach (ColumnDef col in table.Columns)
        {
            string fieldName  = col.Field;
            string label      = string.IsNullOrWhiteSpace(col.Label) ? col.Field : col.Label;

            var deleteItem = new MenuItem { Header = $"Delete column '{label}'" };
            deleteItem.Click += (_, _) => DeleteColumnFromTable(table, fieldName);

            var ctxMenu = new ContextMenu();
            ctxMenu.Items.Add(deleteItem);

            var headerStyle = new Style(typeof(DataGridColumnHeader));
            headerStyle.Setters.Add(new Setter(DataGridColumnHeader.ContextMenuProperty, ctxMenu));

            RowDetailDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header      = label,
                Binding     = new Binding(col.Field),
                HeaderStyle = headerStyle
            });
        }

        RowDetailDataGrid.ItemsSource = dt.DefaultView;
    }

    private void SaveMemoDbButton_Click(object sender, RoutedEventArgs e)
    {
        if (DatasetDataGrid.SelectedItem is not DatasetItem ds) return;
        _repository.SaveMemo(ds.Name, DbMemoTextBox.Text.Trim());
        SetStatus($"Memo for '{ds.Name}' saved.");
    }

    private void SaveDbRowsButton_Click(object sender, RoutedEventArgs e)
    {
        if (_dbEditTable is null || _dbEditTableInfo is null) return;

        RowDetailDataGrid.CommitEdit(DataGridEditingUnit.Row, exitEditingMode: true);

        try
        {
            _repository.UpdateTableRows(_dbEditTableInfo.Id, _dbEditTableInfo.Columns, _dbEditTable);
            SetStatus($"'{_dbEditTableInfo.TableName}' changes saved — {_dbEditTable.Rows.Count:N0} rows");
        }
        catch (Exception ex)
        {
            SetStatus($"Save error: {ex.Message}");
        }
    }

    private void DeleteColumnFromTable(DataTableInfo table, string fieldName)
    {
        string label = table.Columns.FirstOrDefault(c => c.Field == fieldName)?.Label ?? fieldName;
        if (MessageBox.Show($"Delete column '{label}'?\nThis action cannot be undone.",
                "Delete Column", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            return;

        _repository.DeleteColumn(table.Id, fieldName);

        List<DataTableInfo> fresh = _repository.GetTables(table.DatasetName);
        DataTableInfo? updated = fresh.FirstOrDefault(t => t.Id == table.Id);
        if (updated is null) return;

        List<(long Id, Dictionary<string, string> Data)> rows = _repository.GetTableRows(table.Id);
        BuildRowDetailGrid(updated, rows);
        SetStatus($"Column '{label}' deleted.");
    }

    private void DbInfoBorder_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFolderDialog
        {
            Title       = "Select DB Save Folder",
            Multiselect = false
        };

        if (dialog.ShowDialog(Window.GetWindow(this)) != true) return;

        string folder     = dialog.FolderName;
        string fileName   = Path.GetFileName(DataInferenceRepository.DatabasePath);
        string targetPath = Path.Combine(folder, fileName);

        if (string.Equals(targetPath, DataInferenceRepository.DatabasePath, StringComparison.OrdinalIgnoreCase))
        {
            SetStatus("That is the same location as the current one.");
            return;
        }

        try
        {
            File.Copy(DataInferenceRepository.DatabasePath, targetPath, overwrite: true);
            WorkbenchSettingsStore.SaveDataInferenceDatabasePath(targetPath);
            DbPathTextBlock.Text = targetPath;
            SetStatus($"DB moved → {targetPath}  (new location will be used on restart)");
        }
        catch (Exception ex)
        {
            SetStatus($"DB move error: {ex.Message}");
        }
    }

    private void PasteImageFromClipboard()
    {
        BitmapSource? bmp = GetClipboardImageSource();
        if (bmp is null) return;

        bmp = FlattenToWhite(bmp);
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

    // -- Claude API --

    private static async Task<(List<ColumnDef> Columns, List<Dictionary<string, string>> Rows,
                                int InputTokens, int OutputTokens)>
        CallClaudeAsync(string apiKey, string rawData, string glossary,
                        List<PendingImage> images, CancellationToken ct)
    {
        var promptSb = new StringBuilder();

        bool imageOnly = images.Count > 0 && string.IsNullOrWhiteSpace(rawData);

        if (!string.IsNullOrWhiteSpace(glossary))
        {
            promptSb.AppendLine("=== Glossary (domain terms) ===");
            foreach (string line in glossary.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                promptSb.AppendLine($"- {line.Trim()}");
            promptSb.AppendLine();
        }

        if (imageOnly)
        {
            // Image-only prompt: instruct Claude to visually read the table from the screenshot
            promptSb.AppendLine(
                "The attached image(s) are Excel sheet screenshots containing manufacturing inspection or production data.\n" +
                "\n" +
                "【Rules — must follow strictly】\n" +
                "\n" +
                "▶ STEP 1. Merged cell handling — CRITICAL\n" +
                "   Excel merged cells show one value spanning multiple rows/columns visually.\n" +
                "   When unmerging, EVERY cell in the merged range must receive that value — including the FIRST row/column of the range.\n" +
                "   The value belongs to the TOP-LEFT cell of the merge; all other cells in the block copy it.\n" +
                "   - Horizontal merge (spans multiple columns): copy the value into EACH column it covers.\n" +
                "     e.g. 'NG AUDIOBUS' spans columns 7-10 → columns 7, 8, 9, 10 all get 'NG AUDIOBUS'\n" +
                "   - Vertical merge (spans multiple rows): copy the value into EACH row it covers, starting from the FIRST row.\n" +
                "     e.g. 'Model A' visually appears in the middle of rows 1-5 → rows 1, 2, 3, 4, 5 ALL get 'Model A' (including row 1)\n" +
                "   - Combined merge (spans both rows and columns): fill EVERY cell in the entire block.\n" +
                "   DO NOT leave any cell empty that was part of a merged range.\n" +
                "   DO NOT skip the first row of a vertical merge — the first row gets the value too.\n" +
                "\n" +
                "▶ STEP 2. Multi-row headers\n" +
                "   - Apply STEP 1 merge-fill first on each header row independently.\n" +
                "   - Then concatenate header rows top-to-bottom per column, omitting exact duplicate words.\n" +
                "     e.g. col 7: row1='NG AUDIOBUS', row2='SPL' → label='NG AUDIOBUS SPL'\n" +
                "     e.g. col 9: row1='NG AUDIOBUS', row2='No sound' → label='NG AUDIOBUS No sound'\n" +
                "\n" +
                "▶ STEP 3. Data rows\n" +
                "   - Apply STEP 1 merge-fill to every data cell (vertical merges across rows are common).\n" +
                "   - Include rows with actual measurements or row numbers.\n" +
                "   - Exclude total/subtotal/average/grand total/blank rows.\n" +
                "\n" +
                "▶ STEP 4. Output — return JSON only (no ``` or other text)\n" +
                "{\"columns\":[{\"field\":\"camelCaseEnglish\",\"label\":\"OriginalHeaderName\"}],\"rows\":[{...}]}\n" +
                "   - field: English camelCase identifier\n" +
                "   - label: original header text as read from the image\n" +
                "   - all values are strings\n");
        }
        else
        {
            // Text (tab-separated) prompt
            promptSb.AppendLine(
                "The following is tab-separated text copied from Excel, containing manufacturing inspection and production data.\n" +
                "\n" +
                "【Table Parsing Rules — must follow strictly】\n" +
                "\n" +
                "▶ STEP 1. Determine column count by tab position (column index)\n" +
                "   - Split each row by tab (\\t) and assign indices 0, 1, 2 ... N.\n" +
                "   - The total column count (N+1) is determined by the row with the most tabs among all header and data rows.\n" +
                "   - If a row has fewer cells than the max tabs, pad the right side with empty strings.\n" +
                "\n" +
                "▶ STEP 2. Multi-row header handling (key — based on tab index)\n" +
                "   - If there are 2 or more header rows, collect header row values for each index i from top to bottom.\n" +
                "   - If index i in a header row is an empty string, fill it with the last non-empty value to the left in the same row (merged cell forward-fill).\n" +
                "   - Concatenate the filled headers from top to bottom, omitting duplicate words.\n" +
                "     e.g. index 7: row1='NG AUDIOBUS', row2='SPL' → label='NG AUDIOBUS SPL'\n" +
                "     e.g. index 9: row1='NG AUDIOBUS', row2='No sound' → label='NG AUDIOBUS No sound'\n" +
                "     e.g. index 13: row1='NG HEARING', row2='Noise' → label='NG HEARING Noise'\n" +
                "   - The label for each column must be the value computed from the tab index;\n" +
                "     do not change the index based on contextual inference.\n" +
                "\n" +
                "▶ STEP 3. Select data rows\n" +
                "   - Include in rows only rows where the first cell is a number (row number) or an actual measurement.\n" +
                "   - Exclude percentage-decimal helper rows (e.g. rows consisting only of '3.4%\\t0.0%...'), blank rows, and total/subtotal/grand total/average rows.\n" +
                "   - For blank cells in data rows, fill down with the value from the previous data row in the same column.\n" +
                "\n" +
                "▶ STEP 4. Output — return JSON only (no ``` or other text)\n" +
                "{\"columns\":[{\"field\":\"camelCaseEnglish\",\"label\":\"OriginalHeaderName\"}],\"rows\":[{...}]}\n" +
                "   - field: English camelCase (remove spaces and special characters)\n" +
                "   - label: index-based header computed in STEP 2\n" +
                "   - all values are strings (including numbers and percentages)\n" +
                "\n" +
                "Data:\n");

            string limited = rawData.Length > 40000 ? rawData[..40000] + "\n...(truncated)" : rawData;
            promptSb.AppendLine(limited);
        }

        // Build request JSON using JsonNode to avoid anonymous-type-in-List<object> serialization pitfall
        JsonNode messageContent;
        if (images.Count > 0)
        {
            var arr = new JsonArray();
            foreach (PendingImage img in images)
            {
                byte[] imgData = ResizeImageForApi(img.Data);
                arr.Add(new JsonObject
                {
                    ["type"]   = "image",
                    ["source"] = new JsonObject
                    {
                        ["type"]       = "base64",
                        ["media_type"] = "image/png",
                        ["data"]       = Convert.ToBase64String(imgData)
                    }
                });
            }
            arr.Add(new JsonObject
            {
                ["type"] = "text",
                ["text"] = promptSb.ToString()
            });
            messageContent = arr;
        }
        else
        {
            messageContent = JsonValue.Create(promptSb.ToString())!;
        }

        var bodyNode = new JsonObject
        {
            ["model"]      = "claude-haiku-4-5-20251001",
            ["max_tokens"] = 8192,
            ["messages"]   = new JsonArray
            {
                new JsonObject
                {
                    ["role"]    = "user",
                    ["content"] = messageContent
                }
            }
        };

        string requestJson = bodyNode.ToJsonString();
        // Log request (truncate base64 image data for readability)
        ApiLog($"[Claude API REQUEST] images={images.Count} textLen={promptSb.Length}\n" +
               (images.Count == 0 ? requestJson[..Math.Min(500, requestJson.Length)] : "(multipart with images)"));

        using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
        req.Headers.Add("x-api-key", apiKey);
        req.Headers.Add("anthropic-version", "2023-06-01");
        req.Content = new StringContent(requestJson, Encoding.UTF8, "application/json");

        using HttpResponseMessage resp = await HttpClient.SendAsync(req, ct);
        string json = await resp.Content.ReadAsStringAsync(ct);
        ApiLog($"[Claude API] status={resp.StatusCode}\n{json}");
        if (!resp.IsSuccessStatusCode)
        {
            string errMsg = json;
            try
            {
                using JsonDocument errDoc = JsonDocument.Parse(json);
                if (errDoc.RootElement.TryGetProperty("error", out JsonElement errEl) &&
                    errEl.TryGetProperty("message", out JsonElement msgEl))
                    errMsg = msgEl.GetString() ?? json;
            }
            catch { /* use raw body */ }
            throw new HttpRequestException($"Claude API {(int)resp.StatusCode}: {errMsg}");
        }

        using JsonDocument doc = JsonDocument.Parse(json);
        JsonElement root = doc.RootElement;

        string text = string.Empty;
        if (root.TryGetProperty("content", out JsonElement content) && content.GetArrayLength() > 0)
            text = content[0].GetProperty("text").GetString() ?? string.Empty;

        int inTok = 0, outTok = 0;
        if (root.TryGetProperty("usage", out JsonElement usage))
        {
            if (usage.TryGetProperty("input_tokens",  out JsonElement i) && i.ValueKind == JsonValueKind.Number) inTok  = i.GetInt32();
            if (usage.TryGetProperty("output_tokens", out JsonElement o) && o.ValueKind == JsonValueKind.Number) outTok = o.GetInt32();
        }

        System.Diagnostics.Debug.WriteLine("=== Claude Response ===");
        System.Diagnostics.Debug.WriteLine(text);
        System.Diagnostics.Debug.WriteLine($"=== tokens: in={inTok} out={outTok} ===");

        var (cols, rows) = ParseClaudeResponse(text);
        FillMergedCells(cols, rows);
        return (cols, rows, inTok, outTok);
    }

    /// <summary>
    /// Post-processing: for each column, if leading rows are empty but a later row has a value,
    /// fill those empty leading rows with that value.
    /// This fixes the common case where Claude omits the value in the first row of a vertical merge.
    /// Also applies standard fill-down (empty cell inherits the previous non-empty value above).
    /// </summary>
    private static void FillMergedCells(List<ColumnDef> cols, List<Dictionary<string, string>> rows)
    {
        if (rows.Count == 0) return;

        foreach (ColumnDef col in cols)
        {
            string field = col.Field;

            // Pass 1: fill-UP — propagate first non-empty value backward to fill empty leading rows
            string? firstValue = null;
            int firstValueIndex = -1;
            for (int i = 0; i < rows.Count; i++)
            {
                if (rows[i].TryGetValue(field, out string? v) && !string.IsNullOrWhiteSpace(v))
                {
                    firstValue      = v;
                    firstValueIndex = i;
                    break;
                }
            }
            if (firstValue is not null && firstValueIndex > 0)
            {
                for (int i = 0; i < firstValueIndex; i++)
                    rows[i][field] = firstValue;
            }

            // Pass 2: fill-DOWN — standard forward fill for mid-table empty cells
            string? last = null;
            for (int i = 0; i < rows.Count; i++)
            {
                if (rows[i].TryGetValue(field, out string? v) && !string.IsNullOrWhiteSpace(v))
                    last = v;
                else if (last is not null)
                    rows[i][field] = last;
            }
        }
    }

    private static (List<ColumnDef> Columns, List<Dictionary<string, string>> Rows)
        ParseClaudeResponse(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return ([], []);
        string cleaned = text.Trim();

        // Remove code fence: find first ``` block regardless of position and extract just the content
        int fenceOpen = cleaned.IndexOf("```", StringComparison.Ordinal);
        if (fenceOpen >= 0)
        {
            int nl = cleaned.IndexOf('\n', fenceOpen);
            if (nl >= 0)
            {
                int fenceClose = cleaned.IndexOf("```", nl, StringComparison.Ordinal);
                cleaned = fenceClose > nl
                    ? cleaned[(nl + 1)..fenceClose].Trim()
                    : cleaned[(nl + 1)..].Trim();
            }
        }

        // Extract { ... } block (strip surrounding prose text)
        int braceOpen  = cleaned.IndexOf('{');
        int braceClose = cleaned.LastIndexOf('}');
        if (braceOpen >= 0 && braceClose > braceOpen)
            cleaned = cleaned[braceOpen..(braceClose + 1)];

        System.Diagnostics.Debug.WriteLine($"[ParseClaudeResponse] cleaned length={cleaned.Length}, start={cleaned[..Math.Min(80, cleaned.Length)]}");

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

            System.Diagnostics.Debug.WriteLine($"[ParseClaudeResponse] cols={cols.Count}, rows={rows.Count}");
            return (cols, rows);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ParseClaudeResponse] PARSE ERROR: {ex.Message}\ncleaned={cleaned[..Math.Min(200, cleaned.Length)]}");
            return ([], []);
        }
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

    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "JinoWorkHost", "datainference_api.log");

    private static void ApiLog(string message)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
            File.AppendAllText(LogPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n\n");
        }
        catch { /* logging must not crash the app */ }
        System.Diagnostics.Debug.WriteLine(message);
    }

    private static BitmapImage BytesToBitmap(byte[] data)
    {
        using var ms = new MemoryStream(data);
        var bmp = new BitmapImage();
        bmp.BeginInit(); bmp.CacheOption = BitmapCacheOption.OnLoad; bmp.StreamSource = ms; bmp.EndInit();
        bmp.Freeze(); return bmp;
    }

    /// <summary>
    /// Gets a BitmapSource from the clipboard.
    /// Priority: System.Drawing.Bitmap → WPF BitmapSource → PNG → EMF → DIB
    /// </summary>
    private static BitmapSource? GetClipboardImageSource()
    {
        var data = Clipboard.GetDataObject();
        if (data is null) return null;

        ApiLog($"[Clipboard formats] {string.Join(", ", data.GetFormats())}");

        // 1. System.Drawing.Bitmap — Excel places this directly, most reliable
        if (data.GetDataPresent("System.Drawing.Bitmap"))
        {
            try
            {
                if (data.GetData("System.Drawing.Bitmap") is System.Drawing.Bitmap sdBmp)
                    return ConvertSdBitmapToSource(sdBmp);
            }
            catch (Exception ex) { ApiLog($"[SDBitmap failed] {ex.Message}"); }
        }

        // 2. WPF BitmapSource
        if (data.GetDataPresent("System.Windows.Media.Imaging.BitmapSource"))
        {
            try
            {
                if (data.GetData("System.Windows.Media.Imaging.BitmapSource") is BitmapSource bs)
                    return bs;
            }
            catch (Exception ex) { ApiLog($"[WPFBitmapSource failed] {ex.Message}"); }
        }

        // 3. Standard WPF GetImage
        if (Clipboard.ContainsImage())
        {
            try { return Clipboard.GetImage(); }
            catch (Exception ex) { ApiLog($"[WPF GetImage failed] {ex.Message}"); }
        }

        // 4. PNG stream
        if (data.GetDataPresent("PNG") && data.GetData("PNG") is MemoryStream pngMs)
        {
            try
            {
                var bi = new BitmapImage();
                bi.BeginInit();
                bi.CacheOption  = BitmapCacheOption.OnLoad;
                bi.StreamSource = pngMs;
                bi.EndInit();
                bi.Freeze();
                return bi;
            }
            catch (Exception ex) { ApiLog($"[PNG failed] {ex.Message}"); }
        }

        // 5. EMF — render at 2x screen DPI for sharpness
        if (data.GetDataPresent(DataFormats.EnhancedMetafile))
        {
            try
            {
                if (data.GetData(DataFormats.EnhancedMetafile) is System.Drawing.Imaging.Metafile mf)
                using (mf)
                {
                    var header = mf.GetMetafileHeader();
                    // bounds in 0.01mm units; convert to px at 192 DPI (2× for clarity)
                    const float targetDpi = 192f;
                    int w = Math.Max(1, (int)(header.Bounds.Width  / 2540f * targetDpi));
                    int h = Math.Max(1, (int)(header.Bounds.Height / 2540f * targetDpi));
                    using var sdBmp = new System.Drawing.Bitmap(w, h,
                        System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    sdBmp.SetResolution(targetDpi, targetDpi);
                    using (var g = System.Drawing.Graphics.FromImage(sdBmp))
                    {
                        g.Clear(System.Drawing.Color.White);
                        g.DrawImage(mf, 0, 0, w, h);
                    }
                    return ConvertSdBitmapToSource(sdBmp);
                }
            }
            catch (Exception ex) { ApiLog($"[EMF failed] {ex.Message}"); }
        }

        return null;
    }

    private static BitmapSource ConvertSdBitmapToSource(System.Drawing.Bitmap sdBmp)
    {
        var hBitmap = sdBmp.GetHbitmap();
        try
        {
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                hBitmap,
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }
        finally
        {
            DeleteObject(hBitmap);
        }
    }

    [System.Runtime.InteropServices.DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr hObject);

    /// <summary>Composites the image onto a white background, removing any transparency.</summary>
    private static BitmapSource FlattenToWhite(BitmapSource src)
    {
        var visual = new System.Windows.Media.DrawingVisual();
        using (var dc = visual.RenderOpen())
        {
            dc.DrawRectangle(System.Windows.Media.Brushes.White, null,
                new Rect(0, 0, src.PixelWidth, src.PixelHeight));
            dc.DrawImage(src, new Rect(0, 0, src.PixelWidth, src.PixelHeight));
        }
        var rt = new RenderTargetBitmap(src.PixelWidth, src.PixelHeight,
            src.DpiX == 0 ? 96 : src.DpiX,
            src.DpiY == 0 ? 96 : src.DpiY,
            System.Windows.Media.PixelFormats.Pbgra32);
        rt.Render(visual);
        rt.Freeze();
        return rt;
    }

    private static byte[] BitmapSourceToBytes(BitmapSource bmp)
    {
        using var ms = new MemoryStream();
        var enc = new PngBitmapEncoder();
        enc.Frames.Add(BitmapFrame.Create(bmp)); enc.Save(ms); return ms.ToArray();
    }

    /// <summary>
    /// Resizes image bytes so they stay under the Claude API 5 MB per-image limit.
    /// Progressively halves the decode width until the encoded PNG fits.
    /// </summary>
    private static byte[] ResizeImageForApi(byte[] data)
    {
        const long MaxBytes = 4_500_000; // 4.5 MB — safe margin under the 5 MB limit
        if (data.Length <= MaxBytes) return data;

        // Start from a width that should bring us under the limit
        BitmapImage original;
        using (var ms = new MemoryStream(data))
        {
            original = new BitmapImage();
            original.BeginInit();
            original.CacheOption    = BitmapCacheOption.OnLoad;
            original.StreamSource   = ms;
            original.EndInit();
            original.Freeze();
        }

        // Estimate initial scale: pixel area scales with file size (rough heuristic)
        double scale     = Math.Sqrt((double)MaxBytes / data.Length) * 0.85;
        int    tryWidth  = Math.Max(400, (int)(original.PixelWidth * scale));

        for (int attempt = 0; attempt < 6; attempt++)
        {
            BitmapImage resized;
            using (var ms = new MemoryStream(data))
            {
                resized = new BitmapImage();
                resized.BeginInit();
                resized.CacheOption    = BitmapCacheOption.OnLoad;
                resized.StreamSource   = ms;
                resized.DecodePixelWidth = tryWidth;
                resized.EndInit();
                resized.Freeze();
            }

            using var outMs = new MemoryStream();
            var enc = new PngBitmapEncoder();
            enc.Frames.Add(BitmapFrame.Create(resized));
            enc.Save(outMs);
            byte[] result = outMs.ToArray();

            ApiLog($"[ResizeImage] attempt={attempt} width={tryWidth} size={result.Length:N0} bytes");

            if (result.Length <= MaxBytes) return result;
            tryWidth = (int)(tryWidth * 0.7);
        }

        // Fallback: JPEG at 70% quality — always much smaller
        BitmapImage fallback;
        using (var ms = new MemoryStream(data))
        {
            fallback = new BitmapImage();
            fallback.BeginInit();
            fallback.CacheOption    = BitmapCacheOption.OnLoad;
            fallback.StreamSource   = ms;
            fallback.DecodePixelWidth = 1200;
            fallback.EndInit();
            fallback.Freeze();
        }
        using var jpgMs = new MemoryStream();
        var jpgEnc = new JpegBitmapEncoder { QualityLevel = 70 };
        jpgEnc.Frames.Add(BitmapFrame.Create(fallback));
        jpgEnc.Save(jpgMs);
        return jpgMs.ToArray();
    }

    // -- Private model classes --

    private sealed class PendingTable
    {
        public int    TableIndex { get; set; }
        public string TableName  { get; set; } = string.Empty;
        public List<ColumnDef>                  Columns     { get; set; } = [];
        public List<Dictionary<string, string>> Rows        { get; set; } = [];
        public System.Data.DataTable?           SourceTable { get; set; }
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

    private sealed class EditableTableInfo : System.ComponentModel.INotifyPropertyChanged
    {
        private string _tableName;
        public long           Id            { get; }
        public string         DatasetName   { get; }
        public List<ColumnDef> Columns      { get; }
        public int            RowCount      { get; }
        public string         CreatedAtLocal { get; }

        public string TableName
        {
            get => _tableName;
            set { _tableName = value; PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(TableName))); }
        }

        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;

        public EditableTableInfo(DataTableInfo info)
        {
            Id             = info.Id;
            DatasetName    = info.DatasetName;
            Columns        = info.Columns;
            RowCount       = info.RowCount;
            CreatedAtLocal = info.CreatedAtLocal;
            _tableName     = info.TableName;
        }
    }
}