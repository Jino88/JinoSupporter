using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WorkbenchHost.Infrastructure;

namespace JinoSupporter.App.Modules.DataInference;

public partial class DataInferenceReportView : UserControl
{
    private static readonly HttpClient HttpClient = new() { Timeout = TimeSpan.FromSeconds(180) };

    // Haiku 4.5 pricing (per 1M tokens, USD)
    private const double PriceInputPer1M  = 0.80;
    private const double PriceOutputPer1M = 4.00;

    private readonly DataInferenceRepository _repository = new();
    private List<DatasetCheckItem> _datasets = [];
    private readonly HashSet<string> _selectedTags = new(StringComparer.OrdinalIgnoreCase);
    private string? _lastHtml;
    private bool _isBusy;
    private int _sessionInputTokens;
    private int _sessionOutputTokens;

    private string _loadedDbPath = string.Empty;

    public DataInferenceReportView()
    {
        InitializeComponent();
        RefreshAll();
        IsVisibleChanged += (_, e) =>
        {
            if ((bool)e.NewValue && DataInferenceRepository.DatabasePath != _loadedDbPath)
                RefreshAll();
        };
    }

    private void RefreshAll()
    {
        _loadedDbPath = DataInferenceRepository.DatabasePath;
        LoadTagFilter();
        LoadDatasets();
        LoadSavedReports();
        if (!string.IsNullOrWhiteSpace(WorkbenchSettingsStore.GetAnthropicSessionKey()))
            _ = RefreshCreditBalanceAsync();
    }

    // -- Event handlers --

    private void TokenCostBadge_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        => _ = RefreshCreditBalanceAsync();

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
            CreditBalanceTextBlock.Text       = "Balance: Set Session Key (right-click)";
            CreditBalanceTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(97, 112, 132));
            return;
        }

        CreditBalanceTextBlock.Text = "Fetching balance...";
        try
        {
            string? balance = await BillingHelper.FetchBalanceAsync(cookieHeader, Dispatcher);
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
                CreditBalanceTextBlock.Text       = "Balance: Session expired (right-click)";
                CreditBalanceTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(220, 38, 38));
            }
        }
        catch
        {
            CreditBalanceTextBlock.Text = "Balance: Fetch failed";
        }
    }

    private void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        LoadTagFilter();
        LoadDatasets();
        LoadSavedReports();
    }

    private void ClearTagFilterButton_Click(object sender, RoutedEventArgs e)
    {
        _selectedTags.Clear();
        foreach (CheckBox cb in TagFilterPanel.Children.OfType<CheckBox>())
            cb.IsChecked = false;
        LoadDatasets();
    }

    private void SelectAllButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (DatasetCheckItem item in _datasets)
            if (item.CheckBox is not null) item.CheckBox.IsChecked = true;
    }

    private void ClearSelectionButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (DatasetCheckItem item in _datasets)
            if (item.CheckBox is not null) item.CheckBox.IsChecked = false;
    }

    private async void GenerateReportButton_Click(object sender, RoutedEventArgs e)
        => await GenerateReportAsync();

    private void SaveHtmlButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastHtml)) { SetStatus("No report to save."); return; }

        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "HTML files|*.html",
            FileName = $"report_{DateTime.Now:yyyyMMdd_HHmmss}.html"
        };

        if (dlg.ShowDialog() == true)
        {
            File.WriteAllText(dlg.FileName, _lastHtml, Encoding.UTF8);
            SetStatus($"Saved: {dlg.FileName}");
        }
    }

    private void CopyHtmlButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastHtml)) { SetStatus("No report to copy."); return; }
        Clipboard.SetText(_lastHtml);
        SetStatus("HTML copied to clipboard.");
    }

    // -- Core logic --

    private void LoadSavedReports()
    {
        SavedReportsPanel.Children.Clear();
        List<(long Id, string Title, string DatasetNames, string CreatedAt)> reports = _repository.GetReportList();

        if (reports.Count == 0)
        {
            SavedReportsPanel.Children.Add(new TextBlock
            {
                Text = "No saved reports",
                FontSize = 10,
                Foreground = new SolidColorBrush(Color.FromRgb(97, 112, 132)),
                Margin = new Thickness(2, 4, 2, 4)
            });
            return;
        }

        foreach ((long id, string title, string dsNames, string createdAt) in reports)
        {
            long capturedId    = id;
            string capturedTitle = title;

            var row = new Border
            {
                Padding         = new Thickness(6, 4, 6, 4),
                Margin          = new Thickness(0, 0, 0, 2),
                CornerRadius    = new CornerRadius(4),
                Background      = new SolidColorBrush(Color.FromRgb(248, 250, 253)),
                BorderBrush     = new SolidColorBrush(Color.FromRgb(211, 220, 234)),
                BorderThickness = new Thickness(1),
                Cursor          = System.Windows.Input.Cursors.Hand,
                ToolTip         = $"{dsNames}\n{createdAt}"
            };

            var inner = new DockPanel();

            // Delete button
            var delBtn = new Button
            {
                Content         = "✕",
                FontSize        = 9,
                Width           = 16, Height = 16,
                Padding         = new Thickness(0),
                Background      = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground      = new SolidColorBrush(Color.FromRgb(180, 40, 40)),
                Cursor          = System.Windows.Input.Cursors.Hand,
                VerticalAlignment = VerticalAlignment.Center,
                ToolTip         = "Delete"
            };
            delBtn.Click += (_, e) =>
            {
                e.Handled = true;
                if (MessageBox.Show($"Delete this report?\n{capturedTitle}",
                        "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                    return;
                _repository.DeleteReport(capturedId);
                LoadSavedReports();
            };
            DockPanel.SetDock(delBtn, Dock.Right);
            inner.Children.Add(delBtn);

            var txt = new TextBlock
            {
                Text          = $"{createdAt}\n{dsNames}",
                FontSize      = 9,
                Foreground    = new SolidColorBrush(Color.FromRgb(34, 48, 74)),
                TextWrapping  = TextWrapping.NoWrap,
                TextTrimming  = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center
            };
            inner.Children.Add(txt);

            row.Child = inner;
            row.MouseLeftButtonUp += (_, _) =>
            {
                string html = _repository.GetReportHtml(capturedId);
                if (string.IsNullOrWhiteSpace(html)) return;
                _lastHtml = html;
                ReportPlaceholder.Visibility = Visibility.Collapsed;
                ReportBrowser.NavigateToString(html);
                ReportTitleText.Text = capturedTitle;
            };

            SavedReportsPanel.Children.Add(row);
        }
    }

    private void LoadTagFilter()
    {
        TagFilterPanel.Children.Clear();

        List<string> allTags = _repository.GetAllDistinctTags();
        if (allTags.Count == 0)
        {
            TagFilterPanel.Children.Add(new TextBlock
            {
                Text = "No tags",
                FontSize = 10,
                Foreground = new SolidColorBrush(Color.FromRgb(148, 163, 184)),
                Margin = new Thickness(4, 4, 4, 4)
            });
            return;
        }

        foreach (string tag in allTags)
        {
            string captured = tag;
            var cb = new CheckBox
            {
                Content = tag,
                FontSize = 11,
                Margin = new Thickness(0, 2, 0, 2),
                IsChecked = _selectedTags.Contains(tag)
            };
            cb.Checked   += (_, _) => { _selectedTags.Add(captured);    LoadDatasets(); };
            cb.Unchecked += (_, _) => { _selectedTags.Remove(captured); LoadDatasets(); };
            TagFilterPanel.Children.Add(cb);
        }
    }

    private void LoadDatasets()
    {
        DatasetCheckPanel.Children.Clear();
        _datasets.Clear();

        // List of dataset names filtered by selected tags
        List<string> filteredNames = _repository.GetDatasetsByTags(_selectedTags.ToList());
        List<(string Name, int TableCount, int ImageCount)> summary = _repository.GetDatasetSummary();

        // Apply filter
        if (_selectedTags.Count > 0)
            summary = summary.Where(s => filteredNames.Contains(s.Name, StringComparer.OrdinalIgnoreCase)).ToList();

        if (summary.Count == 0)
        {
            DatasetCheckPanel.Children.Add(new TextBlock
            {
                Text = _selectedTags.Count > 0 ? "No datasets match the selected tags." : "No datasets in DB.",
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(97, 112, 132)),
                Margin = new Thickness(4, 6, 4, 6)
            });
            SetStatus("No datasets found.");
            return;
        }

        foreach ((string name, int tableCount, int imageCount) in summary)
        {
            var item = new DatasetCheckItem { Name = name, TableCount = tableCount, ImageCount = imageCount };

            var cb = new CheckBox
            {
                Margin = new Thickness(0, 3, 0, 3),
                IsChecked = false
            };

            var label = new StackPanel { Orientation = Orientation.Horizontal };
            label.Children.Add(new TextBlock
            {
                Text = name,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(34, 48, 74))
            });
            label.Children.Add(new TextBlock
            {
                Text = $"  {tableCount}t",
                FontSize = 10,
                Foreground = new SolidColorBrush(Color.FromRgb(97, 112, 132)),
                VerticalAlignment = VerticalAlignment.Center
            });
            if (imageCount > 0)
            {
                label.Children.Add(new TextBlock
                {
                    Text = $" / {imageCount}img",
                    FontSize = 10,
                    Foreground = new SolidColorBrush(Color.FromRgb(74, 111, 165)),
                    VerticalAlignment = VerticalAlignment.Center
                });
            }

            cb.Content = label;
            item.CheckBox = cb;
            _datasets.Add(item);
            DatasetCheckPanel.Children.Add(cb);
        }

        SetStatus($"{_datasets.Count} dataset(s) loaded." + (_selectedTags.Count > 0 ? $" (Tag filter: {_selectedTags.Count} active)" : string.Empty));
    }

    private async Task GenerateReportAsync()
    {
        if (_isBusy) return;

        List<DatasetCheckItem> selected = _datasets.Where(d => d.CheckBox?.IsChecked == true).ToList();
        if (selected.Count == 0)
        {
            SetStatus("Select at least one dataset.");
            return;
        }

        string? apiKey = WorkbenchSettingsStore.TryGetClaudeApiKey();
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            SetStatus("Claude API key missing. Go to Settings.");
            return;
        }

        _isBusy = true;
        IsEnabled = false;
        try
        {
            SetStatus("Building data summary...");
            string prompt = BuildReportPrompt(selected);

            // Collect images for all selected datasets
            var imageBlocks = new List<object>();
            foreach (DatasetCheckItem ds in selected)
            {
                foreach (DatasetImageRow img in _repository.GetImages(ds.Name))
                {
                    string ext = Path.GetExtension(img.FileName).TrimStart('.').ToLowerInvariant();
                    string mediaType = ext switch
                    {
                        "jpg" or "jpeg" => "image/jpeg",
                        "gif"           => "image/gif",
                        "webp"          => "image/webp",
                        _               => "image/png"
                    };
                    imageBlocks.Add(new
                    {
                        type   = "image",
                        source = new { type = "base64", media_type = mediaType, data = Convert.ToBase64String(img.ImageData) }
                    });
                }
            }

            SetStatus($"Generating report ({selected.Count} dataset(s))... Please wait.");
            int outTokLive = 0;
            (string html, int inTok, int outTok) = await CallClaudeStreamingAsync(
                apiKey, prompt, imageBlocks, CancellationToken.None,
                onProgress: (chars, tokens) =>
                {
                    outTokLive = tokens;
                    Dispatcher.Invoke(() =>
                        SetStatus($"Generating report... {chars:N0} chars written / {tokens:N0} output tokens"));
                });

            if (string.IsNullOrWhiteSpace(html))
            {
                SetStatus("Claude returned an empty response. Try again.");
                return;
            }

            _sessionInputTokens  += inTok;
            _sessionOutputTokens += outTok;
            UpdateTokenDisplay();

            // Append reference images section
            if (imageBlocks.Count > 0)
            {
                string imgSection = BuildImageSection(selected);
                int bodyClose = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                html = bodyClose >= 0 ? html.Insert(bodyClose, imgSection) : html + imgSection;
            }

            _lastHtml = html;
            ReportPlaceholder.Visibility = Visibility.Collapsed;
            ReportBrowser.NavigateToString(html);

            string title = $"{DateTime.Now:yyyy-MM-dd HH:mm} — {string.Join(", ", selected.Select(s => s.Name))}";
            ReportTitleText.Text = $"Report — {title}";

            // Auto-save to DB (with selected dataset name list)
            string dsNames = string.Join(", ", selected.Select(s => s.Name));
            _repository.SaveReport(title, dsNames, html);
            LoadSavedReports();
            SetStatus($"Report generated ({selected.Count} dataset(s)) — saved to DB.");
        }
        catch (Exception ex)
        {
            SetStatus($"Error: {ex.Message}");
        }
        finally
        {
            _isBusy = false;
            IsEnabled = true;
        }
    }

    private string BuildImageSection(List<DatasetCheckItem> selected)
    {
        var sb = new StringBuilder();
        sb.AppendLine("""
            <section style="margin:40px auto;max-width:1000px;padding:0 24px;">
              <h2 style="font-size:18px;font-weight:700;color:#1e293b;border-bottom:2px solid #e2e8f0;padding-bottom:8px;margin-bottom:20px;">Reference Images</h2>
              <div style="display:flex;flex-wrap:wrap;gap:16px;">
            """);
        foreach (DatasetCheckItem ds in selected)
        {
            foreach (DatasetImageRow img in _repository.GetImages(ds.Name))
            {
                string ext = Path.GetExtension(img.FileName).TrimStart('.').ToLowerInvariant();
                string mediaType = ext switch
                {
                    "jpg" or "jpeg" => "image/jpeg",
                    "gif"           => "image/gif",
                    "webp"          => "image/webp",
                    _               => "image/png"
                };
                string base64 = Convert.ToBase64String(img.ImageData);
                string encodedName = System.Net.WebUtility.HtmlEncode(img.FileName);
                sb.AppendLine($"""
                      <figure style="margin:0;text-align:center;">
                        <img src="data:{mediaType};base64,{base64}"
                             alt="{encodedName}"
                             style="max-width:480px;max-height:360px;object-fit:contain;border:1px solid #e2e8f0;border-radius:8px;" />
                        <figcaption style="font-size:11px;color:#64748b;margin-top:4px;">{encodedName}</figcaption>
                      </figure>
                    """);
            }
        }
        sb.AppendLine("  </div></section>");
        return sb.ToString();
    }

    private string BuildReportPrompt(List<DatasetCheckItem> selected)
    {
        var sb = new StringBuilder();
        sb.AppendLine(
            "You are an expert in manufacturing data analysis. Analyze the data below and write an HTML document in Korean.\n" +
            "\n" +
            "Each dataset includes a [Memo] and [Tags].\n" +
            "- Memo: contextual information written by the person in charge, covering the purpose, background, and review content of the data\n" +
            "- Tags: keywords representing the nature of the data\n" +
            "Use this information actively to derive insights that fit the memo/tag context, not just a list of numbers.\n" +
            "\n" +
            "Requirements:\n" +
            "- Return only a complete HTML document with embedded CSS, no markdown or code fences.\n" +
            "- Use a clean, modern style (white background, readable font, alternating row-color tables).\n" +
            "- Structure: three sections:\n" +
            "  1) Consolidated data table: merge similar items into one table, include numeric aggregates (sum, mean, max, min, defect rate, etc.)\n" +
            "  2) Correlation analysis: summarize relationships between numeric columns in a table; highlight notable patterns/outliers briefly in the relevant cell\n" +
            "  3) AI inference & insights: based on context from the memo and tags, bullet-list what the data implies.\n" +
            "     - Describe notable patterns, anomalies, improvement points, and cautions with specific figures\n" +
            "     - If the memo specifies a review objective/question, answer it directly\n" +
            "- Write all text in Korean.\n");
        sb.AppendLine($"Report date: {DateTime.Now:yyyy-MM-dd HH:mm}");

        string glossary = _repository.GetGlossaryText();
        if (!string.IsNullOrWhiteSpace(glossary))
        {
            sb.AppendLine();
            sb.AppendLine("=== Glossary ===");
            foreach (string line in glossary.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                sb.AppendLine($"- {line.Trim()}");
        }

        sb.AppendLine();

        foreach (DatasetCheckItem ds in selected)
        {
            sb.AppendLine($"=== Dataset: {ds.Name} ===");

            List<string> tags = _repository.GetTags(ds.Name);
            if (tags.Count > 0)
                sb.AppendLine($"[Tags] {string.Join(", ", tags)}");

            string memo = _repository.GetMemo(ds.Name);
            if (!string.IsNullOrWhiteSpace(memo))
                sb.AppendLine($"[Memo] {memo}");

            List<DataTableInfo> tables = _repository.GetTables(ds.Name);

            foreach (DataTableInfo table in tables)
            {
                List<(long Id, Dictionary<string, string> Data)> rows = _repository.GetTableRows(table.Id);
                sb.AppendLine($"Table: {table.TableName} ({rows.Count} rows)");

                // Header row
                sb.AppendLine(string.Join("\t", table.Columns.Select(c => c.Label)));

                // Data rows (cap at 200 per table to stay within token budget)
                int limit = Math.Min(rows.Count, 200);
                for (int i = 0; i < limit; i++)
                {
                    Dictionary<string, string> data = rows[i].Data;
                    sb.AppendLine(string.Join("\t",
                        table.Columns.Select(c => data.TryGetValue(c.Field, out string? v) ? v : string.Empty)));
                }

                if (rows.Count > 200)
                    sb.AppendLine($"... ({rows.Count - 200} more rows omitted)");

                sb.AppendLine();
            }

            sb.AppendLine();
        }

        return sb.ToString();
    }

    private static async Task<(string Html, int InputTokens, int OutputTokens)> CallClaudeStreamingAsync(
        string apiKey, string prompt, List<object> imageBlocks, CancellationToken ct,
        Action<int, int>? onProgress = null)
    {
        string limited = prompt.Length > 40000 ? prompt[..40000] + "\n...(truncated)" : prompt;

        object content;
        if (imageBlocks.Count > 0)
        {
            var blocks = new List<object>(imageBlocks) { new { type = "text", text = limited } };
            content = blocks.ToArray();
        }
        else
        {
            content = limited;
        }

        var body = new
        {
            model      = "claude-opus-4-6",
            max_tokens = 16000,
            stream     = true,
            messages   = new[] { new { role = "user", content } }
        };

        using var req = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
        req.Headers.Add("x-api-key", apiKey);
        req.Headers.Add("anthropic-version", "2023-06-01");
        req.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

        using HttpResponseMessage resp = await HttpClient.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
        resp.EnsureSuccessStatusCode();

        var textSb   = new StringBuilder();
        int inTok = 0, outTok = 0;
        int reportInterval = 0;

        using Stream stream = await resp.Content.ReadAsStreamAsync(ct);
        using var reader = new System.IO.StreamReader(stream);

        while (!reader.EndOfStream)
        {
            string? line = await reader.ReadLineAsync();
            if (string.IsNullOrWhiteSpace(line)) continue;
            if (!line.StartsWith("data: ", StringComparison.Ordinal)) continue;

            string data = line[6..];
            if (data == "[DONE]") break;

            try
            {
                using JsonDocument doc = JsonDocument.Parse(data);
                JsonElement root = doc.RootElement;

                string? type = root.TryGetProperty("type", out JsonElement typeEl)
                    ? typeEl.GetString() : null;

                if (type == "content_block_delta")
                {
                    if (root.TryGetProperty("delta", out JsonElement delta)
                        && delta.TryGetProperty("text", out JsonElement textEl))
                    {
                        textSb.Append(textEl.GetString());
                        outTok++;

                        // progress callback every 50 chunks
                        if (onProgress != null && ++reportInterval % 50 == 0)
                            onProgress(textSb.Length, outTok);
                    }
                }
                else if (type == "message_start")
                {
                    if (root.TryGetProperty("message", out JsonElement msg)
                        && msg.TryGetProperty("usage", out JsonElement usage)
                        && usage.TryGetProperty("input_tokens", out JsonElement i))
                        inTok = i.GetInt32();
                }
                else if (type == "message_delta")
                {
                    if (root.TryGetProperty("usage", out JsonElement usage)
                        && usage.TryGetProperty("output_tokens", out JsonElement o))
                        outTok = o.GetInt32();
                }
            }
            catch { /* ignore lines that fail to parse */ }
        }

        string text = textSb.ToString().Trim();
        if (text.StartsWith("```", StringComparison.Ordinal))
        {
            int nl = text.IndexOf('\n');
            if (nl >= 0) text = text[(nl + 1)..];
            int closing = text.LastIndexOf("```", StringComparison.Ordinal);
            if (closing > 0) text = text[..closing].Trim();
        }

        return (text, inTok, outTok);
    }

    private void UpdateTokenDisplay()
    {
        double cost = (_sessionInputTokens / 1_000_000.0 * PriceInputPer1M)
                    + (_sessionOutputTokens / 1_000_000.0 * PriceOutputPer1M);
        TokenCostTextBlock.Text  = $"${cost:F4}";
        TokenUsageTextBlock.Text = $"{_sessionInputTokens + _sessionOutputTokens:N0} tokens";
    }

    private void SetStatus(string message) => StatusTextBlock.Text = message;

    // -- Model --

    private sealed class DatasetCheckItem
    {
        public string  Name       { get; init; } = string.Empty;
        public int     TableCount { get; init; }
        public int     ImageCount { get; init; }
        public CheckBox? CheckBox { get; set; }
    }
}
