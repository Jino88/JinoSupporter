using System;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace JinoSupporter.App.Modules.DataInference;

public partial class ChatGptTableInferenceWindow : Window
{
    private bool _isUpdatingText;

    public ChatGptTableInferenceWindow()
    {
        InitializeComponent();
    }

    private void PasteClipboardButton_Click(object sender, RoutedEventArgs e)
    {
        if (!Clipboard.ContainsText())
        {
            StatusTextBlock.Text = "No text data in the clipboard.";
            return;
        }

        _isUpdatingText = true;
        RawInputTextBox.Text = Clipboard.GetText();
        _isUpdatingText = false;
        ParseRawInput();
    }

    private void ParseButton_Click(object sender, RoutedEventArgs e)
    {
        ParseRawInput();
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        _isUpdatingText = true;
        RawInputTextBox.Clear();
        _isUpdatingText = false;
        ParsedDataGrid.ItemsSource = null;
        StatusTextBlock.Text = "Input data has been cleared.";
    }

    private void RawInputTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isUpdatingText)
        {
            return;
        }

        ParseRawInput();
    }

    private void ParseRawInput()
    {
        string raw = RawInputTextBox.Text;
        if (string.IsNullOrWhiteSpace(raw))
        {
            ParsedDataGrid.ItemsSource = null;
            StatusTextBlock.Text = "No pasted data.";
            return;
        }

        string[] lines = raw
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .ToArray();

        if (lines.Length == 0)
        {
            ParsedDataGrid.ItemsSource = null;
            StatusTextBlock.Text = "No valid rows found.";
            return;
        }

        string[][] rows = lines.Select(line => line.Split('\t')).ToArray();
        int maxColumnCount = rows.Max(row => row.Length);
        bool hasHeader = rows.Length > 1;

        DataTable table = new();

        for (int i = 0; i < maxColumnCount; i++)
        {
            string header = hasHeader && i < rows[0].Length && !string.IsNullOrWhiteSpace(rows[0][i])
                ? rows[0][i].Trim()
                : $"Column{i + 1}";

            string uniqueHeader = header;
            int suffix = 2;
            while (table.Columns.Contains(uniqueHeader))
            {
                uniqueHeader = $"{header}_{suffix++}";
            }

            table.Columns.Add(uniqueHeader);
        }

        int startRowIndex = hasHeader ? 1 : 0;
        for (int rowIndex = startRowIndex; rowIndex < rows.Length; rowIndex++)
        {
            DataRow row = table.NewRow();
            for (int col = 0; col < maxColumnCount; col++)
            {
                row[col] = col < rows[rowIndex].Length ? rows[rowIndex][col].Trim() : string.Empty;
            }

            table.Rows.Add(row);
        }

        ParsedDataGrid.ItemsSource = table.DefaultView;
        StatusTextBlock.Text = $"Table created with {table.Rows.Count:N0} rows and {table.Columns.Count:N0} columns.";
    }
}
