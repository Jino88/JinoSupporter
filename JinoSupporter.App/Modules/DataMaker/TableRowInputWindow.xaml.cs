using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace DataMaker
{
    public partial class TableRowInputWindow : Window
    {
        private sealed class MissingItemOption
        {
            public required IReadOnlyDictionary<string, string> ValuesByColumn { get; init; }
        }

        private readonly List<(string ColumnName, TextBox TextBox)> _fields = new();
        private readonly IReadOnlyList<string> _columns;
        private List<MissingItemOption> _allMissingItems = new();
        private bool _isUpdatingFieldsFromSelection;
        private readonly bool _enableLiveFiltering;

        public IReadOnlyList<string> Values => _fields
            .Select(field => field.TextBox.Text.Trim())
            .ToList();

        public TableRowInputWindow(
            string title,
            string targetDescription,
            IReadOnlyList<string> columns,
            IReadOnlyList<IReadOnlyDictionary<string, string>>? missingItems = null,
            bool enableLiveFiltering = true)
        {
            InitializeComponent();

            Title = title;
            HeaderTextBlock.Text = title;
            PathTextBlock.Text = targetDescription;
            Height = Math.Min(1080, Math.Max(780, 340 + (columns.Count * 92)));
            _columns = columns;
            _enableLiveFiltering = enableLiveFiltering;

            foreach (string column in columns)
            {
                var label = new TextBlock
                {
                    Text = column,
                    Margin = new Thickness(0, 0, 0, 6),
                    Foreground = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#6B7280"))
                };

                var textBox = new TextBox
                {
                    Margin = new Thickness(0, 0, 0, 12),
                    Padding = new Thickness(8, 6, 8, 6),
                    Background = System.Windows.Media.Brushes.White,
                    BorderBrush = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E2E4EA"))
                };
                textBox.TextChanged += FilterTextBox_TextChanged;

                FieldsPanel.Children.Add(label);
                FieldsPanel.Children.Add(textBox);
                _fields.Add((column, textBox));
            }

            BindMissingItems(columns, missingItems);

            Loaded += (_, _) =>
            {
                if (_fields.Count > 0)
                {
                    _fields[0].TextBox.Focus();
                }
            };
        }

        private void BindMissingItems(
            IReadOnlyList<string> columns,
            IReadOnlyList<IReadOnlyDictionary<string, string>>? missingItems)
        {
            if (missingItems == null || missingItems.Count == 0)
            {
                _allMissingItems = new List<MissingItemOption>();
                MissingHeaderTextBlock.Text = "Unregistered Items";
                MissingItemsDataGrid.ItemsSource = null;
                MissingItemsDataGrid.IsEnabled = false;
                return;
            }

            _allMissingItems = missingItems
                .Select(item => new MissingItemOption
                {
                    ValuesByColumn = new Dictionary<string, string>(item, StringComparer.OrdinalIgnoreCase)
                })
                .ToList();

            MissingItemsDataGrid.IsEnabled = true;
            RefreshFilteredMissingItems();
        }

        private void RefreshFilteredMissingItems()
        {
            if (_allMissingItems.Count == 0)
            {
                MissingHeaderTextBlock.Text = "Unregistered Items";
                MissingItemsDataGrid.ItemsSource = null;
                return;
            }

            var filteredItems = _enableLiveFiltering
                ? _allMissingItems.Where(IsMatchForCurrentFilters).ToList()
                : _allMissingItems;

            MissingHeaderTextBlock.Text = $"Unregistered Items ({filteredItems.Count}/{_allMissingItems.Count})";
            MissingItemsDataGrid.ItemsSource = BuildGridTable(filteredItems).DefaultView;
        }

        private bool IsMatchForCurrentFilters(MissingItemOption item)
        {
            foreach (var field in _fields)
            {
                string filter = field.TextBox.Text.Trim();
                if (string.IsNullOrWhiteSpace(filter))
                {
                    continue;
                }

                string candidateValue = item.ValuesByColumn.TryGetValue(field.ColumnName, out string? value)
                    ? value ?? string.Empty
                    : string.Empty;

                if (candidateValue.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }
            }

            return true;
        }

        private DataTable BuildGridTable(IEnumerable<MissingItemOption> items)
        {
            var table = new DataTable();
            foreach (string column in _columns)
            {
                table.Columns.Add(column, typeof(string));
            }

            foreach (MissingItemOption item in items)
            {
                DataRow row = table.NewRow();
                foreach (string column in _columns)
                {
                    row[column] = item.ValuesByColumn.TryGetValue(column, out string? value)
                        ? value ?? string.Empty
                        : string.Empty;
                }

                table.Rows.Add(row);
            }

            return table;
        }

        private void MissingItemsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MissingItemsDataGrid.SelectedItem is not DataRowView selectedRow)
            {
                return;
            }

            _isUpdatingFieldsFromSelection = true;
            try
            {
                foreach (var field in _fields)
                {
                    field.TextBox.Text = selectedRow[field.ColumnName]?.ToString() ?? string.Empty;
                }
            }
            finally
            {
                _isUpdatingFieldsFromSelection = false;
            }
        }

        private void FilterTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingFieldsFromSelection)
            {
                return;
            }

            if (!_enableLiveFiltering)
            {
                return;
            }

            RefreshFilteredMissingItems();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            string? emptyColumn = _fields
                .Select(field => new { field.ColumnName, Value = field.TextBox.Text.Trim() })
                .FirstOrDefault(field => string.IsNullOrWhiteSpace(field.Value))
                ?.ColumnName;

            if (!string.IsNullOrWhiteSpace(emptyColumn))
            {
                MessageBox.Show($"{emptyColumn} 값을 입력해주세요.", "Input Required", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
