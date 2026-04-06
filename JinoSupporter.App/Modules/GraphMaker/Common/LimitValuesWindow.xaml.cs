using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace GraphMaker;

public partial class LimitValuesWindow : Window
{
    private sealed class LimitValueItem
    {
        public string ColumnName { get; init; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    private readonly ObservableCollection<LimitValueItem> _items = new();

    public IReadOnlyDictionary<string, string> Values { get; private set; } =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    public LimitValuesWindow(IEnumerable<string> columnNames, IReadOnlyDictionary<string, string>? existingValues)
    {
        InitializeComponent();
        LimitDataGrid.ItemsSource = _items;

        foreach (string columnName in columnNames)
        {
            _items.Add(new LimitValueItem
            {
                ColumnName = columnName,
                Value = existingValues != null && existingValues.TryGetValue(columnName, out string? value)
                    ? value
                    : string.Empty
            });
        }
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        Values = _items.ToDictionary(
            item => item.ColumnName,
            item => item.Value?.Trim() ?? string.Empty,
            StringComparer.OrdinalIgnoreCase);

        DialogResult = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
