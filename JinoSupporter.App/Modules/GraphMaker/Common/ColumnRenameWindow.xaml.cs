using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace GraphMaker;

public partial class ColumnRenameWindow : Window
{
    private sealed class ColumnRenameItem
    {
        public string OriginalName { get; init; } = string.Empty;
        public string NewName { get; set; } = string.Empty;
    }

    private readonly ObservableCollection<ColumnRenameItem> _items = new();

    public IReadOnlyDictionary<string, string> RenamedColumns { get; private set; } =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    public ColumnRenameWindow(IEnumerable<string> columnNames)
    {
        InitializeComponent();
        ColumnDataGrid.ItemsSource = _items;

        foreach (string columnName in columnNames)
        {
            _items.Add(new ColumnRenameItem
            {
                OriginalName = columnName,
                NewName = columnName
            });
        }
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Dictionary<string, string> renamed = new(StringComparer.OrdinalIgnoreCase);
            HashSet<string> usedNames = new(StringComparer.OrdinalIgnoreCase);

            foreach (ColumnRenameItem item in _items)
            {
                string newName = string.IsNullOrWhiteSpace(item.NewName)
                    ? item.OriginalName
                    : item.NewName.Trim();

                if (!usedNames.Add(newName))
                {
                    throw new InvalidOperationException($"Duplicate column name '{newName}' is not allowed.");
                }

                renamed[item.OriginalName] = newName;
            }

            RenamedColumns = renamed;
            DialogResult = true;
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = ex.Message;
            MessageBox.Show(this, ex.Message, "Rename Columns", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
