using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Text;

namespace GraphMaker;

public sealed class InteractiveDataGrid : DataGrid
{
    public event EventHandler<IReadOnlyList<string>>? FilesDropped;
    public event EventHandler? DeleteRowRequested;
    public event EventHandler? CellValueCommitted;
    public event EventHandler? RowValueCommitted;
    public event EventHandler<DataGridColumnHeader>? ColumnRenameRequested;

    public InteractiveDataGrid()
    {
        AutoGenerateColumns = true;
        IsReadOnly = false;
        AllowDrop = true;
        SelectionMode = DataGridSelectionMode.Extended;
        SelectionUnit = DataGridSelectionUnit.CellOrRowHeader;
        CanUserAddRows = true;
        CanUserDeleteRows = false;
        CanUserReorderColumns = false;
        CanUserResizeRows = true;
        CanUserSortColumns = false;
        IsTabStop = true;
        ScrollViewer.SetHorizontalScrollBarVisibility(this, ScrollBarVisibility.Auto);
        ScrollViewer.SetVerticalScrollBarVisibility(this, ScrollBarVisibility.Auto);
        EnableRowVirtualization = true;
        EnableColumnVirtualization = true;
        VirtualizingPanel.SetIsVirtualizing(this, true);
        VirtualizingPanel.SetVirtualizationMode(this, VirtualizationMode.Recycling);

        PreviewMouseRightButtonDown += HandlePreviewMouseRightButtonDown;
        DragEnter += HandleDragEnter;
        DragLeave += HandleDragLeave;
        DragOver += HandleDragOver;
        Drop += HandleDrop;
        CellEditEnding += HandleCellEditEnding;
        RowEditEnding += HandleRowEditEnding;
        PreviewKeyDown += HandlePreviewKeyDown;

        ColumnHeaderStyle = BuildColumnHeaderStyle();
        ContextMenu = BuildContextMenu();
    }

    public bool TryDeleteSelectedRow(DataTable? table)
    {
        if (table == null || SelectedItem is not DataRowView rowView)
        {
            return false;
        }

        int rowIndex = table.Rows.IndexOf(rowView.Row);
        if (rowIndex < 0)
        {
            return false;
        }

        table.Rows.RemoveAt(rowIndex);
        return true;
    }

    private ContextMenu BuildContextMenu()
    {
        var contextMenu = new ContextMenu();
        var deleteMenuItem = new MenuItem { Header = "Delete Row" };
        deleteMenuItem.Click += (_, _) => DeleteRowRequested?.Invoke(this, EventArgs.Empty);
        contextMenu.Items.Add(deleteMenuItem);
        return contextMenu;
    }

    private Style BuildColumnHeaderStyle()
    {
        var style = new Style(typeof(DataGridColumnHeader));
        style.Setters.Add(new EventSetter(Control.MouseDoubleClickEvent, new MouseButtonEventHandler(HandleColumnHeaderMouseDoubleClick)));

        var contextMenu = new ContextMenu();
        var renameItem = new MenuItem { Header = "Rename Column" };
        renameItem.Click += HandleRenameMenuItemClick;
        contextMenu.Items.Add(renameItem);

        style.Setters.Add(new Setter(ContextMenuService.ContextMenuProperty, contextMenu));
        return style;
    }

    private void HandlePreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        DependencyObject? source = e.OriginalSource as DependencyObject;
        while (source != null && source is not DataGridRow)
        {
            source = VisualTreeHelper.GetParent(source);
        }

        if (source is DataGridRow row)
        {
            row.IsSelected = true;
            SelectedItem = row.Item;
        }
    }

    private void HandleDragEnter(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            return;
        }

        e.Effects = DragDropEffects.Copy;
    }

    private void HandleDragLeave(object sender, DragEventArgs e)
    {
    }

    private void HandleDragOver(object sender, DragEventArgs e)
    {
        e.Handled = true;
    }

    private void HandleDrop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            return;
        }

        if (e.Data.GetData(DataFormats.FileDrop) is not string[] files || files.Length == 0)
        {
            return;
        }

        FilesDropped?.Invoke(this, files.Where(path => !string.IsNullOrWhiteSpace(path)).ToArray());
    }

    private void HandleCellEditEnding(object? sender, DataGridCellEditEndingEventArgs e)
    {
        if (e.EditAction == DataGridEditAction.Commit)
        {
            CellValueCommitted?.Invoke(this, EventArgs.Empty);
        }
    }

    private void HandleRowEditEnding(object? sender, DataGridRowEditEndingEventArgs e)
    {
        if (e.EditAction != DataGridEditAction.Commit)
        {
            return;
        }

        Dispatcher.BeginInvoke(new Action(() =>
        {
            RowValueCommitted?.Invoke(this, EventArgs.Empty);
        }), System.Windows.Threading.DispatcherPriority.Background);
    }

    private void HandlePreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (Keyboard.Modifiers != ModifierKeys.Control)
        {
            return;
        }

        if (e.Key == Key.C)
        {
            CopySelectedCellsToClipboard();
            e.Handled = true;
        }
        else if (e.Key == Key.V)
        {
            if (PasteClipboardIntoGrid())
            {
                e.Handled = true;
            }
        }
    }

    private void CopySelectedCellsToClipboard()
    {
        if (SelectedCells.Count == 0)
        {
            return;
        }

        var orderedCells = SelectedCells
            .Where(cell => cell.Item != null && cell.Column != null)
            .Select(cell => new
            {
                RowIndex = Items.IndexOf(cell.Item),
                ColumnIndex = Columns.IndexOf(cell.Column),
                Value = ExtractCellText(cell.Item, cell.Column)
            })
            .Where(cell => cell.RowIndex >= 0 && cell.ColumnIndex >= 0)
            .OrderBy(cell => cell.RowIndex)
            .ThenBy(cell => cell.ColumnIndex)
            .ToList();

        if (orderedCells.Count == 0)
        {
            return;
        }

        var rowGroups = orderedCells.GroupBy(cell => cell.RowIndex).OrderBy(group => group.Key);
        var sb = new StringBuilder();
        bool firstRow = true;
        foreach (var rowGroup in rowGroups)
        {
            if (!firstRow)
            {
                sb.AppendLine();
            }

            firstRow = false;
            bool firstCell = true;
            foreach (var cell in rowGroup.OrderBy(cell => cell.ColumnIndex))
            {
                if (!firstCell)
                {
                    sb.Append('\t');
                }

                firstCell = false;
                sb.Append(cell.Value);
            }
        }

        if (sb.Length > 0)
        {
            Clipboard.SetText(sb.ToString());
        }
    }

    private bool PasteClipboardIntoGrid()
    {
        if (ItemsSource is not DataView dataView)
        {
            return false;
        }

        string text = Clipboard.ContainsText() ? Clipboard.GetText() : string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string[] rows = text.Replace("\r\n", "\n").Replace('\r', '\n')
            .Split('\n', StringSplitOptions.RemoveEmptyEntries);
        if (rows.Length == 0)
        {
            return false;
        }

        int startRowIndex = ResolveStartRowIndex(dataView);
        int startColumnIndex = ResolveStartColumnIndex();
        if (startColumnIndex < 0)
        {
            return false;
        }

        DataTable table = dataView.Table;
        int requiredColumnCount = startColumnIndex + rows.Max(row => row.Split('\t').Length);
        EnsureTableHasColumns(table, requiredColumnCount);
        Items.Refresh();

        for (int r = 0; r < rows.Length; r++)
        {
            string[] cells = rows[r].Split('\t');
            int targetRowIndex = startRowIndex + r;
            while (targetRowIndex >= table.Rows.Count)
            {
                table.Rows.Add(table.NewRow());
            }

            DataRow row = table.Rows[targetRowIndex];
            for (int c = 0; c < cells.Length; c++)
            {
                int targetColumnIndex = startColumnIndex + c;
                if (targetColumnIndex < 0 || targetColumnIndex >= Columns.Count)
                {
                    continue;
                }

                DataGridColumn column = Columns[targetColumnIndex];
                string columnName = column.Header?.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(columnName) || !table.Columns.Contains(columnName))
                {
                    continue;
                }

                row[columnName] = cells[c];
            }
        }

        Items.Refresh();
        CellValueCommitted?.Invoke(this, EventArgs.Empty);
        RowValueCommitted?.Invoke(this, EventArgs.Empty);
        return true;
    }

    private int ResolveStartRowIndex(DataView dataView)
    {
        if (CurrentCell.Item != null)
        {
            int currentIndex = Items.IndexOf(CurrentCell.Item);
            if (currentIndex >= 0)
            {
                return currentIndex;
            }
        }

        if (SelectedCells.Count > 0 && SelectedCells[0].Item != null)
        {
            int selectedIndex = Items.IndexOf(SelectedCells[0].Item);
            if (selectedIndex >= 0)
            {
                return selectedIndex;
            }
        }

        return dataView.Count > 0 ? 0 : 0;
    }

    private int ResolveStartColumnIndex()
    {
        if (CurrentCell.Column != null)
        {
            int currentIndex = Columns.IndexOf(CurrentCell.Column);
            if (currentIndex >= 0)
            {
                return currentIndex;
            }
        }

        if (SelectedCells.Count > 0 && SelectedCells[0].Column != null)
        {
            int selectedIndex = Columns.IndexOf(SelectedCells[0].Column);
            if (selectedIndex >= 0)
            {
                return selectedIndex;
            }
        }

        return Columns.Count > 0 ? 0 : -1;
    }

    private static string ExtractCellText(object item, DataGridColumn column)
    {
        string columnName = column.Header?.ToString() ?? string.Empty;
        if (item is DataRowView rowView && !string.IsNullOrWhiteSpace(columnName) && rowView.Row.Table.Columns.Contains(columnName))
        {
            return rowView.Row[columnName]?.ToString() ?? string.Empty;
        }

        return string.Empty;
    }

    private void EnsureTableHasColumns(DataTable table, int requiredColumnCount)
    {
        while (table.Columns.Count < requiredColumnCount)
        {
            string columnName = BuildNextColumnName(table.Columns.Count + 1, table);
            table.Columns.Add(columnName, typeof(string));
        }
    }

    private static string BuildNextColumnName(int sequence, DataTable table)
    {
        string baseName = $"Column {sequence}";
        if (!table.Columns.Contains(baseName))
        {
            return baseName;
        }

        int suffix = 2;
        while (table.Columns.Contains($"{baseName} ({suffix})"))
        {
            suffix++;
        }

        return $"{baseName} ({suffix})";
    }

    private void HandleColumnHeaderMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGridColumnHeader header)
        {
            return;
        }

        ColumnRenameRequested?.Invoke(this, header);
        e.Handled = true;
    }

    private void HandleRenameMenuItemClick(object? sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem menuItem ||
            menuItem.Parent is not ContextMenu contextMenu ||
            contextMenu.PlacementTarget is not DataGridColumnHeader header)
        {
            return;
        }

        ColumnRenameRequested?.Invoke(this, header);
    }
}
