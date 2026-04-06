using System;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;

namespace GraphMaker;

public class PreviewColorChoice
{
    public string Name { get; init; } = string.Empty;
    public System.Windows.Media.Color Color { get; init; }
}

public static class PreviewGraphViewBase
{
    public static void InitializeDefaultColorOptions(ComboBox comboBox, IList<PreviewColorChoice> colorChoices)
    {
        colorChoices.Clear();
        colorChoices.Add(new PreviewColorChoice { Name = "Black", Color = System.Windows.Media.Colors.Black });
        colorChoices.Add(new PreviewColorChoice { Name = "Blue", Color = System.Windows.Media.Colors.DodgerBlue });
        colorChoices.Add(new PreviewColorChoice { Name = "Red", Color = System.Windows.Media.Colors.IndianRed });
        colorChoices.Add(new PreviewColorChoice { Name = "Green", Color = System.Windows.Media.Colors.SeaGreen });
        colorChoices.Add(new PreviewColorChoice { Name = "Orange", Color = System.Windows.Media.Colors.DarkOrange });
        colorChoices.Add(new PreviewColorChoice { Name = "Purple", Color = System.Windows.Media.Colors.MediumPurple });
        colorChoices.Add(new PreviewColorChoice { Name = "Gray", Color = System.Windows.Media.Colors.Gray });

        comboBox.ItemsSource = colorChoices;
        comboBox.DisplayMemberPath = "Name";
        comboBox.SelectedIndex = 0;
    }

    public static void BindColorComboBox(ComboBox comboBox, IList<PreviewColorChoice> colorChoices, int selectedIndex = 0)
    {
        comboBox.ItemsSource = colorChoices;
        comboBox.DisplayMemberPath = "Name";
        comboBox.SelectedIndex = Math.Max(0, Math.Min(selectedIndex, Math.Max(0, colorChoices.Count - 1)));
    }

    public static void WirePreviewGrid(
        InteractiveDataGrid grid,
        Action<IReadOnlyList<string>> onFilesDropped,
        Action onDeleteRow,
        Action? onCellCommitted = null,
        Action? onRowCommitted = null,
        Action<DataGridColumnHeader>? onColumnRenameRequested = null)
    {
        grid.FilesDropped += (_, files) => onFilesDropped(files);
        grid.DeleteRowRequested += (_, _) => onDeleteRow();

        if (onCellCommitted != null)
        {
            grid.CellValueCommitted += (_, _) => onCellCommitted();
        }

        if (onRowCommitted != null)
        {
            grid.RowValueCommitted += (_, _) => onRowCommitted();
        }

        if (onColumnRenameRequested != null)
        {
            grid.ColumnRenameRequested += (_, header) => onColumnRenameRequested(header);
        }
    }

    public static void ApplyPreviewSummary(
        Run fileNameRun,
        Run rowCountRun,
        Run columnCountRun,
        TextBlock statusTextBlock,
        string fileName,
        int rowCount,
        int columnCount,
        string statusMessage)
    {
        fileNameRun.Text = fileName;
        rowCountRun.Text = rowCount.ToString("N0");
        columnCountRun.Text = columnCount.ToString();
        statusTextBlock.Text = statusMessage;
    }

    public static void ResetPreviewSummary(
        Run fileNameRun,
        Run rowCountRun,
        Run columnCountRun,
        TextBlock statusTextBlock,
        string emptyFileText,
        string emptyStatusText)
    {
        fileNameRun.Text = emptyFileText;
        rowCountRun.Text = "0";
        columnCountRun.Text = "0";
        statusTextBlock.Text = emptyStatusText;
    }
}
