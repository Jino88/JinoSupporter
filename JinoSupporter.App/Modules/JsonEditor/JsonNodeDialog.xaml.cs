using System;
using System.Windows;
using System.Windows.Controls;

namespace WorkbenchHost.Modules.JsonEditor;

public partial class JsonNodeDialog : Window
{
    public string NodeKey { get; private set; } = string.Empty;
    public JsonTreeNodeKind NodeKind { get; private set; } = JsonTreeNodeKind.String;
    public string NodeValue { get; private set; } = string.Empty;

    public JsonNodeDialog(bool requiresKey)
    {
        InitializeComponent();

        if (!requiresKey)
        {
            KeyTextBox.Visibility = Visibility.Collapsed;
        }
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (KeyTextBox.Visibility == Visibility.Visible && string.IsNullOrWhiteSpace(KeyTextBox.Text))
        {
            MessageBox.Show(this, "Key is required.", "JsonEditor", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (KindComboBox.SelectedItem is not ComboBoxItem item ||
            !Enum.TryParse(item.Tag?.ToString(), ignoreCase: true, out JsonTreeNodeKind kind))
        {
            MessageBox.Show(this, "Type selection is invalid.", "JsonEditor", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        NodeKey = KeyTextBox.Text.Trim();
        NodeKind = kind;
        NodeValue = ValueTextBox.Text;
        DialogResult = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
