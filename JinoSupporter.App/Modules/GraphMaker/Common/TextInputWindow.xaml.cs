using System.Windows;

namespace GraphMaker;

public partial class TextInputWindow : Window
{
    public string InputText => InputTextBox.Text.Trim();

    public TextInputWindow(string title, string prompt, string initialValue)
    {
        InitializeComponent();
        Title = title;
        PromptTextBlock.Text = prompt;
        InputTextBox.Text = initialValue;
        Loaded += (_, _) =>
        {
            InputTextBox.Focus();
            InputTextBox.SelectAll();
        };
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
