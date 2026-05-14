using System.Windows;

namespace BmesNgRateStandalone;

public partial class MainWindow : Window
{
    private readonly StandaloneUpdateService _updateService = new();

    public MainWindow()
    {
        InitializeComponent();
        var app = (App)Application.Current;
        blazorWebView.Services = app.Services;
        Loaded += OnLoaded;
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        Loaded -= OnLoaded;
        await Task.Delay(1000);
        await _updateService.CheckAndPromptAsync(this);
    }

    private async void UpdateButton_Click(object sender, RoutedEventArgs e)
    {
        UpdateButton.IsEnabled = false;
        object? previousContent = UpdateButton.Content;
        UpdateButton.Content = "Checking...";

        try
        {
            await _updateService.CheckAndPromptAsync(this, notifyWhenCurrent: true);
        }
        finally
        {
            UpdateButton.Content = previousContent;
            UpdateButton.IsEnabled = true;
        }
    }
}
