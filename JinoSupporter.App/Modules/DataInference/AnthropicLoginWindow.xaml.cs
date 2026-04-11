using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace JinoSupporter.App.Modules.DataInference;

public partial class AnthropicLoginWindow : Window
{
    public string? SessionKey { get; private set; }

    public AnthropicLoginWindow()
    {
        InitializeComponent();
        Loaded += async (_, _) => await InitWebViewAsync();
    }

    private async Task InitWebViewAsync()
    {
        await BrowserView.EnsureCoreWebView2Async();

        BrowserView.CoreWebView2.NavigationCompleted += async (_, args) =>
        {
            if (!args.IsSuccess) return;
            await TryAutoExtractAsync();
        };
    }

    /// <summary>Auto-detect: extract cookies when arriving at platform.claude.com (outside Google login pages)</summary>
    private async Task TryAutoExtractAsync()
    {
        try
        {
            string url = BrowserView.Source?.ToString() ?? string.Empty;
            if (string.IsNullOrEmpty(url)) return;

            // Ignore while Google auth is in progress
            if (url.Contains("google.com") || url.Contains("accounts.google")) return;

            // Proceed only when on platform.claude.com and not on the login page
            if (!url.Contains("platform.claude.com")) return;
            if (url.TrimEnd('/').EndsWith("/login")) return;

            // Wait briefly for cookies to stabilize
            await Task.Delay(1500);

            string? cookie = await CollectCookiesAsync();
            if (!string.IsNullOrWhiteSpace(cookie))
            {
                SessionKey = cookie;
                Dispatcher.Invoke(Close);
            }
        }
        catch { /* ignore */ }
    }

    /// <summary>Collects cookies from both platform.claude.com and claude.ai and returns them as a Cookie header string</summary>
    private async Task<string?> CollectCookiesAsync()
    {
        try
        {
            var list1 = await BrowserView.CoreWebView2.CookieManager
                .GetCookiesAsync("https://platform.claude.com");
            var list2 = await BrowserView.CoreWebView2.CookieManager
                .GetCookiesAsync("https://claude.ai");

            var merged = list1
                .Concat(list2)
                .GroupBy(c => c.Name)
                .Select(g => g.First())
                .ToList();

            if (merged.Count == 0) return null;

            return string.Join("; ", merged.Select(c => $"{c.Name}={c.Value}"));
        }
        catch { return null; }
    }

    // Manual save button: click when window doesn't close automatically after login
    private async void SaveSessionButton_Click(object sender, RoutedEventArgs e)
    {
        SaveSessionButton.IsEnabled = false;
        SaveSessionButton.Content = "Saving...";

        string? cookie = await CollectCookiesAsync();
        if (!string.IsNullOrWhiteSpace(cookie))
        {
            SessionKey = cookie;
            Close();
        }
        else
        {
            MessageBox.Show(
                "Could not retrieve cookies.\nPlease make sure you are logged in to platform.claude.com.",
                "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            SaveSessionButton.IsEnabled = true;
            SaveSessionButton.Content = "Save Session";
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e) => Close();
}
