using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Web.WebView2.Core;

namespace JinoSupporter.App.Modules.Home;

public partial class UsageLoginWindow : Window
{
    private readonly string _providerKey;
    private readonly string _usageUrl;
    private bool _capturingSession;

    public string? CookieHeader { get; private set; }

    public UsageLoginWindow(string title, string providerKey, string usageUrl)
    {
        InitializeComponent();
        Title = $"{title} Login";
        HeaderTitleTextBlock.Text = $"{title} 로그인";
        HeaderDescriptionTextBlock.Text = $"이 창에서 {title}에 로그인한 뒤 'Use This Session'을 누르면 HOME 화면이 현재 세션을 재사용합니다.";
        _providerKey = providerKey;
        _usageUrl = usageUrl;
        Loaded += UsageLoginWindow_Loaded;
    }

    private async void UsageLoginWindow_Loaded(object sender, RoutedEventArgs e)
    {
        try
        {
            StatusTextBlock.Text = "브라우저 초기화 중입니다...";
            CoreWebView2Environment environment = await ProviderWebViewSession.GetEnvironmentAsync(_providerKey);
            await Browser.EnsureCoreWebView2Async(environment);
            Browser.CoreWebView2.Settings.IsStatusBarEnabled = false;
            Browser.CoreWebView2.NavigationCompleted += CoreWebView2_NavigationCompleted;
            Browser.Source = new Uri(_usageUrl);
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"로그인 브라우저 초기화 실패: {ex.Message}";
        }
    }

    private void OpenUsagePageButton_Click(object sender, RoutedEventArgs e)
    {
        Browser.Source = new Uri(_usageUrl);
    }

    private async void UseThisSessionButton_Click(object sender, RoutedEventArgs e)
    {
        await CaptureSessionAsync(autoClose: true);
    }

    private async Task CaptureSessionAsync(bool autoClose)
    {
        if (Browser.CoreWebView2 is null || _capturingSession)
        {
            return;
        }

        try
        {
            _capturingSession = true;
            StatusTextBlock.Text = "세션 쿠키를 읽는 중입니다...";

            Uri uri = new(_usageUrl);
            var cookies = await Browser.CoreWebView2.CookieManager.GetCookiesAsync($"{uri.Scheme}://{uri.Host}");
            string cookieHeader = string.Join(
                "; ",
                cookies
                    .Where(cookie => !string.IsNullOrWhiteSpace(cookie.Name) && !string.IsNullOrWhiteSpace(cookie.Value))
                    .Select(cookie => $"{cookie.Name}={cookie.Value}"));

            if (string.IsNullOrWhiteSpace(cookieHeader))
            {
                StatusTextBlock.Text = $"{uri.Host} 쿠키를 찾지 못했습니다. 로그인 후 다시 시도하세요.";
                return;
            }

            CookieHeader = cookieHeader;
            StatusTextBlock.Text = "로그인 세션을 확인했습니다.";

            if (autoClose)
            {
                DialogResult = true;
                Close();
            }
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"세션 쿠키 읽기 실패: {ex.Message}";
        }
        finally
        {
            _capturingSession = false;
        }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void CoreWebView2_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        string currentUri = Browser.Source?.ToString() ?? _usageUrl;
        StatusTextBlock.Text = e.IsSuccess
            ? $"현재 페이지: {currentUri} | 로그인 완료 후 'Use This Session'을 눌러주세요."
            : $"페이지 로드 실패: {currentUri}";
    }
}
