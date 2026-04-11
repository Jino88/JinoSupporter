using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Web.WebView2.Core;
using WorkbenchHost.Infrastructure;

namespace JinoSupporter.App.Modules.Home;

public partial class HomeView : UserControl
{
    private const string ChatGptUsageUrl = "https://chatgpt.com/codex/cloud/settings/usage";
    private const string ClaudeUsageUrl = "https://claude.ai/settings/usage";

    private readonly UsageDashboardParser _chatGptParser = new(
        "ChatGPT",
        new UsageDashboardParser.CardDefinition("5시간 사용 한도", "5-hour usage limit", "5시간 사용 한도"),
        new UsageDashboardParser.CardDefinition("주간 사용 한도", "weekly usage limit", "주간 사용 한도"),
        new UsageDashboardParser.CardDefinition("코드 검토", "code review", "코드 검토"),
        new UsageDashboardParser.CardDefinition("남은 크레딧", "remaining credits", "남은 크레딧"));

    private readonly UsageDashboardParser _claudeParser = new(
        "Claude",
        new UsageDashboardParser.CardDefinition("현재 세션", "현재 세션", "current session"),
        new UsageDashboardParser.CardDefinition("주간 한도", "주간 한도", "weekly limit"),
        new UsageDashboardParser.CardDefinition("추가 사용량", "추가 사용량", "additional usage"),
        new UsageDashboardParser.CardDefinition("월간 지출 한도", "월간 지출 한도", "monthly spending limit"));

    private bool _chatGptBrowserReady;
    private bool _claudeBrowserReady;
    private bool _chatGptAutoLoginAttempted;
    private bool _claudeAutoLoginAttempted;

    public HomeView()
    {
        InitializeComponent();
        ApplyChatGptSnapshot(_chatGptParser.CreateDefaultSnapshot());
        ApplyClaudeSnapshot(_claudeParser.CreateDefaultSnapshot());
        Loaded += HomeView_Loaded;
    }

    private async void HomeView_Loaded(object sender, RoutedEventArgs e)
    {
        Loaded -= HomeView_Loaded;
        RememberSessionCheckBox.IsChecked = WorkbenchSettingsStore.IsCodexAutoLoginEnabled();
        await InitializeBrowsersAsync();
        await RefreshAllAsync(allowInteractiveLogin: RememberSessionCheckBox.IsChecked == true);
    }

    private async void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        await RefreshAllAsync(allowInteractiveLogin: true);
    }

    private async void ChatGptLoginButton_Click(object sender, RoutedEventArgs e)
    {
        await PromptLoginAsync("ChatGPT", "ChatGpt", ChatGptUsageUrl);
        await RefreshChatGptAsync(false);
    }

    private async void ClaudeLoginButton_Click(object sender, RoutedEventArgs e)
    {
        await PromptLoginAsync("Claude", "Claude", ClaudeUsageUrl);
        await RefreshClaudeAsync(false);
    }

    private void RememberSessionCheckBox_Checked(object sender, RoutedEventArgs e)
    {
        WorkbenchSettingsStore.SetCodexAutoLoginEnabled(true);
    }

    private void RememberSessionCheckBox_Unchecked(object sender, RoutedEventArgs e)
    {
        WorkbenchSettingsStore.SetCodexAutoLoginEnabled(false);
    }

    private async Task InitializeBrowsersAsync()
    {
        if (!_chatGptBrowserReady)
        {
            CoreWebView2Environment environment = await ProviderWebViewSession.GetEnvironmentAsync("ChatGpt");
            await ChatGptHiddenBrowser.EnsureCoreWebView2Async(environment);
            ChatGptHiddenBrowser.CoreWebView2.Settings.IsStatusBarEnabled = false;
            _chatGptBrowserReady = true;
        }

        if (!_claudeBrowserReady)
        {
            CoreWebView2Environment environment = await ProviderWebViewSession.GetEnvironmentAsync("Claude");
            await ClaudeHiddenBrowser.EnsureCoreWebView2Async(environment);
            ClaudeHiddenBrowser.CoreWebView2.Settings.IsStatusBarEnabled = false;
            _claudeBrowserReady = true;
        }
    }

    private async Task RefreshAllAsync(bool allowInteractiveLogin)
    {
        await Task.WhenAll(
            RefreshChatGptAsync(allowInteractiveLogin),
            RefreshClaudeAsync(allowInteractiveLogin));
    }

    private async Task RefreshChatGptAsync(bool allowInteractiveLogin)
    {
        ChatGptStatusTextBlock.Text = "로그인된 브라우저 세션에서 ChatGPT 사용량을 읽는 중입니다.";
        CodexUsageSnapshot snapshot = await ReadUsageFromBrowserAsync(
            ChatGptHiddenBrowser,
            ChatGptUsageUrl,
            _chatGptParser,
            LooksLikeChatGptUsageDashboardText);

        if (!snapshot.IsAuthenticated && allowInteractiveLogin && !_chatGptAutoLoginAttempted && RememberSessionCheckBox.IsChecked == true)
        {
            _chatGptAutoLoginAttempted = true;
            if (await PromptLoginAsync("ChatGPT", "ChatGpt", ChatGptUsageUrl))
            {
                snapshot = await ReadUsageFromBrowserAsync(ChatGptHiddenBrowser, ChatGptUsageUrl, _chatGptParser, LooksLikeChatGptUsageDashboardText);
            }
        }
        else if (snapshot.IsAuthenticated)
        {
            _chatGptAutoLoginAttempted = false;
        }

        ApplyChatGptSnapshot(snapshot);
    }

    private async Task RefreshClaudeAsync(bool allowInteractiveLogin)
    {
        ClaudeStatusTextBlock.Text = "로그인된 브라우저 세션에서 Claude 사용량을 읽는 중입니다.";
        CodexUsageSnapshot snapshot = await ReadUsageFromBrowserAsync(
            ClaudeHiddenBrowser,
            ClaudeUsageUrl,
            _claudeParser,
            LooksLikeClaudeUsageDashboardText);

        if (!snapshot.IsAuthenticated && allowInteractiveLogin && !_claudeAutoLoginAttempted && RememberSessionCheckBox.IsChecked == true)
        {
            _claudeAutoLoginAttempted = true;
            if (await PromptLoginAsync("Claude", "Claude", ClaudeUsageUrl))
            {
                snapshot = await ReadUsageFromBrowserAsync(ClaudeHiddenBrowser, ClaudeUsageUrl, _claudeParser, LooksLikeClaudeUsageDashboardText);
            }
        }
        else if (snapshot.IsAuthenticated)
        {
            _claudeAutoLoginAttempted = false;
        }

        NormalizeClaudeSnapshot(snapshot);
        ApplyClaudeSnapshot(snapshot);
    }

    private Task<bool> PromptLoginAsync(string title, string providerKey, string usageUrl)
    {
        UsageLoginWindow loginWindow = new(title, providerKey, usageUrl)
        {
            Owner = Window.GetWindow(this)
        };

        bool? result = loginWindow.ShowDialog();
        return Task.FromResult(result == true && !string.IsNullOrWhiteSpace(loginWindow.CookieHeader));
    }

    private static async Task<CodexUsageSnapshot> ReadUsageFromBrowserAsync(
        Microsoft.Web.WebView2.Wpf.WebView2 browser,
        string usageUrl,
        UsageDashboardParser parser,
        Func<string, bool> readyPredicate)
    {
        if (browser.CoreWebView2 is null)
        {
            return parser.CreateDefaultSnapshot();
        }

        TaskCompletionSource<bool> navigationCompletion = new(TaskCreationOptions.RunContinuationsAsynchronously);

        void HandleNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs args)
        {
            browser.CoreWebView2.NavigationCompleted -= HandleNavigationCompleted;
            navigationCompletion.TrySetResult(args.IsSuccess);
        }

        browser.CoreWebView2.NavigationCompleted += HandleNavigationCompleted;
        browser.CoreWebView2.Navigate(usageUrl);

        if (!await navigationCompletion.Task)
        {
            CodexUsageSnapshot failedSnapshot = parser.CreateDefaultSnapshot();
            failedSnapshot.StatusMessage = "사용량 페이지를 여는 데 실패했습니다.";
            return failedSnapshot;
        }

        string? pageText = await WaitForUsageTextAsync(browser, readyPredicate);
        return parser.ParseUsageText(pageText);
    }

    private static async Task<string?> WaitForUsageTextAsync(Microsoft.Web.WebView2.Wpf.WebView2 browser, Func<string, bool> readyPredicate)
    {
        const string script = """
            (() => {
              const text = document.body ? (document.body.innerText || '') : '';
              return text.trim();
            })();
            """;

        for (int attempt = 0; attempt < 80; attempt++)
        {
            string scriptResult = await browser.CoreWebView2.ExecuteScriptAsync(script);
            string? pageText = JsonSerializer.Deserialize<string>(scriptResult);
            if (!string.IsNullOrWhiteSpace(pageText) && readyPredicate(pageText.Trim()))
            {
                return pageText.Trim();
            }

            await Task.Delay(500);
        }

        string fallbackResult = await browser.CoreWebView2.ExecuteScriptAsync(script);
        return JsonSerializer.Deserialize<string>(fallbackResult);
    }

    private static bool LooksLikeChatGptUsageDashboardText(string text)
    {
        return text.Contains("남은 크레딧", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("remaining credits", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("5시간 사용 한도", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("5-hour usage limit", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeClaudeUsageDashboardText(string text)
    {
        return (text.Contains("현재 세션", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("current session", StringComparison.OrdinalIgnoreCase)) &&
               (text.Contains("주간 한도", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("weekly limit", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("추가 사용량", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("monthly spending limit", StringComparison.OrdinalIgnoreCase));
    }

    private void ApplyChatGptSnapshot(CodexUsageSnapshot snapshot)
    {
        ChatGptStatusTextBlock.Text = snapshot.StatusMessage;
        NormalizeChatGptSnapshot(snapshot);
        ApplyCard(ChatGptFiveHourValueTextBlock, ChatGptFiveHourDetailTextBlock, snapshot.Cards, 0);
        ApplyCard(ChatGptWeeklyValueTextBlock, ChatGptWeeklyDetailTextBlock, snapshot.Cards, 1);
        ApplyCard(ChatGptReviewValueTextBlock, ChatGptReviewDetailTextBlock, snapshot.Cards, 2);
        ApplyCard(ChatGptCreditsValueTextBlock, ChatGptCreditsDetailTextBlock, snapshot.Cards, 3);
    }

    private void ApplyClaudeSnapshot(CodexUsageSnapshot snapshot)
    {
        ClaudeStatusTextBlock.Text = snapshot.StatusMessage;
        ApplyCard(ClaudeValue1TextBlock, ClaudeDetail1TextBlock, snapshot.Cards, 0);
        ApplyCard(ClaudeValue2TextBlock, ClaudeDetail2TextBlock, snapshot.Cards, 1);
        ApplyCard(ClaudeValue3TextBlock, ClaudeDetail3TextBlock, snapshot.Cards, 2);
        ApplyCard(ClaudeValue4TextBlock, ClaudeDetail4TextBlock, snapshot.Cards, 3);
    }

    private static void NormalizeClaudeSnapshot(CodexUsageSnapshot snapshot)
    {
        if (!snapshot.IsAuthenticated || string.IsNullOrWhiteSpace(snapshot.DebugText) || snapshot.Cards.Count < 4)
        {
            return;
        }

        string[] lines = snapshot.DebugText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

        ApplyClaudeCard(lines, "현재 세션", snapshot.Cards[0], includeBalanceDetail: false);
        ApplyClaudeCard(lines, "주간 한도", snapshot.Cards[1], includeBalanceDetail: false);
        ApplyClaudeCard(lines, "추가 사용량", snapshot.Cards[2], includeBalanceDetail: false);
        ApplyClaudeCard(lines, "월간 지출 한도", snapshot.Cards[3], includeBalanceDetail: true);

        snapshot.Cards[0].Value = ConvertPercentUsedToRemaining(snapshot.Cards[0].Value);
        snapshot.Cards[0].Detail = string.Empty;
        snapshot.Cards[1].Value = ConvertPercentUsedToRemaining(snapshot.Cards[1].Value);

        string? monthlyLimit = FindValueNearLine(lines, "월간 지출 한도");
        if (!string.IsNullOrWhiteSpace(monthlyLimit))
        {
            snapshot.Cards[3].Value = monthlyLimit;
        }
    }

    private static void NormalizeChatGptSnapshot(CodexUsageSnapshot snapshot)
    {
        if (!snapshot.IsAuthenticated)
        {
            return;
        }

        foreach (CodexUsageCard card in snapshot.Cards)
        {
            if (card.Value.Contains("%", StringComparison.OrdinalIgnoreCase) &&
                !card.Value.Contains("남음", StringComparison.OrdinalIgnoreCase))
            {
                card.Value = $"{card.Value} 남음";
            }
        }
    }

    private static void ApplyClaudeCard(string[] lines, string title, CodexUsageCard card, bool includeBalanceDetail)
    {
        int titleIndex = FindLineIndex(lines, title);
        if (titleIndex < 0)
        {
            return;
        }

        string? value = null;
        string? detail = null;

        for (int i = titleIndex + 1; i < lines.Length && i <= titleIndex + 8; i++)
        {
            string line = lines[i].Trim();
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            if (LooksLikeClaudeSectionTitle(line))
            {
                break;
            }

            if (value is null && LooksLikeClaudeValue(line))
            {
                value = NormalizeClaudeValue(line);
                continue;
            }

            if (detail is null && LooksLikeClaudeDetail(line))
            {
                detail = line;
            }
        }

        if (includeBalanceDetail)
        {
            string? balanceValue = FindValueNearLine(lines, "현재 잔액");
            if (!string.IsNullOrWhiteSpace(balanceValue))
            {
                detail = $"{balanceValue} 남음";
            }
        }

        if (!string.IsNullOrWhiteSpace(value))
        {
            card.Value = value;
        }

        if (!string.IsNullOrWhiteSpace(detail))
        {
            card.Detail = detail;
        }
    }

    private static int FindLineIndex(string[] lines, string target)
    {
        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i].Contains(target, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }

    private static string? FindValueNearLine(string[] lines, string title)
    {
        int index = FindLineIndex(lines, title);
        if (index < 0)
        {
            return null;
        }

        for (int i = Math.Max(0, index - 2); i <= Math.Min(lines.Length - 1, index + 2); i++)
        {
            string line = lines[i].Trim();
            if (LooksLikeClaudeValue(line))
            {
                return NormalizeClaudeValue(line);
            }
        }

        return null;
    }

    private static bool LooksLikeClaudeSectionTitle(string line)
    {
        return line.Equals("현재 세션", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("주간 한도", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("추가 사용량", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("월간 지출 한도", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("한도 조정", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("추가 사용량 구매", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeClaudeValue(string line)
    {
        return line.Contains("% 사용됨", StringComparison.OrdinalIgnoreCase) ||
               line.StartsWith("US$", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeClaudeDetail(string line)
    {
        return line.Contains("재설정", StringComparison.OrdinalIgnoreCase) ||
               line.Contains("마지막 업데이트", StringComparison.OrdinalIgnoreCase) ||
               line.Contains("현재 잔액", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeClaudeValue(string line)
    {
        return line.Trim().Replace(" 사용됨", string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    private static string ConvertPercentUsedToRemaining(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        string raw = value.Replace("%", string.Empty).Replace("남음", string.Empty).Trim();
        if (!int.TryParse(raw, out int usedPercent))
        {
            return value;
        }

        int remainingPercent = Math.Clamp(100 - usedPercent, 0, 100);
        return $"{remainingPercent}% 남음";
    }

    private static void ApplyCard(TextBlock valueBlock, TextBlock detailBlock, IReadOnlyList<CodexUsageCard> cards, int index)
    {
        if (index >= cards.Count)
        {
            valueBlock.Text = "-";
            detailBlock.Text = string.Empty;
            return;
        }

        valueBlock.Text = cards[index].Value;
        detailBlock.Text = cards[index].Detail;
    }
}
