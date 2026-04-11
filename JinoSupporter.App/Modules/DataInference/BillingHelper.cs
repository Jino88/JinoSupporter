using System;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace JinoSupporter.App.Modules.DataInference;

/// <summary>
/// Shared helper that reads the balance from the platform.claude.com billing page.
/// Because it is a Next.js SPA, it renders the page via WebView2 and then extracts text with JS.
/// </summary>
internal static class BillingHelper
{
    public static async Task<string?> FetchBalanceAsync(string cookieHeader, Dispatcher dispatcher)
    {
        var tcs = new TaskCompletionSource<string?>(TaskCreationOptions.RunContinuationsAsynchronously);

        await dispatcher.InvokeAsync(async () =>
        {
            Window? helperWin = null;
            try
            {
                helperWin = new Window
                {
                    Width = 1, Height = 1,
                    Left = -9999, Top = -9999,
                    ShowInTaskbar = false,
                    WindowStyle = WindowStyle.None,
                    AllowsTransparency = true,
                    Opacity = 0,
                    ShowActivated = false
                };

                var wv = new Microsoft.Web.WebView2.Wpf.WebView2();
                helperWin.Content = wv;
                helperWin.Show();

                await wv.EnsureCoreWebView2Async();

                // Parse the saved cookie string and configure the WebView2 cookie manager
                foreach (string pair in cookieHeader.Split(';'))
                {
                    int eq = pair.IndexOf('=');
                    if (eq < 0) continue;
                    string name  = pair[..eq].Trim();
                    string value = pair[(eq + 1)..].Trim();
                    if (string.IsNullOrEmpty(name)) continue;
                    try
                    {
                        foreach (string domain in new[] { "platform.claude.com", ".claude.ai" })
                        {
                            var c = wv.CoreWebView2.CookieManager.CreateCookie(name, value, domain, "/");
                            wv.CoreWebView2.CookieManager.AddOrUpdateCookie(c);
                        }
                    }
                    catch { /* ignore invalid cookies */ }
                }

                bool fired = false;
                wv.CoreWebView2.NavigationCompleted += async (_, args) =>
                {
                    if (fired) return;
                    fired = true;

                    if (!args.IsSuccess) { tcs.TrySetResult(null); return; }

                    // Wait for SPA rendering
                    await Task.Delay(3000);

                    try
                    {
                        string json = await wv.CoreWebView2.ExecuteScriptAsync(
                            "JSON.stringify(document.body ? document.body.innerText : '')");
                        string text = JsonSerializer.Deserialize<string>(json) ?? string.Empty;

                        var m = Regex.Match(text, @"US\$\s*[\d,]+\.?\d*|\$\s*[\d,]+\.?\d*");
                        tcs.TrySetResult(m.Success ? m.Value.Trim() : null);
                    }
                    catch { tcs.TrySetResult(null); }
                };

                wv.CoreWebView2.Navigate("https://platform.claude.com/settings/billing");
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
            finally
            {
                _ = tcs.Task.ContinueWith(_ => dispatcher.Invoke(() => helperWin?.Close()));
            }
        });

        return await tcs.Task.WaitAsync(TimeSpan.FromSeconds(30));
    }
}
