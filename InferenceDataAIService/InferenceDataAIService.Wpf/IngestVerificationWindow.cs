using System.IO;
using System.Net;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Navigation;

namespace InferenceDataAIService.Wpf;

internal sealed class IngestVerificationWindow : Window
{
    private readonly System.Windows.Controls.WebBrowser _browser = new();
    private readonly string _temporaryHtmlPath;

    internal IngestVerificationWindow(
        IngestWorkbookResult result,
        string html,
        string temporaryDirectory)
    {
        Directory.CreateDirectory(temporaryDirectory);
        _temporaryHtmlPath = Path.Combine(
            temporaryDirectory,
            $"InferenceDataAIService-ingest-verification-{Guid.NewGuid():N}.html");
        Title = $"처리 내용 검증 · {Path.GetFileName(result.SourcePath)}";
        Width = 1280;
        Height = 840;
        MinWidth = 900;
        MinHeight = 620;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = Brush("#202020");
        Foreground = Brush("#D6D6D6");

        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition());

        var header = new Grid
        {
            Background = Brush("#252525"),
            Margin = new Thickness(0),
        };
        header.ColumnDefinitions.Add(new ColumnDefinition());
        header.ColumnDefinitions.Add(
            new ColumnDefinition { Width = GridLength.Auto });

        var heading = new StackPanel { Margin = new Thickness(18, 12, 18, 12) };
        heading.Children.Add(new TextBlock
        {
            Text = "처리 내용 검증",
            FontSize = 18,
            FontWeight = FontWeights.SemiBold,
            Foreground = Brush("#EEEEEE"),
        });
        heading.Children.Add(new TextBlock
        {
            Text = $"{result.Status} · Study {result.StudyCount:N0}건 · "
                + $"DB 무결성 {(result.IntegrityOk == true ? "통과" : "확인 필요")}",
            Margin = new Thickness(0, 4, 0, 0),
            Foreground = Brush("#A78BFA"),
        });
        header.Children.Add(heading);

        var close = new System.Windows.Controls.Button
        {
            Content = "닫기",
            Margin = new Thickness(8, 12, 18, 12),
            Padding = new Thickness(18, 6, 18, 6),
            VerticalAlignment = VerticalAlignment.Center,
            Background = Brush("#303030"),
            Foreground = Brush("#EEEEEE"),
            BorderBrush = Brush("#4A4A4A"),
        };
        close.Click += (_, _) => Close();
        Grid.SetColumn(close, 1);
        header.Children.Add(close);

        root.Children.Add(header);
        _browser.Navigating += Browser_Navigating;
        Grid.SetRow(_browser, 1);
        root.Children.Add(_browser);
        Content = root;

        Loaded += (_, _) => LoadHtml(html);
        Closed += (_, _) => DeleteTemporaryHtml();
    }

    private void LoadHtml(string html)
    {
        File.WriteAllText(
            _temporaryHtmlPath,
            html,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        _browser.Navigate(new Uri(_temporaryHtmlPath));
    }

    private async void Browser_Navigating(
        object sender,
        NavigatingCancelEventArgs e)
    {
        if (e.Uri is null
            || !string.Equals(
                e.Uri.Scheme,
                "inference-excel",
                StringComparison.OrdinalIgnoreCase))
            return;
        e.Cancel = true;
        var parameters = ParseQuery(e.Uri.Query);
        parameters.TryGetValue("source", out var sourcePath);
        parameters.TryGetValue("sheet", out var sheet);
        parameters.TryGetValue("range", out var range);
        if (string.IsNullOrWhiteSpace(sourcePath)
            || string.IsNullOrWhiteSpace(sheet)
            || string.IsNullOrWhiteSpace(range))
        {
            System.Windows.MessageBox.Show(
                this,
                "원본 Excel 경로, 시트 또는 셀 범위 정보가 없습니다.",
                "원본 Excel 열기",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }
        try
        {
            await ExcelRangeNavigator.OpenReadOnlyAsync(
                sourcePath,
                sheet,
                range);
        }
        catch (Exception exception)
        {
            System.Windows.MessageBox.Show(
                this,
                exception.Message,
                "원본 Excel 범위 열기 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private static Dictionary<string, string> ParseQuery(string query)
    {
        var result = new Dictionary<string, string>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var part in query.TrimStart('?').Split(
                     '&',
                     StringSplitOptions.RemoveEmptyEntries))
        {
            var separator = part.IndexOf('=');
            var key = separator < 0 ? part : part[..separator];
            var value = separator < 0
                ? string.Empty
                : part[(separator + 1)..];
            result[WebUtility.UrlDecode(key)] =
                WebUtility.UrlDecode(value);
        }
        return result;
    }

    private void DeleteTemporaryHtml()
    {
        try
        {
            if (File.Exists(_temporaryHtmlPath))
                File.Delete(_temporaryHtmlPath);
        }
        catch (IOException)
        {
            // The temporary preview can be cleaned by the OS.
        }
        catch (UnauthorizedAccessException)
        {
            // The temporary preview can be cleaned by the OS.
        }
    }

    private static SolidColorBrush Brush(string color) =>
        new((System.Windows.Media.Color)
            System.Windows.Media.ColorConverter.ConvertFromString(color));
}
