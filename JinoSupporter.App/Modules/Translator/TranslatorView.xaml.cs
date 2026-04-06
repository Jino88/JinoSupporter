using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using CustomKeyboardCSharp.Models;
using CustomKeyboardCSharp.Services;
using Microsoft.Win32;

namespace JinoSupporter.App.Modules.Translator;

public partial class TranslatorView : UserControl
{
    public event Action? WebModuleSnapshotChanged;

    private readonly SettingsService _settingsService = new();
    private readonly TranslationService _translationService = new();
    private readonly TranslationHistoryRepository _historyRepository = new();
    private readonly GoogleDriveSyncService _googleDriveSyncService = new();
    private readonly ScreenCaptureService _screenCaptureService = new();

    private AppSettings _settings = new();
    private string _detectedText = string.Empty;
    private bool _loadingSettings;

    public TranslatorView()
    {
        InitializeComponent();
        ProviderComboBox.ItemsSource = Enum.GetValues(typeof(AiProvider)).Cast<AiProvider>();
        DirectionComboBox.ItemsSource = Enum.GetValues(typeof(TranslationDirection)).Cast<TranslationDirection>();
        ScreenDirectionComboBox.ItemsSource = Enum.GetValues(typeof(TranslationDirection)).Cast<TranslationDirection>();
        LoadSettings();
        RefreshCollections();
        NotifyWebModuleSnapshotChanged();
    }

    public object GetWebModuleSnapshot()
    {
        return new
        {
            moduleType = "Translator",
            provider = _settings.SelectedProvider.ToString(),
            sourceLength = SourceTextBox.Text.Length,
            optionCount = OptionListBox.Items.Count,
            historyCount = HistoryListBox.Items.Count,
            vocabularyCount = VocabularyListBox.Items.Count,
            driveStatus = DriveStatusTextBlock.Text,
            statusMessage = StatusTextBlock.Text
        };
    }

    public object UpdateWebModuleState(JsonElement payload) => GetWebModuleSnapshot();

    public object InvokeWebModuleAction(string action)
    {
        switch (action)
        {
            case "translate-text":
                _ = TranslateCurrentTextAsync();
                break;
            case "swap-text":
                SwapTexts();
                break;
            case "clear-text":
                ClearAll();
                break;
        }

        return GetWebModuleSnapshot();
    }

    private void LoadSettings()
    {
        _loadingSettings = true;
        _settings = _settingsService.Load();
        ProviderComboBox.SelectedItem = _settings.SelectedProvider;
        DirectionComboBox.SelectedItem = TranslationDirection.AutoToVietnamese;
        ScreenDirectionComboBox.SelectedItem = _settings.ScreenTranslationDirection;
        TimeoutTextBox.Text = _settings.TranslationTimeoutSeconds.ToString();
        OpenAiApiKeyBox.Password = _settings.OpenAiApiKey;
        GeminiApiKeyBox.Password = _settings.GeminiApiKey;
        GoogleClientIdTextBox.Text = _settings.GoogleClientId;
        GoogleClientSecretTextBox.Text = _settings.GoogleClientSecret;
        DbPathTextBlock.Text = _historyRepository.DatabasePath;
        DriveStatusTextBlock.Text = _googleDriveSyncService.HasCachedLogin() ? "Google Drive login cache found." : "No Google Drive session cached.";
        _loadingSettings = false;
    }

    private void SaveSettings()
    {
        if (_loadingSettings)
        {
            return;
        }

        _settings.SelectedProvider = ProviderComboBox.SelectedItem is AiProvider provider ? provider : AiProvider.OpenAi;
        _settings.ScreenTranslationDirection = ScreenDirectionComboBox.SelectedItem is TranslationDirection screenDirection
            ? screenDirection
            : TranslationDirection.KoreanToVietnamese;
        _settings.OpenAiApiKey = OpenAiApiKeyBox.Password.Trim();
        _settings.GeminiApiKey = GeminiApiKeyBox.Password.Trim();
        _settings.GoogleClientId = GoogleClientIdTextBox.Text.Trim();
        _settings.GoogleClientSecret = GoogleClientSecretTextBox.Text.Trim();
        _settings.TranslationTimeoutSeconds = int.TryParse(TimeoutTextBox.Text, out int timeout) ? timeout : _settings.TranslationTimeoutSeconds;
        _settingsService.Save(_settings);
        NotifyWebModuleSnapshotChanged();
    }

    private async void TranslateButton_Click(object sender, RoutedEventArgs e)
    {
        await TranslateCurrentTextAsync();
    }

    private async Task TranslateCurrentTextAsync()
    {
        string source = SourceTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(source))
        {
            SetStatus("Enter text to translate.");
            return;
        }

        SaveSettings();
        await RunBusyAsync(async cancellationToken =>
        {
            TranslationDirection direction = DirectionComboBox.SelectedItem is TranslationDirection selected
                ? selected
                : TranslationDirection.AutoToVietnamese;
            TranslationBundle bundle = await _translationService.TranslateAsync(_settings, source, direction, cancellationToken);
            RenderBundle(bundle, source, string.Empty);
            _historyRepository.SaveTranslation(_settings.SelectedProvider, "text", source, direction.ToString(), bundle);
            RefreshCollections();
            SetStatus("Text translation completed.");
        });
    }

    private async void CaptureTranslateButton_Click(object sender, RoutedEventArgs e)
    {
        Window? owner = Window.GetWindow(this);
        if (owner is null)
        {
            SetStatus("Screen capture owner window was not found.");
            return;
        }

        SaveSettings();
        var capture = await _screenCaptureService.CaptureAsync(owner);
        if (capture is null)
        {
            SetStatus("Screen capture canceled.");
            return;
        }

        CapturedImage.Source = capture.Value.Image;
        await RunBusyAsync(async cancellationToken =>
        {
            TranslationDirection direction = ScreenDirectionComboBox.SelectedItem is TranslationDirection selected
                ? selected
                : TranslationDirection.KoreanToVietnamese;
            ImageTranslationResult result = await _translationService.TranslateImageAsync(_settings, capture.Value.ImageBytes, direction, cancellationToken);
            _detectedText = result.DetectedText;
            SourceTextBox.Text = result.DetectedText;
            RenderBundle(result.ToBundle(), result.DetectedText, result.DetectedText);
            _historyRepository.SaveTranslation(_settings.SelectedProvider, "image", result.DetectedText, direction.ToString(), result.ToBundle());
            RefreshCollections();
            SetStatus("Image translation completed.");
        });
    }

    private void SwapButton_Click(object sender, RoutedEventArgs e)
    {
        SwapTexts();
    }

    private void SwapTexts()
    {
        if (OptionListBox.Items.Count == 0)
        {
            return;
        }

        string firstOption = OptionListBox.Items[0]?.ToString() ?? string.Empty;
        string currentSource = SourceTextBox.Text;
        SourceTextBox.Text = firstOption;
        _detectedText = currentSource;
        DetectedTextBlock.Text = string.IsNullOrWhiteSpace(currentSource) ? string.Empty : $"Previous source: {currentSource}";
        SetStatus("Swapped source with the first translated option.");
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        ClearAll();
    }

    private void ClearAll()
    {
        SourceTextBox.Clear();
        OptionListBox.ItemsSource = null;
        DetectedTextBlock.Text = string.Empty;
        CapturedImage.Source = null;
        _detectedText = string.Empty;
        SetStatus("Translator workspace cleared.");
    }

    private void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshCollections();
        SetStatus("History and vocabulary refreshed.");
    }

    private void ExportDbButton_Click(object sender, RoutedEventArgs e)
    {
        string exportedPath = _historyRepository.ExportDatabaseToDownloads();
        SetStatus($"DB exported to {exportedPath}");
    }

    private void DeleteHistoryButton_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryListBox.SelectedItem is not HistoryListItem selected)
        {
            SetStatus("Choose a history row first.");
            return;
        }

        _historyRepository.DeleteHistory(selected.Id);
        RefreshCollections();
        SetStatus("Selected history deleted.");
    }

    private void DeleteVocabularyButton_Click(object sender, RoutedEventArgs e)
    {
        if (VocabularyListBox.SelectedItem is not VocabularyListItem selected)
        {
            SetStatus("Choose a vocabulary row first.");
            return;
        }

        _historyRepository.DeleteVocabulary(selected.Id);
        RefreshCollections();
        SetStatus("Selected vocabulary deleted.");
    }

    private async void SignInButton_Click(object sender, RoutedEventArgs e)
    {
        SaveSettings();
        await RunBusyAsync(async cancellationToken =>
        {
            var session = await _googleDriveSyncService.SignInAsync(_settings, cancellationToken);
            DriveStatusTextBlock.Text = $"{session.Email} connected.";
            SetStatus("Google Drive sign-in completed.");
        });
    }

    private async void SignOutButton_Click(object sender, RoutedEventArgs e)
    {
        SaveSettings();
        await RunBusyAsync(async cancellationToken =>
        {
            await _googleDriveSyncService.SignOutAsync(_settings, cancellationToken);
            DriveStatusTextBlock.Text = "Google Drive session cleared.";
            SetStatus("Google Drive sign-out completed.");
        });
    }

    private async void UploadDbButton_Click(object sender, RoutedEventArgs e)
    {
        SaveSettings();
        await RunBusyAsync(async cancellationToken =>
        {
            var result = await _googleDriveSyncService.UploadDatabaseAsync(_settings, _historyRepository.DatabasePath, cancellationToken);
            DriveStatusTextBlock.Text = result.Updated ? "Desktop DB updated on Drive." : "Desktop DB uploaded to Drive.";
            SetStatus("Drive upload completed.");
        });
    }

    private async void DownloadDbButton_Click(object sender, RoutedEventArgs e)
    {
        SaveSettings();
        await RunBusyAsync(async cancellationToken =>
        {
            bool downloaded = await _googleDriveSyncService.DownloadDatabaseAsync(_settings, _historyRepository.DatabasePath, cancellationToken);
            DriveStatusTextBlock.Text = downloaded ? "Desktop DB downloaded from Drive." : "No Drive DB found.";
            RefreshCollections();
            SetStatus(downloaded ? "Drive download completed." : "No Drive DB available.");
        });
    }

    private async void MergeDbButton_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog dialog = new()
        {
            Filter = "SQLite DB (*.db)|*.db|All files (*.*)|*.*",
            Title = "Choose another translator DB to merge"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        await RunBusyAsync(_ =>
        {
            DriveSyncSnapshot primary = _historyRepository.ExportSharedSnapshot();
            DriveSyncSnapshot external = _historyRepository.ExportSharedSnapshotFrom(dialog.FileName);
            DriveSyncSnapshot merged = TranslationHistoryRepository.MergeSnapshots(primary, external);
            _historyRepository.ReplaceAllWithSnapshot(merged);
            RefreshCollections();
            SetStatus("Translator DB merge completed.");
            return Task.CompletedTask;
        });
    }

    private void ProviderComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e) => SaveSettings();
    private void DirectionComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e) => NotifyWebModuleSnapshotChanged();
    private void ScreenDirectionComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e) => SaveSettings();
    private void SettingsField_OnChanged(object sender, RoutedEventArgs e) => SaveSettings();
    private void SourceTextBox_OnTextChanged(object sender, TextChangedEventArgs e) => NotifyWebModuleSnapshotChanged();

    private void RenderBundle(TranslationBundle bundle, string sourceText, string detectedText)
    {
        OptionListBox.ItemsSource = bundle.Options
            .Select((option, index) => $"{index + 1}. {option.Text}{Environment.NewLine}{option.Nuance}")
            .ToArray();
        DetectedTextBlock.Text = string.IsNullOrWhiteSpace(detectedText) ? string.Empty : $"Detected text: {detectedText}";
        _detectedText = detectedText;
        NotifyWebModuleSnapshotChanged();
    }

    private void RefreshCollections()
    {
        HistoryListBox.ItemsSource = _historyRepository
            .GetRecentHistories()
            .Select(item => new HistoryListItem(
                item.Id,
                $"{item.Direction} | {item.Provider} | {FromUnix(item.CreatedAt):yyyy-MM-dd HH:mm}{Environment.NewLine}{item.SourceText}"))
            .ToArray();

        VocabularyListBox.ItemsSource = _historyRepository
            .GetRecentVocabulary()
            .Select(item => new VocabularyListItem(
                item.Id,
                $"{item.SourceWord} -> {item.TargetMeaning}{Environment.NewLine}{item.Direction} | {item.Provider}"))
            .ToArray();

        DbPathTextBlock.Text = _historyRepository.DatabasePath;
        NotifyWebModuleSnapshotChanged();
    }

    private async Task RunBusyAsync(Func<CancellationToken, Task> operation)
    {
        try
        {
            IsEnabled = false;
            SetStatus("Working...");
            await operation(CancellationToken.None);
        }
        catch (Exception ex)
        {
            SetStatus(ex.Message);
        }
        finally
        {
            IsEnabled = true;
            NotifyWebModuleSnapshotChanged();
        }
    }

    private void SetStatus(string message)
    {
        StatusTextBlock.Text = message;
        NotifyWebModuleSnapshotChanged();
    }

    private void NotifyWebModuleSnapshotChanged()
    {
        WebModuleSnapshotChanged?.Invoke();
    }

    private static DateTimeOffset FromUnix(long value)
    {
        return DateTimeOffset.FromUnixTimeMilliseconds(value);
    }

    private sealed record HistoryListItem(long Id, string Text)
    {
        public override string ToString() => Text;
    }

    private sealed record VocabularyListItem(long Id, string Text)
    {
        public override string ToString() => Text;
    }
}
