using System.Collections.ObjectModel;
using System.Windows.Media;
using CustomKeyboardCSharp.Helpers;
using CustomKeyboardCSharp.Models;

namespace CustomKeyboardCSharp.ViewModels;

public sealed class MainViewModel : ObservableObject
{
    private string _statusMessage = "Ready";
    private string _inputText = string.Empty;
    private string _capturedText = string.Empty;
    private string _overlayStatus = "NOT";
    private string _apiStatus = "NOT";
    private ImageSource? _capturedImage;
    private AiProvider _selectedProvider = AiProvider.OpenAi;
    private TranslationDirection _selectedDirection = TranslationDirection.AutoToVietnamese;
    private TranslationDirection _screenDirection = TranslationDirection.KoreanToVietnamese;
    private string _openAiApiKey = string.Empty;
    private string _geminiApiKey = string.Empty;
    private int _translationTimeoutSeconds = 45;
    private bool _isBusy;

    public ObservableCollection<OptionViewModel> TranslationOptions { get; } = [];
    public ObservableCollection<GlossaryEntryViewModel> GlossaryEntries { get; } = [];
    public ObservableCollection<HistoryCardViewModel> HistoryItems { get; } = [];
    public ObservableCollection<VocabularyCardViewModel> VocabularyItems { get; } = [];

    public Array Providers => Enum.GetValues(typeof(AiProvider));
    public Array Directions => Enum.GetValues(typeof(TranslationDirection));
    public Array ScreenDirections => Enum.GetValues(typeof(TranslationDirection));

    public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
    public string InputText { get => _inputText; set => SetProperty(ref _inputText, value); }
    public string CapturedText { get => _capturedText; set => SetProperty(ref _capturedText, value); }
    public string OverlayStatus { get => _overlayStatus; set => SetProperty(ref _overlayStatus, value); }
    public string ApiStatus { get => _apiStatus; set => SetProperty(ref _apiStatus, value); }
    public ImageSource? CapturedImage { get => _capturedImage; set => SetProperty(ref _capturedImage, value); }
    public AiProvider SelectedProvider { get => _selectedProvider; set => SetProperty(ref _selectedProvider, value); }
    public TranslationDirection SelectedDirection { get => _selectedDirection; set => SetProperty(ref _selectedDirection, value); }
    public TranslationDirection ScreenDirection { get => _screenDirection; set => SetProperty(ref _screenDirection, value); }
    public string OpenAiApiKey { get => _openAiApiKey; set => SetProperty(ref _openAiApiKey, value); }
    public string GeminiApiKey { get => _geminiApiKey; set => SetProperty(ref _geminiApiKey, value); }
    public int TranslationTimeoutSeconds { get => _translationTimeoutSeconds; set => SetProperty(ref _translationTimeoutSeconds, value); }
    public bool IsBusy { get => _isBusy; set => SetProperty(ref _isBusy, value); }
}
