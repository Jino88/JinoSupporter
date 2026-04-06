namespace CustomKeyboardCSharp.Models;

public sealed class AppSettings
{
    public AiProvider SelectedProvider { get; set; } = AiProvider.OpenAi;
    public string OpenAiApiKey { get; set; } = string.Empty;
    public string GeminiApiKey { get; set; } = string.Empty;
    public string GoogleClientId { get; set; } = string.Empty;
    public string GoogleClientSecret { get; set; } = string.Empty;
    public int TranslationTimeoutSeconds { get; set; } = 45;
    public TranslationDirection ScreenTranslationDirection { get; set; } = TranslationDirection.KoreanToVietnamese;
}
