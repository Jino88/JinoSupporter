namespace CustomKeyboardCSharp.Models;

public sealed class HistoryItem
{
    public long Id { get; set; }
    public long CreatedAt { get; set; }
    public string Provider { get; set; } = string.Empty;
    public string Mode { get; set; } = string.Empty;
    public string Direction { get; set; } = string.Empty;
    public string SourceText { get; set; } = string.Empty;
    public List<TranslationOptionItem> Options { get; set; } = [];
}

public sealed class TranslationOptionItem
{
    public int Position { get; set; }
    public string TranslatedText { get; set; } = string.Empty;
    public string Nuance { get; set; } = string.Empty;
}
