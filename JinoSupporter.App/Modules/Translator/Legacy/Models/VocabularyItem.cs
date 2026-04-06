namespace CustomKeyboardCSharp.Models;

public sealed class VocabularyItem
{
    public long Id { get; set; }
    public string SourceWord { get; set; } = string.Empty;
    public string TargetMeaning { get; set; } = string.Empty;
    public string Direction { get; set; } = string.Empty;
    public string Provider { get; set; } = string.Empty;
}
