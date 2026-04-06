namespace CustomKeyboardCSharp.ViewModels;

public sealed class HistoryCardViewModel
{
    public long Id { get; set; }
    public string SourceText { get; set; } = string.Empty;
    public string Provider { get; set; } = string.Empty;
    public string Direction { get; set; } = string.Empty;
    public string Mode { get; set; } = string.Empty;
    public string CreatedAtText { get; set; } = string.Empty;
    public List<OptionViewModel> Options { get; set; } = [];
}
