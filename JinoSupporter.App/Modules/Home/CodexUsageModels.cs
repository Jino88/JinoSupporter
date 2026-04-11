namespace JinoSupporter.App.Modules.Home;

public sealed class CodexUsageCard
{
    public string Title { get; set; } = string.Empty;
    public string Value { get; set; } = "-";
    public string Detail { get; set; } = string.Empty;
}

public sealed class CodexUsageSnapshot
{
    public bool IsAuthenticated { get; set; }
    public string StatusMessage { get; set; } = string.Empty;
    public string SourceUrl { get; set; } = "https://chatgpt.com/codex/cloud/settings/usage";
    public string DebugText { get; set; } = string.Empty;
    public List<CodexUsageCard> Cards { get; } = new();
}
