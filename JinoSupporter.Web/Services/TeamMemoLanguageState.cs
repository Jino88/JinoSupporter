namespace JinoSupporter.Web.Services;

/// <summary>
/// The display language picked on the team memo pages, kept for the browser session (circuit) so
/// moving between Work Memo and Board keeps it.
/// </summary>
public sealed class TeamMemoLanguageState
{
    public string Lang { get; private set; } = TeamMemoLanguage.Original;

    public event Action? Changed;

    public void Set(string lang)
    {
        if (lang == Lang) return;
        Lang = lang;
        Changed?.Invoke();
    }
}
