namespace JinoSupporter.Web.Services;

/// <summary>
/// Fetches material master records (MES020010/SearchList) from BMES and persists
/// them to the BmesMaterials table. Login and the BMES call live in BmesWebSession / BmesGateway.
/// </summary>
public sealed class BmesMaterialService(
    NgRateSettingsService settings,
    WebRepository repo)
{
    /// <summary>
    /// Fetches all materials from MES020010/SearchList and upserts them.
    /// Returns the number of rows saved, or -1 on failure.
    /// </summary>
    public async Task<int> FetchAllAsync(IProgress<string>? progress = null)
    {
        string loginId  = settings.LoginId;
        string password = settings.Password;
        if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
        {
            progress?.Report("[ERROR] BMES credentials not configured.");
            return -1;
        }

        var opened = await BmesWebSession.OpenAsync(loginId, password, new BmesWebSessionOptions
        {
            Timeout      = TimeSpan.FromSeconds(300),
            TokenMessage = "Fetching verification token…",
            LoginMessage = "Logging in…",
        }, progress);
        if (opened.Session is null)
        {
            progress?.Report(opened.Failure == BmesLoginFailure.Token
                ? "[ERROR] Failed to obtain token."
                : "[ERROR] Login failed — check credentials in Setting.");
            return -1;
        }

        List<BmesMaterial>? list;
        using (var session = opened.Session)
            list = await BmesGateway.FetchMaterialsAsync(session, progress);
        if (list is null) return -1;

        progress?.Report($"Parsed {list.Count:N0} rows. Saving to DB…");
        int saved = await Task.Run(() => repo.UpsertBmesMaterials(list));
        progress?.Report($"✓ Saved {saved:N0} material(s).");
        return saved;
    }
}
