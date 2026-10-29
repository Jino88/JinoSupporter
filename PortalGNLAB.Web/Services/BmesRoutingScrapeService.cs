namespace PortalGNLAB.Web.Services;

/// <summary>
/// Fills RoutingTable and ReasonTable from PLATFORM_API (the old portal's NG type refresh):
/// AI_TOOL_00055 view R (GN raw data, last 3 months) gives material × process × line stage and the
/// process × NG codes that occurred; AI_TOOL_00056 gives the process / NG code / line stage names.
/// Both merges only add rows that are missing; existing rows and their Reason values stay.
/// </summary>
public sealed class BmesRoutingScrapeService(NgRateSettingsService settings)
{
    /// <summary>
    /// Adds missing RoutingTable rows. Returns the number of routing rows received from BMES, or -1 on failure.
    /// </summary>
    public async Task<int> FetchAllAsync(IProgress<string>? progress = null)
    {
        using BmesWebSession? session = await OpenAsync(progress);
        if (session is null) return -1;
        BmesRoutingSource? source = await FetchSourceAsync(session, progress);
        if (source is null) return -1;

        Dictionary<string, string> names = await Task.Run(settings.GetKnownMaterialNames);
        var unnamed = source.Routes.Where(r => !names.ContainsKey(r.Matnr))
            .GroupBy(r => r.Werks, StringComparer.OrdinalIgnoreCase).ToList();
        foreach (var plant in unnamed)
        {
            var matnrs = plant.Select(r => r.Matnr).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            progress?.Report($"Calling AI_TOOL_00055 view S for {matnrs.Count} material name(s) (WERKS {plant.Key})…");
            try
            {
                foreach (var (code, name) in await BmesGateway.FetchMaterialNamesAsync(session, plant.Key, matnrs, DateTime.Today))
                    names.TryAdd(code, name);
            }
            catch (BmesApiException ex)
            {
                progress?.Report($"[WARN] Material names (WERKS {plant.Key}): {ex.Message}");
            }
        }

        var rows = source.Routes.Where(r => names.ContainsKey(r.Matnr))
            .Select(r => (names[r.Matnr], r.ProcessCode, r.ProcessName, r.StageName)).ToList();
        int noName = source.Routes.Count - rows.Count;
        if (noName > 0) progress?.Report($"[WARN] {noName:N0} routing row(s) skipped: no material name.");
        if (source.UnnamedRoutes > 0) progress?.Report($"[WARN] {source.UnnamedRoutes:N0} routing row(s) skipped: no process name in AI_TOOL_00056.");

        progress?.Report($"Received {source.Routes.Count:N0} routing row(s). Merging into RoutingTable…");
        int added = await Task.Run(() => settings.MergeRoutingRows(rows, progress));
        progress?.Report(added > 0
            ? $"✓ Done. {source.Routes.Count:N0} routing row(s) from BMES API / added {added:N0} new RoutingTable row(s)."
            : $"✓ Done. {source.Routes.Count:N0} routing row(s) from BMES API / no new RoutingTable rows.");
        return source.Routes.Count;
    }

    /// <summary>
    /// Adds missing (ProcessName, NgName) pairs to ReasonTable with an empty Reason.
    /// Returns the number of added rows, or -1 on failure.
    /// </summary>
    public async Task<int> FetchReasonPairsAsync(IProgress<string>? progress = null)
    {
        using BmesWebSession? session = await OpenAsync(progress);
        if (session is null) return -1;
        BmesRoutingSource? source = await FetchSourceAsync(session, progress);
        if (source is null) return -1;

        if (source.UnnamedNgCodes > 0)
            progress?.Report($"[WARN] {source.UnnamedNgCodes:N0} NG code(s) skipped: no name in AI_TOOL_00056.");
        progress?.Report($"Received {source.NgCodes.Count:N0} process / NG code pair(s). Merging into ReasonTable…");
        int added = await Task.Run(() => settings.MergeReasonPairs(
            source.NgCodes.Select(c => (c.ProcessName, c.NgName)), progress));
        return added;
    }

    private async Task<BmesWebSession?> OpenAsync(IProgress<string>? progress)
    {
        BmesLoginOutcome login = await BmesWebSession.OpenAsync(settings.LoginId, settings.Password, progress: progress);
        if (login.Succeeded) return login.Session;
        progress?.Report("[ERROR] BMES API login not configured. Set it in Setting.");
        return null;
    }

    private static async Task<BmesRoutingSource?> FetchSourceAsync(BmesWebSession session, IProgress<string>? progress)
    {
        try
        {
            return await BmesGateway.FetchRoutingSourceAsync(session, DateTime.Today, progress);
        }
        catch (BmesApiException ex)
        {
            progress?.Report($"[ERROR] BMES API: {ex.Message}" + (ex.TraceId is null ? "" : $" (trace {ex.TraceId})"));
            return null;
        }
    }
}
