using System.Net.Http.Json;

namespace JinoSupporter.Web.Services;

public sealed class BmesSettingsSyncService(NgRateSettingsService settings, WebRepository repo)
{
    private static readonly Uri ServerBaseUri = new("http://10.6.4.54:5050/");
    private static readonly HttpClient Http = new()
    {
        BaseAddress = ServerBaseUri,
        Timeout = TimeSpan.FromSeconds(15),
    };
    private readonly NgRateSettingsService _settings = settings;
    private readonly WebRepository _repo = repo;

    public Task<BmesSettingsSyncResult> SyncRoutingRowsAsync(IReadOnlyList<RoutingRow> rows) =>
        PostRowsAsync("standalone/sync/routing-table", rows);

    public Task<BmesSettingsSyncResult> SyncReasonRowsAsync(IReadOnlyList<ReasonRow> rows) =>
        PostRowsAsync("standalone/sync/reason-table", rows);

    public Task<BmesSettingsSyncResult> SyncModelGroupsAsync(IReadOnlyList<ModelGroupRecord> groups) =>
        PostRowsAsync("standalone/sync/model-groups", groups);

    public Task<BmesSettingsSyncResult> SyncBmesMaterialsAsync(IReadOnlyList<BmesMaterial> rows) =>
        PostRowsAsync("standalone/sync/bmes-materials", rows);

    public async Task<BmesSettingsSyncResult> PullRoutingRowsFromServerAsync()
    {
        var result = await GetRowsAsync<RoutingRow>("standalone/sync/routing-table");
        if (!result.Succeeded)
            return new BmesSettingsSyncResult(false, result.Message, 0);

        _settings.ReplaceRoutingRows(result.Rows);
        return new BmesSettingsSyncResult(true, "Routing table loaded from server DB.", result.Rows.Count);
    }

    public async Task<BmesSettingsSyncResult> PullReasonRowsFromServerAsync()
    {
        var result = await GetRowsAsync<ReasonRow>("standalone/sync/reason-table");
        if (!result.Succeeded)
            return new BmesSettingsSyncResult(false, result.Message, 0);

        _settings.ReplaceReasonRows(result.Rows);
        return new BmesSettingsSyncResult(true, "Reason table loaded from server DB.", result.Rows.Count);
    }

    public async Task<BmesSettingsSyncResult> PullModelGroupsFromServerAsync()
    {
        var result = await GetRowsAsync<ModelGroupRecord>("standalone/sync/model-groups");
        if (!result.Succeeded)
            result = await GetRowsAsync<ModelGroupRecord>("standalone/model-groups.json");

        if (!result.Succeeded)
            return new BmesSettingsSyncResult(false, result.Message, 0);

        _repo.SaveModelGroups(result.Rows);
        return new BmesSettingsSyncResult(true, "Model groups loaded from server DB.", result.Rows.Count);
    }

    public async Task<BmesSettingsSyncResult> PullBmesMaterialsFromServerAsync()
    {
        var result = await GetRowsAsync<BmesMaterial>("standalone/sync/bmes-materials");
        if (!result.Succeeded)
            return new BmesSettingsSyncResult(false, result.Message, 0);

        int saved = _repo.UpsertBmesMaterials(result.Rows);
        return new BmesSettingsSyncResult(true, "BMES materials loaded from server DB.", saved);
    }

    public async Task<BmesSettingsSyncResult> PullAllRowsFromServerAsync()
    {
        var routing = await PullRoutingRowsFromServerAsync();
        var reason = await PullReasonRowsFromServerAsync();

        if (routing.Succeeded && reason.Succeeded)
            return new BmesSettingsSyncResult(true, "Routing/Reason tables loaded from server DB.", routing.Rows + reason.Rows);

        string message = $"Routing: {routing.Message} Reason: {reason.Message}";
        return new BmesSettingsSyncResult(false, message, routing.Rows + reason.Rows);
    }

    private static async Task<BmesSettingsSyncResult> PostRowsAsync<T>(string path, IReadOnlyList<T> rows)
    {
        try
        {
            using HttpResponseMessage response = await Http.PostAsJsonAsync(path, rows);
            if (response.IsSuccessStatusCode)
                return new BmesSettingsSyncResult(true, "Server synced.", rows.Count);

            string body = await response.Content.ReadAsStringAsync();
            string message = string.IsNullOrWhiteSpace(body)
                ? $"Server returned {(int)response.StatusCode}."
                : body.Trim();
            return new BmesSettingsSyncResult(false, message, rows.Count);
        }
        catch (Exception ex)
        {
            return new BmesSettingsSyncResult(false, ex.Message, rows.Count);
        }
    }

    private static async Task<ServerRowsResult<T>> GetRowsAsync<T>(string path)
    {
        try
        {
            using HttpResponseMessage response = await Http.GetAsync(path);
            if (!response.IsSuccessStatusCode)
            {
                string body = await response.Content.ReadAsStringAsync();
                string message = string.IsNullOrWhiteSpace(body)
                    ? $"Server returned {(int)response.StatusCode}."
                    : body.Trim();
                return new ServerRowsResult<T>(false, message, []);
            }

            var rows = await response.Content.ReadFromJsonAsync<List<T>>() ?? [];
            return new ServerRowsResult<T>(true, "Server loaded.", rows);
        }
        catch (Exception ex)
        {
            return new ServerRowsResult<T>(false, ex.Message, []);
        }
    }

    private sealed record ServerRowsResult<T>(bool Succeeded, string Message, List<T> Rows);
}

public sealed record BmesSettingsSyncResult(bool Succeeded, string Message, int Rows);
