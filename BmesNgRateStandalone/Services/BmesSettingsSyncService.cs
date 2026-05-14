using System.Net.Http.Json;

namespace BmesNgRateStandalone.Services;

public sealed class BmesSettingsSyncService
{
    private static readonly Uri ServerBaseUri = new("http://10.6.4.54:5050/");
    private static readonly HttpClient Http = new()
    {
        BaseAddress = ServerBaseUri,
        Timeout = TimeSpan.FromSeconds(15),
    };

    public Task<BmesSettingsSyncResult> SyncRoutingRowsAsync(IReadOnlyList<RoutingRow> rows) =>
        PostRowsAsync("standalone/sync/routing-table", rows);

    public Task<BmesSettingsSyncResult> SyncReasonRowsAsync(IReadOnlyList<ReasonRow> rows) =>
        PostRowsAsync("standalone/sync/reason-table", rows);

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
}

public sealed record BmesSettingsSyncResult(bool Succeeded, string Message, int Rows);
