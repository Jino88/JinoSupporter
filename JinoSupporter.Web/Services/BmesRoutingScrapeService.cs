using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Fetches routing master records from BMES MES020240/SearchList for each WERKS
/// (3200, 3210, 3220) and dumps all rows verbatim to a standalone SQLite DB
/// (bmes_routing_raw.db) next to ngrate_settings.db. Column set mirrors the API
/// response; all values are stored as TEXT. Full refresh — table is dropped and
/// recreated on every call.
/// </summary>
public sealed class BmesRoutingScrapeService(NgRateSettingsService settings)
{
    private static string[] Columns => BmesGateway.RoutingColumns;

    public string RawDbPath =>
        Path.Combine(settings.SettingsDbDirectory, "bmes_routing_raw.db");

    /// <summary>
    /// Returns number of rows saved across all WERKS requests, or -1 on failure.
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

        List<Dictionary<string, string>>? allRows;
        using (var session = opened.Session)
            allRows = await BmesGateway.FetchRoutingRowsAsync(session, progress);
        if (allRows is null) return -1;

        progress?.Report($"Parsed {allRows.Count:N0} total rows. Saving to {Path.GetFileName(RawDbPath)}…");
        int saved = await Task.Run(() => SaveToSqlite(allRows));
        progress?.Report($"✓ Saved {saved:N0} row(s) to bmes_routing_raw.db");

        progress?.Report("Merging into RoutingTable…");
        int added = await Task.Run(() => settings.MergeRoutingFromRawDb(RawDbPath, progress));
        progress?.Report(added > 0
            ? $"✓ Done. Saved {saved:N0} raw / added {added:N0} new RoutingTable row(s)."
            : $"✓ Done. Saved {saved:N0} raw / no new RoutingTable rows.");
        return saved;
    }

    private int SaveToSqlite(List<Dictionary<string, string>> rows)
    {
        using var conn = new SqliteConnection($"Data Source={RawDbPath}");
        conn.Open();

        string colDefs = string.Join(", ", Columns.Select(c => $"[{c}] TEXT"));
        using (var drop = conn.CreateCommand())
        {
            drop.CommandText = "DROP TABLE IF EXISTS [BmesRouting];";
            drop.ExecuteNonQuery();
        }
        using (var create = conn.CreateCommand())
        {
            create.CommandText = $"CREATE TABLE [BmesRouting] ({colDefs}, [FetchedAt] TEXT);";
            create.ExecuteNonQuery();
        }

        string colList   = string.Join(", ", Columns.Select(c => $"[{c}]"));
        string paramList = string.Join(", ", Columns.Select((_, i) => $"@p{i}"));
        string fetchedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        using var tx  = conn.BeginTransaction();
        using var ins = conn.CreateCommand();
        ins.CommandText = $"INSERT INTO [BmesRouting] ({colList}, [FetchedAt]) VALUES ({paramList}, @fetchedAt);";
        for (int i = 0; i < Columns.Length; i++) ins.Parameters.Add(new SqliteParameter($"@p{i}", string.Empty));
        ins.Parameters.Add(new SqliteParameter("@fetchedAt", fetchedAt));

        int saved = 0;
        foreach (var row in rows)
        {
            for (int i = 0; i < Columns.Length; i++)
                ins.Parameters[i].Value = row.TryGetValue(Columns[i], out var v) ? v : string.Empty;
            ins.Parameters[Columns.Length].Value = fetchedAt;
            ins.ExecuteNonQuery();
            saved++;
        }
        tx.Commit();
        return saved;
    }
}
