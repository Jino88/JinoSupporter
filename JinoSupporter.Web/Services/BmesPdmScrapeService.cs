using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Fetches the drawing/spec master from BMES MES073300/SearchList and dumps every row
/// verbatim to a standalone SQLite DB (bmes_pdm_raw.db) next to ngrate_settings.db.
/// All values are stored as TEXT. Full refresh — the table is dropped and recreated on
/// every call.
/// </summary>
public sealed class BmesPdmScrapeService(NgRateSettingsService settings)
{
    private static string[] Columns => BmesGateway.PdmColumns;

    public string RawDbPath =>
        Path.Combine(settings.SettingsDbDirectory, "bmes_pdm_raw.db");

    /// <summary>
    /// Returns number of rows saved, or -1 on failure.
    /// </summary>
    public async Task<int> FetchAllAsync(IProgress<string>? progress = null)
    {
        string loginId = settings.LoginId;
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

        List<Dictionary<string, string>>? rows;
        using (var session = opened.Session)
            rows = await BmesGateway.FetchPdmRowsAsync(session, progress);
        if (rows is null) return -1;

        progress?.Report($"Parsed {rows.Count:N0} rows. Saving to {Path.GetFileName(RawDbPath)}…");
        int saved = await Task.Run(() => SaveToSqlite(rows));
        progress?.Report($"✓ Saved {saved:N0} row(s) to bmes_pdm_raw.db");
        return saved;
    }

    /// <summary>
    /// Rows whose PDRTX (part name) or BMONO_TX (model name) contains <paramref name="search"/>.
    /// Returns an empty list when the raw DB has not been fetched yet.
    /// </summary>
    public List<BmesPdmRow> Search(string search, int maxRows = 300)
    {
        if (!File.Exists(RawDbPath))
            return [];

        using var conn = new SqliteConnection($"Data Source={RawDbPath};Mode=ReadOnly");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            SELECT PDMNO, PDMVR, PDTNO_TX, PDRTX, BMONO_TX, USEYN, ERDAT
            FROM BmesPdm
            WHERE (@search = '' OR PDRTX LIKE @like OR BMONO_TX LIKE @like)
            -- BMONO_TX is populated on only a small fraction of rows, so ordering by it first
            -- would bury every linked row behind thousands of blanks.
            ORDER BY PDRTX, BMONO_TX
            LIMIT @maxRows;
            """;
        string term = (search ?? string.Empty).Trim();
        cmd.Parameters.AddWithValue("@search", term);
        cmd.Parameters.AddWithValue("@like", "%" + term + "%");
        cmd.Parameters.AddWithValue("@maxRows", Math.Clamp(maxRows, 1, 5000));

        var result = new List<BmesPdmRow>();
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new BmesPdmRow
            {
                Pdmno = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                Pdmvr = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                PdtnoTx = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                Pdrtx = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                BmonoTx = reader.IsDBNull(4) ? string.Empty : reader.GetString(4),
                Useyn = reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                Erdat = reader.IsDBNull(6) ? string.Empty : reader.GetString(6),
            });
        }
        return result;
    }

    /// <summary>
    /// Resolves the drawing attached to a PDMNO and returns its bytes.
    /// Chain: SearchDetail (find the file) → DownloadCheck (server-side permission check) →
    /// the file endpoint itself.
    /// </summary>
    public async Task<BmesPdmDownload> DownloadDrawingAsync(
        string pdmno,
        string pdmvr = "0",
        IProgress<string>? progress = null)
    {
        string docNo = (pdmno ?? string.Empty).Trim();
        if (docNo.Length == 0)
            return BmesPdmDownload.Failed("PDMNO is empty.");

        string loginId = settings.LoginId;
        string password = settings.Password;
        if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
            return BmesPdmDownload.Failed("BMES credentials not configured.");

        // No default X-Requested-With here: the file action is reached by navigation in the
        // BMES UI, and the header is only added to the two JSON calls in the gateway.
        var opened = await BmesWebSession.OpenAsync(loginId, password, new BmesWebSessionOptions
        {
            Timeout    = TimeSpan.FromSeconds(300),
            Referrer   = new Uri($"{BmesWebSession.BaseUrl}/MES073300"),
            AjaxHeader = false,
        });
        if (opened.Session is null)
        {
            return BmesPdmDownload.Failed(opened.Failure == BmesLoginFailure.Token
                ? "Failed to obtain verification token."
                : "Login failed — check credentials in Setting.");
        }

        using var session = opened.Session;
        return await BmesGateway.DownloadPdmDrawingAsync(session, docNo, pdmvr, progress);
    }

    /// <summary>
    /// Rows whose PDRTX matches one of <paramref name="partNames"/> or whose BMONO_TX matches
    /// one of <paramref name="modelNames"/>. Used to narrow the PDM master down to what a
    /// loaded BOM actually contains. Matching is exact (trimmed, case-insensitive).
    /// </summary>
    public List<BmesPdmRow> SearchByNames(
        IEnumerable<string> partNames,
        IEnumerable<string> modelNames,
        int maxRows = 2000)
    {
        if (!File.Exists(RawDbPath))
            return [];

        var parts = Normalize(partNames);
        var models = Normalize(modelNames);
        if (parts.Count == 0 && models.Count == 0)
            return [];

        using var conn = new SqliteConnection($"Data Source={RawDbPath};Mode=ReadOnly");
        conn.Open();

        var found = new Dictionary<string, BmesPdmRow>(StringComparer.OrdinalIgnoreCase);

        // Batched so the parameter count stays well inside SQLite's limit.
        foreach (var batch in Batch(parts, 400))
            Collect(conn, "PDRTX", batch, found);
        foreach (var batch in Batch(models, 400))
            Collect(conn, "BMONO_TX", batch, found);

        return found.Values
            .OrderBy(r => r.Pdrtx, StringComparer.OrdinalIgnoreCase)
            .ThenBy(r => r.Pdmno, StringComparer.OrdinalIgnoreCase)
            .Take(Math.Clamp(maxRows, 1, 20000))
            .ToList();
    }

    private static void Collect(
        SqliteConnection conn,
        string column,
        List<string> values,
        Dictionary<string, BmesPdmRow> into)
    {
        string placeholders = string.Join(", ", values.Select((_, i) => $"@v{i}"));
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            $"""
            SELECT PDMNO, PDMVR, PDTNO_TX, PDRTX, BMONO_TX, USEYN, ERDAT
            FROM BmesPdm
            WHERE TRIM(UPPER({column})) IN ({placeholders});
            """;
        for (int i = 0; i < values.Count; i++)
            cmd.Parameters.AddWithValue($"@v{i}", values[i]);

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var row = new BmesPdmRow
            {
                Pdmno = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                Pdmvr = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                PdtnoTx = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                Pdrtx = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                BmonoTx = reader.IsDBNull(4) ? string.Empty : reader.GetString(4),
                Useyn = reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                Erdat = reader.IsDBNull(6) ? string.Empty : reader.GetString(6),
            };
            into[row.Pdmno + "\t" + row.Pdrtx] = row;
        }
    }

    private static List<string> Normalize(IEnumerable<string> values)
        => values
            .Select(v => (v ?? string.Empty).Trim().ToUpperInvariant())
            .Where(v => v.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .ToList();

    private static IEnumerable<List<string>> Batch(List<string> values, int size)
    {
        for (int i = 0; i < values.Count; i += size)
            yield return values.GetRange(i, Math.Min(size, values.Count - i));
    }

    /// <summary>Row count and last fetch time, for showing whether a refresh is needed.</summary>
    public (int Rows, string FetchedAt) GetStatus()
    {
        if (!File.Exists(RawDbPath))
            return (0, string.Empty);

        try
        {
            using var conn = new SqliteConnection($"Data Source={RawDbPath};Mode=ReadOnly");
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*), MAX(FetchedAt) FROM BmesPdm;";
            using var reader = cmd.ExecuteReader();
            if (!reader.Read())
                return (0, string.Empty);
            return (
                reader.IsDBNull(0) ? 0 : reader.GetInt32(0),
                reader.IsDBNull(1) ? string.Empty : reader.GetString(1));
        }
        catch
        {
            return (0, string.Empty);
        }
    }

    private int SaveToSqlite(List<Dictionary<string, string>> rows)
    {
        using var conn = new SqliteConnection($"Data Source={RawDbPath}");
        conn.Open();

        string colDefs = string.Join(", ", Columns.Select(c => $"[{c}] TEXT"));
        using (var drop = conn.CreateCommand())
        {
            drop.CommandText = "DROP TABLE IF EXISTS [BmesPdm];";
            drop.ExecuteNonQuery();
        }
        using (var create = conn.CreateCommand())
        {
            create.CommandText = $"CREATE TABLE [BmesPdm] ({colDefs}, [FetchedAt] TEXT);";
            create.ExecuteNonQuery();
        }

        string colList = string.Join(", ", Columns.Select(c => $"[{c}]"));
        string paramList = string.Join(", ", Columns.Select((_, i) => $"@p{i}"));
        string fetchedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        using var tx = conn.BeginTransaction();
        using var ins = conn.CreateCommand();
        ins.CommandText = $"INSERT INTO [BmesPdm] ({colList}, [FetchedAt]) VALUES ({paramList}, @fetchedAt);";
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

        using (var index = conn.CreateCommand())
        {
            index.CommandText =
                "CREATE INDEX IF NOT EXISTS IX_BmesPdm_Search ON BmesPdm (BMONO_TX, PDRTX);";
            index.ExecuteNonQuery();
        }

        tx.Commit();
        return saved;
    }
}

/// <summary>Result of a drawing download attempt.</summary>
public sealed class BmesPdmDownload
{
    public bool Success { get; init; }
    public string FileName { get; init; } = string.Empty;
    public byte[] Bytes { get; init; } = [];
    public string Error { get; init; } = string.Empty;

    public static BmesPdmDownload Ok(string fileName, byte[] bytes)
        => new() { Success = true, FileName = fileName, Bytes = bytes };

    public static BmesPdmDownload Failed(string error)
        => new() { Success = false, Error = error };
}

/// <summary>One drawing/spec master row, trimmed to the fields used for picking a model.</summary>
public sealed class BmesPdmRow
{
    public string Pdmno { get; init; } = string.Empty;

    /// <summary>Document revision. Not always "0" — it is required to download the right file.</summary>
    public string Pdmvr { get; init; } = string.Empty;

    public string PdtnoTx { get; init; } = string.Empty;
    public string Pdrtx { get; init; } = string.Empty;
    public string BmonoTx { get; init; } = string.Empty;
    public string Useyn { get; init; } = string.Empty;
    public string Erdat { get; init; } = string.Empty;

    /// <summary>
    /// Name used as the BOM root. BMONO_TX (the model link) is filled on only a small
    /// fraction of the master, so PDRTX is the practical fallback.
    /// </summary>
    public string BomTarget =>
        string.IsNullOrWhiteSpace(BmonoTx) ? Pdrtx.Trim() : BmonoTx.Trim();
}
