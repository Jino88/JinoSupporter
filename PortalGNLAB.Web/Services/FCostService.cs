using System.Globalization;
using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// BMES F-Cost (MES072400) data collection → SQLite save.
/// Mirrors NgRateService's auth flow (token → login → GET) but hits a different endpoint
/// that returns Input Amount / F-Cost / F-Cost Rate per (Model, Material) plus a Total row,
/// each carrying 14 pre-aggregated period columns (COL0001..COL0014).
/// Re-uses NgRateSettingsService for credentials + save directory so the user only has
/// to configure BMES login once.
/// </summary>
public sealed class FCostService(NgRateSettingsService settings, AppActivityLogger activity)
{
    private readonly NgRateSettingsService _settings = settings;
    private readonly AppActivityLogger _activity = activity;

    private void Log(IProgress<string>? progress, string msg)
    {
        _activity.Log("FCost", msg);
        progress?.Report(msg);
    }

    /// <summary>Routes the shared BMES session's step messages through <see cref="Log"/>.</summary>
    private IProgress<string> LogTo(IProgress<string>? progress) => new LogProgress(this, progress);

    private sealed class LogProgress(FCostService owner, IProgress<string>? inner) : IProgress<string>
    {
        public void Report(string value) => owner.Log(inner, value);
    }

    // ── HTTP ─────────────────────────────────────────────────────────────────────

    public string GetRawDbPath()
        => PortalDb.FilePath;

    public FCostRawStatus GetRawStatus()
    {
        string dbPath = GetRawDbPath();
        var status = new FCostRawStatus
        {
            DbPath = dbPath,
            Exists = File.Exists(dbPath),
        };
        if (!status.Exists) return status;

        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
            conn.Open();

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    """
                    SELECT
                        COUNT(*),
                        COALESCE(SUM(CASE WHEN Status = 'OK' THEN 1 ELSE 0 END), 0),
                        COALESCE(SUM(CASE WHEN Status <> 'OK' THEN 1 ELSE 0 END), 0),
                        COALESCE(MIN(QueryDate), ''),
                        COALESCE(MAX(QueryDate), '')
                    FROM FCostRawPulls;
                    """;
                using var r = cmd.ExecuteReader();
                if (r.Read())
                {
                    status.PullCount = Convert.ToInt32(r.GetValue(0));
                    status.SuccessCount = Convert.ToInt32(r.GetValue(1));
                    status.FailedCount = Convert.ToInt32(r.GetValue(2));
                    status.FirstDate = r.IsDBNull(3) ? string.Empty : r.GetString(3);
                    status.LastDate = r.IsDBNull(4) ? string.Empty : r.GetString(4);
                }
            }

            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT COUNT(*) FROM FCostRawRows;";
                status.TotalRows = Convert.ToInt32(cmd.ExecuteScalar() ?? 0);
            }
        }
        catch
        {
            status.PullCount = 0;
            status.SuccessCount = 0;
            status.FailedCount = 0;
            status.TotalRows = 0;
            status.FirstDate = string.Empty;
            status.LastDate = string.Empty;
        }

        return status;
    }

    public async Task<FCostRawBackfillResult> BackfillRawAsync(
        DateTime startDate,
        DateTime endDate,
        bool force = false,
        int delayMs = 1200,
        DateTime? forceFromDate = null,
        TimeSpan? forceRefreshTtl = null,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        startDate = startDate.Date;
        endDate = endDate.Date;
        DateTime today = DateTime.Today;
        if (endDate > today) endDate = today;
        if (startDate > endDate)
            throw new ArgumentException("Start date must be before or equal to end date.");

        string saveDir = _settings.FCostDbSaveDirectory;
        Directory.CreateDirectory(saveDir);
        string dbPath = GetRawDbPath();
        await Task.Run(() => EnsureRawDb(dbPath), cancellationToken);

        var result = new FCostRawBackfillResult
        {
            DbPath = dbPath,
            StartDate = startDate,
            EndDate = endDate,
        };

        int totalDays = (endDate - startDate).Days + 1;
        var datesToFetch = new List<(DateTime Day, int Ordinal)>();
        int scanOrdinal = 0;
        for (DateTime day = startDate; day <= endDate; day = day.AddDays(1))
        {
            cancellationToken.ThrowIfCancellationRequested();
            scanOrdinal++;
            result.AttemptedDays++;

            // forceFromDate/forceRefreshTtl are kept for callers; HistoryCachePolicy decides reuse.
            if (ShouldSkipRaw(force, RawPullSucceeded(dbPath, day, DateTime.Now)))
            {
                result.SkippedDays++;
                continue;
            }

            datesToFetch.Add((day, scanOrdinal));
        }

        // One summary line instead of one line per cached day: every progress message costs a
        // Blazor re-render (and a log-scroll JS round trip), which for a 90-day range added up.
        if (result.SkippedDays > 0)
        {
            Log(
                progress,
                $"RAW cache: {result.SkippedDays:N0} of {totalDays:N0} day(s) already OK, {datesToFetch.Count:N0} to fetch");
        }

        if (datesToFetch.Count == 0)
        {
            Log(
                progress,
                $"RAW backfill done from cache: skipped={result.SkippedDays:N0}, no BMES login required");
            return result;
        }

        string id = _settings.LoginId;
        string pwd = _settings.Password;
        if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(pwd))
        {
            string msg = "BMES credentials are not configured.";
            Log(progress, "[ERROR] " + msg);
            result.Failures.Add(msg);
            return result;
        }

        BmesLoginOutcome login = await BmesWebSession.OpenAsync(
            id, pwd,
            new BmesWebSessionOptions
            {
                TokenMessage = "RAW backfill: login token",
                LoginMessage = $"RAW backfill: login as {id}",
            },
            LogTo(progress));
        if (!login.Succeeded)
        {
            string msg = login.Failure == BmesLoginFailure.Token
                ? "Failed to read BMES login token."
                : "BMES login failed — " + login.Diagnostic;
            Log(progress, "[ERROR] " + msg);
            result.Failures.Add(msg);
            return result;
        }
        using BmesWebSession session = login.Session!;

        for (int fetchIndex = 0; fetchIndex < datesToFetch.Count; fetchIndex++)
        {
            (DateTime day, int ordinal) = datesToFetch[fetchIndex];
            cancellationToken.ThrowIfCancellationRequested();

            Log(progress, $"[{ordinal:N0}/{totalDays:N0}] {day:yyyy-MM-dd}: fetch");
            try
            {
                var fetched = await BmesGateway.FetchFCostAsync(session, day, progress);
                if (fetched is null)
                {
                    result.FailedDays++;
                    string msg = $"{day:yyyy-MM-dd}: fetch failed";
                    result.Failures.Add(msg);
                    await Task.Run(() => SaveRawFailure(dbPath, day, msg), cancellationToken);
                    continue;
                }

                await Task.Run(
                    () => SaveRawSuccess(dbPath, day, fetched.Rows, fetched.Columns),
                    cancellationToken);
                result.FetchedDays++;
                result.TotalRows += fetched.Rows.Count;
                Log(progress, $"{day:yyyy-MM-dd}: saved {fetched.Rows.Count:N0} rows");
            }
            catch (Exception ex)
            {
                result.FailedDays++;
                string msg = $"{day:yyyy-MM-dd}: {ex.Message}";
                result.Failures.Add(msg);
                await Task.Run(() => SaveRawFailure(dbPath, day, msg), cancellationToken);
                Log(progress, "[WARN] " + msg);
            }

            if (delayMs > 0 && fetchIndex < datesToFetch.Count - 1)
                await Task.Delay(delayMs, cancellationToken);
        }

        Log(progress, $"RAW backfill done: fetched={result.FetchedDays:N0}, skipped={result.SkippedDays:N0}, failed={result.FailedDays:N0}");
        return result;
    }

    // ── SQLite ───────────────────────────────────────────────────────────────────

    internal static bool ShouldSkipRaw(bool force, bool cachedUsable) => !force && cachedUsable;

    internal static bool RawPullSucceeded(
        string dbPath,
        DateTime queryDate,
        DateTime now)
    {
        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText =
                "SELECT Status, RowCount, FetchedAt FROM FCostRawPulls WHERE QueryDate = @qdate;";
            cmd.Parameters.AddWithValue("@qdate", queryDate.ToString("yyyy-MM-dd"));
            using var r = cmd.ExecuteReader();
            if (!r.Read()) return false;
            string status = r.IsDBNull(0) ? string.Empty : r.GetString(0);
            int rowCount = r.IsDBNull(1) ? 0 : Convert.ToInt32(r.GetValue(1));
            if (!string.Equals(status, "OK", StringComparison.OrdinalIgnoreCase) || rowCount <= 0)
                return false;

            string fetchedText = r.IsDBNull(2) ? string.Empty : r.GetString(2);
            return DateTimeOffset.TryParse(
                       fetchedText,
                       CultureInfo.InvariantCulture,
                       DateTimeStyles.RoundtripKind,
                       out DateTimeOffset fetchedAt) &&
                   HistoryCachePolicy.CanReuse(queryDate, fetchedAt.LocalDateTime, now);
        }
        catch
        {
            return false;
        }
    }

    private static void EnsureRawDb(string dbPath)
    {
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            PRAGMA journal_mode=WAL;

            CREATE TABLE IF NOT EXISTS FCostRawPulls (
                QueryDate   TEXT PRIMARY KEY,
                Status      TEXT NOT NULL,
                FetchedAt   TEXT NOT NULL,
                RowCount    INTEGER NOT NULL DEFAULT 0,
                ColumnCount INTEGER NOT NULL DEFAULT 0,
                ErrorMessage TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS FCostRawRows (
                QueryDate TEXT NOT NULL,
                RowNo INTEGER NOT NULL,
                FetchedAt TEXT NOT NULL,
                RawJson TEXT NOT NULL DEFAULT '',
                FaccoTx TEXT, PrdGrTx TEXT, VeridTx TEXT, ModNoTx TEXT,
                AssemTx TEXT, AbChgTx TEXT, MCodeTx TEXT, MatnrTx TEXT,
                ZTypeTx TEXT, ZType TEXT, TSort TEXT, ZSort TEXT,
                Facco TEXT, Werks TEXT, PrdGr TEXT, ModNo TEXT, Verid TEXT,
                AbChg TEXT, Cat01 TEXT, Matnr TEXT, Assem TEXT,
                TwSyn TEXT, ZValu TEXT,
                Col0001 REAL, Col0002 REAL, Col0003 REAL, Col0004 REAL,
                Col0005 REAL, Col0006 REAL, Col0007 REAL, Col0008 REAL,
                Col0009 REAL, Col0010 REAL, Col0011 REAL, Col0012 REAL,
                Col0013 REAL, Col0014 REAL,
                PRIMARY KEY (QueryDate, RowNo)
            );
            CREATE INDEX IF NOT EXISTS IX_FCostRawRows_QueryDate ON FCostRawRows(QueryDate);
            CREATE INDEX IF NOT EXISTS IX_FCostRawRows_Matnr ON FCostRawRows(Matnr);
            CREATE INDEX IF NOT EXISTS IX_FCostRawRows_ZType ON FCostRawRows(ZType);
            CREATE INDEX IF NOT EXISTS IX_FCostRawRows_Werks ON FCostRawRows(Werks);

            CREATE TABLE IF NOT EXISTS FCostRawColumns (
                QueryDate TEXT NOT NULL,
                ColIndex INTEGER NOT NULL,
                Code TEXT NOT NULL,
                Header TEXT NOT NULL,
                PDate TEXT NOT NULL,
                Kind TEXT NOT NULL,
                PRIMARY KEY (QueryDate, ColIndex)
            );
            CREATE INDEX IF NOT EXISTS IX_FCostRawColumns_PDate ON FCostRawColumns(PDate);
            """;
        cmd.ExecuteNonQuery();
    }

    private static void SaveRawFailure(string dbPath, DateTime queryDate, string errorMessage)
    {
        EnsureRawDb(dbPath);
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            INSERT INTO FCostRawPulls (QueryDate, Status, FetchedAt, RowCount, ColumnCount, ErrorMessage)
            VALUES (@qdate, 'FAILED', @fetched, 0, 0, @error)
            ON CONFLICT(QueryDate) DO UPDATE SET
                Status = excluded.Status,
                FetchedAt = excluded.FetchedAt,
                ErrorMessage = excluded.ErrorMessage;
            """;
        cmd.Parameters.AddWithValue("@qdate", queryDate.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@fetched", DateTime.Now.ToString("O"));
        cmd.Parameters.AddWithValue("@error", errorMessage);
        cmd.ExecuteNonQuery();
    }

    private static void SaveRawSuccess(
        string dbPath,
        DateTime queryDate,
        List<FCostRow> rows,
        List<FCostColumnMeta> columns)
    {
        EnsureRawDb(dbPath);
        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();
        using var tx = conn.BeginTransaction();
        string qIso = queryDate.ToString("yyyy-MM-dd");
        string nowIso = DateTime.Now.ToString("O");

        using (var del = conn.CreateCommand())
        {
            del.Transaction = tx;
            del.CommandText =
                """
                DELETE FROM FCostRawRows WHERE QueryDate = @qdate;
                DELETE FROM FCostRawColumns WHERE QueryDate = @qdate;
                """;
            del.Parameters.AddWithValue("@qdate", qIso);
            del.ExecuteNonQuery();
        }

        using (var pull = conn.CreateCommand())
        {
            pull.Transaction = tx;
            pull.CommandText =
                """
                INSERT INTO FCostRawPulls (QueryDate, Status, FetchedAt, RowCount, ColumnCount, ErrorMessage)
                VALUES (@qdate, 'OK', @fetched, @rows, @cols, '')
                ON CONFLICT(QueryDate) DO UPDATE SET
                    Status = excluded.Status,
                    FetchedAt = excluded.FetchedAt,
                    RowCount = excluded.RowCount,
                    ColumnCount = excluded.ColumnCount,
                    ErrorMessage = excluded.ErrorMessage;
                """;
            pull.Parameters.AddWithValue("@qdate", qIso);
            pull.Parameters.AddWithValue("@fetched", nowIso);
            pull.Parameters.AddWithValue("@rows", rows.Count);
            pull.Parameters.AddWithValue("@cols", columns.Count);
            pull.ExecuteNonQuery();
        }

        if (columns.Count > 0)
        {
            using var colIns = conn.CreateCommand();
            colIns.Transaction = tx;
            colIns.CommandText =
                """
                INSERT INTO FCostRawColumns (QueryDate, ColIndex, Code, Header, PDate, Kind)
                VALUES (@qdate, @idx, @code, @header, @pdate, @kind);
                """;
            var cpQ = colIns.Parameters.Add("@qdate", SqliteType.Text);
            var cpI = colIns.Parameters.Add("@idx", SqliteType.Integer);
            var cpC = colIns.Parameters.Add("@code", SqliteType.Text);
            var cpH = colIns.Parameters.Add("@header", SqliteType.Text);
            var cpP = colIns.Parameters.Add("@pdate", SqliteType.Text);
            var cpK = colIns.Parameters.Add("@kind", SqliteType.Text);
            colIns.Prepare();

            foreach (var c in columns)
            {
                cpQ.Value = qIso;
                cpI.Value = c.Index;
                cpC.Value = c.Code;
                cpH.Value = c.Header;
                cpP.Value = c.PDate;
                cpK.Value = c.Kind.ToString();
                colIns.ExecuteNonQuery();
            }
        }

        using var ins = conn.CreateCommand();
        ins.Transaction = tx;
        ins.CommandText =
            """
            INSERT INTO FCostRawRows
            (QueryDate, RowNo, FetchedAt, RawJson, FaccoTx, PrdGrTx, VeridTx, ModNoTx,
             AssemTx, AbChgTx, MCodeTx, MatnrTx, ZTypeTx, ZType, TSort, ZSort,
             Facco, Werks, PrdGr, ModNo, Verid, AbChg, Cat01, Matnr, Assem,
             TwSyn, ZValu,
             Col0001, Col0002, Col0003, Col0004, Col0005, Col0006, Col0007,
             Col0008, Col0009, Col0010, Col0011, Col0012, Col0013, Col0014)
            VALUES
            (@qdate, @rowno, @fetched, @rawjson, @facco_tx, @prdgr_tx, @verid_tx, @modno_tx,
             @assem_tx, @abchg_tx, @mcode_tx, @matnr_tx, @ztype_tx, @ztype, @tsort, @zsort,
             @facco, @werks, @prdgr, @modno, @verid, @abchg, @cat01, @matnr, @assem,
             @twsyn, @zvalu,
             @c01, @c02, @c03, @c04, @c05, @c06, @c07, @c08, @c09, @c10,
             @c11, @c12, @c13, @c14);
            """;

        SqliteParameter P(string n, SqliteType t) => ins.Parameters.Add(n, t);
        var pQDate    = P("@qdate",    SqliteType.Text);
        var pRowNo    = P("@rowno",    SqliteType.Integer);
        var pFetched  = P("@fetched",  SqliteType.Text);
        var pRawJson  = P("@rawjson",  SqliteType.Text);
        var pFaccoTx  = P("@facco_tx", SqliteType.Text);
        var pPrdGrTx  = P("@prdgr_tx", SqliteType.Text);
        var pVeridTx  = P("@verid_tx", SqliteType.Text);
        var pModNoTx  = P("@modno_tx", SqliteType.Text);
        var pAssemTx  = P("@assem_tx", SqliteType.Text);
        var pAbChgTx  = P("@abchg_tx", SqliteType.Text);
        var pMCodeTx  = P("@mcode_tx", SqliteType.Text);
        var pMatnrTx  = P("@matnr_tx", SqliteType.Text);
        var pZTypeTx  = P("@ztype_tx", SqliteType.Text);
        var pZType    = P("@ztype",    SqliteType.Text);
        var pTSort    = P("@tsort",    SqliteType.Text);
        var pZSort    = P("@zsort",    SqliteType.Text);
        var pFacco    = P("@facco",    SqliteType.Text);
        var pWerks    = P("@werks",    SqliteType.Text);
        var pPrdGr    = P("@prdgr",    SqliteType.Text);
        var pModNo    = P("@modno",    SqliteType.Text);
        var pVerid    = P("@verid",    SqliteType.Text);
        var pAbChg    = P("@abchg",    SqliteType.Text);
        var pCat01    = P("@cat01",    SqliteType.Text);
        var pMatnr    = P("@matnr",    SqliteType.Text);
        var pAssem    = P("@assem",    SqliteType.Text);
        var pTwSyn    = P("@twsyn",    SqliteType.Text);
        var pZValu    = P("@zvalu",    SqliteType.Text);
        var pC = new SqliteParameter[14];
        for (int i = 0; i < 14; i++) pC[i] = P($"@c{i + 1:D2}", SqliteType.Real);
        ins.Prepare();

        for (int i = 0; i < rows.Count; i++)
        {
            var r = rows[i];
            pQDate.Value = qIso;
            pRowNo.Value = i + 1;
            pFetched.Value = nowIso;
            pRawJson.Value = r.RawJson;
            pFaccoTx.Value = r.FaccoTx; pPrdGrTx.Value = r.PrdGrTx; pVeridTx.Value = r.VeridTx;
            pModNoTx.Value = r.ModNoTx; pAssemTx.Value = r.AssemTx; pAbChgTx.Value = r.AbChgTx;
            pMCodeTx.Value = r.MCodeTx; pMatnrTx.Value = r.MatnrTx; pZTypeTx.Value = r.ZTypeTx;
            pZType.Value   = r.ZType;   pTSort.Value   = r.TSort;   pZSort.Value   = r.ZSort;
            pFacco.Value   = r.Facco;   pWerks.Value   = r.Werks;   pPrdGr.Value   = r.PrdGr;
            pModNo.Value   = r.ModNo;   pVerid.Value   = r.Verid;   pAbChg.Value   = r.AbChg;
            pCat01.Value   = r.Cat01;   pMatnr.Value   = r.Matnr;   pAssem.Value   = r.Assem;
            pTwSyn.Value   = r.TwSyn;   pZValu.Value   = r.ZValu;
            pC[0].Value  = r.Col0001;  pC[1].Value  = r.Col0002;  pC[2].Value  = r.Col0003;
            pC[3].Value  = r.Col0004;  pC[4].Value  = r.Col0005;  pC[5].Value  = r.Col0006;
            pC[6].Value  = r.Col0007;  pC[7].Value  = r.Col0008;  pC[8].Value  = r.Col0009;
            pC[9].Value  = r.Col0010;  pC[10].Value = r.Col0011;  pC[11].Value = r.Col0012;
            pC[12].Value = r.Col0013;  pC[13].Value = r.Col0014;
            ins.ExecuteNonQuery();
        }

        tx.Commit();
    }

}
