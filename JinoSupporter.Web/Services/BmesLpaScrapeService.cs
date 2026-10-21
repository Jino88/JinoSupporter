namespace JinoSupporter.Web.Services;

/// <summary>
/// Fetches LPA (Layered Process Audit) records from BMES MES073260/SearchList.
///
/// Unlike the other BMES scrapers this one posts a JSON body rather than a query string,
/// and it does not persist anything — the page just renders what comes back.
///
/// The response schema is not pinned down anywhere, so rows are read generically: every
/// property of every returned object becomes a column, ordered by first appearance. That
/// way a server-side schema change shows up as an extra column instead of silently
/// dropping data.
/// </summary>
public sealed class BmesLpaScrapeService(NgRateSettingsService settings)
{
        /// <summary>Search conditions. Defaults mirror the "all records" query.</summary>
    public sealed record LpaQuery
    {
        public string Facco   { get; init; } = "GN";
        public string DateFrom { get; init; } = "19000101";
        public string DateTo   { get; init; } = "20501231";
        public string Lqrno   { get; init; } = "";
        public string Lqbno   { get; init; } = "";
        public string Auloc   { get; init; } = "";
        public string Implv   { get; init; } = "";
        public string Dicno   { get; init; } = "";
        public string Laseq   { get; init; } = "";
        public string Chker   { get; init; } = "";
        public string Zstat   { get; init; } = "";
        public string UseYn   { get; init; } = "Y";
    }

    public sealed record LpaResult(
        IReadOnlyList<string> Columns,
        IReadOnlyList<IReadOnlyDictionary<string, string>> Rows,
        string? Error)
    {
        public bool IsSuccess => Error is null;
        public static LpaResult Fail(string message) => new([], [], message);
    }

    /// <summary>Conditions for the per-audit checklist behind one LQRNO.</summary>
    public sealed record LpaDetailQuery
    {
        public required string Lqrno { get; init; }
        public string Facco { get; init; } = "GN";
        public string Dicno { get; init; } = "";
        public string Auloc { get; init; } = "";
        public string Implv { get; init; } = "";
        public string Lcitm { get; init; } = "";
    }

    public async Task<LpaResult> FetchAsync(LpaQuery query, IProgress<string>? progress = null)
    {
        var (session, error) = await OpenSessionAsync(progress);
        if (session is null) return LpaResult.Fail(error!);
        using (session)
            return await BmesGateway.FetchLpaListAsync(session, query, progress);
    }

    public async Task<LpaResult> FetchDetailAsync(LpaDetailQuery query, IProgress<string>? progress = null)
    {
        var (session, error) = await OpenSessionAsync(progress);
        if (session is null) return LpaResult.Fail(error!);
        using (session)
            return await BmesGateway.FetchLpaDetailAsync(session, query, progress);
    }

    /// <summary>
    /// How many detail lookups run at once inside one batch. The token+login handshake is
    /// already paid once for the whole batch, so this only bounds how hard BMES gets hit —
    /// kept low deliberately: MES073261 is a legacy endpoint and the point of the batch is
    /// to avoid hammering it, not to race it.
    /// </summary>
    private const int DetailBatchConcurrency = 4;

    /// <summary>
    /// Fetch the checklist behind many LQRNOs over a SINGLE authenticated session.
    ///
    /// <see cref="FetchDetailAsync"/> opens a fresh session per call, and that handshake
    /// (GET token → POST login) costs more than the search itself. Prefetching a whole
    /// result set one row at a time would therefore mean one login per row; this logs in
    /// once and reuses the cookie for every lookup.
    ///
    /// Returns one entry per distinct LQRNO. A row that failed carries its own
    /// <see cref="LpaResult.Error"/> instead of failing the batch — a partial prefetch is
    /// still worth showing.
    /// </summary>
    public async Task<IReadOnlyDictionary<string, LpaResult>> FetchDetailsAsync(
        IReadOnlyList<LpaDetailQuery> queries, IProgress<string>? progress = null,
        IProgress<(int Done, int Total)>? counter = null)
    {
        var results = new Dictionary<string, LpaResult>(StringComparer.Ordinal);
        if (queries.Count == 0) return results;

        var (session, error) = await OpenSessionAsync(progress);
        if (session is null)
        {
            foreach (var q in queries) results[q.Lqrno] = LpaResult.Fail(error!);
            return results;
        }

        using (session)
        {
            int done = 0;
            using var gate = new SemaphoreSlim(DetailBatchConcurrency);
            var tasks = queries.Select(async q =>
            {
                await gate.WaitAsync();
                try
                {
                    // No per-call progress: the batch reports its own count instead, which
                    // is the only number that means anything across hundreds of lookups.
                    LpaResult r = await BmesGateway.FetchLpaDetailAsync(session, q, null);
                    int n = Interlocked.Increment(ref done);
                    if (n % 10 == 0 || n == queries.Count)
                    {
                        progress?.Report($"Fetching detail {n:N0} / {queries.Count:N0}…");
                        counter?.Report((n, queries.Count));
                    }
                    return (q.Lqrno, r);
                }
                finally { gate.Release(); }
            }).ToList();

            foreach ((string lqrno, LpaResult r) in await Task.WhenAll(tasks))
                results[lqrno] = r;   // duplicate LQRNOs collapse; they carry the same detail
        }

        return results;
    }

    /// <summary>Opens the shared BMES session. Returns the error text on failure.</summary>
    private async Task<(BmesWebSession? Session, string? Error)> OpenSessionAsync(IProgress<string>? progress)
    {
        var opened = await BmesWebSession.OpenAsync(settings.LoginId, settings.Password, new BmesWebSessionOptions
        {
            Timeout      = TimeSpan.FromSeconds(300),
            TokenMessage = "Fetching verification token…",
            LoginMessage = "Logging in…",
        }, progress);
        if (opened.Session is not null) return (opened.Session, null);

        return (null, opened.Failure switch
        {
            BmesLoginFailure.MissingCredentials => "BMES credentials are not configured. Set them in Setting.",
            BmesLoginFailure.Token              => "Failed to obtain the verification token from BMES.",
            _                                   => "BMES login failed — check the credentials in Setting.",
        });
    }
}
