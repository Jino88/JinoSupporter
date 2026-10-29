namespace PortalGNLAB.Web.Services;

/// <summary>One NG record as the report services read it. Every value is the raw text the
/// BMES API / NgRateCache row carried (quantities included); readers parse and normalize
/// exactly as they did when the rows went through the old temp_*.db file.</summary>
public sealed class NgRateDataRow
{
    public required string MaterialName   { get; init; }
    public required string ProductionLine { get; init; }
    /// <summary>MATERIALNAME + "_" + PRODUCTION_LINE (what SetLineShift used to write).</summary>
    public required string LineShift      { get; init; }
    public required string ProcessCode    { get; init; }
    public required string ProcessName    { get; init; }
    public required string NgName         { get; init; }
    public required string QtyInput       { get; init; }
    public required string QtyNg          { get; init; }
    public required string ProductDate    { get; init; }
}

/// <summary>Routing row frozen at fetch time; ProcessName / ProcessType already normalized.</summary>
public sealed record NgRateRoutingEntry(string ModelName, string ProcessCode, string ProcessName, string ProcessType);

/// <summary>Reason row frozen at fetch time; ProcessName / NgName already normalized.</summary>
public sealed record NgRateReasonEntry(string ProcessName, string NgName, string Reason);

/// <summary>
/// The merged NG rows of one NG Rate fetch (API rows + NgRateCache rows) plus the Routing /
/// Reason tables as they were at fetch time. Replaces the per-request temp_*.db file: the rows
/// stay in memory and the report services read them from here.
/// </summary>
public sealed class NgRateDataSet
{
    public required string Handle { get; init; }
    public required DateTime CreatedAt { get; init; }
    public required IReadOnlyList<NgRateDataRow> Rows { get; init; }
    public required IReadOnlyList<NgRateRoutingEntry> Routing { get; init; }
    public required IReadOnlyList<NgRateReasonEntry> Reasons { get; init; }

    private static long _sequence;

    /// <summary>Builds the data set from fetched rows (column names as in NgRateCacheSchema).
    /// A missing column reads as empty text, like the old temp DB insert. Repeated values are
    /// shared so a few hundred thousand rows stay small.</summary>
    public static NgRateDataSet Create(
        IReadOnlyList<Dictionary<string, string>> rows,
        IReadOnlyList<NgRateRoutingEntry> routing,
        IReadOnlyList<NgRateReasonEntry> reasons)
    {
        var pool = new Dictionary<string, string>(StringComparer.Ordinal);
        string Share(string value)
        {
            if (pool.TryGetValue(value, out string? shared)) return shared;
            pool[value] = value;
            return value;
        }
        string Col(Dictionary<string, string> row, string column) =>
            Share(row.TryGetValue(column, out string? value) && value is not null ? value : string.Empty);

        var dataRows = new List<NgRateDataRow>(rows.Count);
        foreach (var row in rows)
        {
            string material = Col(row, "MATERIALNAME");
            string line = Col(row, "PRODUCTION_LINE");
            dataRows.Add(new NgRateDataRow
            {
                MaterialName   = material,
                ProductionLine = line,
                LineShift      = Share(material + "_" + line),
                ProcessCode    = Col(row, "PROCESSCODE"),
                ProcessName    = Col(row, "PROCESSNAME"),
                NgName         = Col(row, "NGNAME"),
                QtyInput       = Col(row, "QTYINPUT"),
                QtyNg          = Col(row, "QTYNG"),
                ProductDate    = Col(row, "PRODUCT_DATE"),
            });
        }

        DateTime now = DateTime.Now;
        long sequence = Interlocked.Increment(ref _sequence);
        return new NgRateDataSet
        {
            Handle = $"ngrate-data/{now:yyyyMMdd-HHmmss}-{sequence}",
            CreatedAt = now,
            Rows = dataRows,
            Routing = routing,
            Reasons = reasons,
        };
    }
}

/// <summary>
/// Keeps the most recent NG Rate data sets in memory, keyed by <see cref="NgRateDataSet.Handle"/>.
/// Pages and report services pass the handle around where they used to pass the temp_*.db path.
/// </summary>
public static class NgRateDataStore
{
    /// <summary>How many fetches stay readable. The old flow deleted every earlier temp DB on each
    /// fetch; keeping a few lets a page that fetched just before another page still re-run its report.</summary>
    public const int Capacity = 4;

    private static readonly object Gate = new();
    private static readonly List<NgRateDataSet> Sets = [];

    public static void Add(NgRateDataSet set)
    {
        lock (Gate)
        {
            Sets.RemoveAll(s => s.Handle == set.Handle);
            Sets.Add(set);
            while (Sets.Count > Capacity)
                Sets.RemoveAt(0);
        }
    }

    public static bool Contains(string? handle) => TryGet(handle, out _);

    public static bool TryGet(string? handle, out NgRateDataSet set)
    {
        lock (Gate)
        {
            set = Sets.LastOrDefault(s => string.Equals(s.Handle, handle, StringComparison.Ordinal))!;
            return set is not null;
        }
    }

    /// <summary>The data set for <paramref name="handle"/>; throws when it has been dropped
    /// (more than <see cref="Capacity"/> newer fetches) or the server restarted.</summary>
    public static NgRateDataSet Get(string? handle) =>
        TryGet(handle, out var set)
            ? set
            : throw new InvalidOperationException("NG Rate data is no longer in memory. Run the query again.");

    public static string? LatestHandle()
    {
        lock (Gate)
            return Sets.Count == 0 ? null : Sets[^1].Handle;
    }
}
