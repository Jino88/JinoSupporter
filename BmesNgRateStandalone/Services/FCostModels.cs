namespace BmesNgRateStandalone.Services;

/// <summary>One row from the BMES MES072400 (F-Cost) endpoint. The first three rows of a
/// typical response are the GN-factory totals (ZType = INAMT / FCOST / FRATE with empty
/// model+material). Subsequent rows repeat the same three ZTypes per (ModNo, Material).
/// COL0001..COL0014 are pre-aggregated period values whose meaning is defined server-side
/// (today / week-to-date / month-to-date / year-to-date / etc.).
///
/// JOIN KEYS to NG Rate DB (per user, 2026-05-05):
///   MatnrTx  ↔ MaterialName  (NG Rate's MAKTX)
///   VeridTx  ↔ LineShift     (NG Rate's PRODUCTION_LINE)
/// Both fields are forward-filled across the INAMT/FCOST/FRATE triplet at fetch time
/// (see FCostService.ForwardFillTriplets) so every row carries its own join keys.
/// </summary>
public sealed class FCostRow
{
    public long   Id        { get; set; }
    public string FetchedAt { get; set; } = string.Empty;
    public string QueryDate { get; set; } = string.Empty;
    public string RawJson   { get; set; } = string.Empty;

    // Display-side metadata (suffix _TX in the BMES JSON).
    public string FaccoTx { get; set; } = string.Empty;
    public string PrdGrTx { get; set; } = string.Empty;
    public string VeridTx { get; set; } = string.Empty;
    public string ModNoTx { get; set; } = string.Empty;
    public string AssemTx { get; set; } = string.Empty;
    public string AbChgTx { get; set; } = string.Empty;
    public string MCodeTx { get; set; } = string.Empty;
    public string MatnrTx { get; set; } = string.Empty;
    public string ZTypeTx { get; set; } = string.Empty;   // "Input Amount" / "F-Cost" / "F-Cost Rate(%)"

    public string TSort   { get; set; } = string.Empty;
    public string ZSort   { get; set; } = string.Empty;
    public string ZType   { get; set; } = string.Empty;   // INAMT / FCOST / FRATE

    // Code-side metadata (no _TX suffix in the BMES JSON).
    public string Facco { get; set; } = string.Empty;
    public string Werks { get; set; } = string.Empty;
    public string PrdGr { get; set; } = string.Empty;
    public string ModNo { get; set; } = string.Empty;
    public string Verid { get; set; } = string.Empty;
    public string AbChg { get; set; } = string.Empty;
    public string Cat01 { get; set; } = string.Empty;
    public string Matnr { get; set; } = string.Empty;
    public string Assem { get; set; } = string.Empty;
    public string TwSyn { get; set; } = string.Empty;
    public string ZValu { get; set; } = string.Empty;

    public double Col0001 { get; set; }
    public double Col0002 { get; set; }
    public double Col0003 { get; set; }
    public double Col0004 { get; set; }
    public double Col0005 { get; set; }
    public double Col0006 { get; set; }
    public double Col0007 { get; set; }
    public double Col0008 { get; set; }
    public double Col0009 { get; set; }
    public double Col0010 { get; set; }
    public double Col0011 { get; set; }
    public double Col0012 { get; set; }
    public double Col0013 { get; set; }
    public double Col0014 { get; set; }

    public double[]? Values { get; set; }

    public int ValueCount => Values?.Length ?? 14;

    public double GetCol(int index)
    {
        if (Values is { Length: > 0 })
            return index >= 1 && index <= Values.Length ? Values[index - 1] : 0;

        return index switch
        {
            1 => Col0001, 2 => Col0002, 3 => Col0003, 4 => Col0004, 5 => Col0005,
            6 => Col0006, 7 => Col0007, 8 => Col0008, 9 => Col0009, 10 => Col0010,
            11 => Col0011, 12 => Col0012, 13 => Col0013, 14 => Col0014,
            _ => 0,
        };
    }

    /// <summary>User-requested key for SubGroup mapping: <c>MatnrTx + "_" + ModNoTx</c>.
    /// Empty when either side is blank (e.g., GN total rows).</summary>
    public string JoinKey =>
        string.IsNullOrEmpty(MatnrTx) || string.IsNullOrEmpty(ModNoTx)
            ? string.Empty
            : $"{MatnrTx}_{ModNoTx}";

    /// <summary>Backup key matching NG Rate's actual LineShift format
    /// (<c>MaterialName + "_" + PRODUCTION_LINE</c> = <c>MatnrTx + "_" + VeridTx</c>).
    /// Tried as a fallback when JoinKey doesn't match anything in ModelGroups.</summary>
    public string AltJoinKey =>
        string.IsNullOrEmpty(MatnrTx) || string.IsNullOrEmpty(VeridTx)
            ? string.Empty
            : $"{MatnrTx}_{VeridTx}";
}

/// <summary>Period metadata for one of the 14 COL columns. BMES emits a sibling
/// <c>BottomGridColumnList</c> array per response that maps each <c>COLnnnn</c> code to
/// the actual date / week / month it represents — captured here so the UI can render
/// real headers ("05-04", "26-W19", "26-05") instead of generic "C1..C14".</summary>
public sealed class FCostColumnMeta
{
    /// <summary>1..14 — matches FCostRow.GetCol(int) and ColXXXX numbering.</summary>
    public int    Index   { get; init; }
    public string Code    { get; init; } = string.Empty;   // ZPVTT — e.g. "20260504", "2026-19", "202605"
    public string Header  { get; init; } = string.Empty;   // ZPVTT_TX — e.g. "05-04", "26-W19", "26-05"
    public string PDate   { get; init; } = string.Empty;   // PDATE — normalized "YYYYMMDD" / "YYYYWW" / "YYYYMM"

    /// <summary>Inferred from the code shape: dates are 8 digits, weeks contain '-', months are 6 digits.</summary>
    public FCostPeriodKind Kind
    {
        get
        {
            if (string.IsNullOrEmpty(Code)) return FCostPeriodKind.Unknown;
            if (Code.Contains('-'))         return FCostPeriodKind.Week;
            if (Code.Length == 8)           return FCostPeriodKind.Day;
            if (Code.Length == 6)           return FCostPeriodKind.Month;
            return FCostPeriodKind.Unknown;
        }
    }
}

public enum FCostPeriodKind
{
    Unknown,
    Day,
    Week,
    Month,
}

public sealed class FCostRawBackfillResult
{
    public string   DbPath        { get; set; } = string.Empty;
    public DateTime StartDate     { get; set; }
    public DateTime EndDate       { get; set; }
    public int      AttemptedDays { get; set; }
    public int      FetchedDays   { get; set; }
    public int      SkippedDays   { get; set; }
    public int      FailedDays    { get; set; }
    public int      TotalRows     { get; set; }
    public List<string> Failures  { get; set; } = new();
}

public sealed class FCostRawStatus
{
    public string DbPath       { get; set; } = string.Empty;
    public bool   Exists       { get; set; }
    public int    PullCount    { get; set; }
    public int    SuccessCount { get; set; }
    public int    FailedCount  { get; set; }
    public int    TotalRows    { get; set; }
    public string FirstDate    { get; set; } = string.Empty;
    public string LastDate     { get; set; } = string.Empty;
}

/// <summary>One Material's three F-Cost rows (Input/F-Cost/Rate) collapsed into a single
/// view-model so the page can render a single table row per Material — Input across the
/// 14 period columns on one line, F-Cost on the next, Rate on the third (similar to how
/// the BMES screen presents it).</summary>
public sealed class FCostMaterialBlock
{
    /// <summary>Row label combining model + material so identical material codes under
    /// different models stay distinguishable in the table.</summary>
    public string DisplayName  { get; init; } = string.Empty;

    public string ProductGroup { get; init; } = string.Empty;
    public string ModelNo      { get; init; } = string.Empty;
    public string Material     { get; init; } = string.Empty;
    public string Verid        { get; init; } = string.Empty;

    /// <summary>Total of the 14-column INAMT line — used for sorting biggest-volume items first.</summary>
    public double InputTotal =>
        Input?.Values is { Length: > 0 } vals ? vals.Sum() : Input?.GetCol(8) ?? 0;

    public FCostRow? Input { get; set; }   // ZType = INAMT
    public FCostRow? Cost  { get; set; }   // ZType = FCOST
    public FCostRow? Rate  { get; set; }   // ZType = FRATE

    public bool IsTotal => string.IsNullOrEmpty(Material) && string.IsNullOrEmpty(ModelNo);
}

/// <summary>SubGroup-level aggregate: sums Input and F-Cost across every F-Cost row whose
/// JoinKey (MatnrTx_ModNoTx or MatnrTx_VeridTx — see <see cref="FCostRow.JoinKey"/> /
/// <see cref="FCostRow.AltJoinKey"/>) matches a LineShift inside the SubGroup's subtree.
/// Rate% is recomputed AFTER summing so it reflects the weighted total, not an average of
/// per-Material rates.</summary>
public sealed class FCostSubGroupAggregate
{
    public string Display      { get; init; } = string.Empty;   // SubGroup name (or material fallback)
    public string GroupName    { get; init; } = string.Empty;
    public string MidGroupName { get; init; } = string.Empty;
    public string SubGroupKey  { get; init; } = string.Empty;

    /// <summary>How many F-Cost rows actually contributed (count of distinct Material rows
    /// matched into this SubGroup). 0 means the SubGroup has no matching rows in this pull.</summary>
    public int    MatchedMaterialCount { get; set; }

    /// <summary>14-column sums. Index 0 = COL0001, index 13 = COL0014.</summary>
    public double[] InputByCol { get; init; } = new double[14];
    public double[] CostByCol  { get; init; } = new double[14];

    /// <summary>Rate% = Cost / Input * 100 per column. 0 if Input=0.</summary>
    public double[] RateByCol
    {
        get
        {
            var arr = new double[14];
            for (int i = 0; i < 14; i++)
                arr[i] = InputByCol[i] > 0 ? CostByCol[i] / InputByCol[i] * 100.0 : 0;
            return arr;
        }
    }
}

/// <summary>Top-level F-Cost report shape returned to the page.</summary>
public sealed class FCostReport
{
    public string                   DbPath      { get; init; } = string.Empty;
    public string                   QueryDate   { get; init; } = string.Empty;
    public DateTime                 GeneratedAt { get; init; } = DateTime.Now;
    public int                      TotalRows   { get; init; }

    /// <summary>Three-row "Total" block at the top — directly from the GN factory totals.</summary>
    public FCostMaterialBlock?      Total       { get; init; }

    /// <summary>Per-Material blocks ordered by BMES TSORT (input-volume desc by default).</summary>
    public List<FCostMaterialBlock> Materials   { get; init; } = new();

    /// <summary>Per-SubGroup aggregates rolled up from the matching Material rows. Computed
    /// only when the caller passes a model-group list into the report generator.</summary>
    public List<FCostSubGroupAggregate> SubGroups { get; init; } = new();

    /// <summary>Materials that had F-Cost data but no matching SubGroup in the supplied
    /// ModelGroups — surfaced so the user can spot mapping gaps (typo / new model / etc.).</summary>
    public List<FCostMaterialBlock>     UnmappedMaterials { get; init; } = new();

    /// <summary>Period metadata for each of the 14 COL columns, parsed from the response's
    /// BottomGridColumnList. Header/Code/Kind drive UI rendering — the report layer can
    /// group consecutive Day/Week/Month columns visually using <see cref="FCostColumnMeta.Kind"/>.</summary>
    public IReadOnlyList<FCostColumnMeta> Columns { get; init; } = Array.Empty<FCostColumnMeta>();

    /// <summary>Convenience: just the human-readable headers in column order.</summary>
    public IReadOnlyList<string> ColumnHeaders =>
        Columns.Count > 0
            ? Columns.Select(c => c.Header).ToArray()
            : new[] { "C1","C2","C3","C4","C5","C6","C7","C8","C9","C10","C11","C12","C13","C14" };
}
