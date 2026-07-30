using System.Net;
using System.Text;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Renders one LPA search result — the row list AND the date × AULOC NG pivot — into a
/// SINGLE self-contained static HTML file, the same way
/// <see cref="BmesReportHtmlExportService"/> does for the BMES report. The Blazor page
/// keeps only the search form and shows this file in an iframe; tab switching, folding,
/// filtering and detail expansion all happen inside the HTML with no server round-trip.
///
/// Everything the viewer can show is baked in at generation time, including each row's
/// MES073261 checklist — the file is a complete snapshot, so it stays readable after the
/// session ends and can be saved or mailed as-is. That is what makes the prefetch in
/// <see cref="BmesLpaScrapeService.FetchDetailsAsync"/> necessary.
///
/// Memory: detail blocks live in inert &lt;template&gt; elements and are cloned into the
/// table only while open, so a few hundred checklists cost DOM nodes only on demand.
/// </summary>
public sealed class BmesLpaHtmlExportService
{
    private const string ReportFileName = "lpa.html";

    private static string ExportRoot => AppStoragePaths.Combine("_temp", "bmes-lpa");

    /// <summary>Everything one generated file needs. Rows and details come straight from
    /// <see cref="BmesLpaScrapeService"/>; the query fields are shown in the header so a
    /// saved file still says what it was a search for.</summary>
    public sealed record ExportInput
    {
        public required IReadOnlyList<string> Columns { get; init; }
        public required IReadOnlyList<IReadOnlyDictionary<string, string>> Rows { get; init; }

        /// <summary>LQRNO → its checklist result. Missing or failed entries degrade to a
        /// message inside that row's detail block.</summary>
        public required IReadOnlyDictionary<string, BmesLpaScrapeService.LpaResult> Details { get; init; }

        public DateTime From { get; init; }
        public DateTime To { get; init; }
        public string Facco { get; init; } = string.Empty;

        /// <summary>Where the viewer's photos come from. The web app serves downscaled copies
        /// from its own <c>/bmes/lpa/img</c> route (default); the WPF standalone has no such
        /// route (BlazorWebView, no server endpoints), so it points each photo straight at the
        /// anonymous BMES GetImage URL instead.</summary>
        public bool DirectBmesImages { get; init; }
    }

    /// <summary>Validate token and return the on-disk file path, or null.</summary>
    public string? ResolveReportFile(string token)
    {
        if (!IsValidToken(token)) return null;
        string path = Path.GetFullPath(Path.Combine(ExportRoot, token, ReportFileName));
        string root = Path.GetFullPath(ExportRoot);
        if (!path.StartsWith(root, StringComparison.OrdinalIgnoreCase)) return null; // path-escape guard
        return File.Exists(path) ? path : null;
    }

    private static bool IsValidToken(string token) =>
        !string.IsNullOrEmpty(token) &&
        token.Length is > 0 and <= 64 &&
        token.All(c => (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f'));

    /// <summary>Write the file under a fresh token folder (dropping earlier runs) and
    /// return the token the viewer route is keyed by.</summary>
    public async Task<string> GenerateAsync(ExportInput input)
    {
        string token = Guid.NewGuid().ToString("N");
        string dir = Path.Combine(ExportRoot, token);
        Directory.CreateDirectory(dir);

        string html = BuildHtml(input);
        await File.WriteAllTextAsync(Path.Combine(dir, ReportFileName), html, new UTF8Encoding(false));

        CleanupOldTokens(token);
        return token;
    }

    /// <summary>The viewer HTML as a string, for a host with no file/route to serve it (the WPF
    /// standalone drops it straight into an <c>&lt;iframe srcdoc&gt;</c>).</summary>
    public string RenderHtml(ExportInput input) => BuildHtml(input);

    // ── Column rules (moved off the page: the table now only exists here) ────────────

    /// <summary>
    /// Columns dropped from the table on request. They are hidden from the view only — the
    /// values stay in each row because the export still reads several of them: LQRNO keys
    /// the detail lookup, LASEQ and AUDAT/AULOC drive the sort and grouping, and FACCO/DICNO
    /// feed the detail query's conditions.
    /// </summary>
    private static readonly HashSet<string> HiddenColumns = new(StringComparer.OrdinalIgnoreCase)
    {
        "LQRNO", "LQBNO", "LOBVE",
        "FACCO", "DICNO",
        "LASEQ", "DICAD",
        "MATNR", "WERKS",
        "ZSTAT", "USEYN",
        "ERNAM", "ERDAT",
        // Dropped on request. VERID stays hidden but VERID_TX (its label) is kept.
        "NACNT", "ERNAM_TX", "VERID",
    };

    /// <summary>Synthetic column: NG rate in ppm, inserted right after NG (NGCNT). It has no
    /// backing field in the row — the cell value is computed from NGCNT and TOTAL.</summary>
    private const string NgRateColumn = "NG RATE";

    private static List<string> DisplayColumns(IReadOnlyList<string> columns)
    {
        var cols = columns.Where(c => !HiddenColumns.Contains(c)).ToList();
        int ng = cols.IndexOf("NGCNT");
        if (ng >= 0) cols.Insert(ng + 1, NgRateColumn);
        return cols;
    }

    /// <summary>The numeric columns the model total row fills in (and that get right-aligned).</summary>
    private static bool IsTotalColumn(string col)
        => col is "TOTAL" or "OKCNT" or "NGCNT" or NgRateColumn;

    /// <summary>Display header: OKCNT→OK, NGCNT→NG; everything else verbatim.</summary>
    private static string ColumnHeader(string col) => col switch
    {
        "OKCNT" => "OK",
        "NGCNT" => "NG",
        _ => col,
    };

    private static string Cell(IReadOnlyDictionary<string, string> row, string key)
        => row.TryGetValue(key, out string? v) ? v ?? string.Empty : string.Empty;

    private static int Count(IReadOnlyDictionary<string, string> row, string col)
        => int.TryParse(Cell(row, col).Trim(), out int v) ? v : 0;

    /// <summary>NG rate in ppm = NG / TOTAL × 1,000,000. "-" when TOTAL is 0 (no denominator).</summary>
    public static string NgRateText(int ng, int total)
        => total > 0 ? Math.Round(ng * 1_000_000.0 / total).ToString("N0") : "-";

    /// <summary>The audit number a detail lookup keys off. The list endpoint's exact column
    /// name is not documented, so fall back to any column whose name contains LQRNO.</summary>
    public static string RowLqrno(IReadOnlyDictionary<string, string> row)
    {
        if (row.TryGetValue("LQRNO", out string? direct) && !string.IsNullOrWhiteSpace(direct))
            return direct;

        foreach (var kv in row)
            if (kv.Key.Contains("LQRNO", StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(kv.Value))
                return kv.Value;

        return string.Empty;
    }

    /// <summary>Rows ordered audit date → checklist name → layer, newest date first.
    /// AUDAT is ISO and LASEQ is zero-padded (LA01/LA02/LA03), so ordinal order is already
    /// chronological and layer order.</summary>
    private static List<IReadOnlyDictionary<string, string>> Ordered(
        IReadOnlyList<IReadOnlyDictionary<string, string>> rows)
        => rows
            .OrderByDescending(r => Cell(r, "AUDAT"), StringComparer.Ordinal)
            .ThenBy(r => Cell(r, "AULOC"), StringComparer.Ordinal)
            .ThenBy(r => Cell(r, "LASEQ"), StringComparer.Ordinal)
            .ThenBy(r => Cell(r, "LQRNO"), StringComparer.Ordinal)
            .ToList();

    private static string Enc(string? s) => WebUtility.HtmlEncode(s ?? string.Empty);

    // ── List tab ────────────────────────────────────────────────────────────────────

    /// <summary>
    /// The row list, with a group row per audit date and a roll-up row per AULOC inside it.
    ///
    /// Every row carries its date/model key and its raw TOTAL/OK/NG as data attributes: the
    /// viewer re-derives group counts and the model totals from whatever survives the text
    /// filter, so a filtered view totals only what it shows (which is what the interactive
    /// page did server-side).
    /// </summary>
    private static string BuildListTab(
        ExportInput input,
        List<IReadOnlyDictionary<string, string>> rows,
        List<string> cols,
        StringBuilder templates)
    {
        var sb = new StringBuilder();
        int span = cols.Count + 2;                       // + index + detail button
        int firstNumIdx = cols.FindIndex(IsTotalColumn);
        int labelSpan = firstNumIdx < 0 ? span : firstNumIdx + 2;

        // The same LQRNO can appear on several rows; its checklist is emitted once.
        var emitted = new HashSet<string>(StringComparer.Ordinal);

        sb.Append("<table class=\"lpa-table\" id=\"lpa-list\"><thead><tr>")
          .Append("<th style=\"width:52px;\">#</th><th style=\"width:56px;\"></th>");
        foreach (string c in cols)
            sb.Append("<th class=\"").Append(IsTotalColumn(c) ? "lpa-num" : "")
              .Append("\">").Append(Enc(ColumnHeader(c))).Append("</th>");
        sb.Append("</tr></thead><tbody>");

        string? lastDate = null, lastModel = null;
        foreach (var row in rows)
        {
            string date = Cell(row, "AUDAT");
            string model = Cell(row, "AULOC");
            string modelKey = date + " " + model;

            // Rows are ordered date → AULOC → layer, so a change of value is where a group
            // starts; no separate grouping pass needed.
            if (!string.Equals(date, lastDate, StringComparison.Ordinal))
            {
                lastDate = date;
                lastModel = null;                        // restart model runs per date
                sb.Append("<tr class=\"lpa-group-row\" data-group=\"date\" data-date=\"")
                  .Append(Enc(date)).Append("\"><td colspan=\"").Append(span).Append("\">")
                  .Append("<span class=\"lpa-caret\">▼</span><strong>")
                  .Append(Enc(date.Length > 0 ? date : "(No date)"))
                  .Append("</strong><span class=\"lpa-muted lpa-count\"></span></td></tr>");
            }

            if (!string.Equals(model, lastModel, StringComparison.Ordinal))
            {
                lastModel = model;
                sb.Append("<tr class=\"lpa-model-row\" data-group=\"model\" data-date=\"")
                  .Append(Enc(date)).Append("\" data-model=\"").Append(Enc(modelKey))
                  .Append("\"><td class=\"lpa-model-label\" colspan=\"").Append(labelSpan).Append("\">")
                  .Append("<span class=\"lpa-model-indent\"></span><span class=\"lpa-caret\">▼</span><strong>")
                  .Append(Enc(model.Length > 0 ? model : "(No model)"))
                  .Append("</strong><span class=\"lpa-muted lpa-count\"></span></td>");
                if (firstNumIdx >= 0)
                    for (int ci = firstNumIdx; ci < cols.Count; ci++)
                        sb.Append("<td class=\"lpa-num lpa-total-cell\" data-tot=\"")
                          .Append(Enc(cols[ci])).Append("\"></td>");
                sb.Append("</tr>");
            }

            string lqrno = RowLqrno(row);
            int total = Count(row, "TOTAL"), ok = Count(row, "OKCNT"), ng = Count(row, "NGCNT");

            sb.Append("<tr class=\"lpa-data-row\" data-date=\"").Append(Enc(date))
              .Append("\" data-model=\"").Append(Enc(modelKey))
              .Append("\" data-lqrno=\"").Append(Enc(lqrno))
              .Append("\" data-total=\"").Append(total)
              .Append("\" data-ok=\"").Append(ok)
              .Append("\" data-ng=\"").Append(ng)
              .Append("\"><td class=\"lpa-muted lpa-seq-cell\"></td>")
              .Append("<td><button type=\"button\" class=\"lpa-detail-btn\">Detail</button></td>");

            foreach (string c in cols)
            {
                string text = c == NgRateColumn ? NgRateText(ng, total) : Cell(row, c);
                sb.Append("<td class=\"").Append(IsTotalColumn(c) ? "lpa-num" : "")
                  .Append("\">").Append(Enc(text)).Append("</td>");
            }
            sb.Append("</tr>");

            if (lqrno.Length > 0 && emitted.Add(lqrno))
                AppendDetailTemplate(templates, lqrno, input.Details, input.DirectBmesImages);
        }

        sb.Append("</tbody></table>");
        return sb.ToString();
    }

    /// <summary>One row's checklist, parked in an inert template until it is opened.</summary>
    private static void AppendDetailTemplate(
        StringBuilder templates,
        string lqrno,
        IReadOnlyDictionary<string, BmesLpaScrapeService.LpaResult> details,
        bool directImages)
    {
        templates.Append("<template data-detail=\"").Append(Enc(lqrno)).Append("\">");
        templates.Append("<div class=\"lpa-detail-inline\"><div class=\"lpa-detail-inline-head\">")
                 .Append("<span class=\"lpa-strong\">Detail - ").Append(Enc(lqrno)).Append("</span>");

        details.TryGetValue(lqrno, out BmesLpaScrapeService.LpaResult? result);

        if (result is { IsSuccess: true })
            templates.Append("<span class=\"lpa-muted\">MES073261 · ")
                     .Append(result.Rows.Count.ToString("N0")).Append(" item(s)</span>");
        templates.Append("<button type=\"button\" class=\"lpa-detail-close\">Close</button></div>");

        if (result is null)
            templates.Append("<div class=\"lpa-note\">Detail was not prefetched.</div>");
        else if (!result.IsSuccess)
            templates.Append("<div class=\"lpa-error\">").Append(Enc(result.Error)).Append("</div>");
        else if (result.Rows.Count == 0)
            templates.Append("<div class=\"lpa-note\">No detail rows.</div>");
        else
            foreach (var item in result.Rows)
                AppendDetailItem(templates, item, lqrno, directImages);

        templates.Append("</div></template>");
    }

    /// <summary>Fields already shown in an item's header or rendered as photos, so the field
    /// grid skips them.</summary>
    private static readonly string[] DetailHeaderKeys = ["LORSQ", "TYPRC", "IMPLV", "RESUT", "ZIMAG_TX"];

    private static void AppendDetailItem(
        StringBuilder sb,
        IReadOnlyDictionary<string, string> item,
        string lqrno,
        bool directImages)
    {
        sb.Append("<div class=\"lpa-detail-item\"><div class=\"lpa-detail-head\">")
          .Append("<span class=\"lpa-seq\">").Append(Enc(Cell(item, "LORSQ"))).Append("</span>")
          .Append("<strong>").Append(Enc(Cell(item, "TYPRC"))).Append("</strong>");

        if (Cell(item, "IMPLV").Length > 0)
            sb.Append("<span class=\"lpa-badge\">").Append(Enc(Cell(item, "IMPLV"))).Append("</span>");
        string resut = Cell(item, "RESUT");
        if (resut.Trim().Length > 0)
        {
            // The old page painted every result green, which made a failed check read as a
            // pass at a glance. Coloured by code now (see ResutBadgeClass) and shown with
            // the BMES label, because a bare "B" says nothing on its own.
            sb.Append("<span class=\"lpa-badge ").Append(ResutBadgeClass(resut)).Append("\">")
              .Append(Enc(ResutLabel(resut))).Append("</span>");
        }

        // Photos ride in the item header, next to the result they are evidence for.
        if (ItemImagePaths(item).Count > 0)
        {
            sb.Append("<span class=\"lpa-thumbs\">");
            AppendThumbs(sb, item, ImageCaption(lqrno, string.Empty, string.Empty, item), directImages);
            sb.Append("</span>");
        }

        sb.Append("</div><dl class=\"lpa-detail-grid\">");
        // Field names are shown verbatim: the response schema is not documented anywhere,
        // so inventing Korean labels would be a guess.
        foreach (var kv in item)
        {
            if (string.IsNullOrWhiteSpace(kv.Value) || DetailHeaderKeys.Contains(kv.Key)) continue;
            sb.Append("<dt>").Append(Enc(kv.Key)).Append("</dt><dd>").Append(Enc(kv.Value)).Append("</dd>");
        }
        sb.Append("</dl></div>");
    }

    // ── Result photos (ZIMAG_TX) ────────────────────────────────────────────────────

    /// <summary>
    /// The photo paths on one checklist item. ZIMAG_TX holds the primary images — up to
    /// two, comma separated (the BMES grid's <c>maxPrimaryCount</c>) — while ZIMAG is just
    /// how many were uploaded, so the paths are the only thing worth reading.
    /// </summary>
    public static List<string> ItemImagePaths(IReadOnlyDictionary<string, string> item)
        => Cell(item, "ZIMAG_TX")
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();

    /// <summary>
    /// Thumbnails for one item, each opening the popup. The photo is not embedded: the
    /// <c>src</c> points at <c>/bmes/lpa/img</c>, which downloads and downscales it the first
    /// time it is shown and serves the stored copy afterwards. <c>loading="lazy"</c> means a
    /// thumbnail inside a collapsed detail or below the fold never leaves BMES until it is
    /// actually scrolled to (or its checklist opened) — which is what stopped Search from
    /// blocking on every photo of a search. If the photo is gone, <c>onerror</c> swaps in a
    /// link to the live original.
    /// </summary>
    private static void AppendThumbs(
        StringBuilder sb,
        IReadOnlyDictionary<string, string> item,
        string caption,
        bool directImages)
    {
        foreach (string path in ItemImagePaths(item))
        {
            // Web: the downscaled thumbnail from our own route. WPF: the anonymous BMES original
            // directly (no route to serve a resized copy). Both keep data-img so the lightbox and
            // the onerror fallback resolve the same original.
            string src = directImages
                ? BmesLpaImageService.ImageUrlPrefix + Uri.EscapeDataString(path)
                : "/bmes/lpa/img?size=thumb&path=" + Uri.EscapeDataString(path);
            sb.Append("<img class=\"lpa-thumb\" loading=\"lazy\" src=\"").Append(Enc(src))
              .Append("\" data-img=\"").Append(Enc(path))
              .Append("\" data-cap=\"").Append(Enc(caption))
              .Append("\" alt=\"\" title=\"Click to enlarge\" onerror=\"lpaImgFail(this)\">");
        }
    }

    /// <summary>What the popup prints under the photo: enough to say which audit and which
    /// check item it belongs to once it fills the screen on its own.</summary>
    private static string ImageCaption(string lqrno, string date, string model, IReadOnlyDictionary<string, string> item)
    {
        var parts = new List<string>(4);
        if (lqrno.Length > 0) parts.Add(lqrno);
        if (date.Length > 0) parts.Add(date);
        if (model.Length > 0) parts.Add(model);
        string seq = Cell(item, "LORSQ"), type = Cell(item, "TYPRC");
        if (seq.Length > 0 || type.Length > 0) parts.Add((seq.Length > 0 ? seq + ". " : "") + type);
        return string.Join(" · ", parts);
    }

    // ── Pivot tab ───────────────────────────────────────────────────────────────────

    private readonly record struct Bucket(int Total, int Ok, int Ng)
    {
        public Bucket Add(int total, int ok, int ng) => new(Total + total, Ok + ok, Ng + ng);
    }

    /// <summary>Period bucket keys, matching the conventions
    /// <see cref="NgRateReportService"/> already uses so a week here means the same week
    /// it does on the BMES report (FirstDay rule, Monday start — not ISO).</summary>
    private static string WeekKey(DateTime d) =>
        $"W:{d.Year:0000}{System.Globalization.CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(d, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Monday):00}";

    private static string MonthKey(DateTime d) => $"M:{d.Year:0000}{d.Month:00}";

    private static string PeriodHeader(string key)
    {
        if (key.StartsWith("W:", StringComparison.Ordinal))
            return int.TryParse(key[6..], out int w) ? "W" + w : key;
        if (key.StartsWith("M:", StringComparison.Ordinal))
            return int.TryParse(key[6..], out int m) ? "M" + m : key;
        // InvariantCulture matters: "/" in a format string is the *culture's* date
        // separator, so on this ko-KR server "MM/dd" would quietly render as "07-23".
        return DateTime.TryParse(key, out DateTime d)
             ? d.ToString("MM/dd", System.Globalization.CultureInfo.InvariantCulture)
             : key.Length >= 5 ? key[^5..] : key;
    }

    /// <summary>The synthetic leading column: the model's figure over the whole search
    /// range. It is also the row sort key, so it is shown rather than left implicit.</summary>
    private const string OverallKey = "ALL";

    private readonly record struct PeriodBlock(string Name, string Label, List<string> Keys);

    /// <summary>The model pivot's markup plus the two things the Check Item pivot underneath has
    /// to reuse: the exact period columns (so both tables scroll the same days/weeks/months) and
    /// the per-period Total, which is the denominator every check-item rate is measured against.</summary>
    private sealed record PivotRender(string Html, List<PeriodBlock> Blocks, Dictionary<string, Bucket> ColTotals);

    /// <summary>
    /// AULOC × period cross-tab. Rows are models ordered by their total inspected count
    /// (largest first, then name); columns run newest → oldest within three blocks:
    /// Date, Week, Month.
    ///
    /// Each cell shows NG RATE in ppm with the raw NG/TOTAL underneath, because a rate
    /// alone hides how much it rests on — 1/1 and 1000/1000000 are both "1,000,000 ppm"
    /// and "0/50" is not the same evidence as "0/50000".
    ///
    /// A blank cell means "no audit for that model in that period" — distinct from an
    /// audit that found nothing, which is a real 0.
    /// </summary>
    private static PivotRender BuildPivotTab(List<IReadOnlyDictionary<string, string>> rows)
    {
        var cells = new Dictionary<(string Auloc, string Col), Bucket>();
        var rowTotals = new Dictionary<string, Bucket>(StringComparer.Ordinal);
        var colTotals = new Dictionary<string, Bucket>(StringComparer.Ordinal);
        var dateKeys = new SortedSet<string>(StringComparer.Ordinal);
        var weekKeys = new SortedSet<string>(StringComparer.Ordinal);
        var monthKeys = new SortedSet<string>(StringComparer.Ordinal);

        void Accumulate(string auloc, string col, int total, int ok, int ng)
        {
            cells.TryGetValue((auloc, col), out Bucket b);
            cells[(auloc, col)] = b.Add(total, ok, ng);
            colTotals.TryGetValue(col, out Bucket c);
            colTotals[col] = c.Add(total, ok, ng);
        }

        foreach (var r in rows)
        {
            string auloc = Cell(r, "AULOC");
            if (auloc.Length == 0) auloc = "(No model)";
            string date = Cell(r, "AUDAT");
            int total = Count(r, "TOTAL"), ok = Count(r, "OKCNT"), ng = Count(r, "NGCNT");

            rowTotals.TryGetValue(auloc, out Bucket rt);
            rowTotals[auloc] = rt.Add(total, ok, ng);
            Accumulate(auloc, OverallKey, total, ok, ng);

            if (date.Length > 0)
            {
                dateKeys.Add(date);
                Accumulate(auloc, date, total, ok, ng);
            }

            // A row whose AUDAT will not parse still counts in the date column (the raw
            // string keys it) but cannot be placed in a week or month.
            if (DateTime.TryParse(date, out DateTime dt))
            {
                string wk = WeekKey(dt), mk = MonthKey(dt);
                weekKeys.Add(wk);
                monthKeys.Add(mk);
                Accumulate(auloc, wk, total, ok, ng);
                Accumulate(auloc, mk, total, ok, ng);
            }
        }

        if (rowTotals.Count == 0)
            return new PivotRender("<div class=\"lpa-note\">No rows to aggregate.</div>", [], colTotals);

        // Newest first in every block.
        var blocks = new List<PeriodBlock>
        {
            new("date",  "Date", dateKeys.Reverse().ToList()),
            new("week",  "Week", weekKeys.Reverse().ToList()),
            new("month", "Month",   monthKeys.Reverse().ToList()),
        };

        var models = rowTotals.Keys
            .OrderByDescending(a => rowTotals[a].Total)
            .ThenBy(a => a, StringComparer.Ordinal)
            .ToList();

        var sb = new StringBuilder();
        sb.Append("<table class=\"lpa-table lpa-pivot\" id=\"lpa-pivot\">");

        // ── header: block band, then the period columns ──
        sb.Append("<thead><tr class=\"lpa-pivot-band\"><th class=\"lpa-pivot-corner\" rowspan=\"2\">Model</th>")
          .Append("<th class=\"lpa-num lpa-pivot-sum\" rowspan=\"2\">Overall</th>");
        foreach (var blk in blocks)
        {
            if (blk.Keys.Count == 0) continue;
            sb.Append("<th class=\"sep-th\" data-block=\"").Append(blk.Name).Append("\" rowspan=\"2\"></th>")
              .Append("<th class=\"lpa-num lpa-band\" data-block=\"").Append(blk.Name)
              .Append("\" colspan=\"").Append(blk.Keys.Count).Append("\">").Append(blk.Label).Append("</th>");
        }
        sb.Append("</tr><tr class=\"lpa-pivot-head\">");
        foreach (var blk in blocks)
            for (int i = 0; i < blk.Keys.Count; i++)
                sb.Append("<th class=\"lpa-num\" data-block=\"").Append(blk.Name)
                  .Append("\" data-idx=\"").Append(i).Append("\">")
                  .Append(Enc(PeriodHeader(blk.Keys[i]))).Append("</th>");
        sb.Append("</tr>");

        // Total leads instead of trailing: it is the number that gets read first, and as a
        // third header row it stays pinned while the model rows scroll under it. It sits in
        // <thead> so the text filter, which only walks tBodies[0], cannot hide the total.
        sb.Append("<tr class=\"lpa-pivot-total\"><th class=\"lpa-pivot-row-head\">Total</th>");
        AppendPivotCell(sb, colTotals.GetValueOrDefault(OverallKey), "lpa-num lpa-pivot-sum", null, null);
        foreach (var blk in blocks)
        {
            if (blk.Keys.Count == 0) continue;
            sb.Append("<td class=\"sep-td\" data-block=\"").Append(blk.Name).Append("\"></td>");
            for (int i = 0; i < blk.Keys.Count; i++)
                AppendPivotCell(sb, colTotals.TryGetValue(blk.Keys[i], out Bucket b) ? b : null,
                                "lpa-num lpa-pivot-sum", blk.Name, i);
        }
        sb.Append("</tr></thead><tbody>");

        foreach (string m in models)
        {
            sb.Append("<tr><th class=\"lpa-pivot-row-head\">").Append(Enc(m)).Append("</th>");
            AppendPivotCell(sb, cells.GetValueOrDefault((m, OverallKey)), "lpa-num lpa-pivot-sum", null, null);
            foreach (var blk in blocks)
            {
                if (blk.Keys.Count == 0) continue;
                sb.Append("<td class=\"sep-td\" data-block=\"").Append(blk.Name).Append("\"></td>");
                for (int i = 0; i < blk.Keys.Count; i++)
                    AppendPivotCell(sb, cells.TryGetValue((m, blk.Keys[i]), out Bucket b) ? b : null,
                                    "lpa-num", blk.Name, i);
            }
            sb.Append("</tr>");
        }

        sb.Append("</tbody></table>");

        return new PivotRender(sb.ToString(), blocks, colTotals);
    }

    /// <summary>
    /// Check Item × period cross-tab, sitting under the model pivot with the same columns.
    /// Rows are the distinct Check Item (LCITM) texts, worst first, merged across every model —
    /// the model pivot says which model is bad, this says which check is failing.
    ///
    /// The denominator is how many times THAT check item was actually inspected — not the
    /// inspected piece quantity the model pivot divides by. A checklist line is answered once
    /// per audit, so "16/40" means the check ran 40 times and failed 16; dividing it by the
    /// pieces produced (78 on a day) would be comparing unlike things.
    ///
    /// Counts come from <see cref="BuildCheckPivotRows"/>, which de-duplicates by LQRNO the same
    /// way Defect Detail does (one checklist can appear on several list rows).
    /// </summary>
    private static string BuildCheckPivotTab(
        List<IReadOnlyDictionary<string, string>> rows,
        IReadOnlyDictionary<string, BmesLpaScrapeService.LpaResult> details,
        List<PeriodBlock> blocks,
        Dictionary<string, Bucket> colTotals)
    {
        IReadOnlyList<CheckPivotRow> checkRows = BuildCheckPivotRows(rows, details);
        if (checkRows.Count == 0 || blocks.Count == 0)
            return string.Empty;

        // Blank = that check was not run in that period. Run-and-passed is a real 0.
        static Bucket? CellFor(CheckPivotRow row, string col)
        {
            int checkedCount = row.CheckedByKey.GetValueOrDefault(col);
            return checkedCount == 0 ? null : new Bucket(checkedCount, 0, row.NgByKey.GetValueOrDefault(col));
        }

        var sb = new StringBuilder();
        sb.Append("<div class=\"lpa-subhead\"><span class=\"lpa-strong\">NG Rate by Check Item</span>")
          .Append("<span class=\"lpa-muted\">ppm = NG count / how often that check ran (a different ")
          .Append("denominator from the table above, which counts pieces) · ")
          .Append(checkRows.Count.ToString("N0")).Append(" checks (including those with no NG)</span></div>");

        sb.Append("<table class=\"lpa-table lpa-pivot lpa-chkpivot\" id=\"lpa-chkpivot\">");

        sb.Append("<thead><tr class=\"lpa-pivot-band\"><th class=\"lpa-pivot-corner\" rowspan=\"2\">")
          .Append("Model · Check Point · Check Item</th>")
          .Append("<th class=\"lpa-num lpa-pivot-sum\" rowspan=\"2\">Overall</th>");
        foreach (var blk in blocks)
        {
            if (blk.Keys.Count == 0) continue;
            sb.Append("<th class=\"sep-th\" data-block=\"").Append(blk.Name).Append("\" rowspan=\"2\"></th>")
              .Append("<th class=\"lpa-num lpa-band\" data-block=\"").Append(blk.Name)
              .Append("\" colspan=\"").Append(blk.Keys.Count).Append("\">").Append(blk.Label).Append("</th>");
        }
        sb.Append("</tr><tr class=\"lpa-pivot-head\">");
        foreach (var blk in blocks)
            for (int i = 0; i < blk.Keys.Count; i++)
                sb.Append("<th class=\"lpa-num\" data-block=\"").Append(blk.Name)
                  .Append("\" data-idx=\"").Append(i).Append("\">")
                  .Append(Enc(PeriodHeader(blk.Keys[i]))).Append("</th>");
        sb.Append("</tr>");

        sb.Append("<tr class=\"lpa-pivot-total\"><th class=\"lpa-pivot-row-head\">Total</th>");
        AppendPivotCell(sb, TotalCell(null), "lpa-num lpa-pivot-sum", null, null);
        foreach (var blk in blocks)
        {
            if (blk.Keys.Count == 0) continue;
            sb.Append("<td class=\"sep-td\" data-block=\"").Append(blk.Name).Append("\"></td>");
            for (int i = 0; i < blk.Keys.Count; i++)
                AppendPivotCell(sb, TotalCell(blk.Keys[i]), "lpa-num lpa-pivot-sum", blk.Name, i);
        }
        sb.Append("</tr></thead><tbody>");

        foreach (var row in checkRows)
        {
            // Model → Check Point → Check Item, the order the Excel sheet's columns run in.
            sb.Append("<tr><th class=\"lpa-pivot-row-head lpa-chk-head\">");
            if (row.Models.Length > 0)
                sb.Append("<span class=\"lpa-chk-typrc\"><b>Model</b> ")
                  .Append(Enc(row.Models)).Append("</span>");
            if (row.CheckPoints.Length > 0)
                sb.Append("<span class=\"lpa-chk-typrc\"><b>Check Point</b> ")
                  .Append(Enc(row.CheckPoints)).Append("</span>");
            sb.Append("<span class=\"lpa-chk-item\">").Append(Enc(row.Item)).Append("</span>")
              .Append("</th>");

            AppendPivotCell(sb, OverallCell(row), "lpa-num lpa-pivot-sum", null, null);
            foreach (var blk in blocks)
            {
                if (blk.Keys.Count == 0) continue;
                sb.Append("<td class=\"sep-td\" data-block=\"").Append(blk.Name).Append("\"></td>");
                for (int i = 0; i < blk.Keys.Count; i++)
                    AppendPivotCell(sb, CellFor(row, blk.Keys[i]), "lpa-num", blk.Name, i);
            }
            sb.Append("</tr>");
        }

        sb.Append("</tbody></table>");
        return sb.ToString();

        static Bucket? OverallCell(CheckPivotRow row) =>
            row.OverallChecked == 0 ? null : new Bucket(row.OverallChecked, 0, row.OverallNg);

        // Total = all NG over all checks run in that period, across every row of this table.
        // A null key means the Overall column.
        Bucket? TotalCell(string? col)
        {
            int checkedCount = col is null
                ? checkRows.Sum(r => r.OverallChecked)
                : checkRows.Sum(r => r.CheckedByKey.GetValueOrDefault(col));
            if (checkedCount == 0) return null;
            int ng = col is null
                ? checkRows.Sum(r => r.OverallNg)
                : checkRows.Sum(r => r.NgByKey.GetValueOrDefault(col));
            return new Bucket(checkedCount, 0, ng);
        }
    }

    /// <summary>
    /// Every Check Item (LCITM) that was answered at least once, merged across models. For each
    /// period key (raw AUDAT for date, <c>W:</c>/<c>M:</c> for week/month)
    /// it carries BOTH how many times the check was answered and how many of those were NG, so a
    /// rate can be read straight off the row.
    ///
    /// A checklist line is answered once per audit, so the denominator counts checklist rows —
    /// deliberately NOT the inspected piece quantity the model pivot uses, which measures
    /// something else entirely. De-duplicated by LQRNO like <see cref="BuildNgMatrix"/>, since
    /// one checklist can appear on several list rows.
    /// </summary>
    public static IReadOnlyList<CheckPivotRow> BuildCheckPivotRows(
        List<IReadOnlyDictionary<string, string>> rows,
        IReadOnlyDictionary<string, BmesLpaScrapeService.LpaResult> details)
    {
        var ngByKey = new Dictionary<string, Dictionary<string, int>>(StringComparer.Ordinal);
        var checkedByKey = new Dictionary<string, Dictionary<string, int>>(StringComparer.Ordinal);
        var overallNg = new Dictionary<string, int>(StringComparer.Ordinal);
        var overallChecked = new Dictionary<string, int>(StringComparer.Ordinal);
        // One Check Item text can be reached from more than one Check Point, and from more than
        // one model; keep them all so the merged row still says where it came from.
        var checkPoints = new Dictionary<string, SortedSet<string>>(StringComparer.Ordinal);
        var models = new Dictionary<string, SortedSet<string>>(StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.Ordinal);

        static void Bump(Dictionary<string, Dictionary<string, int>> map, string item, string key)
        {
            if (!map.TryGetValue(item, out var byKey))
                map[item] = byKey = new Dictionary<string, int>(StringComparer.Ordinal);
            byKey[key] = byKey.GetValueOrDefault(key) + 1;
        }

        static void AddTo(Dictionary<string, SortedSet<string>> map, string item, string value)
        {
            if (value.Length == 0) return;
            if (!map.TryGetValue(item, out var set))
                map[item] = set = new SortedSet<string>(StringComparer.Ordinal);
            set.Add(value);
        }

        foreach (var r in rows)
        {
            string lqrno = RowLqrno(r);
            if (lqrno.Length == 0 || !seen.Add(lqrno)) continue;
            if (!details.TryGetValue(lqrno, out BmesLpaScrapeService.LpaResult? res) || !res.IsSuccess)
                continue;

            string date = Cell(r, "AUDAT");
            string model = Cell(r, "AULOC");
            if (model.Length == 0) model = "(no model)";

            bool parsed = DateTime.TryParse(date, out DateTime dt);
            string weekKey = parsed ? WeekKey(dt) : "";
            string monthKey = parsed ? MonthKey(dt) : "";

            foreach (var checklistRow in res.Rows)
            {
                string lcitm = Cell(checklistRow, "LCITM");
                string item = lcitm.Length > 0 ? lcitm : "(no text)";
                bool isNg = IsNgResult(Cell(checklistRow, "RESUT"));

                AddTo(checkPoints, item, Cell(checklistRow, "TYPRC"));
                AddTo(models, item, model);

                overallChecked[item] = overallChecked.GetValueOrDefault(item) + 1;
                if (isNg) overallNg[item] = overallNg.GetValueOrDefault(item) + 1;

                if (date.Length > 0)
                {
                    Bump(checkedByKey, item, date);
                    if (isNg) Bump(ngByKey, item, date);
                }
                if (parsed)
                {
                    Bump(checkedByKey, item, weekKey);
                    Bump(checkedByKey, item, monthKey);
                    if (isNg)
                    {
                        Bump(ngByKey, item, weekKey);
                        Bump(ngByKey, item, monthKey);
                    }
                }
            }
        }

        // Every check that was answered, not just the ones that failed — a check sitting at 0
        // is evidence too (it ran and passed), and hiding it makes the table look like the
        // checklist is shorter than it is.
        //
        // Ordered by the three label columns in the order they are displayed — Model, then Check Point,
        // then Check Item — so one model's failing checks sit together instead of the rows
        // hopping between models by NG count.
        string ModelsOf(string k) => models.TryGetValue(k, out var m) ? string.Join(" · ", m) : "";
        string PointsOf(string k) => checkPoints.TryGetValue(k, out var c) ? string.Join(" · ", c) : "";

        return overallChecked.Keys
            .OrderBy(ModelsOf, StringComparer.Ordinal)
            .ThenBy(PointsOf, StringComparer.Ordinal)
            .ThenBy(k => k, StringComparer.Ordinal)
            .Select(k => new CheckPivotRow(
                k,
                PointsOf(k),
                ModelsOf(k),
                overallNg.GetValueOrDefault(k),
                overallChecked[k],
                ngByKey.TryGetValue(k, out var ng) ? ng : new Dictionary<string, int>(StringComparer.Ordinal),
                checkedByKey.TryGetValue(k, out var ck) ? ck : new Dictionary<string, int>(StringComparer.Ordinal)))
            .ToList();
    }

    /// <summary>One Check Item merged across models: the Check Point(s) and Model(s) it appears
    /// under, and — over the whole range and per period key — how often it was checked and how
    /// often that check came back NG.</summary>
    public sealed record CheckPivotRow(
        string Item, string CheckPoints, string Models,
        int OverallNg, int OverallChecked,
        IReadOnlyDictionary<string, int> NgByKey,
        IReadOnlyDictionary<string, int> CheckedByKey);

    // ── Shared data for the Excel export ───────────────────────────────────────────────
    // The pivot HTML above and the pivot Excel sheet are built from the same accumulation
    // rules; BuildPivotData exposes the numbers (BuildPivotTab keeps its own copy of the loop
    // so the interactive table is not disturbed — the two must stay in step).

    public readonly record struct PivotBucket(int Ng, int Total);

    /// <summary>One period block (Date/Week/Month) limited to its newest columns, with the header
    /// labels already resolved.</summary>
    public sealed record PivotBlock(string Name, string Label, IReadOnlyList<string> Keys, IReadOnlyList<string> Headers);

    /// <summary>The Model × period pivot as numbers: models (worst first), the period blocks, and
    /// lookups for the Total row (<see cref="ColTotals"/> + <see cref="GrandTotal"/>), each
    /// model's Overall column (<see cref="ModelOverall"/>) and every (model, period) cell.</summary>
    public sealed record PivotData(
        IReadOnlyList<string> Models,
        IReadOnlyList<PivotBlock> Blocks,
        PivotBucket GrandTotal,
        IReadOnlyDictionary<string, PivotBucket> ColTotals,
        IReadOnlyDictionary<string, PivotBucket> ModelOverall,
        IReadOnlyDictionary<(string Model, string Key), PivotBucket> Cells);

    /// <summary>Aggregate the pivot and keep only the newest <paramref name="dateN"/> days,
    /// <paramref name="weekN"/> weeks and <paramref name="monthN"/> months (0 = all), matching
    /// the viewer's Date/Week/Month boxes.</summary>
    public static PivotData BuildPivotData(
        List<IReadOnlyDictionary<string, string>> rows, int dateN, int weekN, int monthN)
    {
        var cells = new Dictionary<(string Auloc, string Col), Bucket>();
        var rowTotals = new Dictionary<string, Bucket>(StringComparer.Ordinal);
        var colTotals = new Dictionary<string, Bucket>(StringComparer.Ordinal);
        var dateKeys = new SortedSet<string>(StringComparer.Ordinal);
        var weekKeys = new SortedSet<string>(StringComparer.Ordinal);
        var monthKeys = new SortedSet<string>(StringComparer.Ordinal);

        void Accumulate(string auloc, string col, int total, int ok, int ng)
        {
            cells.TryGetValue((auloc, col), out Bucket b);
            cells[(auloc, col)] = b.Add(total, ok, ng);
            colTotals.TryGetValue(col, out Bucket c);
            colTotals[col] = c.Add(total, ok, ng);
        }

        foreach (var r in rows)
        {
            string auloc = Cell(r, "AULOC");
            if (auloc.Length == 0) auloc = "(No model)";
            string date = Cell(r, "AUDAT");
            int total = Count(r, "TOTAL"), ok = Count(r, "OKCNT"), ng = Count(r, "NGCNT");

            rowTotals.TryGetValue(auloc, out Bucket rt);
            rowTotals[auloc] = rt.Add(total, ok, ng);
            Accumulate(auloc, OverallKey, total, ok, ng);

            if (date.Length > 0) { dateKeys.Add(date); Accumulate(auloc, date, total, ok, ng); }
            if (DateTime.TryParse(date, out DateTime dt))
            {
                string wk = WeekKey(dt), mk = MonthKey(dt);
                weekKeys.Add(wk); monthKeys.Add(mk);
                Accumulate(auloc, wk, total, ok, ng);
                Accumulate(auloc, mk, total, ok, ng);
            }
        }

        static List<string> Newest(SortedSet<string> keys, int n)
        {
            IEnumerable<string> newest = keys.Reverse();
            return (n > 0 ? newest.Take(n) : newest).ToList();
        }

        PivotBlock Blk(string name, string label, SortedSet<string> keys, int n)
        {
            var k = Newest(keys, n);
            return new PivotBlock(name, label, k, k.Select(PeriodHeader).ToList());
        }

        var blocks = new List<PivotBlock>
        {
            Blk("date",  "Date", dateKeys,  dateN),
            Blk("week",  "Week", weekKeys,  weekN),
            Blk("month", "Month",   monthKeys, monthN),
        };

        var models = rowTotals.Keys
            .OrderByDescending(a => rowTotals[a].Total)
            .ThenBy(a => a, StringComparer.Ordinal)
            .ToList();

        static PivotBucket B(Bucket b) => new(b.Ng, b.Total);

        var cellOut = cells.ToDictionary(kv => kv.Key, kv => B(kv.Value));
        var colTotOut = colTotals.ToDictionary(kv => kv.Key, kv => B(kv.Value), StringComparer.Ordinal);
        var modelOverall = models.ToDictionary(
            m => m,
            m => cells.TryGetValue((m, OverallKey), out Bucket b) ? B(b) : default,
            StringComparer.Ordinal);
        PivotBucket grand = colTotals.TryGetValue(OverallKey, out Bucket g) ? B(g) : default;

        return new PivotData(models, blocks, grand, colTotOut, modelOverall, cellOut);
    }

    /// <summary>The List tab as a flat table: the visible columns (with the synthetic NG RATE)
    /// and one string row per audit, newest first — what the Excel List sheet writes verbatim.</summary>
    public static (IReadOnlyList<string> Headers, IReadOnlyList<IReadOnlyList<string>> Rows) BuildListTable(
        IReadOnlyList<string> columns, IReadOnlyList<IReadOnlyDictionary<string, string>> rows)
    {
        var cols = DisplayColumns(columns);
        var headers = cols.Select(ColumnHeader).ToList();
        var outRows = new List<IReadOnlyList<string>>();
        foreach (var r in Ordered(rows))
        {
            int ng = Count(r, "NGCNT"), total = Count(r, "TOTAL");
            outRows.Add(cols.Select(c => c == NgRateColumn ? NgRateText(ng, total) : Cell(r, c)).ToList());
        }
        return (headers, outRows);
    }

    /// <summary>One pivot cell: ppm on top, the NG/TOTAL it came from underneath.
    /// A null bucket is a period the model was not audited in and renders empty.</summary>
    private static void AppendPivotCell(
        StringBuilder sb, Bucket? bucket, string cls, string? block, int? idx)
    {
        sb.Append("<td class=\"").Append(cls).Append(bucket is null ? " lpa-pivot-empty" : " lpa-pv");
        if (block is not null) sb.Append("\" data-block=\"").Append(block).Append("\" data-idx=\"").Append(idx);
        sb.Append("\">");

        if (bucket is { } b)
            sb.Append("<span class=\"pv-rate\">").Append(Enc(NgRateText(b.Ng, b.Total)))
              .Append("</span><span class=\"pv-cnt\">").Append(b.Ng.ToString("N0"))
              .Append('/').Append(b.Total.ToString("N0")).Append("</span>");

        sb.Append("</td>");
    }

    // ── NG detail table (sits under the pivot, same tab) ────────────────────────────

    /// <summary>
    /// RESUT is a CODE, not text: MES073261 returns "A"/"B"/"C" and the BMES grid renders
    /// the label from <c>GRID_SELECT_OPT.RESUT</c>, which is
    /// <c>A : OK</c>, <c>B : NG</c>, <c>C : N/A</c>. Verified against live data
    /// (2026-06-08~07-24, GN, 600 audits): count of "B" equalled the list's NGCNT on every
    /// single audit — 138 = 138 — while A/C/blank were 6,056/238/14.
    /// </summary>
    private static readonly Dictionary<string, string> ResutLabels = new(StringComparer.OrdinalIgnoreCase)
    {
        ["A"] = "A : OK",
        ["B"] = "B : NG",
        ["C"] = "C : N/A",
    };

    /// <summary>The RESUT code shown the way BMES shows it. Unknown codes pass through
    /// verbatim rather than being swallowed, so a new code shows up instead of hiding.</summary>
    private static string ResutLabel(string resut)
    {
        string v = resut.Trim();
        return v.Length == 0 ? string.Empty
             : ResutLabels.TryGetValue(v, out string? label) ? label
             : v;
    }

    /// <summary>Whether a checklist item's RESUT means "not OK" — code "B". The text rule is
    /// kept as a fallback in case the endpoint ever starts returning labels instead of
    /// codes; it cannot fire on A/C. One place, so the inline badge and the NG table can
    /// never disagree.</summary>
    private static bool IsNgResult(string resut)
    {
        string v = resut.Trim();
        return v.Equals("B", StringComparison.OrdinalIgnoreCase)
            || v.Contains("NG", StringComparison.OrdinalIgnoreCase)
            || v.Contains("FAIL", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Badge colour for a result: red for NG, green for OK, neutral for N/A and
    /// anything unrecognised — painting N/A green would read as a pass.</summary>
    private static string ResutBadgeClass(string resut)
    {
        if (IsNgResult(resut)) return "lpa-badge-ng";
        return resut.Trim().Equals("A", StringComparison.OrdinalIgnoreCase)
            || resut.Contains("OK", StringComparison.OrdinalIgnoreCase)
             ? "lpa-badge-result"
             : "lpa-badge-na";
    }

    /// <summary>One failed checklist item, carrying the list row it came from.</summary>
    public readonly record struct NgDetailItem(
        string Date, string Model, string Laseq, string Lqrno, IReadOnlyDictionary<string, string> Item);

    /// <summary>The Defect Detail matrix as data, shared by the HTML table and the Excel export so
    /// both show the same rows, columns and counts. Rows are a (Check Point · Check Item) pair;
    /// columns are <see cref="Dates"/> (the dates an NG fell on, newest first).</summary>
    public sealed record NgMatrix(
        int NgCount, int ListNg,
        IReadOnlyList<string> Dates,
        IReadOnlyList<NgMatrixGroup> Groups);

    public sealed record NgMatrixGroup(string Model, IReadOnlyList<NgMatrixRow> Rows);

    /// <summary>One check (TYPRC × LCITM) and its NG items bucketed by date.</summary>
    public sealed record NgMatrixRow(
        string Typrc, string Lcitm, int Ng,
        IReadOnlyDictionary<string, IReadOnlyList<NgDetailItem>> ItemsByDate);

    /// <summary>A date column's header label — MM/dd where the raw AUDAT parses, else verbatim.
    /// Invariant culture so the separator stays "/" on a ko-KR host.</summary>
    public static string NgDateLabel(string d) =>
        DateTime.TryParse(d, out DateTime dt)
            ? dt.ToString("MM/dd", System.Globalization.CultureInfo.InvariantCulture)
            : d;

    /// <summary>
    /// Fold every failed checklist item into the matrix: one row per (Check Point, Check Item),
    /// one column per date an NG fell on (newest first), each cell the NG items for that
    /// check on that day. Models (AULOC) are the row sections, worst first; within a model
    /// rows cluster by Check Point and then run worst first.
    /// </summary>
    public static NgMatrix BuildNgMatrix(
        List<IReadOnlyDictionary<string, string>> rows,
        IReadOnlyDictionary<string, BmesLpaScrapeService.LpaResult> details)
    {
        var items = new List<NgDetailItem>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        int listNg = 0;

        foreach (var r in rows)
        {
            listNg += Count(r, "NGCNT");

            // One checklist per LQRNO: the same audit can appear on several list rows and
            // they share a single detail object, which would otherwise be listed twice.
            string lqrno = RowLqrno(r);
            if (lqrno.Length == 0 || !seen.Add(lqrno)) continue;
            if (!details.TryGetValue(lqrno, out BmesLpaScrapeService.LpaResult? res) || !res.IsSuccess)
                continue;

            string model = Cell(r, "AULOC");
            if (model.Length == 0) model = "(No model)";

            foreach (var item in res.Rows)
                if (IsNgResult(Cell(item, "RESUT")))
                    items.Add(new NgDetailItem(Cell(r, "AUDAT"), model, Cell(r, "LASEQ"), lqrno, item));
        }

        var dates = items.Select(i => i.Date)
            .Where(d => d.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .OrderByDescending(d => d, StringComparer.Ordinal)
            .ToList();

        var groups = items
            .GroupBy(i => i.Model, StringComparer.Ordinal)
            .OrderByDescending(g => g.Count())
            .ThenBy(g => g.Key, StringComparer.Ordinal)
            .Select(g => new NgMatrixGroup(g.Key, g
                .GroupBy(i => (Typrc: Cell(i.Item, "TYPRC"), Lcitm: Cell(i.Item, "LCITM")))
                .OrderBy(rg => rg.Key.Typrc, StringComparer.Ordinal)
                .ThenByDescending(rg => rg.Count())
                .ThenBy(rg => rg.Key.Lcitm, StringComparer.Ordinal)
                .Select(rg => new NgMatrixRow(rg.Key.Typrc, rg.Key.Lcitm, rg.Count(),
                    rg.GroupBy(i => i.Date, StringComparer.Ordinal)
                      .ToDictionary(
                          x => x.Key,
                          x => (IReadOnlyList<NgDetailItem>)x.ToList(),
                          StringComparer.Ordinal)))
                .ToList() as IReadOnlyList<NgMatrixRow>))
            .ToList();

        return new NgMatrix(items.Count, listNg, dates, groups);
    }

    /// <summary>
    /// The Defect Detail matrix as an HTML table, foldable per model. The pivot above says how
    /// much NG there is; this says which check failed, when, and what it looked like.
    ///
    /// Returns the markup and whether a table was produced at all (the caller splits the
    /// panel height only when there is a second table to split it with).
    /// </summary>
    private static (string Html, bool HasTable) BuildNgDetailTab(
        List<IReadOnlyDictionary<string, string>> rows,
        IReadOnlyDictionary<string, BmesLpaScrapeService.LpaResult> details,
        bool directImages)
    {
        NgMatrix m = BuildNgMatrix(rows, details);

        var sb = new StringBuilder();
        sb.Append("<div class=\"lpa-subhead\"><span class=\"lpa-strong\">Defect Detail</span>")
          .Append("<span class=\"lpa-muted\">NG (RESUT=B) items only · ")
          .Append(m.NgCount.ToString("N0"));
        if (m.ListNg != m.NgCount)
            sb.Append(" · List NG Total ").Append(m.ListNg.ToString("N0"));
        sb.Append("</span>");

        if (m.NgCount == 0)
            return (sb.Append("</div><div class=\"lpa-note\">No items marked as NG.")
                      .Append(m.ListNg > 0
                          ? " The list NG total is " + m.ListNg.ToString("N0") +
                            ", so the detail RESUT may not be code B."
                          : "")
                      .Append("</div>").ToString(), false);

        sb.Append("<span class=\"lpa-subhead-gap\"></span>")
          .Append("<button type=\"button\" id=\"tb-ng-expand\">Expand all</button>")
          .Append("<button type=\"button\" id=\"tb-ng-fold\">Collapse all</button></div>");

        int span = 2 + m.Dates.Count;   // Check Point + Check Item labels, then one column per date
        sb.Append("<div class=\"lpa-scroll lpa-scroll-ng\">")
          .Append("<table class=\"lpa-table lpa-ngdetail\" id=\"lpa-ngdetail\"><thead><tr>")
          .Append("<th class=\"lpa-ng-c1\">Check Point</th><th class=\"lpa-ng-c2\">Check Item</th>");
        foreach (string d in m.Dates)
            sb.Append("<th class=\"lpa-num\">").Append(Enc(NgDateLabel(d))).Append("</th>");
        sb.Append("</tr></thead><tbody>");

        foreach (var g in m.Groups)
        {
            sb.Append("<tr class=\"lpa-ng-group\" data-model=\"").Append(Enc(g.Model))
              .Append("\"><td colspan=\"").Append(span).Append("\">")
              .Append("<span class=\"lpa-caret\">▼</span><strong>").Append(Enc(g.Model))
              .Append("</strong><span class=\"lpa-muted lpa-count\"></span></td></tr>");

            foreach (var row in g.Rows)
            {
                // data-ng: the model's NG count is summed from these so folding/filtering keeps
                // it honest; data-model lets the filter narrow to a model the group row names.
                sb.Append("<tr class=\"lpa-ng-row\" data-model=\"").Append(Enc(g.Model))
                  .Append("\" data-ng=\"").Append(row.Ng).Append("\">")
                  .Append("<td class=\"lpa-ng-c1\" title=\"").Append(Enc(row.Typrc)).Append("\">")
                  .Append(Enc(row.Typrc)).Append("</td>")
                  .Append("<td class=\"lpa-ng-c2\" title=\"").Append(Enc(row.Lcitm)).Append("\">")
                  .Append(Enc(row.Lcitm)).Append("</td>");

                foreach (string d in m.Dates)
                {
                    if (row.ItemsByDate.TryGetValue(d, out var cellItems) && cellItems.Count > 0)
                    {
                        sb.Append("<td class=\"lpa-ng-cell\"><span class=\"lpa-ng-qty\">")
                          .Append(cellItems.Count).Append("</span><span class=\"lpa-thumbs\">");
                        foreach (var ci in cellItems)
                            AppendThumbs(sb, ci.Item, ImageCaption(ci.Lqrno, ci.Date, g.Model, ci.Item), directImages);
                        sb.Append("</span></td>");
                    }
                    else
                        sb.Append("<td class=\"lpa-ng-cell lpa-ng-empty\"></td>");
                }
                sb.Append("</tr>");
            }
        }

        sb.Append("</tbody></table></div>");
        return (sb.ToString(), true);
    }

    // ── Page ────────────────────────────────────────────────────────────────────────

    private static string BuildHtml(ExportInput input)
    {
        var rows = Ordered(input.Rows);
        var cols = DisplayColumns(input.Columns);
        var templates = new StringBuilder();

        string listTab = rows.Count == 0
            ? "<div class=\"lpa-note\">No rows to show.</div>"
            : BuildListTab(input, rows, cols, templates);
        PivotRender pivot = BuildPivotTab(rows);
        string pivotTab = pivot.Html;
        string checkPivotTab = BuildCheckPivotTab(rows, input.Details, pivot.Blocks, pivot.ColTotals);
        string checkPivotSection = checkPivotTab.Length == 0
            ? ""
            : $"<div class=\"lpa-scroll lpa-scroll-chk\">{checkPivotTab}</div>";
        var (ngDetailTab, hasNgTable) = BuildNgDetailTab(rows, input.Details, input.DirectBmesImages);
        // Defect Detail moved to its own tab, so the pivot panel only ever stacks the two pivots.
        _ = hasNgTable;
        string pivotPanelClass = checkPivotTab.Length > 0 ? " has-chk" : "";

        int detailOk = input.Details.Values.Count(d => d.IsSuccess);
        int detailBad = input.Details.Count - detailOk;

        string header =
            $"{input.From:yyyy-MM-dd} ~ {input.To:yyyy-MM-dd} · FACCO {(input.Facco.Length > 0 ? input.Facco : "-")}" +
            $" · {rows.Count:N0} rows · Detail {detailOk:N0}" +
            (detailBad > 0 ? $" (failed {detailBad:N0})" : "") +
            $" · Generated {DateTime.Now:yyyy-MM-dd HH:mm}";

        // Lightbox photo source: the resized /bmes/lpa/img route (web) or the BMES original
        // directly (WPF, no route). Thumbnails already differ via AppendThumbs; this is the popup.
        string imgDirect = input.DirectBmesImages ? "true" : "false";

        return $$"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>BMES LPA</title>
<style>
  body { margin: 0; padding: 10px 14px 24px; background: #fff;
         font: 12px/1.35 'Malgun Gothic',system-ui,sans-serif; color: #0f172a; }
  .lpa-muted { color: #64748b; font-size: 11px; margin-left: 8px; }
  .lpa-strong { font-weight: 600; }
  .lpa-note { padding: 24px; color: #64748b; }
  .lpa-error { padding: 10px 12px; border-radius: 6px; background: #fef2f2; color: #b91c1c; }

  /* Tabs — same chrome as the BMES report viewer. */
  .lpa-tabs { display: flex; flex-wrap: wrap; gap: 4px; border-bottom: 1px solid #d7dee8; margin-bottom: 10px; }
  .lpa-tab { border: 1px solid transparent; border-bottom: 0; background: transparent; color: #475569;
             padding: 8px 18px; font-size: 13px; font-weight: 600; border-radius: 6px 6px 0 0; cursor: pointer; }
  .lpa-tab:hover { background: #f8fafc; color: #0f172a; }
  .lpa-tab.active { background: #fff; border-color: #d7dee8; color: #0f172a; margin-bottom: -1px; }

  .lpa-head { color: #64748b; font-size: 11px; margin: 0 0 8px; }

  .lpa-toolbar { position: sticky; top: 0; z-index: 30; display: flex; flex-wrap: wrap; align-items: center;
                 gap: 6px 12px; padding: 7px 10px; margin-bottom: 10px; background: #f8fafc;
                 border: 1px solid #e2e8f0; border-radius: 8px; color: #334155; }
  .lpa-toolbar label { display: inline-flex; align-items: center; gap: 5px; font-weight: 600; white-space: nowrap; }
  .lpa-toolbar input, .lpa-toolbar select { font: inherit; font-weight: 400; padding: 2px 6px;
                 border: 1px solid #cbd5e1; border-radius: 4px; background: #fff; color: #0f172a; }
  .lpa-toolbar button { font: inherit; font-weight: 600; padding: 3px 10px; cursor: pointer;
                 border: 1px solid #cbd5e1; border-radius: 4px; background: #fff; color: #334155; }
  .lpa-toolbar button:hover { background: #eef2f7; }
  .lpa-toolbar .tb-sep { width: 1px; align-self: stretch; background: #dbe3ec; }
  @media print { .lpa-toolbar, .lpa-tabs { display: none !important; } }

  .lpa-panel { display: none; }
  .lpa-panel.active { display: block; }
  .lpa-scroll { overflow: auto; max-height: 78vh; border: 1px solid #e2e8f0; border-radius: 8px; }

  /* zoom scales the hard-coded px sizes proportionally, which a font-size override
     cannot do without flattening the deliberate differences between them. */
  .lpa-scroll { zoom: var(--tb-zoom, 1); }

  .lpa-table { border-collapse: collapse; width: 100%; font-size: 11px; white-space: nowrap; }
  .lpa-table th, .lpa-table td { border-bottom: 1px solid #eef2f7; padding: 3px 8px; vertical-align: middle; }
  .lpa-table thead th { position: sticky; top: 0; z-index: 2; background: #f1f5f9;
                        border-bottom: 1px solid #cbd5e1; font-weight: 600; text-align: left; }
  .lpa-table tbody tr.lpa-data-row:hover > td { background: #f8fafc; }
  .lpa-data-row { cursor: pointer; }
  .lpa-row-selected > td { background: #e0f2fe !important; }
  .lpa-table td.lpa-num, .lpa-table th.lpa-num { text-align: right; font-variant-numeric: tabular-nums; }

  .lpa-group-row { cursor: pointer; }
  .lpa-group-row > td { position: sticky; left: 0; background: #f1f5f9 !important;
                        border-top: 1px solid #cbd5e1; font-size: 12px; padding: 5px 10px; }
  .lpa-group-row:hover > td { background: #e2e8f0 !important; }
  .lpa-caret { display: inline-block; width: 14px; color: #64748b; font-size: 10px; }
  .lpa-model-row { cursor: pointer; }
  .lpa-model-row > td { background: #f8fafc !important; border-top: 1px solid #e2e8f0;
                        font-size: 11px; padding: 3px 10px; color: #334155; }
  /* Only the label sticks to the left; the total cells sit under their own columns. */
  .lpa-model-row > td.lpa-model-label { position: sticky; left: 0; }
  .lpa-model-row > td.lpa-total-cell { font-weight: 700; color: #0f172a; }
  .lpa-model-row:hover > td { background: #eef2f7 !important; }
  .lpa-model-indent { display: inline-block; width: 18px; }
  .lpa-detail-btn { font-size: 10px; padding: 0 6px; border: 1px solid #93c5fd; border-radius: 4px;
                    background: #fff; color: #1d4ed8; cursor: pointer; }
  .lpa-detail-btn:hover { background: #eff6ff; }

  /* Inline detail: expands right under the clicked row, inside its own full-width cell. */
  .lpa-detail-row > td { padding: 0 !important; background: #fff; border-top: 0; }
  .lpa-detail-inline { padding: 12px 14px 14px 40px; background: #f8fafc;
                       border-bottom: 2px solid #cbd5e1; white-space: normal; }
  .lpa-detail-inline-head { display: flex; flex-wrap: wrap; align-items: center; gap: 8px;
                            margin-bottom: 10px; font-size: 13px; }
  .lpa-detail-close { margin-left: auto; font-size: 11px; padding: 1px 8px; border: 1px solid #cbd5e1;
                      border-radius: 4px; background: #fff; color: #475569; cursor: pointer; }
  .lpa-detail-item { border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 12px;
                     margin-bottom: 10px; background: #fff; }
  .lpa-detail-item:last-child { margin-bottom: 0; }
  .lpa-detail-head { display: flex; flex-wrap: wrap; align-items: center; gap: 8px;
                     margin-bottom: 8px; font-size: 13px; }
  .lpa-seq { display: inline-flex; align-items: center; justify-content: center; min-width: 24px;
             height: 24px; padding: 0 6px; border-radius: 999px; background: #0f172a; color: #fff;
             font-size: 11px; font-weight: 700; }
  .lpa-badge { padding: 1px 8px; border-radius: 999px; background: #eef2f7; color: #475569;
               font-size: 11px; font-weight: 600; }
  .lpa-badge-result { background: #dcfce7; color: #166534; }
  .lpa-badge-ng { background: #fee2e2; color: #b91c1c; }
  /* N/A (RESUT=C) and unknown codes: neutral, never the green that reads as a pass. */
  .lpa-badge-na { background: #e2e8f0; color: #475569; }
  .lpa-detail-grid { display: grid; grid-template-columns: 120px 1fr; gap: 4px 12px; margin: 0; font-size: 12px; }
  .lpa-detail-grid dt { color: #64748b; font-weight: 600; }
  /* Checklist text arrives as multi-line Vietnamese + Korean pairs — keep the breaks. */
  .lpa-detail-grid dd { margin: 0; white-space: pre-wrap; word-break: break-word; }
  @media (max-width: 720px) { .lpa-detail-grid { grid-template-columns: 1fr; } }

  /* Sized to content rather than stretched to 100%: with only a few AULOCs a full-width
     pivot turns the model column into a huge empty gutter. Wide ones scroll instead. */
  .lpa-pivot { width: auto; }
  .lpa-pivot th, .lpa-pivot td { min-width: 74px; padding: 3px 10px; }
  .lpa-pivot th.lpa-pivot-corner, .lpa-pivot th.lpa-pivot-row-head {
             position: sticky; left: 0; z-index: 3; background: #f1f5f9; text-align: left; font-weight: 600; }
  .lpa-pivot thead th.lpa-pivot-corner { z-index: 4; }
  .lpa-pivot td.lpa-pivot-empty { background: #fbfcfe; }
  .lpa-pivot .lpa-pivot-sum { font-weight: 700; background: #f8fafc; }
  /* Total is a third header row, not a footer: pinned under the period row so it stays put
     while the models scroll. z-index above the sticky row heads (3), which would otherwise
     paint over it on the way past. --head-h is measured for the same reason --band-h is. */
  .lpa-pivot thead tr.lpa-pivot-total > th, .lpa-pivot thead tr.lpa-pivot-total > td {
             position: sticky; top: calc(var(--band-h, 26px) + var(--head-h, 20px)); z-index: 5;
             background: #f1f5f9; font-weight: 700; border-bottom: 2px solid #cbd5e1; }
  .lpa-pivot thead tr.lpa-pivot-total > th { left: 0; z-index: 6; text-align: left; }
  .lpa-pivot thead tr.lpa-pivot-total > td.sep-td { background: #eef2f7; }
  /* ppm on top, the NG/TOTAL it rests on underneath. */
  .pv-rate { display: block; font-variant-numeric: tabular-nums; }
  .pv-cnt  { display: block; font-size: 9px; font-weight: 400; color: #64748b;
             font-variant-numeric: tabular-nums; }
  .lpa-pivot th.lpa-band { text-align: center; background: #eef2f7; color: #475569; font-size: 11px; }
  /* Two header rows: the band sticks at the top and the period row directly under it.
     --band-h is measured at runtime (see syncPivotHeader) because the band's real height
     moves with font and zoom, and a guessed offset makes the rows overlap when scrolled. */
  .lpa-pivot thead tr.lpa-pivot-band > th { top: 0; padding-top: 2px; padding-bottom: 2px; }
  .lpa-pivot thead tr.lpa-pivot-head > th { top: var(--band-h, 26px); }
  /* Block dividers: after the sizing rule above so they can stay hairline-thin. */
  .lpa-pivot .sep-th, .lpa-pivot .sep-td {
             min-width: 0; width: 7px; padding: 0; background: #eef2f7;
             border-left: 1px solid #dbe3ec; border-right: 1px solid #dbe3ec; }

  /* NG detail table: what the pivot's numbers were actually made of. */
  .lpa-subhead { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 10px;
                 margin: 14px 0 6px; font-size: 12px; color: #334155; }
  .lpa-subhead .lpa-muted { margin-left: 0; }
  .lpa-subhead-gap { flex: 1 1 auto; }
  .lpa-subhead button { font: inherit; font-weight: 600; padding: 2px 9px; cursor: pointer;
                 border: 1px solid #cbd5e1; border-radius: 4px; background: #fff; color: #334155; }
  .lpa-subhead button:hover { background: #eef2f7; }
  @media print { .lpa-subhead button { display: none !important; } }
  /* Two stacked tables in one tab: split the height rather than let the pivot push the
     detail table off screen. */
  /* Defect Detail has its own tab, so it keeps the default full height; only the pivot tab
     stacks two tables and has to split. */
  /* Three stacked tables (model pivot · check-item pivot · defect detail) need a smaller slice. */
  .lpa-panel.has-chk .lpa-scroll { max-height: 40vh; }
  .lpa-scroll-chk { margin-top: 10px; }
  /* Check Item texts are sentences, not codes: give the row head real width and let it wrap,
     otherwise one long Vietnamese line pushes every period column off screen. */
  .lpa-chkpivot th.lpa-pivot-corner, .lpa-chkpivot th.lpa-pivot-row-head {
    min-width: 260px; max-width: 380px; white-space: normal; text-align: left; }
  .lpa-chkpivot th.lpa-chk-head { padding: 4px 10px; font-weight: 500; }
  /* Check Item reads last but stays the row's identity, so it keeps the darker/larger type. */
  .lpa-chk-item { display: block; margin-top: 2px; }
  .lpa-chk-typrc { display: block; color: #64748b; font-size: 10px; font-weight: 400; }
  .lpa-chk-typrc b { font-weight: 600; color: #94a3b8; margin-right: 4px; }
  /* One rule per date column: without them the photos of adjacent days run together. */
  .lpa-ngdetail thead th.lpa-num, .lpa-ngdetail td.lpa-ng-cell { border-left: 1px solid #dbe3ec; }
  .lpa-ng-group { cursor: pointer; }
  .lpa-ng-group > td { position: sticky; left: 0; background: #f1f5f9 !important;
                       border-top: 1px solid #cbd5e1; font-size: 12px; padding: 5px 10px; }
  .lpa-ng-group:hover > td { background: #e2e8f0 !important; }
  .lpa-ngdetail tbody tr.lpa-ng-row:hover > td { background: #f8fafc; }
  /* NG matrix: rows are (Check Point · Check Item), columns are the dates NG fell on. The two
     label columns stay pinned as the date columns scroll sideways; their fixed widths are
     what the second column's left offset is measured from. */
  .lpa-ngdetail th.lpa-ng-c1, .lpa-ngdetail td.lpa-ng-c1,
  .lpa-ngdetail th.lpa-ng-c2, .lpa-ngdetail td.lpa-ng-c2 {
               position: sticky; background: #fff; text-align: left; font-weight: 400;
               overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .lpa-ngdetail th.lpa-ng-c1, .lpa-ngdetail td.lpa-ng-c1 { left: 0;
               width: 150px; min-width: 150px; max-width: 150px; }
  .lpa-ngdetail th.lpa-ng-c2, .lpa-ngdetail td.lpa-ng-c2 { left: 150px;
               width: 200px; min-width: 200px; max-width: 200px; border-right: 1px solid #cbd5e1; }
  .lpa-ngdetail thead th.lpa-ng-c1, .lpa-ngdetail thead th.lpa-ng-c2 {
               z-index: 4; background: #f1f5f9; font-weight: 600; }
  .lpa-ngdetail tbody td.lpa-ng-c1, .lpa-ngdetail tbody td.lpa-ng-c2 { z-index: 1; }
  .lpa-ngdetail tr.lpa-ng-row:hover > td.lpa-ng-c1,
  .lpa-ngdetail tr.lpa-ng-row:hover > td.lpa-ng-c2 { background: #f8fafc; }
  /* Cells: the NG count on top, its photos below; an audited-but-clean date is left blank. */
  .lpa-ng-cell { text-align: center; vertical-align: top; }
  .lpa-ng-empty { background: #fbfcfe; }
  .lpa-ng-qty { display: block; font-weight: 700; color: #b91c1c;
               font-variant-numeric: tabular-nums; }
  .lpa-ng-cell .lpa-thumbs { flex-wrap: wrap; justify-content: center; margin-top: 3px; }

  /* Result photos: thumbnails in the table, full view in the popup. */
  .lpa-thumbs { display: inline-flex; gap: 3px; align-items: center; }
  td.lpa-thumbs { white-space: nowrap; }
  .lpa-thumb { height: 34px; width: 34px; object-fit: cover; border: 1px solid #cbd5e1;
               border-radius: 3px; cursor: zoom-in; background: #f8fafc; vertical-align: middle; }
  .lpa-thumb:hover { border-color: #64748b; }
  /* Past the embedding budget the photo is a link to BMES instead of its own pixels. */
  .lpa-thumb-link { font-size: 11px; color: #2563eb; text-decoration: none; border: 1px solid #cbd5e1;
               border-radius: 3px; padding: 1px 5px; }
  .lpa-lightbox { position: fixed; inset: 0; z-index: 100; display: none;
               background: rgba(15,23,42,.88); }
  .lpa-lightbox.open { display: flex; flex-direction: column; align-items: center; justify-content: center; }
  .lpa-lightbox img { max-width: 94vw; max-height: 82vh; object-fit: contain;
               background: #fff; border-radius: 4px; }
  .lpa-lb-bar { display: flex; align-items: center; gap: 12px; margin-top: 10px;
               color: #e2e8f0; font-size: 12px; max-width: 94vw; }
  .lpa-lb-bar button, .lpa-lb-bar a { font: inherit; color: #e2e8f0; background: rgba(255,255,255,.12);
               border: 1px solid rgba(255,255,255,.25); border-radius: 4px; padding: 3px 10px;
               cursor: pointer; text-decoration: none; }
  .lpa-lb-bar button:hover, .lpa-lb-bar a:hover { background: rgba(255,255,255,.24); }
  .lpa-lb-cap { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .lpa-lb-close { position: absolute; top: 12px; right: 16px; }
  @media print { .lpa-lightbox { display: none !important; } }
</style>
</head>
<body>
<div class="lpa-tabs">
  <button type="button" class="lpa-tab active" data-tab="list">List</button>
  <button type="button" class="lpa-tab" data-tab="pivot">NG Summary</button>
  <button type="button" class="lpa-tab" data-tab="ng">Defect Detail</button>
</div>
<div class="lpa-head">{{header}}</div>

<div class="lpa-toolbar">
  <label>Filter <input type="search" id="tb-filter" placeholder="Search all columns…" style="width:180px;"></label>
  <span class="tb-sep" data-only="list"></span>
  <button type="button" id="tb-expand" data-only="list">Expand all</button>
  <button type="button" id="tb-fold-model" data-only="list">Collapse models</button>
  <button type="button" id="tb-fold-all" data-only="list">Collapse all</button>
  <span class="tb-sep" data-only="pivot"></span>
  <label data-only="pivot">Date <input type="number" id="tb-date" min="0" step="1" value="7" placeholder="Overall" style="width:58px;"> d</label>
  <label data-only="pivot">Week <input type="number" id="tb-week" min="0" step="1" value="4" placeholder="Overall" style="width:58px;"> w</label>
  <label data-only="pivot">Month <input type="number" id="tb-month" min="0" step="1" value="3" placeholder="Overall" style="width:58px;"> mo</label>
  <span class="tb-sep"></span>
  <label>Size <input type="number" id="tb-zoom" min="50" max="250" step="5" value="100" style="width:60px;">%</label>
  <span id="tb-stat" class="lpa-muted"></span>
</div>

<div class="lpa-panel active" data-panel="list"><div class="lpa-scroll">{{listTab}}</div></div>
<div class="lpa-panel{{pivotPanelClass}}" data-panel="pivot"><div class="lpa-scroll">{{pivotTab}}</div>{{checkPivotSection}}</div>
<div class="lpa-panel" data-panel="ng">{{ngDetailTab}}</div>
{{templates}}

<div class="lpa-lightbox" id="lpa-lightbox">
  <img id="lpa-lb-img" src="" alt="">
  <div class="lpa-lb-bar">
    <button type="button" id="lpa-lb-prev">‹ Prev</button>
    <span id="lpa-lb-pos"></span>
    <button type="button" id="lpa-lb-next">Next ›</button>
    <span class="lpa-lb-cap" id="lpa-lb-cap"></span>
    <a id="lpa-lb-orig" href="#" target="_blank" rel="noopener">Original</a>
  </div>
  <button type="button" class="lpa-lb-close" id="lpa-lb-close">Close ✕</button>
</div>
<script>
(function () {
  var listTable = document.getElementById('lpa-list');
  var ngTable = document.getElementById('lpa-ngdetail');
  var stat = document.getElementById('tb-stat');
  var filter = document.getElementById('tb-filter');
  var collapsedDates = {}, collapsedModels = {}, collapsedNg = {};

  function fmt(n) { return n.toLocaleString('en-US'); }
  function rate(ng, total) { return total > 0 ? fmt(Math.round(ng * 1000000 / total)) : '-'; }

  // ── List view ──────────────────────────────────────────────────────────────
  // One pass decides visibility for every row and re-derives the group counts and
  // model totals from what survives, so a filtered view totals only what it shows.
  function applyList() {
    if (!listTable) return;
    var needle = (filter.value || '').trim().toLowerCase();
    var rows = listTable.tBodies[0].rows;
    var groups = {}, models = {}, shown = 0, all = 0;

    for (var i = 0; i < rows.length; i++) {
      var tr = rows[i];
      if (!tr.classList.contains('lpa-data-row')) continue;
      all++;
      var match = !needle || tr.textContent.toLowerCase().indexOf(needle) >= 0;
      tr.dataset.match = match ? '1' : '0';
      if (!match) continue;
      shown++;

      var d = tr.dataset.date, m = tr.dataset.model;
      var g = groups[d] || (groups[d] = { n: 0 });
      g.n++;
      var mm = models[m] || (models[m] = { n: 0, total: 0, ok: 0, ng: 0 });
      mm.n++;
      mm.total += +tr.dataset.total; mm.ok += +tr.dataset.ok; mm.ng += +tr.dataset.ng;
    }

    for (var j = 0; j < rows.length; j++) {
      var r = rows[j];

      if (r.classList.contains('lpa-group-row')) {
        var gi = groups[r.dataset.date];
        r.hidden = !gi;
        if (gi) {
          r.querySelector('.lpa-caret').textContent = collapsedDates[r.dataset.date] ? '▶' : '▼';
          r.querySelector('.lpa-count').textContent = fmt(gi.n);
        }
        continue;
      }

      if (r.classList.contains('lpa-model-row')) {
        var mi = models[r.dataset.model];
        r.hidden = !mi || collapsedDates[r.dataset.date];
        if (mi && !r.hidden) {
          r.querySelector('.lpa-caret').textContent = collapsedModels[r.dataset.model] ? '▶' : '▼';
          r.querySelector('.lpa-count').textContent = fmt(mi.n);
          r.querySelectorAll('[data-tot]').forEach(function (td) {
            var k = td.dataset.tot;
            td.textContent = k === 'TOTAL' ? fmt(mi.total)
                           : k === 'OKCNT' ? fmt(mi.ok)
                           : k === 'NGCNT' ? fmt(mi.ng)
                           : k === 'NG RATE' ? rate(mi.ng, mi.total) : '';
          });
        }
        continue;
      }

      if (r.classList.contains('lpa-detail-row')) continue;   // handled with its owner row

      var visible = r.dataset.match === '1' &&
                    !collapsedDates[r.dataset.date] && !collapsedModels[r.dataset.model];
      r.hidden = !visible;
      // The open detail sits right after its row and follows it out of view rather
      // than dangling under an unrelated group.
      var next = r.nextElementSibling;
      if (next && next.classList.contains('lpa-detail-row')) next.hidden = !visible;
    }

    // Number only the rows actually on screen, so # stays 1..n with no gaps.
    var n = 0;
    for (var k = 0; k < rows.length; k++) {
      var dr = rows[k];
      if (!dr.classList.contains('lpa-data-row') || dr.hidden) continue;
      dr.cells[0].textContent = ++n;
    }
    stat.textContent = fmt(shown) + ' / ' + fmt(all) + ' rows';
  }

  function closeDetail() {
    var open = listTable && listTable.querySelector('.lpa-detail-row');
    if (open) open.remove();
    var sel = listTable && listTable.querySelector('.lpa-row-selected');
    if (sel) sel.classList.remove('lpa-row-selected');
  }

  function toggleDetail(tr) {
    var wasOpen = tr.classList.contains('lpa-row-selected');
    closeDetail();
    if (wasOpen) return;

    var tpl = document.querySelector('template[data-detail="' + (tr.dataset.lqrno || '') + '"]');
    var row = listTable.tBodies[0].insertRow(tr.sectionRowIndex + 1);
    row.className = 'lpa-detail-row';
    var td = row.insertCell(0);
    td.colSpan = tr.cells.length;
    if (tpl) td.appendChild(tpl.content.cloneNode(true));
    else td.innerHTML = '<div class="lpa-detail-inline"><div class="lpa-note">' +
                        'This row has no LQRNO, so no detail was prefetched.</div></div>';
    tr.classList.add('lpa-row-selected');
  }

  // ── Result photos ──────────────────────────────────────────────────────────
  // A thumbnail carries only its BMES path in data-img; both its own src and the popup's
  // src point at /bmes/lpa/img, which downscales on first request and serves the stored
  // copy afterwards. Nothing is embedded, so the file stays small however many photos a
  // search has. IMG_URL is the untouched original, used for the "Original" link and the
  // fallback when a photo is gone.
  var IMG_URL = '{{BmesLpaImageService.ImageUrlPrefix}}';
  var IMG_DIRECT = {{imgDirect}};
  var lb = document.getElementById('lpa-lightbox');
  var lbImg = document.getElementById('lpa-lb-img');
  var lbCap = document.getElementById('lpa-lb-cap');
  var lbPos = document.getElementById('lpa-lb-pos');
  var lbOrig = document.getElementById('lpa-lb-orig');
  var lbList = [], lbAt = -1;

  function viewSrc(path) {
    return IMG_DIRECT ? (IMG_URL + encodeURIComponent(path))
                      : ('/bmes/lpa/img?size=view&path=' + encodeURIComponent(path));
  }

  // A photo that 404s (deleted on BMES) drops its broken <img> for a link to the original.
  window.lpaImgFail = function (el) {
    var a = document.createElement('a');
    a.className = 'lpa-thumb-link';
    a.textContent = 'Photo';
    a.target = '_blank';
    a.rel = 'noopener';
    a.href = IMG_URL + encodeURIComponent(el.dataset.img || '');
    el.replaceWith(a);
  };

  function showPhoto() {
    var t = lbList[lbAt], path = t.dataset.img;
    lbImg.src = viewSrc(path);
    lbCap.textContent = t.dataset.cap || path;
    lbPos.textContent = (lbAt + 1) + ' / ' + lbList.length;
    lbOrig.href = IMG_URL + encodeURIComponent(path);
  }

  function openPhoto(thumb) {
    // ‹ › walks what is actually on screen, in document order: with the table filtered
    // or a model folded away, the photos behind hidden rows are not part of the review.
    lbList = [].slice.call(document.querySelectorAll('.lpa-thumb')).filter(function (t) {
      return t.offsetParent !== null;
    });
    lbAt = lbList.indexOf(thumb);
    if (lbAt < 0) { lbList = [thumb]; lbAt = 0; }
    showPhoto();
    lb.classList.add('open');
  }

  function stepPhoto(d) {
    if (lbList.length < 2) return;
    lbAt = (lbAt + d + lbList.length) % lbList.length;
    showPhoto();
  }

  function closePhoto() {
    lb.classList.remove('open');
    lbImg.removeAttribute('src');   // let the decoded bitmap go
  }

  lb.addEventListener('click', function (e) {
    if (e.target === lb) closePhoto();   // backdrop only: clicking the photo should not lose it
  });
  document.getElementById('lpa-lb-close').addEventListener('click', closePhoto);
  document.getElementById('lpa-lb-prev').addEventListener('click', function () { stepPhoto(-1); });
  document.getElementById('lpa-lb-next').addEventListener('click', function () { stepPhoto(1); });
  document.addEventListener('keydown', function (e) {
    if (!lb.classList.contains('open')) return;
    if (e.key === 'Escape') closePhoto();
    else if (e.key === 'ArrowLeft') stepPhoto(-1);
    else if (e.key === 'ArrowRight') stepPhoto(1);
  });

  document.addEventListener('click', function (e) {
    var tab = e.target.closest('.lpa-tab');
    if (tab) { selectTab(tab.dataset.tab); return; }

    // Before the row handlers: a photo in an open checklist must not also toggle the row.
    if (e.target.classList.contains('lpa-thumb')) { openPhoto(e.target); return; }

    if (e.target.closest('.lpa-detail-close')) { closeDetail(); return; }

    var ngGroup = e.target.closest('.lpa-ng-group');
    if (ngGroup) {
      var nm = ngGroup.dataset.model;
      if (collapsedNg[nm]) delete collapsedNg[nm]; else collapsedNg[nm] = 1;
      applyNgDetail(); return;
    }

    var group = e.target.closest('.lpa-group-row');
    if (group) {
      var d = group.dataset.date;
      if (collapsedDates[d]) delete collapsedDates[d]; else collapsedDates[d] = 1;
      closeDetail(); applyList(); return;
    }

    var model = e.target.closest('.lpa-model-row');
    if (model) {
      var m = model.dataset.model;
      if (collapsedModels[m]) delete collapsedModels[m]; else collapsedModels[m] = 1;
      closeDetail(); applyList(); return;
    }

    var dataRow = e.target.closest('.lpa-data-row');
    if (dataRow) toggleDetail(dataRow);
  });

  // ── Pivot view ─────────────────────────────────────────────────────────────
  // Every period column is in the file; the Date/Week/Month boxes only decide how many of
  // the newest ones stay on screen. Blank means all of them.
  var BLOCKS = ['date', 'week', 'month'];
  var limitInput = {
    date:  document.getElementById('tb-date'),
    week:  document.getElementById('tb-week'),
    month: document.getElementById('tb-month')
  };

  function limitOf(block) {
    var v = parseInt(limitInput[block].value, 10);
    return isNaN(v) ? null : Math.max(0, v);
  }

  // Both pivots (model and check item) carry the same data-block/data-idx columns, so the
  // date/week/month boxes and the filter drive them through one routine.
  var PIVOT_IDS = ['lpa-pivot', 'lpa-chkpivot'];

  function applyPivotTable(pivot) {
    if (!pivot) return;

    pivot.querySelectorAll('[data-idx]').forEach(function (el) {
      var n = limitOf(el.dataset.block);
      el.hidden = n !== null && +el.dataset.idx >= n;
    });

    BLOCKS.forEach(function (b) {
      var shown = pivot.querySelectorAll('thead th[data-block="' + b + '"][data-idx]:not([hidden])').length;
      // A band header spanning zero columns would still paint a stray cell, so the
      // whole block — divider included — goes when nothing in it is left.
      pivot.querySelectorAll('.lpa-band[data-block="' + b + '"]').forEach(function (th) {
        th.hidden = shown === 0;
        th.colSpan = Math.max(1, shown);
      });
      pivot.querySelectorAll('.sep-th[data-block="' + b + '"], .sep-td[data-block="' + b + '"]')
           .forEach(function (el) { el.hidden = shown === 0; });
    });

    // The filter is a list-tab control; on the pivot it narrows visible rows.
    var needle = (filter.value || '').trim().toLowerCase();
    Array.prototype.forEach.call(pivot.tBodies[0].rows, function (tr) {
      tr.hidden = !!needle && tr.textContent.toLowerCase().indexOf(needle) < 0;
    });

    syncPivotHeader(pivot);
  }

  function applyPivot() {
    PIVOT_IDS.forEach(function (id) { applyPivotTable(document.getElementById(id)); });
  }

  // Pin the period header row directly below the band row, and Total directly below that.
  // Measured rather than assumed: row heights follow the font and the zoom setting, and an
  // offset that is even a few px short makes the header rows overlap once the table scrolls.
  function syncPivotHeader(pivot) {
    if (!pivot || !pivot.tHead) return;
    var band = pivot.tHead.rows[0].offsetHeight;
    if (band > 0) pivot.style.setProperty('--band-h', band + 'px');  // 0 while the tab is hidden
    var head = pivot.tHead.rows[1] ? pivot.tHead.rows[1].offsetHeight : 0;
    if (head > 0) pivot.style.setProperty('--head-h', head + 'px');
  }

  // ── NG detail view ─────────────────────────────────────────────────────────
  // Same shape as the list: one pass marks what the filter keeps, a second decides
  // visibility per model block. A model's NG count is the sum of its rows' data-ng (each
  // row's NG occurrences across the date columns), not a row count, so it stays equal to
  // the pivot's model NG even though rows collapse many occurrences onto one line.
  function applyNgDetail() {
    if (!ngTable) return;
    var needle = (filter.value || '').trim().toLowerCase();
    var rows = ngTable.tBodies[0].rows, counts = {};

    for (var i = 0; i < rows.length; i++) {
      var tr = rows[i];
      if (!tr.classList.contains('lpa-ng-row')) continue;
      // The model lives on the group row, not in the cells, so it is matched explicitly:
      // typing a model name should narrow this table to it the way it does the pivot,
      // not empty it.
      var hay = (tr.dataset.model + ' ' + tr.textContent).toLowerCase();
      var match = !needle || hay.indexOf(needle) >= 0;
      tr.dataset.match = match ? '1' : '0';
      if (match) counts[tr.dataset.model] =
        (counts[tr.dataset.model] || 0) + (parseInt(tr.dataset.ng, 10) || 0);
    }

    for (var j = 0; j < rows.length; j++) {
      var r = rows[j];
      if (r.classList.contains('lpa-ng-group')) {
        var c = counts[r.dataset.model];
        r.hidden = !c;
        if (c) {
          r.querySelector('.lpa-caret').textContent = collapsedNg[r.dataset.model] ? '▶' : '▼';
          r.querySelector('.lpa-count').textContent = fmt(c);
        }
        continue;
      }
      var visible = r.dataset.match === '1' && !collapsedNg[r.dataset.model];
      r.hidden = !visible;
    }
  }

  // ── Chrome ─────────────────────────────────────────────────────────────────
  function selectTab(name) {
    document.querySelectorAll('.lpa-tab').forEach(function (b) {
      b.classList.toggle('active', b.dataset.tab === name);
    });
    document.querySelectorAll('.lpa-panel').forEach(function (p) {
      p.classList.toggle('active', p.dataset.panel === name);
    });
    // Per-tab toolbar controls: hidden rather than disabled so the bar stays short.
    document.querySelectorAll('[data-only]').forEach(function (el) {
      el.style.display = el.dataset.only === name ? '' : 'none';
    });
    if (name === 'pivot') applyPivot();
    else if (name === 'ng') applyNgDetail();
    else applyList();
  }

  filter.addEventListener('input', function () {
    closeDetail();
    var active = document.querySelector('.lpa-panel.active').dataset.panel;
    if (active === 'pivot') applyPivot();
    else if (active === 'ng') applyNgDetail();
    else applyList();
  });
  BLOCKS.forEach(function (b) { limitInput[b].addEventListener('input', applyPivot); });

  var zoom = document.getElementById('tb-zoom');
  function applyZoom() {
    document.querySelectorAll('.lpa-scroll').forEach(function (el) {
      el.style.setProperty('--tb-zoom', String((parseFloat(zoom.value) || 100) / 100));
    });
    // zoom changes the band's height, in both pivots
    PIVOT_IDS.forEach(function (id) { syncPivotHeader(document.getElementById(id)); });
  }
  zoom.addEventListener('input', applyZoom);

  document.getElementById('tb-expand').addEventListener('click', function () {
    collapsedDates = {}; collapsedModels = {}; applyList();
  });
  document.getElementById('tb-fold-model').addEventListener('click', function () {
    collapsedDates = {};
    document.querySelectorAll('.lpa-model-row').forEach(function (r) { collapsedModels[r.dataset.model] = 1; });
    closeDetail(); applyList();
  });
  document.getElementById('tb-fold-all').addEventListener('click', function () {
    document.querySelectorAll('.lpa-group-row').forEach(function (r) { collapsedDates[r.dataset.date] = 1; });
    closeDetail(); applyList();
  });

  // Absent when no checklist item failed — the table is not emitted at all then.
  var ngExpand = document.getElementById('tb-ng-expand');
  if (ngExpand) ngExpand.addEventListener('click', function () { collapsedNg = {}; applyNgDetail(); });
  var ngFold = document.getElementById('tb-ng-fold');
  if (ngFold) ngFold.addEventListener('click', function () {
    ngTable.querySelectorAll('.lpa-ng-group').forEach(function (r) { collapsedNg[r.dataset.model] = 1; });
    applyNgDetail();
  });

  applyZoom();
  selectTab('list');
})();
</script>
</body>
</html>
""";
    }

    private static void CleanupOldTokens(string keepToken)
    {
        try
        {
            if (!Directory.Exists(ExportRoot)) return;
            foreach (string sub in Directory.GetDirectories(ExportRoot))
            {
                if (string.Equals(Path.GetFileName(sub), keepToken, StringComparison.OrdinalIgnoreCase))
                    continue;
                try { Directory.Delete(sub, recursive: true); }
                catch { /* a folder may be momentarily locked; ignore */ }
            }
        }
        catch { /* best-effort cleanup */ }
    }
}
