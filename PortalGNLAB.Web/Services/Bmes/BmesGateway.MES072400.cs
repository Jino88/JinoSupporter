using System.Globalization;
using System.Text.Json;

namespace PortalGNLAB.Web.Services;

public static partial class BmesGateway
{
    /// <summary>Parsed BMES F-Cost response: data rows + the per-response column metadata
    /// from BottomGridColumnList that tells us what date / week / month each COLnnnn means.</summary>
    public sealed record BmesFCostResponse(List<FCostRow> Rows, List<FCostColumnMeta> Columns);

    /// <summary>
    /// MES072400 — F-Cost (input amount, F-Cost and F-Cost rate per model and material)
    /// for one query date, with 14 pre-aggregated period columns (COL0001..COL0014).
    /// Returns the rows (INAMT/FCOST/FRATE triplets, metadata copied down) and the
    /// per-response column metadata, or null when the call or the parse failed.
    /// </summary>
    public static async Task<BmesFCostResponse?> FetchFCostAsync(
        BmesWebSession session, DateTime sdate, IProgress<string>? progress)
    {
        // Query string mirrors what the BMES UI sends (captured via DevTools).
        // ZGUBN=D (day), ZTYPE=A (all), L_FACCO=GN (factory) — all empty filters left blank.
        string url =
            $"{BmesWebSession.BaseUrl}/MES072400/SearchList?perPage=" +
            $"&Condition.SDATE={sdate:yyyy-MM-dd}" +
            "&Condition.ZGUBN=D" +
            "&Condition.ZTYPE=A" +
            "&L_FACCO%5B0%5D=GN" +
            "&L_DIVIS%5B0%5D=" +
            "&L_CATEG%5B0%5D=" +
            "&L_PRODU%5B0%5D=" +
            "&L_ITEMS%5B0%5D=" +
            "&L_MODEL%5B0%5D=" +
            "&L_SERIE%5B0%5D=" +
            "&L_MODEL2%5B0%5D=" +
            "&L_VERID%5B0%5D=" +
            "&L_MATNR%5B0%5D=" +
            "&REMEM=N&page=1";

        try
        {
            var response = await session.Client.GetAsync(url);
            if (!response.IsSuccessStatusCode)
            {
                progress?.Report($"[WARN] F-Cost response error: {response.StatusCode}");
                return null;
            }

            string json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);

            // Server returns { result, message, data: { contents: [...] } }
            var root = doc.RootElement;
            if (root.TryGetProperty("result", out var ok) && ok.ValueKind == JsonValueKind.False)
            {
                progress?.Report("[WARN] BMES returned result=false");
                return null;
            }

            var dataObj  = root.GetProperty("data");
            var contents = dataObj.GetProperty("contents");
            var rows     = new List<FCostRow>(capacity: contents.GetArrayLength());

            foreach (var item in contents.EnumerateArray())
            {
                rows.Add(ParseFCostRow(item));
            }

            // BMES emits each Material as a strict INAMT → FCOST → FRATE triplet but blanks
            // out the metadata text fields on the FCOST/FRATE continuation rows (visual de-dup
            // for their own UI). Forward-fill from the most recent INAMT so every saved row
            // is self-contained and any later DB query / export sees the full Material context.
            ForwardFillFCostTriplets(rows);

            // BottomGridColumnList tells us what date / week / month each COLnnnn represents
            // for THIS response. Older / newer queries shift the mapping (today moves), so we
            // persist it alongside the rows rather than hard-coding it.
            var columns = ParseFCostColumnMeta(dataObj);

            progress?.Report($"  F-Cost: {rows.Count:N0} rows · {columns.Count} column meta");
            return new BmesFCostResponse(rows, columns);
        }
        catch (Exception ex)
        {
            progress?.Report($"[WARN] F-Cost parse error: {ex.Message}");
            return null;
        }
    }

    /// <summary>Reads the response's BottomGridColumnList — one entry per COLnnnn — and
    /// returns the 14-row metadata. Falls back to an empty list if the response shape is
    /// older / different (the report layer then uses generic "C1..C14" headers).</summary>
    private static List<FCostColumnMeta> ParseFCostColumnMeta(JsonElement dataObj)
    {
        var list = new List<FCostColumnMeta>();
        if (!dataObj.TryGetProperty("BottomGridColumnList", out var arr) ||
            arr.ValueKind != JsonValueKind.Array)
        {
            return list;
        }

        foreach (var item in arr.EnumerateArray())
        {
            string code   = item.TryGetProperty("ZPVTT",    out var z)  ? (z.GetString()  ?? "") : "";
            string header = item.TryGetProperty("ZPVTT_TX", out var zt) ? (zt.GetString() ?? "") : "";
            string pdate  = item.TryGetProperty("PDATE",    out var pd) ? (pd.GetString() ?? "") : "";
            string colName = item.TryGetProperty("ZCOLN",   out var cn) ? (cn.GetString() ?? "") : "";

            // ZCOLN is "COL0001".."COL0014" — strip prefix to recover the 1..14 index.
            int index = 0;
            if (colName.Length > 3 && int.TryParse(colName[3..], out int parsed))
                index = parsed;

            if (index < 1 || index > 14) continue;

            list.Add(new FCostColumnMeta
            {
                Index  = index,
                Code   = code,
                Header = header,
                PDate  = pdate,
            });
        }

        // Sort by index so the page can iterate in 1..14 order.
        list.Sort((a, b) => a.Index.CompareTo(b.Index));
        return list;
    }

    // ── Forward-fill blank metadata across INAMT/FCOST/FRATE triplets ───────────

    /// <summary>BMES blanks out the _TX (and many code) fields on FCOST/FRATE continuation
    /// rows because their own UI relies on Excel-style "see the row above" rendering. We
    /// copy them down so each saved row stands on its own — useful for direct SQL queries,
    /// CSV exports, or anyone reading rows out of order.</summary>
    private static void ForwardFillFCostTriplets(List<FCostRow> rows)
    {
        FCostRow? anchor = null;
        foreach (var r in rows)
        {
            // INAMT marks the start of a new Material — it always carries full metadata.
            if (string.Equals(r.ZType, "INAMT", StringComparison.OrdinalIgnoreCase))
            {
                anchor = r;
                continue;
            }
            if (anchor is null) continue;

            // Only fill blanks; never overwrite a value the row legitimately carries.
            if (string.IsNullOrEmpty(r.FaccoTx)) r.FaccoTx = anchor.FaccoTx;
            if (string.IsNullOrEmpty(r.PrdGrTx)) r.PrdGrTx = anchor.PrdGrTx;
            if (string.IsNullOrEmpty(r.VeridTx)) r.VeridTx = anchor.VeridTx;
            if (string.IsNullOrEmpty(r.ModNoTx)) r.ModNoTx = anchor.ModNoTx;
            if (string.IsNullOrEmpty(r.AssemTx)) r.AssemTx = anchor.AssemTx;
            if (string.IsNullOrEmpty(r.AbChgTx)) r.AbChgTx = anchor.AbChgTx;
            if (string.IsNullOrEmpty(r.MCodeTx)) r.MCodeTx = anchor.MCodeTx;
            if (string.IsNullOrEmpty(r.MatnrTx)) r.MatnrTx = anchor.MatnrTx;

            if (string.IsNullOrEmpty(r.Facco)) r.Facco = anchor.Facco;
            if (string.IsNullOrEmpty(r.Werks)) r.Werks = anchor.Werks;
            if (string.IsNullOrEmpty(r.PrdGr)) r.PrdGr = anchor.PrdGr;
            if (string.IsNullOrEmpty(r.ModNo)) r.ModNo = anchor.ModNo;
            if (string.IsNullOrEmpty(r.Verid)) r.Verid = anchor.Verid;
            if (string.IsNullOrEmpty(r.AbChg)) r.AbChg = anchor.AbChg;
            if (string.IsNullOrEmpty(r.Cat01)) r.Cat01 = anchor.Cat01;
            if (string.IsNullOrEmpty(r.Matnr)) r.Matnr = anchor.Matnr;
            if (string.IsNullOrEmpty(r.Assem)) r.Assem = anchor.Assem;
            if (string.IsNullOrEmpty(r.TwSyn)) r.TwSyn = anchor.TwSyn;
        }
    }

    // ── JSON → row ───────────────────────────────────────────────────────────────

    private static FCostRow ParseFCostRow(JsonElement item)
    {
        string S(string n) =>
            item.TryGetProperty(n, out var p) && p.ValueKind != JsonValueKind.Null
                ? p.ToString()
                : string.Empty;

        // BMES emits numbers as strings ("364672.000"). Parse with invariant culture.
        double D(string n)
        {
            if (!item.TryGetProperty(n, out var p) || p.ValueKind == JsonValueKind.Null)
                return 0;
            string raw = p.ToString();
            return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double v)
                ? v : 0;
        }

        return new FCostRow
        {
            RawJson  = item.GetRawText(),
            FaccoTx  = S("FACCO_TX"),  PrdGrTx = S("PRDGR_TX"), VeridTx = S("VERID_TX"),
            ModNoTx  = S("MODNO_TX"),  AssemTx = S("ASSEM_TX"), AbChgTx = S("ABCHG_TX"),
            MCodeTx  = S("MCODE_TX"),  MatnrTx = S("MATNR_TX"), ZTypeTx = S("ZTYPE_TX"),
            TSort    = S("TSORT"),     ZSort   = S("ZSORT"),    ZType   = S("ZTYPE"),
            Facco    = S("FACCO"),     Werks   = S("WERKS"),    PrdGr   = S("PRDGR"),
            ModNo    = S("MODNO"),     Verid   = S("VERID"),    AbChg   = S("ABCHG"),
            Cat01    = S("CAT01"),     Matnr   = S("MATNR"),    Assem   = S("ASSEM"),
            TwSyn    = S("TWSYN"),     ZValu   = S("ZVALU"),
            Col0001  = D("COL0001"), Col0002 = D("COL0002"), Col0003 = D("COL0003"),
            Col0004  = D("COL0004"), Col0005 = D("COL0005"), Col0006 = D("COL0006"),
            Col0007  = D("COL0007"), Col0008 = D("COL0008"), Col0009 = D("COL0009"),
            Col0010  = D("COL0010"), Col0011 = D("COL0011"), Col0012 = D("COL0012"),
            Col0013  = D("COL0013"), Col0014 = D("COL0014"),
        };
    }
}
