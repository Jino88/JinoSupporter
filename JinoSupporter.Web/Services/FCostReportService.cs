using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>Reads a fcost_*.db saved by FCostService and shapes the rows into
/// per-Material blocks (Input + Cost + Rate triplets) ready for table rendering.</summary>
public sealed class FCostReportService(NgRateSettingsService settings)
{
    private readonly NgRateSettingsService _settings = settings;

    public string? FindMostRecentDb()
    {
        string dir = _settings.DbSaveDirectory;
        if (!Directory.Exists(dir)) return null;
        return Directory.GetFiles(dir, "fcost_*.db")
            .OrderByDescending(File.GetLastWriteTime)
            .FirstOrDefault();
    }

    public Task<FCostReport> GenerateReportAsync(
        string dbPath,
        IReadOnlyList<ModelGroupRecord>? groupsForSubGroupRollup = null,
        IReadOnlyCollection<string>?     selectedGroupNames      = null,
        IProgress<string>? progress = null)
        => Task.Run(() => GenerateReport(dbPath, groupsForSubGroupRollup, selectedGroupNames, progress));

    private FCostReport GenerateReport(
        string dbPath,
        IReadOnlyList<ModelGroupRecord>? groups,
        IReadOnlyCollection<string>?     selectedGroupNames,
        IProgress<string>? progress)
    {
        progress?.Report("Loading fcost rows…");
        var rows    = LoadRows(dbPath);
        var columns = LoadColumns(dbPath);

        // BMES emits rows in strict sequential triplets: INAMT → FCOST → FRATE for the same
        // (FACCO/MODNO/MATNR) bucket, then on to the next bucket. The continuation rows often
        // come back with blank _TX fields (the BMES UI relies on the user "seeing" the prior
        // row's label) so a naive GroupBy on (Mod, Mat) breaks. Walk the rows in their stored
        // order instead and pair each FCOST/FRATE with the nearest preceding INAMT.
        FCostMaterialBlock? totalBlock = null;
        var perMaterial = new List<FCostMaterialBlock>();
        FCostMaterialBlock? cur = null;

        foreach (var r in rows)
        {
            if (string.Equals(r.ZType, "INAMT", StringComparison.OrdinalIgnoreCase))
            {
                // INAMT row starts a new block. The total block is the very first INAMT
                // we encounter that has empty Matnr (BMES emits totals first).
                bool isTotal = string.IsNullOrEmpty(r.Matnr) && string.IsNullOrEmpty(r.MatnrTx)
                            && totalBlock is null;

                string matLabel = !string.IsNullOrEmpty(r.MatnrTx) ? r.MatnrTx
                                : !string.IsNullOrEmpty(r.Matnr)   ? r.Matnr
                                : "(no material)";
                string display = isTotal
                    ? "Total (GN)"
                    : (string.IsNullOrEmpty(r.ModNoTx) ? matLabel : $"{r.ModNoTx} / {matLabel}");

                cur = new FCostMaterialBlock
                {
                    DisplayName  = display,
                    ProductGroup = r.PrdGrTx,
                    ModelNo      = r.ModNoTx,
                    Material     = string.IsNullOrEmpty(r.MatnrTx) ? r.Matnr : r.MatnrTx,
                    Verid        = r.VeridTx,
                    Input        = r,
                };

                if (isTotal) totalBlock = cur;
                else         perMaterial.Add(cur);
            }
            else if (string.Equals(r.ZType, "FCOST", StringComparison.OrdinalIgnoreCase))
            {
                if (cur is not null) cur.Cost = r;
            }
            else if (string.Equals(r.ZType, "FRATE", StringComparison.OrdinalIgnoreCase))
            {
                if (cur is not null) cur.Rate = r;
            }
        }

        // Sort biggest input first so the rows the user cares about are at the top.
        perMaterial = perMaterial
            .OrderByDescending(b => b.InputTotal)
            .ThenBy(b => b.DisplayName, StringComparer.Ordinal)
            .ToList();

        // ── SubGroup roll-up ────────────────────────────────────────────────────
        var subGroupAggs    = new List<FCostSubGroupAggregate>();
        var unmappedBlocks  = new List<FCostMaterialBlock>();
        if (groups is not null && groups.Count > 0)
        {
            BuildSubGroupAggregates(
                perMaterial, groups, selectedGroupNames, columns,
                out subGroupAggs, out unmappedBlocks);
        }

        return new FCostReport
        {
            DbPath            = dbPath,
            QueryDate         = rows.FirstOrDefault()?.QueryDate ?? "",
            TotalRows         = rows.Count,
            Total             = totalBlock,
            Materials         = perMaterial,
            SubGroups         = subGroupAggs,
            UnmappedMaterials = unmappedBlocks,
            Columns           = columns,
        };
    }

    private static List<FCostColumnMeta> LoadColumns(string dbPath)
    {
        var list = new List<FCostColumnMeta>();
        using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
        conn.Open();

        // FCostColumns table is optional — older fcost_*.db files (pre-column-meta) won't
        // have it. Probe sqlite_master once and bail with an empty list when missing so the
        // report layer falls back to generic "C1..C14" headers via FCostReport.ColumnHeaders.
        using (var probe = conn.CreateCommand())
        {
            probe.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='FCostColumns';";
            if (probe.ExecuteScalar() is null) return list;
        }

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT ColIndex, Code, Header, PDate FROM FCostColumns ORDER BY ColIndex;";
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new FCostColumnMeta
            {
                Index  = r.GetInt32(0),
                Code   = r.IsDBNull(1) ? "" : r.GetString(1),
                Header = r.IsDBNull(2) ? "" : r.GetString(2),
                PDate  = r.IsDBNull(3) ? "" : r.GetString(3),
            });
        }
        return list;
    }

    /// <summary>Walks the supplied ModelGroups (filtered by <paramref name="selectedGroupNames"/>
    /// when provided) and builds a LineShift→SubGroupKey map. Each Material block in
    /// <paramref name="materials"/> is then matched by its INAMT row's JoinKey first
    /// (MatnrTx_ModNoTx — what the user asked for), falling back to AltJoinKey
    /// (MatnrTx_VeridTx — NG Rate's actual LineShift format) so we still find a
    /// SubGroup if the user's preferred key shape doesn't match what's in ModelGroups.</summary>
    private static void BuildSubGroupAggregates(
        IReadOnlyList<FCostMaterialBlock>     materials,
        IReadOnlyList<ModelGroupRecord>       groups,
        IReadOnlyCollection<string>?          selectedGroupNames,
        IReadOnlyList<FCostColumnMeta>        columns,
        out List<FCostSubGroupAggregate>      aggregates,
        out List<FCostMaterialBlock>          unmapped)
    {
        // 1. Build LineShift → (group, mid, subKey, sub-display) lookup.
        var lsToSub = new Dictionary<string, (string Grp, string Mid, string Key, string Disp)>(
            StringComparer.OrdinalIgnoreCase);
        var subAccu = new Dictionary<string, FCostSubGroupAggregate>(StringComparer.Ordinal);

        bool IncludeGroup(string n) =>
            selectedGroupNames is null || selectedGroupNames.Count == 0 ||
            selectedGroupNames.Contains(n);

        foreach (var g in groups)
        {
            if (!IncludeGroup(g.Name)) continue;
            foreach (var mid in g.MidGroups)
            {
                foreach (var sub in mid.SubGroups)
                    CollectSubLeaves(g, mid, sub, parentPath: string.Empty, lsToSub);
            }
        }

        // 2. Walk Material blocks, sum Input + Cost into matching SubGroup buckets.
        unmapped = new List<FCostMaterialBlock>();
        foreach (var mat in materials)
        {
            if (mat.Input is null) continue;

            string primary = mat.Input.JoinKey;
            string fallback = mat.Input.AltJoinKey;

            // Primary key first (user-requested MatnrTx_ModNoTx). If that misses, retry
            // with the NG Rate LineShift format (MatnrTx_VeridTx) so we still match when
            // ModelGroups stores LineShifts in the canonical NG Rate shape.
            if (!lsToSub.TryGetValue(primary, out var hit) && !string.IsNullOrEmpty(fallback))
                lsToSub.TryGetValue(fallback, out hit);

            if (string.IsNullOrEmpty(hit.Key))
            {
                unmapped.Add(mat);
                continue;
            }

            if (!subAccu.TryGetValue(hit.Key, out var agg))
            {
                agg = new FCostSubGroupAggregate
                {
                    Display      = hit.Disp,
                    GroupName    = hit.Grp,
                    MidGroupName = hit.Mid,
                    SubGroupKey  = hit.Key,
                };
                subAccu[hit.Key] = agg;
            }

            // Sum each of the 14 columns from the Input/Cost rows. FRATE row is ignored —
            // we recompute Rate% from the SUMS, not by averaging existing rates (that's
            // what the user explicitly asked for: 더한뒤 백분율).
            for (int i = 1; i <= 14; i++)
            {
                agg.InputByCol[i - 1] += mat.Input?.GetCol(i) ?? 0;
                agg.CostByCol[i  - 1] += mat.Cost ?.GetCol(i) ?? 0;
            }
            agg.MatchedMaterialCount++;
        }

        // 3. Sort SubGroups by previous-week Rate% desc (highest defect cost-rate first), so
        // the worst-performing groups float to the top — same pattern NG Rate's Top-10 view
        // uses. Falls back to current-week Rate% / first-day Rate% / Display name if the
        // BottomGridColumnList didn't load (older fcost_*.db files).
        int sortColIndex0 = FindPreviousWeekColumnIndex0(columns);
        aggregates = subAccu.Values
            .OrderByDescending(a =>
            {
                if (sortColIndex0 < 0 || sortColIndex0 >= a.RateByCol.Length) return 0;
                return a.RateByCol[sortColIndex0];
            })
            .ThenByDescending(a => a.RateByCol.Length > 7 ? a.RateByCol[7] : 0) // 2nd: current-week Rate
            .ThenBy(a => a.Display, StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>Walks the (potentially-nested) sub-group tree and emits one rollup key per
    /// node that holds direct LineShifts. A sub with only nested children produces no key
    /// itself — its descendants are reported individually so e.g. UNIT's `_D3 / L` and
    /// `_D3 / R` show as separate F-Cost rows instead of collapsing into one `_D3` total.</summary>
    private static void CollectSubLeaves(
        ModelGroupRecord g,
        MidGroupRecord   mid,
        SubGroupRecord   sub,
        string           parentPath,
        Dictionary<string, (string Grp, string Mid, string Key, string Disp)> lsToSub)
    {
        string path = string.IsNullOrEmpty(sub.Name)
            ? parentPath
            : (string.IsNullOrEmpty(parentPath) ? sub.Name : $"{parentPath} / {sub.Name}");

        if (sub.LineShifts.Count > 0)
        {
            // Display shows only the leaf name (e.g. `TIU-C11-20_D3_L`), not the full
            // path — the page already prints the Mid context underneath the bold label.
            // Key keeps the path so two leaves with identical names under different
            // parents don't collide in the rollup accumulator.
            string subKey  = $"{g.Name}::{mid.Material}::{path}";
            string subDisp = !string.IsNullOrEmpty(sub.Name) ? sub.Name
                             : !string.IsNullOrEmpty(path)   ? path
                             : mid.Material;
            foreach (var ls in sub.LineShifts)
            {
                if (string.IsNullOrEmpty(ls)) continue;
                // Last-write-wins is fine — duplicate LineShifts inside ModelGroups
                // would be a config error, not something we silently mask here.
                lsToSub[ls] = (g.Name, mid.Material, subKey, subDisp);
            }
        }

        foreach (var child in sub.SubGroups)
            CollectSubLeaves(g, mid, child, path, lsToSub);
    }

    /// <summary>Finds the second Week-kind column (= previous week relative to the query
    /// date) in the BottomGridColumnList ordering. Returns 0-based index, or -1 if the
    /// metadata is missing / has no Week columns. Mirrors NG Rate's SortRefKey logic.</summary>
    private static int FindPreviousWeekColumnIndex0(IReadOnlyList<FCostColumnMeta> columns)
    {
        if (columns is null || columns.Count == 0) return -1;

        var weekIdx = columns
            .Where(c => c.Kind == FCostPeriodKind.Week)
            .Select(c => c.Index - 1)   // 1-based → 0-based for array lookup
            .OrderBy(i => i)
            .ToList();

        // BottomGridColumnList lists Week columns in DESC date order: [W19, W18, W17, W16].
        // Sorted ascending by index that becomes [W19_idx, W18_idx, ...] — index 1 in this
        // list is "previous week" (W18) which is exactly what we want.
        if (weekIdx.Count >= 2) return weekIdx[1];
        if (weekIdx.Count == 1) return weekIdx[0];
        return -1;
    }

    private static List<FCostRow> LoadRows(string dbPath)
    {
        var rows = new List<FCostRow>();
        using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            SELECT Id, FetchedAt, QueryDate,
                   FaccoTx, PrdGrTx, VeridTx, ModNoTx, AssemTx, AbChgTx,
                   MCodeTx, MatnrTx, ZTypeTx, ZType, TSort, ZSort,
                   Facco, Werks, PrdGr, ModNo, Verid, AbChg, Cat01, Matnr, Assem,
                   TwSyn, ZValu,
                   Col0001, Col0002, Col0003, Col0004, Col0005, Col0006, Col0007,
                   Col0008, Col0009, Col0010, Col0011, Col0012, Col0013, Col0014
            FROM FCostRows
            ORDER BY Id;
            """;

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            // Manual column-index walker — i++ inside a ternary skips past DBNull columns
            // so we'd lose alignment; advance once per cell with explicit Sx()/Dx() locals.
            int i = 0;
            long   id  = r.GetInt64(i++);
            string fea = r.IsDBNull(i) ? "" : r.GetString(i); i++;
            string qd  = r.IsDBNull(i) ? "" : r.GetString(i); i++;

            string Sx() { string v = r.IsDBNull(i) ? "" : r.GetString(i); i++; return v; }
            double Dx() { double v = r.IsDBNull(i) ? 0  : r.GetDouble(i); i++; return v; }

            var row = new FCostRow
            {
                Id        = id,
                FetchedAt = fea,
                QueryDate = qd,
                FaccoTx   = Sx(), PrdGrTx = Sx(), VeridTx = Sx(), ModNoTx = Sx(),
                AssemTx   = Sx(), AbChgTx = Sx(), MCodeTx = Sx(), MatnrTx = Sx(),
                ZTypeTx   = Sx(), ZType   = Sx(), TSort   = Sx(), ZSort   = Sx(),
                Facco     = Sx(), Werks   = Sx(), PrdGr   = Sx(), ModNo   = Sx(),
                Verid     = Sx(), AbChg   = Sx(), Cat01   = Sx(), Matnr   = Sx(),
                Assem     = Sx(), TwSyn   = Sx(), ZValu   = Sx(),
                Col0001   = Dx(), Col0002 = Dx(), Col0003 = Dx(), Col0004 = Dx(),
                Col0005   = Dx(), Col0006 = Dx(), Col0007 = Dx(), Col0008 = Dx(),
                Col0009   = Dx(), Col0010 = Dx(), Col0011 = Dx(), Col0012 = Dx(),
                Col0013   = Dx(), Col0014 = Dx(),
            };
            rows.Add(row);
        }

        return rows;
    }
}
