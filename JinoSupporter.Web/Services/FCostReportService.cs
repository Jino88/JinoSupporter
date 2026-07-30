using Microsoft.Data.Sqlite;
using System.Globalization;

namespace JinoSupporter.Web.Services;

/// <summary>Reads a fcost_*.db saved by FCostService and shapes the rows into
/// per-Material blocks (Input + Cost + Rate triplets) ready for table rendering.</summary>
public sealed class FCostReportService(NgRateSettingsService settings)
{
    private readonly NgRateSettingsService _settings = settings;

    public string? FindMostRecentDb()
    {
        string dir = _settings.FCostDbSaveDirectory;
        if (!Directory.Exists(dir)) return null;
        return Directory.GetFiles(dir, "fcost_????????_??????.db")
            .OrderByDescending(File.GetLastWriteTime)
            .FirstOrDefault();
    }

    public Task<FCostReport> GenerateReportAsync(
        string dbPath,
        IReadOnlyList<ModelGroupRecord>? groupsForSubGroupRollup = null,
        IReadOnlyCollection<string>?     selectedGroupNames      = null,
        IProgress<string>? progress = null)
        => Task.Run(() => GenerateReport(dbPath, groupsForSubGroupRollup, selectedGroupNames, progress));

    public Task<FCostReport> GenerateRawRangeReportAsync(
        string rawDbPath,
        DateTime startDate,
        DateTime endDate,
        IReadOnlyList<ModelGroupRecord>? groupsForSubGroupRollup = null,
        IReadOnlyCollection<string>?     selectedGroupNames      = null,
        IProgress<string>? progress = null)
        => Task.Run(() =>
        {
            startDate = startDate.Date;
            endDate = endDate.Date;
            progress?.Report("Loading F-COST RAW range...");
            var columns = BuildPeriodColumns(startDate, endDate);
            var rows = LoadRawRangeRows(rawDbPath, startDate, endDate, columns);
            string label = $"{startDate:yyyy-MM-dd} - {endDate:yyyy-MM-dd}";
            return BuildReport(rawDbPath, rows, columns, label, groupsForSubGroupRollup, selectedGroupNames);
        });

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
        var (totalBlock, perMaterial) = BuildMaterialBlocks(rows);

        // Sort biggest input first so the rows the user cares about are at the top.
        perMaterial = perMaterial
            .OrderByDescending(MaterialSortInput)
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

    private static FCostReport BuildReport(
        string dbPath,
        List<FCostRow> rows,
        List<FCostColumnMeta> columns,
        string queryLabel,
        IReadOnlyList<ModelGroupRecord>? groups,
        IReadOnlyCollection<string>? selectedGroupNames)
    {
        var (totalBlock, perMaterial) = BuildMaterialBlocks(rows);

        perMaterial = perMaterial
            .OrderByDescending(MaterialSortInput)
            .ThenBy(b => b.DisplayName, StringComparer.Ordinal)
            .ToList();

        var subGroupAggs = new List<FCostSubGroupAggregate>();
        var unmappedBlocks = new List<FCostMaterialBlock>();
        if (groups is not null && groups.Count > 0)
        {
            BuildSubGroupAggregates(
                perMaterial, groups, selectedGroupNames, columns,
                out subGroupAggs, out unmappedBlocks);
        }

        return new FCostReport
        {
            DbPath = dbPath,
            QueryDate = queryLabel,
            TotalRows = rows.Count,
            Total = totalBlock,
            Materials = perMaterial,
            SubGroups = subGroupAggs,
            UnmappedMaterials = unmappedBlocks,
            Columns = columns,
        };
    }

    private static double MaterialSortInput(FCostMaterialBlock block)
        => block.Input?.Values is { Length: > 0 } vals ? vals.Sum() : block.InputTotal;

    private static (FCostMaterialBlock? TotalBlock, List<FCostMaterialBlock> PerMaterial) BuildMaterialBlocks(
        IReadOnlyList<FCostRow> rows)
    {
        const string totalKey = "__TOTAL__";
        var blocks = new Dictionary<string, FCostMaterialBlock>(StringComparer.Ordinal);
        var order = new Dictionary<string, int>(StringComparer.Ordinal);
        FCostMaterialBlock? totalBlock = null;

        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            string ztype = row.ZType.Trim();
            if (!string.Equals(ztype, "INAMT", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(ztype, "FCOST", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(ztype, "FRATE", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            bool isTotal = IsBmesTotalRow(row);
            string key = isTotal ? totalKey : MaterialBlockKey(row);
            if (string.IsNullOrWhiteSpace(key))
                key = "__ROW__" + i.ToString(CultureInfo.InvariantCulture);

            if (!blocks.TryGetValue(key, out var block))
            {
                block = CreateMaterialBlock(row, isTotal);
                blocks[key] = block;
                order[key] = i;
            }

            if (isTotal && totalBlock is null)
                totalBlock = block;

            if (string.Equals(ztype, "INAMT", StringComparison.OrdinalIgnoreCase))
                block.Input = row;
            else if (string.Equals(ztype, "FCOST", StringComparison.OrdinalIgnoreCase))
                block.Cost = row;
            else if (string.Equals(ztype, "FRATE", StringComparison.OrdinalIgnoreCase))
                block.Rate = row;
        }

        var perMaterial = blocks
            .Where(kv => !string.Equals(kv.Key, totalKey, StringComparison.Ordinal) && !ReferenceEquals(kv.Value, totalBlock))
            .OrderBy(kv => order.GetValueOrDefault(kv.Key, int.MaxValue))
            .Select(kv => kv.Value)
            .ToList();

        return (totalBlock, perMaterial);
    }

    private static FCostMaterialBlock CreateMaterialBlock(FCostRow row, bool isTotal)
    {
        string matLabel = !string.IsNullOrEmpty(row.MatnrTx) ? row.MatnrTx
                        : !string.IsNullOrEmpty(row.Matnr)   ? row.Matnr
                        : "(no material)";
        string display = isTotal
            ? "Total (GN)"
            : (string.IsNullOrEmpty(row.ModNoTx) ? matLabel : $"{row.ModNoTx} / {matLabel}");

        return new FCostMaterialBlock
        {
            DisplayName = display,
            ProductGroup = row.PrdGrTx,
            ModelNo = row.ModNoTx,
            Material = string.IsNullOrEmpty(row.MatnrTx) ? row.Matnr : row.MatnrTx,
            Verid = row.VeridTx,
        };
    }

    private static string MaterialBlockKey(FCostRow row)
        => string.Join('\t', new[]
        {
            row.Facco, row.Werks, row.PrdGr, row.ModNo, row.Verid, row.AbChg,
            row.Cat01, row.Matnr, row.Assem, row.TwSyn, row.ZValu, row.FaccoTx,
            row.PrdGrTx, row.VeridTx, row.ModNoTx, row.AssemTx, row.AbChgTx,
            row.MCodeTx, row.MatnrTx,
        });

    private static bool IsBmesTotalRow(FCostRow row)
        => string.IsNullOrWhiteSpace(row.Matnr)
        && string.IsNullOrWhiteSpace(row.ModNo)
        && string.IsNullOrWhiteSpace(row.ModNoTx)
        && (string.IsNullOrWhiteSpace(row.MatnrTx) ||
            string.Equals(row.MatnrTx.Trim(), "Total", StringComparison.OrdinalIgnoreCase));

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

    private static List<FCostColumnMeta> BuildPeriodColumns(DateTime startDate, DateTime endDate)
    {
        var list = new List<FCostColumnMeta>();
        int index = 1;
        var days = Enumerable.Range(0, (endDate.Date - startDate.Date).Days + 1)
            .Select(offset => startDate.Date.AddDays(offset))
            .ToList();

        foreach (DateTime day in days.OrderByDescending(d => d))
        {
            list.Add(new FCostColumnMeta
            {
                Index = index++,
                Code = day.ToString("yyyyMMdd"),
                Header = day.ToString("MM-dd"),
                PDate = day.ToString("yyyyMMdd"),
            });
        }

        foreach (var week in days
            .Select(d => new { Year = ISOWeek.GetYear(d), Week = ISOWeek.GetWeekOfYear(d) })
            .Distinct()
            .OrderByDescending(x => x.Year)
            .ThenByDescending(x => x.Week))
        {
            list.Add(new FCostColumnMeta
            {
                Index = index++,
                Code = $"{week.Year}-{week.Week:D2}",
                Header = $"{week.Year % 100:D2}-W{week.Week:D2}",
                PDate = $"{week.Year}{week.Week:D2}",
            });
        }

        foreach (DateTime month in days
            .Select(d => new DateTime(d.Year, d.Month, 1))
            .Distinct()
            .OrderByDescending(d => d))
        {
            list.Add(new FCostColumnMeta
            {
                Index = index++,
                Code = month.ToString("yyyyMM"),
                Header = month.ToString("yy-MM"),
                PDate = month.ToString("yyyyMM"),
            });
        }

        return list;
    }

    private static List<FCostRow> LoadRawRangeRows(
        string rawDbPath,
        DateTime startDate,
        DateTime endDate,
        IReadOnlyList<FCostColumnMeta> reportColumns)
    {
        if (!File.Exists(rawDbPath)) return new List<FCostRow>();

        string startIso = startDate.ToString("yyyy-MM-dd");
        string endIso = endDate.ToString("yyyy-MM-dd");

        using var conn = new SqliteConnection($"Data Source={rawDbPath};Mode=ReadOnly");
        conn.Open();

        var sourcesByQueryDate = LoadPeriodColumnSources(conn, startIso, endIso, reportColumns)
            .GroupBy(s => s.QueryDate, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.ToArray(), StringComparer.Ordinal);
        if (sourcesByQueryDate.Count == 0) return new List<FCostRow>();

        var byKey = new Dictionary<string, (int Order, FCostRow Row)>(StringComparer.Ordinal);

        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            SELECT QueryDate, RowNo, FetchedAt, RawJson,
                   FaccoTx, PrdGrTx, VeridTx, ModNoTx, AssemTx, AbChgTx,
                   MCodeTx, MatnrTx, ZTypeTx, ZType, TSort, ZSort,
                   Facco, Werks, PrdGr, ModNo, Verid, AbChg, Cat01, Matnr, Assem,
                   TwSyn, ZValu,
                   Col0001, Col0002, Col0003, Col0004, Col0005, Col0006, Col0007,
                   Col0008, Col0009, Col0010, Col0011, Col0012, Col0013, Col0014
            FROM FCostRawRows
            WHERE QueryDate BETWEEN @start AND @end
            ORDER BY QueryDate, RowNo;
            """;
        cmd.Parameters.AddWithValue("@start", startIso);
        cmd.Parameters.AddWithValue("@end", endIso);

        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            int i = 0;
            string queryDate = r.IsDBNull(i) ? "" : r.GetString(i); i++;
            int rowNo = r.IsDBNull(i) ? 0 : r.GetInt32(i); i++;
            string fetched = r.IsDBNull(i) ? "" : r.GetString(i); i++;
            string rawJson = r.IsDBNull(i) ? "" : r.GetString(i); i++;

            string Sx() { string v = r.IsDBNull(i) ? "" : r.GetString(i); i++; return v; }
            double Dx() { double v = r.IsDBNull(i) ? 0 : r.GetDouble(i); i++; return v; }

            var raw = new FCostRow
            {
                Id = rowNo,
                FetchedAt = fetched,
                QueryDate = queryDate,
                RawJson = rawJson,
                FaccoTx = Sx(), PrdGrTx = Sx(), VeridTx = Sx(), ModNoTx = Sx(),
                AssemTx = Sx(), AbChgTx = Sx(), MCodeTx = Sx(), MatnrTx = Sx(),
                ZTypeTx = Sx(), ZType = Sx(), TSort = Sx(), ZSort = Sx(),
                Facco = Sx(), Werks = Sx(), PrdGr = Sx(), ModNo = Sx(),
                Verid = Sx(), AbChg = Sx(), Cat01 = Sx(), Matnr = Sx(),
                Assem = Sx(), TwSyn = Sx(), ZValu = Sx(),
                Col0001 = Dx(), Col0002 = Dx(), Col0003 = Dx(), Col0004 = Dx(),
                Col0005 = Dx(), Col0006 = Dx(), Col0007 = Dx(), Col0008 = Dx(),
                Col0009 = Dx(), Col0010 = Dx(), Col0011 = Dx(), Col0012 = Dx(),
                Col0013 = Dx(), Col0014 = Dx(),
            };

            if (!sourcesByQueryDate.TryGetValue(queryDate, out var sources))
                continue;

            string key = RangeGroupKey(raw);
            if (!byKey.TryGetValue(key, out var entry))
            {
                entry = (rowNo, CloneForRange(raw, reportColumns.Count));
                byKey[key] = entry;
            }

            foreach (var source in sources)
            {
                if (entry.Row.Values is null ||
                    source.ReportIndex < 0 ||
                    source.ReportIndex >= entry.Row.Values.Length)
                {
                    continue;
                }

                entry.Row.Values[source.ReportIndex] += StoredColValue(raw, source.ColIndex);
            }
        }

        return byKey.Values
            .OrderBy(x => x.Order)
            .Select(x => x.Row)
            .ToList();
    }

    private sealed record PeriodColumnSource(int ReportIndex, string QueryDate, int ColIndex);

    private static List<PeriodColumnSource> LoadPeriodColumnSources(
        SqliteConnection conn,
        string startIso,
        string endIso,
        IReadOnlyList<FCostColumnMeta> reportColumns)
    {
        var sources = new List<PeriodColumnSource>(reportColumns.Count);

        for (int i = 0; i < reportColumns.Count; i++)
        {
            var col = reportColumns[i];
            if (col.Kind == FCostPeriodKind.Day)
            {
                if (!DateTime.TryParseExact(col.PDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime day))
                    continue;

                string qdate = day.ToString("yyyy-MM-dd");
                if (string.CompareOrdinal(qdate, startIso) < 0 || string.CompareOrdinal(qdate, endIso) > 0)
                    continue;

                int colIndex = FindExactRawColumn(conn, qdate, "Day", col.PDate);
                if (colIndex == 0) colIndex = FindFirstRawColumn(conn, qdate, "Day");
                if (colIndex > 0) sources.Add(new PeriodColumnSource(i, qdate, colIndex));
                continue;
            }

            if (col.Kind is FCostPeriodKind.Week or FCostPeriodKind.Month)
            {
                var source = FindLatestRawPeriodColumn(conn, startIso, endIso, col.Kind.ToString(), col.PDate);
                if (source is not null)
                    sources.Add(new PeriodColumnSource(i, source.Value.QueryDate, source.Value.ColIndex));
            }
        }

        return sources;
    }

    private static int FindExactRawColumn(SqliteConnection conn, string queryDate, string kind, string pdate)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            SELECT ColIndex
            FROM FCostRawColumns
            WHERE QueryDate = @qdate
              AND Kind = @kind
              AND (PDate = @pdate OR Code = @pdate)
            ORDER BY ColIndex
            LIMIT 1;
            """;
        cmd.Parameters.AddWithValue("@qdate", queryDate);
        cmd.Parameters.AddWithValue("@kind", kind);
        cmd.Parameters.AddWithValue("@pdate", pdate);
        return ScalarToInt(cmd.ExecuteScalar());
    }

    private static int FindFirstRawColumn(SqliteConnection conn, string queryDate, string kind)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            SELECT MIN(ColIndex)
            FROM FCostRawColumns
            WHERE QueryDate = @qdate
              AND Kind = @kind;
            """;
        cmd.Parameters.AddWithValue("@qdate", queryDate);
        cmd.Parameters.AddWithValue("@kind", kind);
        return ScalarToInt(cmd.ExecuteScalar());
    }

    private static int ScalarToInt(object? value)
        => value is null or DBNull ? 0 : Convert.ToInt32(value, CultureInfo.InvariantCulture);

    private static (string QueryDate, int ColIndex)? FindLatestRawPeriodColumn(
        SqliteConnection conn,
        string startIso,
        string endIso,
        string kind,
        string pdate)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            SELECT QueryDate, ColIndex
            FROM FCostRawColumns
            WHERE QueryDate BETWEEN @start AND @end
              AND Kind = @kind
              AND PDate = @pdate
            ORDER BY QueryDate DESC
            LIMIT 1;
            """;
        cmd.Parameters.AddWithValue("@start", startIso);
        cmd.Parameters.AddWithValue("@end", endIso);
        cmd.Parameters.AddWithValue("@kind", kind);
        cmd.Parameters.AddWithValue("@pdate", pdate);
        using var r = cmd.ExecuteReader();
        return r.Read() ? (r.GetString(0), r.GetInt32(1)) : null;
    }

    private static Dictionary<string, int> LoadRawDayColumnIndexes(
        SqliteConnection conn,
        string startIso,
        string endIso)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);

        using (var exact = conn.CreateCommand())
        {
            exact.CommandText =
                """
                SELECT QueryDate, ColIndex
                FROM FCostRawColumns
                WHERE QueryDate BETWEEN @start AND @end
                  AND Kind = 'Day'
                  AND (PDate = REPLACE(QueryDate, '-', '') OR Code = REPLACE(QueryDate, '-', ''));
                """;
            exact.Parameters.AddWithValue("@start", startIso);
            exact.Parameters.AddWithValue("@end", endIso);
            using var r = exact.ExecuteReader();
            while (r.Read())
                map[r.GetString(0)] = r.GetInt32(1);
        }

        using (var fallback = conn.CreateCommand())
        {
            fallback.CommandText =
                """
                SELECT QueryDate, MIN(ColIndex)
                FROM FCostRawColumns
                WHERE QueryDate BETWEEN @start AND @end
                  AND Kind = 'Day'
                GROUP BY QueryDate;
                """;
            fallback.Parameters.AddWithValue("@start", startIso);
            fallback.Parameters.AddWithValue("@end", endIso);
            using var r = fallback.ExecuteReader();
            while (r.Read())
            {
                string qdate = r.GetString(0);
                if (!map.ContainsKey(qdate))
                    map[qdate] = r.GetInt32(1);
            }
        }

        return map;
    }

    private static FCostRow CloneForRange(FCostRow source, int columnCount) => new()
    {
        Id = source.Id,
        FetchedAt = source.FetchedAt,
        QueryDate = source.QueryDate,
        RawJson = source.RawJson,
        FaccoTx = source.FaccoTx,
        PrdGrTx = source.PrdGrTx,
        VeridTx = source.VeridTx,
        ModNoTx = source.ModNoTx,
        AssemTx = source.AssemTx,
        AbChgTx = source.AbChgTx,
        MCodeTx = source.MCodeTx,
        MatnrTx = source.MatnrTx,
        ZTypeTx = source.ZTypeTx,
        TSort = source.TSort,
        ZSort = source.ZSort,
        ZType = source.ZType,
        Facco = source.Facco,
        Werks = source.Werks,
        PrdGr = source.PrdGr,
        ModNo = source.ModNo,
        Verid = source.Verid,
        AbChg = source.AbChg,
        Cat01 = source.Cat01,
        Matnr = source.Matnr,
        Assem = source.Assem,
        TwSyn = source.TwSyn,
        ZValu = source.ZValu,
        Values = new double[columnCount],
    };

    private static string RangeGroupKey(FCostRow r)
        => string.Join('\t', new[]
        {
            r.ZType, r.Facco, r.Werks, r.PrdGr, r.ModNo, r.Verid, r.AbChg,
            r.Cat01, r.Matnr, r.Assem, r.TwSyn, r.ZValu, r.FaccoTx, r.PrdGrTx,
            r.VeridTx, r.ModNoTx, r.AssemTx, r.AbChgTx, r.MCodeTx, r.MatnrTx,
            r.ZTypeTx, r.TSort, r.ZSort,
        });

    private static double StoredColValue(FCostRow row, int colIndex) => colIndex switch
    {
        1 => row.Col0001, 2 => row.Col0002, 3 => row.Col0003, 4 => row.Col0004,
        5 => row.Col0005, 6 => row.Col0006, 7 => row.Col0007, 8 => row.Col0008,
        9 => row.Col0009, 10 => row.Col0010, 11 => row.Col0011, 12 => row.Col0012,
        13 => row.Col0013, 14 => row.Col0014,
        _ => 0,
    };

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
        var sourceRateNumeratorBySubGroup = new Dictionary<string, double[]>(StringComparer.Ordinal);
        var sourceRateInputBySubGroup     = new Dictionary<string, double[]>(StringComparer.Ordinal);
        int columnCount = columns.Count > 0 ? columns.Count : 14;

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
                double[] numerator = new double[columnCount];
                double[] weight = new double[columnCount];
                agg = new FCostSubGroupAggregate
                {
                    Display      = hit.Disp,
                    GroupName    = hit.Grp,
                    MidGroupName = hit.Mid,
                    SubGroupKey  = hit.Key,
                    InputByCol   = new double[columnCount],
                    CostByCol    = new double[columnCount],
                    SourceRateByCol = new double[columnCount],
                };
                subAccu[hit.Key] = agg;
                sourceRateNumeratorBySubGroup[hit.Key] = numerator;
                sourceRateInputBySubGroup[hit.Key] = weight;
            }

            double[] rateNumerator = sourceRateNumeratorBySubGroup[hit.Key];
            double[] rateWeight    = sourceRateInputBySubGroup[hit.Key];

            // Sum each of the 14 columns from the Input/Cost rows.
            // Keep FRATE for source-rate aggregation separately so we can show
            // weighted FRATE at display time instead of recomputing it.
            for (int i = 1; i <= columnCount; i++)
            {
                double input = mat.Input?.GetCol(i) ?? 0;
                double cost  = mat.Cost ?.GetCol(i) ?? 0;
                double rate  = mat.Rate?.GetCol(i) ?? 0;

                agg.InputByCol[i - 1] += input;
                agg.CostByCol[i  - 1] += cost;

                rateNumerator[i - 1] += input * rate;
                rateWeight[i - 1] += input;
            }
            agg.MatchedMaterialCount++;
        }

        // 2.5. Build weighted FRATE from child rows:
        //     sum(INPUT * FRATE) / sum(INPUT).
        foreach (var agg in subAccu.Values)
        {
            if (!sourceRateNumeratorBySubGroup.TryGetValue(agg.SubGroupKey, out var numerator) ||
                !sourceRateInputBySubGroup.TryGetValue(agg.SubGroupKey, out var weight))
                continue;

            int count = Math.Min(agg.SourceRateByCol.Length, columnCount);
            for (int i = 0; i < count; i++)
            {
                if (weight[i] > 0)
                    agg.SourceRateByCol[i] = numerator[i] / weight[i];
                else if (agg.InputByCol[i] > 0)
                    agg.SourceRateByCol[i] = agg.CostByCol[i] / agg.InputByCol[i] * 100.0;
                else
                    agg.SourceRateByCol[i] = 0;
            }
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
