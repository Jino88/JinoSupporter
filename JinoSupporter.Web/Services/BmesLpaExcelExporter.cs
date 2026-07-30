using ClosedXML.Excel;

namespace JinoSupporter.Web.Services;

/// <summary>
/// The LPA result as a multi-sheet .xlsx, mirroring the tabs of the viewer:
///   1. <b>NG Summary</b> — the Model × Date/Week/Month pivot (ppm + NG/TOTAL), the first sheet.
///   2. <b>NG by Check Item</b> — the same periods by Check Item, but rated against how often each check ran.
///   3. <b>Defect Detail</b> — the AULOC · Check Point · Check Item × Date matrix with the NG photos embedded.
///   4. <b>List</b> — the flat audit list.
///
/// Every sheet renders from the same structures the HTML tabs are built from
/// (<see cref="BmesLpaHtmlExportService.PivotData"/>, <see cref="BmesLpaHtmlExportService.NgMatrix"/>,
/// <see cref="BmesLpaHtmlExportService.BuildListTable"/>) so the file cannot drift from the
/// screen. Photo bytes are passed in (the caller fetches the popup-size copy through
/// <see cref="BmesLpaImageService"/>): the picture is embedded at that full resolution and only
/// <i>displayed</i> small, so a cell stays compact yet enlarging it in Excel keeps it sharp.
/// </summary>
public static class BmesLpaExcelExporter
{
    private static readonly XLColor HeaderBg = XLColor.FromHtml("#F1F5F9");
    private static readonly XLColor BandBg    = XLColor.FromHtml("#EEF2F7");
    private static readonly XLColor GroupBg   = XLColor.FromHtml("#E2E8F0");
    private static readonly XLColor TotalBg   = XLColor.FromHtml("#EFF6FF");
    private static readonly XLColor QtyFg     = XLColor.FromHtml("#B91C1C");

    private const int ThumbPx = 64;   // on-sheet display size of each photo (embedded bytes are higher-res)
    private const int GapPx   = 2;
    private const int PerRow  = 3;    // photos across a cell before wrapping to the next line
    private const int TextPx  = 16;   // room kept above the photos for the count

    /// <summary>Every ZIMAG_TX photo the matrix references, so the caller knows exactly which
    /// bytes to fetch before exporting (nothing more, nothing less than the sheet shows).</summary>
    public static List<string> ImagePaths(BmesLpaHtmlExportService.NgMatrix matrix) =>
        matrix.Groups
            .SelectMany(g => g.Rows)
            .SelectMany(r => r.ItemsByDate.Values)
            .SelectMany(items => items)
            .SelectMany(it => BmesLpaHtmlExportService.ItemImagePaths(it.Item))
            .Distinct(StringComparer.Ordinal)
            .ToList();

    public static byte[] Export(
        BmesLpaHtmlExportService.PivotData pivot,
        IReadOnlyList<BmesLpaHtmlExportService.CheckPivotRow> checkRows,
        BmesLpaHtmlExportService.NgMatrix matrix,
        IReadOnlyDictionary<string, byte[]> images,
        (IReadOnlyList<string> Headers, IReadOnlyList<IReadOnlyList<string>> Rows) list,
        string title)
    {
        using var wb = new XLWorkbook();
        WritePivotSheet(wb.AddWorksheet("NG Summary"), pivot, title);
        WriteCheckPivotSheet(wb.AddWorksheet("NG by Check Item"), pivot, checkRows, title);
        WriteMatrixSheet(wb.AddWorksheet("Defect Detail"), matrix, images, title);
        WriteListSheet(wb.AddWorksheet("List"), list);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    // ── Sheet 1: NG Summary pivot ──────────────────────────────────────────────────────────
    // Each model (and the Total) takes TWO rows: a NG Rate(ppm) row and a Count(NG/TOTAL) row, the
    // pair labelled in a Type column, with the model name merged down both. Columns:
    // A=Model, B=Type, C=Overall, D.. = the period columns.
    private static void WritePivotSheet(IXLWorksheet ws, BmesLpaHtmlExportService.PivotData pivot, string title)
    {
        var periods = new List<(string Key, string Header)>();
        foreach (var blk in pivot.Blocks)
            for (int i = 0; i < blk.Keys.Count; i++)
                periods.Add((blk.Keys[i], blk.Headers[i]));
        const int firstPeriodCol = 4;
        int lastCol = 3 + periods.Count;

        SetText(ws.Cell(1, 1), title, bold: true, size: 11);
        ws.Range(1, 1, 1, Math.Max(1, lastCol)).Merge();

        if (pivot.Models.Count == 0)
        {
            SetText(ws.Cell(2, 1), "No rows to aggregate.");
            return;
        }

        // Row 2 = band, Row 3 = period headers. Model/Type/Overall span both header rows.
        Merge(ws, 2, 1, 3, 1, "Model", HeaderBg);
        Merge(ws, 2, 2, 3, 2, "Type", HeaderBg);
        Merge(ws, 2, 3, 3, 3, "Overall", HeaderBg);

        int col = firstPeriodCol;
        foreach (var blk in pivot.Blocks)
        {
            if (blk.Keys.Count == 0) continue;
            int start = col;
            for (int i = 0; i < blk.Keys.Count; i++)
            {
                var h = ws.Cell(3, col);
                SetText(h, blk.Headers[i], bold: true);
                h.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                h.Style.Fill.BackgroundColor = HeaderBg;
                col++;
            }
            Merge(ws, 2, start, 2, col - 1, blk.Label, BandBg);
        }

        int r = 4;
        WritePivotEntity(ws, r, "Total", periods, k => pivot.ColTotals.GetValueOrDefault(k),
                         pivot.GrandTotal, isTotal: true);
        r += 2;

        foreach (string m in pivot.Models)
        {
            string model = m;
            WritePivotEntity(ws, r, model, periods,
                             k => pivot.Cells.TryGetValue((model, k), out var b) ? b : null,
                             pivot.ModelOverall.GetValueOrDefault(model), isTotal: false);
            r += 2;
        }

        ws.Column(1).Width = 16;   // Model
        ws.Column(2).Width = 8;    // Type
        ws.Column(3).Width = 11;   // Overall
        for (int i = 0; i < periods.Count; i++) ws.Column(firstPeriodCol + i).Width = 10;
        DrawBlockBorders(ws, pivot, firstPeriodCol, 2, r - 1);
        ws.SheetView.FreezeRows(5);        // title + 2 header rows + Total's two rows
        ws.SheetView.FreezeColumns(3);
    }

    /// <summary>One entity across two rows: NG Rate(ppm) then Count(NG/TOTAL), model name merged
    /// down both. <paramref name="bucketAt"/> yields the period cell, null = not audited.</summary>
    private static void WritePivotEntity(
        IXLWorksheet ws, int r, string name,
        List<(string Key, string Header)> periods,
        Func<string, BmesLpaHtmlExportService.PivotBucket?> bucketAt,
        BmesLpaHtmlExportService.PivotBucket overall, bool isTotal)
    {
        int lastCol = 3 + periods.Count;

        SetText(ws.Cell(r, 1), name, bold: true);
        ws.Range(r, 1, r + 1, 1).Merge().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

        SetText(ws.Cell(r, 2), "NG Rate");
        SetText(ws.Cell(r + 1, 2), "Count");

        WriteRate(ws.Cell(r, 3), overall);
        WriteCount(ws.Cell(r + 1, 3), overall);
        for (int i = 0; i < periods.Count; i++)
        {
            var b = bucketAt(periods[i].Key);
            WriteRate(ws.Cell(r, 4 + i), b);
            WriteCount(ws.Cell(r + 1, 4 + i), b);
        }

        if (isTotal)
        {
            ws.Range(r, 1, r + 1, lastCol).Style.Fill.BackgroundColor = TotalBg;
            ws.Range(r, 1, r + 1, lastCol).Style.Font.Bold = true;
        }
    }

    private static void WriteRate(IXLCell cell, BmesLpaHtmlExportService.PivotBucket? bucket)
    {
        if (bucket is not { } v) return;
        cell.Value = BmesLpaHtmlExportService.NgRateText(v.Ng, v.Total);
        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }

    private static void WriteCount(IXLCell cell, BmesLpaHtmlExportService.PivotBucket? bucket)
    {
        if (bucket is not { } v) return;
        cell.Value = $"{v.Ng:N0}/{v.Total:N0}";
        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        cell.Style.Font.FontColor = XLColor.FromHtml("#64748B");
        cell.Style.Font.FontSize = 9;
    }

    // ── Sheet 2: NG by Check Item ──────────────────────────────────────────────────────
    // The viewer's second pivot: one row pair per Check Item, merged across models, over the same
    // period columns as 'NG Summary' but a DIFFERENT denominator — how often that check ran, not
    // how many pieces were inspected, so its ppm is NOT comparable with sheet 1's.
    // Columns: A=Model, B=Check Point, C=Check Item, D=Type, E=Overall, F.. = the period columns.
    private static void WriteCheckPivotSheet(
        IXLWorksheet ws,
        BmesLpaHtmlExportService.PivotData pivot,
        IReadOnlyList<BmesLpaHtmlExportService.CheckPivotRow> rows,
        string title)
    {
        var periods = new List<(string Key, string Header)>();
        foreach (var blk in pivot.Blocks)
            for (int i = 0; i < blk.Keys.Count; i++)
                periods.Add((blk.Keys[i], blk.Headers[i]));
        const int firstPeriodCol = 6;
        int lastCol = 5 + periods.Count;

        SetText(ws.Cell(1, 1), title, bold: true, size: 11);
        ws.Range(1, 1, 1, Math.Max(1, lastCol)).Merge();

        if (rows.Count == 0)
        {
            SetText(ws.Cell(2, 1), "No items marked as NG.");
            return;
        }

        Merge(ws, 2, 1, 3, 1, "Model", HeaderBg);
        Merge(ws, 2, 2, 3, 2, "Check Point", HeaderBg);
        Merge(ws, 2, 3, 3, 3, "Check Item", HeaderBg);
        Merge(ws, 2, 4, 3, 4, "Type", HeaderBg);
        Merge(ws, 2, 5, 3, 5, "Overall", HeaderBg);

        int col = firstPeriodCol;
        foreach (var blk in pivot.Blocks)
        {
            if (blk.Keys.Count == 0) continue;
            int start = col;
            for (int i = 0; i < blk.Keys.Count; i++)
            {
                var h = ws.Cell(3, col);
                SetText(h, blk.Headers[i], bold: true);
                h.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                h.Style.Fill.BackgroundColor = HeaderBg;
                col++;
            }
            Merge(ws, 2, start, 2, col - 1, blk.Label, BandBg);
        }

        // Total first, like the viewer: it is what gets read before any single check item.
        int r = 4;
        WriteCheckEntity(ws, r, "Total", "", "", periods,
                         k => (rows.Sum(x => x.NgByKey.GetValueOrDefault(k)),
                               rows.Sum(x => x.CheckedByKey.GetValueOrDefault(k))),
                         rows.Sum(x => x.OverallNg), rows.Sum(x => x.OverallChecked), isTotal: true);
        r += 2;

        foreach (var row in rows)
        {
            var captured = row;
            WriteCheckEntity(ws, r, captured.Models, captured.CheckPoints, captured.Item, periods,
                             k => (captured.NgByKey.GetValueOrDefault(k),
                                   captured.CheckedByKey.GetValueOrDefault(k)),
                             captured.OverallNg, captured.OverallChecked, isTotal: false);
            r += 2;
        }

        ws.Column(1).Width = 16;   // Model
        ws.Column(2).Width = 22;   // Check Point
        ws.Column(3).Width = 52;   // Check Item — sentences, not codes
        ws.Column(4).Width = 8;    // Type
        ws.Column(5).Width = 11;   // Overall
        for (int i = 0; i < periods.Count; i++) ws.Column(firstPeriodCol + i).Width = 10;
        DrawBlockBorders(ws, pivot, firstPeriodCol, 2, r - 1);
        ws.SheetView.FreezeRows(5);        // title + 2 header rows + Total's two rows
        ws.SheetView.FreezeColumns(5);
    }

    /// <summary>One Check Item across two rows (NG Rate / Count), its labels merged down both.
    /// The denominator is how often that check ran, so a period where it never ran stays blank
    /// while one where it ran and passed is a real 0.</summary>
    private static void WriteCheckEntity(
        IXLWorksheet ws, int r, string models, string checkPoints, string item,
        List<(string Key, string Header)> periods,
        Func<string, (int Ng, int Checked)> at, int overallNg, int overallChecked, bool isTotal)
    {
        int lastCol = 5 + periods.Count;

        // Total has no model/check point, so its label rides in column A like a model name would.
        SetText(ws.Cell(r, 1), models, top: true, wrap: true, bold: isTotal);
        SetText(ws.Cell(r, 2), checkPoints, top: true, wrap: true);
        SetText(ws.Cell(r, 3), item, top: true, wrap: true);
        for (int c = 1; c <= 3; c++)
            ws.Range(r, c, r + 1, c).Merge().Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

        SetText(ws.Cell(r, 4), "NG Rate");
        SetText(ws.Cell(r + 1, 4), "Count");

        WriteRate(ws.Cell(r, 5), Bucket(overallNg, overallChecked));
        WriteCount(ws.Cell(r + 1, 5), Bucket(overallNg, overallChecked));
        for (int i = 0; i < periods.Count; i++)
        {
            var (ng, checkedCount) = at(periods[i].Key);
            BmesLpaHtmlExportService.PivotBucket? b = Bucket(ng, checkedCount);
            WriteRate(ws.Cell(r, 6 + i), b);
            WriteCount(ws.Cell(r + 1, 6 + i), b);
        }

        if (isTotal)
        {
            ws.Range(r, 1, r + 1, lastCol).Style.Fill.BackgroundColor = TotalBg;
            ws.Range(r, 1, r + 1, lastCol).Style.Font.Bold = true;
        }
    }

    /// <summary>NG over how often the check ran; null (blank cell) when it never ran.</summary>
    private static BmesLpaHtmlExportService.PivotBucket? Bucket(int ng, int checkedCount) =>
        checkedCount == 0 ? null : new BmesLpaHtmlExportService.PivotBucket(ng, checkedCount);

    // ── Sheet 3: Defect Detail matrix (with photos) ─────────────────────────────────────────
    private static void WriteMatrixSheet(
        IXLWorksheet ws, BmesLpaHtmlExportService.NgMatrix matrix,
        IReadOnlyDictionary<string, byte[]> images, string title)
    {
        int dateCount    = matrix.Dates.Count;
        int firstDateCol = 4;                 // A=AULOC, B=Check Point, C=Check Item, D.. = dates
        int lastCol      = Math.Max(3, 3 + dateCount);

        SetText(ws.Cell(1, 1), title, bold: true, size: 11);
        ws.Range(1, 1, 1, lastCol).Merge();

        if (matrix.NgCount == 0)
        {
            SetText(ws.Cell(2, 1), "No items marked as NG.");
            return;
        }

        int headerRow = 2;
        SetText(ws.Cell(headerRow, 1), "AULOC", bold: true);
        SetText(ws.Cell(headerRow, 2), "Check Point", bold: true);
        SetText(ws.Cell(headerRow, 3), "Check Item", bold: true);
        for (int i = 0; i < dateCount; i++)
            SetText(ws.Cell(headerRow, firstDateCol + i), BmesLpaHtmlExportService.NgDateLabel(matrix.Dates[i]), bold: true);
        var hdr = ws.Range(headerRow, 1, headerRow, lastCol);
        hdr.Style.Fill.BackgroundColor = HeaderBg;
        hdr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        int r = headerRow + 1;
        foreach (var g in matrix.Groups)
        {
            SetText(ws.Cell(r, 1), $"{g.Model}  ({g.Rows.Sum(x => x.Ng):N0})", bold: true);
            Merge(ws, r, 1, r, lastCol, $"{g.Model}  ({g.Rows.Sum(x => x.Ng):N0})", GroupBg);
            r++;

            foreach (var row in g.Rows)
            {
                // AULOC repeats on every check row (not merged) so the sheet stays sortable/filterable.
                SetText(ws.Cell(r, 1), g.Model, top: true, wrap: true);
                SetText(ws.Cell(r, 2), row.Typrc, top: true, wrap: true);
                SetText(ws.Cell(r, 3), row.Lcitm, top: true, wrap: true);

                int maxPhotos = 0;
                for (int i = 0; i < dateCount; i++)
                {
                    if (!row.ItemsByDate.TryGetValue(matrix.Dates[i], out var cellItems) || cellItems.Count == 0)
                        continue;

                    int c = firstDateCol + i;
                    IXLCell cell = ws.Cell(r, c);
                    cell.Value = cellItems.Count;
                    cell.Style.Font.FontColor = QtyFg;
                    cell.Style.Font.Bold = true;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

                    int k = 0;
                    foreach (var ci in cellItems)
                        foreach (string path in BmesLpaHtmlExportService.ItemImagePaths(ci.Item))
                        {
                            if (!images.TryGetValue(path, out byte[]? bytes) || bytes is null || bytes.Length == 0)
                                continue;
                            int x = GapPx + (k % PerRow) * (ThumbPx + GapPx);
                            int y = TextPx + (k / PerRow) * (ThumbPx + GapPx);
                            using var pic = new MemoryStream(bytes);
                            ws.AddPicture(pic).MoveTo(cell, x, y).WithSize(ThumbPx, ThumbPx);
                            k++;
                        }
                    maxPhotos = Math.Max(maxPhotos, k);
                }

                if (maxPhotos > 0)
                {
                    int lines = (maxPhotos + PerRow - 1) / PerRow;
                    int heightPx = TextPx + lines * (ThumbPx + GapPx) + GapPx;
                    ws.Row(r).Height = heightPx * 0.75;   // px → pt
                }
                r++;
            }
        }

        ws.Column(1).Width = 16;   // AULOC
        ws.Column(2).Width = 26;   // Check Point
        ws.Column(3).Width = 34;   // Check Item
        double dateWidth = Math.Max(10, (PerRow * (ThumbPx + GapPx) + GapPx - 5) / 7.0);
        for (int i = 0; i < dateCount; i++) ws.Column(firstDateCol + i).Width = dateWidth;
        ws.SheetView.FreezeRows(headerRow);
        ws.SheetView.FreezeColumns(3);
    }

    // ── Sheet 4: List ──────────────────────────────────────────────────────────────────
    private static void WriteListSheet(
        IXLWorksheet ws,
        (IReadOnlyList<string> Headers, IReadOnlyList<IReadOnlyList<string>> Rows) list)
    {
        for (int c = 0; c < list.Headers.Count; c++)
            SetText(ws.Cell(1, c + 1), list.Headers[c], bold: true);
        if (list.Headers.Count > 0)
        {
            var hdr = ws.Range(1, 1, 1, list.Headers.Count);
            hdr.Style.Fill.BackgroundColor = HeaderBg;
        }

        for (int i = 0; i < list.Rows.Count; i++)
        {
            var row = list.Rows[i];
            for (int c = 0; c < row.Count; c++)
                ws.Cell(i + 2, c + 1).Value = row[c];
        }

        for (int c = 0; c < list.Headers.Count; c++) ws.Column(c + 1).Width = 14;
        ws.SheetView.FreezeRows(1);
    }

    /// <summary>Vertical rules where Date → Week → Month meet, plus one closing off the label
    /// columns. The viewer separates the blocks with empty divider columns; a sheet has no room
    /// for those, so the boundary has to be drawn or the three periods read as one long run.</summary>
    private static void DrawBlockBorders(
        IXLWorksheet ws, BmesLpaHtmlExportService.PivotData pivot,
        int firstPeriodCol, int firstRow, int lastRow)
    {
        if (lastRow < firstRow) return;

        ws.Range(firstRow, firstPeriodCol - 1, lastRow, firstPeriodCol - 1)
          .Style.Border.RightBorder = XLBorderStyleValues.Medium;

        int col = firstPeriodCol;
        foreach (var blk in pivot.Blocks)
        {
            if (blk.Keys.Count == 0) continue;
            ws.Range(firstRow, col, lastRow, col).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
            col += blk.Keys.Count;
        }

        if (col > firstPeriodCol)
            ws.Range(firstRow, col - 1, lastRow, col - 1)
              .Style.Border.RightBorder = XLBorderStyleValues.Medium;
    }

    // ── helpers ────────────────────────────────────────────────────────────────────────
    private static void Merge(IXLWorksheet ws, int r1, int c1, int r2, int c2, string text, XLColor bg)
    {
        var cell = ws.Cell(r1, c1);
        cell.Value = text;
        cell.Style.Font.Bold = true;
        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        var range = ws.Range(r1, c1, r2, c2).Merge();
        range.Style.Fill.BackgroundColor = bg;
    }

    private static void SetText(IXLCell cell, string text, bool bold = false, double? size = null,
                                bool top = false, bool wrap = false)
    {
        cell.Value = text;
        if (bold) cell.Style.Font.Bold = true;
        if (size is { } s) cell.Style.Font.FontSize = s;
        if (top) cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        if (wrap) cell.Style.Alignment.WrapText = true;
    }
}
