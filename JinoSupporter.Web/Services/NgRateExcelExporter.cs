using ClosedXML.Excel;

namespace JinoSupporter.Web.Services;

public static class NgRateExcelExporter
{
    private static readonly XLColor HeaderBg   = XLColor.FromHtml("#F1F5F9");
    private static readonly XLColor TitleBg    = XLColor.FromHtml("#E2E8F0");
    private static readonly XLColor TotalBg    = XLColor.FromHtml("#EFF6FF");
    private static readonly XLColor SubRowBg   = XLColor.FromHtml("#FAFAFA");
    private static readonly XLColor SectionFg  = XLColor.FromHtml("#334155");

    public static byte[] Export(
        List<(string Label, NgRateReportService.NgRateReport Report)> reports)
    {
        using var wb = new XLWorkbook();

        foreach (var (label, report) in reports)
        {
            var ws  = wb.Worksheets.Add(SanitizeName(label));
            int row = 1;

            row = WriteSummary(ws, report, row);        row++;
            row = WriteTop10Process(ws, report, row);   row++;
            row = WriteTop10Ng(ws, report, row);

            ws.Columns().AdjustToContents(8, 80);
        }

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    // ── Summary ───────────────────────────────────────────────────────────────────

    private static int WriteSummary(
        IXLWorksheet ws, NgRateReportService.NgRateReport r, int row)
    {
        var allCols = AllCols(r);
        int lastCol = 1 + allCols.Count;

        // Title
        var titleCell = ws.Cell(row, 1);
        titleCell.Value = "Summary — NG PPM by Process Type";
        titleCell.Style.Font.Bold      = true;
        titleCell.Style.Font.FontColor = SectionFg;
        ws.Range(row, 1, row, lastCol).Merge().Style.Fill.BackgroundColor = TitleBg;
        row++;

        // Header
        ws.Cell(row, 1).Value = "Process Type";
        WriteColHeaders(ws, row, 2, allCols);
        StyleHeaderRow(ws, row, 1, lastCol);
        row++;

        // Data
        foreach (var dataRow in r.Summary)
        {
            ws.Cell(row, 1).Value = dataRow.ProcessType;
            WritePpmCells(ws, row, 2, allCols, dataRow.Ppm);
            if (dataRow.IsTotal)
            {
                ws.Range(row, 1, row, lastCol).Style.Font.Bold = true;
                ws.Range(row, 1, row, lastCol).Style.Fill.BackgroundColor = TotalBg;
            }
            row++;
        }

        return row;
    }

    // ── Top 10 Process ────────────────────────────────────────────────────────────

    private static int WriteTop10Process(
        IXLWorksheet ws, NgRateReportService.NgRateReport r, int row)
    {
        var allCols = AllCols(r);
        int lastCol = 4 + allCols.Count;  // #, Type, Process Name, <cols>

        // Title
        var titleCell = ws.Cell(row, 1);
        titleCell.Value = "Process 10 NG — Top 10 Worst Processes by PPM";
        titleCell.Style.Font.Bold      = true;
        titleCell.Style.Font.FontColor = SectionFg;
        ws.Range(row, 1, row, lastCol).Merge().Style.Fill.BackgroundColor = TitleBg;
        row++;

        // Header
        ws.Cell(row, 1).Value = "#";
        ws.Cell(row, 2).Value = "Type";
        ws.Cell(row, 3).Value = "Process Name";
        WriteColHeaders(ws, row, 4, allCols);
        StyleHeaderRow(ws, row, 1, lastCol);
        row++;

        // Data
        foreach (var proc in r.Top10Process)
        {
            ws.Cell(row, 1).Value = proc.Rank;
            ws.Cell(row, 2).Value = proc.ProcessType;
            ws.Cell(row, 3).Value = proc.ProcessName;
            WritePpmCells(ws, row, 4, allCols, proc.Ppm);
            ws.Range(row, 1, row, lastCol).Style.Font.Bold = true;
            row++;

            foreach (var grp in proc.Groups)
            {
                ws.Cell(row, 3).Value = "  └ " + grp.GroupName;
                WritePpmCells(ws, row, 4, allCols, grp.Ppm);
                ws.Range(row, 1, row, lastCol).Style.Fill.BackgroundColor = SubRowBg;
                ws.Range(row, 1, row, lastCol).Style.Font.FontColor =
                    XLColor.FromHtml("#64748B");
                row++;
            }
        }

        return row;
    }

    // ── Top 10 NG Names ───────────────────────────────────────────────────────────

    private static int WriteTop10Ng(
        IXLWorksheet ws, NgRateReportService.NgRateReport r, int row)
    {
        var allCols = AllCols(r);
        int lastCol = 5 + allCols.Count;  // #, Type, Process Name, NG Name, <cols>

        // Title
        var titleCell = ws.Cell(row, 1);
        titleCell.Value = "Worst 10 NG Names";
        titleCell.Style.Font.Bold      = true;
        titleCell.Style.Font.FontColor = SectionFg;
        ws.Range(row, 1, row, lastCol).Merge().Style.Fill.BackgroundColor = TitleBg;
        row++;

        // Header
        ws.Cell(row, 1).Value = "#";
        ws.Cell(row, 2).Value = "Type";
        ws.Cell(row, 3).Value = "Process Name";
        ws.Cell(row, 4).Value = "NG Name";
        WriteColHeaders(ws, row, 5, allCols);
        StyleHeaderRow(ws, row, 1, lastCol);
        row++;

        // Data
        foreach (var ng in r.Top10Ng)
        {
            ws.Cell(row, 1).Value = ng.Rank;
            ws.Cell(row, 2).Value = ng.ProcessType;
            ws.Cell(row, 3).Value = ng.ProcessName;
            ws.Cell(row, 4).Value = ng.NgName;
            WritePpmCells(ws, row, 5, allCols, ng.Ppm);
            ws.Range(row, 1, row, lastCol).Style.Font.Bold = true;
            row++;

            foreach (var grp in ng.Groups)
            {
                ws.Cell(row, 4).Value = "  └ " + grp.GroupName;
                WritePpmCells(ws, row, 5, allCols, grp.Ppm);
                ws.Range(row, 1, row, lastCol).Style.Fill.BackgroundColor = SubRowBg;
                ws.Range(row, 1, row, lastCol).Style.Font.FontColor =
                    XLColor.FromHtml("#64748B");
                row++;
            }
        }

        return row;
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────────────

    private static List<NgRateReportService.PeriodColumn> AllCols(
        NgRateReportService.NgRateReport r)
        => [.. r.DateCols, .. r.WeekCols, .. r.MonthCols];

    private static void WriteColHeaders(
        IXLWorksheet ws, int row, int startCol,
        List<NgRateReportService.PeriodColumn> cols)
    {
        int col = startCol;
        foreach (var c in cols)
        {
            ws.Cell(row, col).Value = c.Header;
            col++;
        }
    }

    private static void WritePpmCells(
        IXLWorksheet ws, int row, int startCol,
        List<NgRateReportService.PeriodColumn> cols,
        Dictionary<string, double> ppm)
    {
        int col = startCol;
        foreach (var c in cols)
        {
            double v = ppm.GetValueOrDefault(c.Key);
            if (v > 0)
            {
                ws.Cell(row, col).Value = v;
                ws.Cell(row, col).Style.NumberFormat.Format = "#,##0";
            }
            col++;
        }
    }

    private static void StyleHeaderRow(
        IXLWorksheet ws, int row, int startCol, int endCol)
    {
        var range = ws.Range(row, startCol, row, endCol);
        range.Style.Font.Bold            = true;
        range.Style.Fill.BackgroundColor = HeaderBg;
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }

    private static string SanitizeName(string name)
    {
        foreach (var c in new[] { '/', '\\', '?', '*', '[', ']', ':' })
            name = name.Replace(c, '_');
        return name.Length > 31 ? name[..31] : name;
    }
}
