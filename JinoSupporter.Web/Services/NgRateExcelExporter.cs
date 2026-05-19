using ClosedXML.Excel;

namespace JinoSupporter.Web.Services;

public static class NgRateExcelExporter
{
    private static readonly XLColor HeaderBg  = XLColor.FromHtml("#F1F5F9");
    private static readonly XLColor TitleBg   = XLColor.FromHtml("#E2E8F0");
    private static readonly XLColor TotalBg   = XLColor.FromHtml("#EFF6FF");
    private static readonly XLColor SubRowBg  = XLColor.FromHtml("#FAFAFA");
    private static readonly XLColor SectionFg = XLColor.FromHtml("#334155");

    public static byte[] Export(
        List<(string Label, NgRateReportService.NgRateReport Report)> reports,
        List<(string Label, NgRateReportService.LineShiftNgReport Report)>? detailReports = null,
        bool includeDateColumns = true)
    {
        using var wb = new XLWorkbook();

        foreach (var (label, report) in reports)
        {
            var ws = AddWorksheet(wb, label);
            int row = 1;

            row = WriteSummary(ws, report, row, includeDateColumns);      row++;
            row = WriteTop10Process(ws, report, row, includeDateColumns); row++;
            row = WriteTop10Ng(ws, report, row, includeDateColumns);      row++;
            row = WriteReason(ws, report, row, includeDateColumns);

            ws.Columns().AdjustToContents(8, 80);
        }

        if (detailReports is not null)
        {
            foreach (var (label, report) in detailReports)
            {
                var ws = AddWorksheet(wb, label);
                WriteLineShiftDetails(ws, report, includeDateColumns);
                ws.Columns().AdjustToContents(8, 80);
            }
        }

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    private static int WriteSummary(
        IXLWorksheet ws,
        NgRateReportService.NgRateReport r,
        int row,
        bool includeDateColumns)
    {
        var allCols = AllCols(r, includeDateColumns);
        int lastCol = 1 + allCols.Count;

        WriteSectionTitle(ws, row, 1, lastCol, "Summary - NG PPM by Process Type");
        row++;

        ws.Cell(row, 1).Value = "Process Type";
        WriteColHeaders(ws, row, 2, allCols);
        StyleHeaderRow(ws, row, 1, lastCol);
        row++;

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

    private static int WriteTop10Process(
        IXLWorksheet ws,
        NgRateReportService.NgRateReport r,
        int row,
        bool includeDateColumns)
    {
        var allCols = AllCols(r, includeDateColumns);
        int lastCol = 3 + allCols.Count;

        WriteSectionTitle(ws, row, 1, lastCol, "Process 10 NG - Top 10 Worst Processes by PPM");
        row++;

        ws.Cell(row, 1).Value = "#";
        ws.Cell(row, 2).Value = "Type";
        ws.Cell(row, 3).Value = "Process Name";
        WriteColHeaders(ws, row, 4, allCols);
        StyleHeaderRow(ws, row, 1, lastCol);
        row++;

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
                ws.Cell(row, 3).Value = "  - " + grp.GroupName;
                WritePpmCells(ws, row, 4, allCols, grp.Ppm);
                StyleSubRow(ws, row, 1, lastCol);
                row++;
            }
        }

        return row;
    }

    private static int WriteTop10Ng(
        IXLWorksheet ws,
        NgRateReportService.NgRateReport r,
        int row,
        bool includeDateColumns)
    {
        var allCols = AllCols(r, includeDateColumns);
        int lastCol = 4 + allCols.Count;

        WriteSectionTitle(ws, row, 1, lastCol, "Worst 10 NG Names");
        row++;

        ws.Cell(row, 1).Value = "#";
        ws.Cell(row, 2).Value = "Type";
        ws.Cell(row, 3).Value = "Process Name";
        ws.Cell(row, 4).Value = "NG Name";
        WriteColHeaders(ws, row, 5, allCols);
        StyleHeaderRow(ws, row, 1, lastCol);
        row++;

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
                ws.Cell(row, 4).Value = "  - " + grp.GroupName;
                WritePpmCells(ws, row, 5, allCols, grp.Ppm);
                StyleSubRow(ws, row, 1, lastCol);
                row++;
            }
        }

        return row;
    }

    private static int WriteReason(
        IXLWorksheet ws,
        NgRateReportService.NgRateReport r,
        int row,
        bool includeDateColumns)
    {
        if (r.ReasonRows.Count == 0) return row;

        var allCols = AllCols(r, includeDateColumns);
        int lastCol = 5 + allCols.Count;

        WriteSectionTitle(ws, row, 1, lastCol, "Reason");
        row++;

        ws.Cell(row, 1).Value = "Reason";
        ws.Cell(row, 2).Value = "#";
        ws.Cell(row, 3).Value = "Type";
        ws.Cell(row, 4).Value = "Process Name";
        ws.Cell(row, 5).Value = "NG Name";
        WriteColHeaders(ws, row, 6, allCols);
        StyleHeaderRow(ws, row, 1, lastCol);
        row++;

        foreach (var reason in r.ReasonRows)
        {
            ws.Cell(row, 1).Value = reason.Reason;
            ws.Cell(row, 2).Value = reason.IsTotal ? "Total" : reason.No;
            ws.Cell(row, 3).Value = reason.ProcessType;
            ws.Cell(row, 4).Value = reason.ProcessName;
            ws.Cell(row, 5).Value = reason.NgName;
            WritePpmCells(ws, row, 6, allCols, reason.Ppm);
            ws.Range(row, 1, row, lastCol).Style.Font.Bold = true;
            if (reason.IsTotal)
                ws.Range(row, 1, row, lastCol).Style.Fill.BackgroundColor = TotalBg;
            row++;

            foreach (var grp in reason.Groups)
            {
                ws.Cell(row, 5).Value = "  - " + grp.GroupName;
                WritePpmCells(ws, row, 6, allCols, grp.Ppm);
                StyleSubRow(ws, row, 1, lastCol);
                row++;
            }
        }

        return row;
    }

    private static void WriteLineShiftDetails(
        IXLWorksheet ws,
        NgRateReportService.LineShiftNgReport report,
        bool includeDateColumns)
    {
        var allCols = AllCols(report, includeDateColumns);
        int row = 1;
        int lastCol = 4 + allCols.Count;

        WriteSectionTitle(ws, row, 1, lastCol, "Detail - Top NG by Model / Group");
        row++;

        ws.Cell(row, 1).Value = "Model / Group";
        ws.Cell(row, 2).Value = "Type";
        ws.Cell(row, 3).Value = "Process Name";
        ws.Cell(row, 4).Value = "NG Name";
        WriteColHeaders(ws, row, 5, allCols);
        StyleHeaderRow(ws, row, 1, lastCol);
        row++;

        foreach (var detail in report.Details
                     .OrderBy(d => d.LineShift, StringComparer.Ordinal)
                     .ThenBy(d => d.ProcessType, StringComparer.Ordinal)
                     .ThenBy(d => d.ProcessName, StringComparer.Ordinal)
                     .ThenBy(d => d.NgName, StringComparer.Ordinal))
        {
            ws.Cell(row, 1).Value = detail.LineShift;
            ws.Cell(row, 2).Value = detail.ProcessType;
            ws.Cell(row, 3).Value = detail.ProcessName;
            ws.Cell(row, 4).Value = detail.NgName;
            WriteDetailPpmCells(ws, row, 5, allCols, detail);
            row++;
        }
    }

    private static IXLWorksheet AddWorksheet(XLWorkbook wb, string label)
    {
        string baseName = SanitizeName(label);
        if (string.IsNullOrWhiteSpace(baseName)) baseName = "Sheet";

        string name = baseName;
        int n = 2;
        while (wb.Worksheets.Any(ws => string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase)))
        {
            string suffix = "_" + n++;
            int maxBase = 31 - suffix.Length;
            name = (baseName.Length > maxBase ? baseName[..maxBase] : baseName) + suffix;
        }

        return wb.Worksheets.Add(name);
    }

    private static List<NgRateReportService.PeriodColumn> AllCols(
        NgRateReportService.NgRateReport r,
        bool includeDateColumns)
        => includeDateColumns
            ? [.. r.DateCols, .. r.WeekCols, .. r.MonthCols]
            : [.. r.WeekCols, .. r.MonthCols];

    private static List<NgRateReportService.PeriodColumn> AllCols(
        NgRateReportService.LineShiftNgReport r,
        bool includeDateColumns)
        => includeDateColumns
            ? [.. r.DateCols, .. r.WeekCols, .. r.MonthCols]
            : [.. r.WeekCols, .. r.MonthCols];

    private static void WriteSectionTitle(
        IXLWorksheet ws,
        int row,
        int startCol,
        int endCol,
        string title)
    {
        var titleCell = ws.Cell(row, startCol);
        titleCell.Value = title;
        titleCell.Style.Font.Bold = true;
        titleCell.Style.Font.FontColor = SectionFg;
        ws.Range(row, startCol, row, endCol).Merge().Style.Fill.BackgroundColor = TitleBg;
    }

    private static void WriteColHeaders(
        IXLWorksheet ws,
        int row,
        int startCol,
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
        IXLWorksheet ws,
        int row,
        int startCol,
        List<NgRateReportService.PeriodColumn> cols,
        Dictionary<string, double> ppm)
    {
        int col = startCol;
        foreach (var c in cols)
        {
            double v = ppm.GetValueOrDefault(c.Key);
            WritePpmCell(ws, row, col, v);
            col++;
        }
    }

    private static void WriteDetailPpmCells(
        IXLWorksheet ws,
        int row,
        int startCol,
        List<NgRateReportService.PeriodColumn> cols,
        NgRateReportService.LineShiftNgDetail detail)
    {
        int col = startCol;
        foreach (var c in cols)
        {
            double v = detail.DatePpm.GetValueOrDefault(c.Key);
            if (v == 0) v = detail.WeekPpm.GetValueOrDefault(c.Key);
            if (v == 0) v = detail.MonthPpm.GetValueOrDefault(c.Key);
            WritePpmCell(ws, row, col, v);
            col++;
        }
    }

    private static void WritePpmCell(IXLWorksheet ws, int row, int col, double value)
    {
        if (value <= 0) return;
        ws.Cell(row, col).Value = value;
        ws.Cell(row, col).Style.NumberFormat.Format = "#,##0";
    }

    private static void StyleHeaderRow(IXLWorksheet ws, int row, int startCol, int endCol)
    {
        var range = ws.Range(row, startCol, row, endCol);
        range.Style.Font.Bold = true;
        range.Style.Fill.BackgroundColor = HeaderBg;
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }

    private static void StyleSubRow(IXLWorksheet ws, int row, int startCol, int endCol)
    {
        var range = ws.Range(row, startCol, row, endCol);
        range.Style.Fill.BackgroundColor = SubRowBg;
        range.Style.Font.FontColor = XLColor.FromHtml("#64748B");
    }

    private static string SanitizeName(string name)
    {
        foreach (char c in new[] { '/', '\\', '?', '*', '[', ']', ':' })
            name = name.Replace(c, '_');
        return name.Length > 31 ? name[..31] : name;
    }
}
