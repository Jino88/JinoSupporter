using ClosedXML.Excel;
using System.Globalization;

namespace BmesNgRateStandalone.Services;

public static class WorkerStatusExcelExporter
{
    private static readonly XLColor HeaderBg = XLColor.FromHtml("#F1F5F9");
    private static readonly XLColor HeaderFg = XLColor.FromHtml("#334155");
    private static readonly XLColor LateBg   = XLColor.FromHtml("#FEF2F2");
    private static readonly XLColor LateFg   = XLColor.FromHtml("#B91C1C");
    private static readonly XLColor GroupAlt = XLColor.FromHtml("#FAFAFA");

    public static byte[] Export(
        IReadOnlyList<WorkerStatusService.WorkerRecord> records,
        DateTime startDate,
        DateTime endDate)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Worker Status");

        const int lastCol = 11;

        // Title
        var title = ws.Cell(1, 1);
        title.Value = $"Worker Status  {startDate:yyyy-MM-dd} ~ {endDate:yyyy-MM-dd}";
        title.Style.Font.Bold     = true;
        title.Style.Font.FontSize = 13;
        ws.Range(1, 1, 1, lastCol).Merge();

        // Header
        const int headerRow = 3;
        string[] headers =
        {
            "#", "Dept Code", "Dept No", "Emp No", "Name", "Date", "Day",
            "Status", "Day Type", "Check In", "Check Out"
        };
        for (int i = 0; i < headers.Length; i++)
        {
            var c = ws.Cell(headerRow, i + 1);
            c.Value = headers[i];
            c.Style.Font.Bold              = true;
            c.Style.Font.FontColor         = HeaderFg;
            c.Style.Fill.BackgroundColor   = HeaderBg;
            c.Style.Alignment.Horizontal   = XLAlignmentHorizontalValues.Left;
            c.Style.Border.BottomBorder    = XLBorderStyleValues.Medium;
        }

        // Data — group by employee, then by date
        int row     = headerRow + 1;
        int groupNo = 0;

        var groups = records
            .GroupBy(r => new { r.DepartmentCode, r.DepartmentNo, r.EmpNo, r.Name })
            .OrderBy(g => g.Key.DepartmentCode)
            .ThenBy(g => g.Key.DepartmentNo)
            .ThenBy(g => g.Key.Name)
            .ThenBy(g => g.Key.EmpNo);

        foreach (var g in groups)
        {
            groupNo++;
            bool altShade = groupNo % 2 == 0;
            var ordered   = g.OrderBy(r => r.Date).ToList();
            int firstRow  = row;

            foreach (var rec in ordered)
            {
                bool late = IsLate(rec);

                ws.Cell(row, 1).Value = groupNo;
                ws.Cell(row, 2).Value = rec.DepartmentCode;
                ws.Cell(row, 3).Value = rec.DepartmentNo;
                ws.Cell(row, 4).Value = rec.EmpNo;
                ws.Cell(row, 5).Value = rec.Name;
                ws.Cell(row, 6).Value = rec.Date;
                ws.Cell(row, 6).Style.DateFormat.Format = "yyyy-MM-dd";
                ws.Cell(row, 7).Value = rec.Date.ToString("ddd", CultureInfo.InvariantCulture);
                ws.Cell(row, 8).Value = rec.WorkStatus;
                ws.Cell(row, 9).Value = rec.DayType;
                ws.Cell(row, 10).Value = TimeWithSchedule(rec.CheckIn, rec.SchedStart);
                ws.Cell(row, 11).Value = TimeWithSchedule(rec.CheckOut, rec.SchedEnd);

                if (late)
                {
                    var inCell = ws.Cell(row, 10);
                    inCell.Style.Fill.BackgroundColor = LateBg;
                    inCell.Style.Font.FontColor       = LateFg;
                    inCell.Style.Font.Bold            = true;
                }

                if (altShade)
                {
                    ws.Range(row, 1, row, lastCol).Style.Fill.BackgroundColor = GroupAlt;
                    if (late)
                    {
                        var inCell = ws.Cell(row, 10);
                        inCell.Style.Fill.BackgroundColor = LateBg;
                    }
                }

                row++;
            }

            // Merge group cells (#, Dept, EmpNo, Name) when more than one row
            if (ordered.Count > 1)
            {
                ws.Range(firstRow, 1, row - 1, 1).Merge()
                    .Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                ws.Range(firstRow, 2, row - 1, 2).Merge()
                    .Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                ws.Range(firstRow, 3, row - 1, 3).Merge()
                    .Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                ws.Range(firstRow, 4, row - 1, 4).Merge()
                    .Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                ws.Range(firstRow, 5, row - 1, 5).Merge()
                    .Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            }

            // Group separator border
            ws.Range(row - 1, 1, row - 1, lastCol).Style.Border.BottomBorder =
                XLBorderStyleValues.Thin;
        }

        // Freeze header, autosize, finish
        ws.SheetView.FreezeRows(headerRow);
        ws.Columns().AdjustToContents(8, 60);
        ws.Column(5).Width = Math.Max(ws.Column(5).Width, 16);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    private static bool IsLate(WorkerStatusService.WorkerRecord rec)
    {
        if (string.IsNullOrEmpty(rec.CheckIn) || string.IsNullOrEmpty(rec.SchedStart))
            return false;
        return string.Compare(rec.CheckIn, rec.SchedStart, StringComparison.Ordinal) > 0;
    }

    private static string TimeWithSchedule(string actual, string scheduled)
    {
        if (string.IsNullOrEmpty(scheduled)) return actual;
        return string.IsNullOrEmpty(actual)
            ? $"({scheduled})"
            : $"{actual} ({scheduled})";
    }
}
