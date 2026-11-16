using ClosedXML.Excel;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// Working > Request: a PROCESSING REQUEST as an .xlsx laid out like the paper form (RequestPage.tsx), 13 columns A..M:
/// the sign boxes and the header fields on rows 1-10, the item heading on row 11, the items under it and the total after
/// them when there is more than one. A4, one page wide.
/// </summary>
public static class WorkingRequestExcelExporter
{
    public const int ItemHeadRow = 11;

    /// <summary>The column widths of the form on the screen (px), as Excel widths.</summary>
    private static readonly int[] WidthsPx = [110, 223, 182, 182, 191, 87, 85, 87, 173, 169, 170, 169, 170];
    private static readonly XLColor LabelBg = XLColor.FromHtml("#F1F5F9");

    public static byte[] Export(WorkingRequestSave request)
    {
        WorkingRequestSave r = WorkingRequestService.Normalize(request);
        using var wb = new XLWorkbook();
        IXLWorksheet ws = wb.Worksheets.Add("Request");
        ws.Style.Font.FontName = "Arial";
        ws.Style.Font.FontSize = 10;
        for (int c = 0; c < WidthsPx.Length; c++) ws.Column(c + 1).Width = Math.Round(WidthsPx[c] / 8.0, 1);

        // sign boxes and the title
        Label(ws.Range("A1:A2"), "Request\ndepart");
        string[] heads = ["Request", "Phone", "Checker", "Confirm"];
        for (int i = 0; i < heads.Length; i++) Label(ws.Range(1, i + 2, 1, i + 2), heads[i]);
        Value(ws.Range("B2"), r.Requester); Value(ws.Range("C2"), r.Phone); Value(ws.Range("D2"), r.Checker); Value(ws.Range("E2"), r.Confirm);
        IXLRange title = Value(ws.Range("F1:I2"), "PROCESSING REQUEST");
        title.Style.Font.Bold = true; title.Style.Font.FontSize = 18;
        title.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        Label(ws.Range("J1:J2"), "Manufacture\nDepartment");
        Label(ws.Range("K1"), "Design"); Label(ws.Range("L1"), "Approval"); Label(ws.Range("M1"), "Confirm");
        Value(ws.Range("K2:K2"), ""); Value(ws.Range("L2:L2"), ""); Value(ws.Range("M2:M2"), "");

        // header fields
        Label(ws.Range("A3:B3"), "Processing method");
        Value(ws.Range("C3:I3"), MethodText(r.Method));
        Label(ws.Range("J3:K3"), "Department name");
        Value(ws.Range("L3:M3"), r.DepartmentName).Style.Font.Bold = true;
        Label(ws.Range("A4:B4"), "*Equipment Name");
        Value(ws.Range("C4:I4"), r.EquipmentName);
        Label(ws.Range("J4:K5"), "*Quantity");
        Label(ws.Range("L4"), "1st time"); Label(ws.Range("M4"), "2nd time");
        Label(ws.Range("A5:B5"), "*Part Name");
        Value(ws.Range("C5:I5"), r.PartName).Style.Font.Bold = true;
        Value(ws.Range("L5"), r.Qty1); Value(ws.Range("M5"), r.Qty2);
        Label(ws.Range("A6:B7"), "*Item Application");
        Value(ws.Range("C6:I7"), r.ItemApplication).Style.Font.FontSize = 14;
        Label(ws.Range("J6:K6"), "*Request date");
        Value(ws.Range("L6:M6"), DotDate(r.RequestDate)).Style.Font.Bold = true;
        Label(ws.Range("J7:K7"), "*Purpose");
        Value(ws.Range("L7:M7"), r.Purpose).Style.Font.Bold = true;
        Label(ws.Range("A8:B8"), "*Using place");
        Value(ws.Range("C8:I8"), r.UsingPlace).Style.Font.FontSize = 14;
        Label(ws.Range("J8:K9"), "*Finish date 2nd time");
        Value(ws.Range("L8:M9"), DotDate(r.FinishDate2)).Style.Font.Bold = true;
        Label(ws.Range("A9:B9"), "*Finish date 1st time");
        Value(ws.Range("C9:I9"), DotDate(r.FinishDate1)).Style.Font.Bold = true;
        Label(ws.Range("A10:B10"), "*Others");
        Value(ws.Range("C10:I10"), r.Others);
        Label(ws.Range("J10:K10"), "*Dep't code");
        Value(ws.Range("L10:M10"), r.DeptCode);

        // items: No., Model, Name (C:E), Line Spec, Stock, Order, Picture/Drawing no (I:J), PICTURE, Note (L:M)
        int h = ItemHeadRow;
        Label(ws.Range(h, 1, h, 1), "No."); Label(ws.Range(h, 2, h, 2), "Model"); Label(ws.Range(h, 3, h, 5), "Name");
        Label(ws.Range(h, 6, h, 6), "Line\nSpec"); Label(ws.Range(h, 7, h, 7), "Stock"); Label(ws.Range(h, 8, h, 8), "Order");
        Label(ws.Range(h, 9, h, 10), "Picture/Drawing no"); Label(ws.Range(h, 11, h, 11), "PICTURE"); Label(ws.Range(h, 12, h, 13), "Note");
        ws.Range(h, 1, h, 13).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        int row = h;
        foreach (WorkingRequestItem i in r.Items)
        {
            row++;
            Number(ws.Range(row, 1, row, 1), i.No);
            Value(ws.Range(row, 2, row, 2), i.Model);
            Value(ws.Range(row, 3, row, 5), i.Name);
            Number(ws.Range(row, 6, row, 6), i.LineSpec);
            Number(ws.Range(row, 7, row, 7), i.Stock);
            Number(ws.Range(row, 8, row, 8), OrderOf(i));
            Value(ws.Range(row, 9, row, 10), i.DrawingNo);
            Value(ws.Range(row, 11, row, 11), i.Picture);
            Value(ws.Range(row, 12, row, 13), i.Note).Style.Alignment.WrapText = true;
        }
        if (r.Items.Count > 1)
        {
            row++;
            Label(ws.Range(row, 1, row, 5), "Total");
            Number(ws.Range(row, 6, row, 6), r.Items.Sum(i => i.LineSpec)).Style.Font.Bold = true;
            Number(ws.Range(row, 7, row, 7), r.Items.Sum(i => i.Stock)).Style.Font.Bold = true;
            Number(ws.Range(row, 8, row, 8), r.Items.Sum(OrderOf)).Style.Font.Bold = true;
            Value(ws.Range(row, 9, row, 13), "");
        }

        ws.Row(2).Height = 42;
        ws.Rows(3, 10).Height = 22;
        ws.Row(h).Height = 30;
        IXLRange form = ws.Range(1, 1, row, 13);
        form.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        form.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

        ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
        ws.PageSetup.FitToPages(1, 0);
        ws.PageSetup.Margins.SetLeft(0.3).SetRight(0.3).SetTop(0.3).SetBottom(0.3);
        ws.PageSetup.SetRowsToRepeatAtTop(h, h);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    /// <summary>The order of a row: the one typed, else what the line needs beyond the stock (as on the form).</summary>
    public static double OrderOf(WorkingRequestItem i) => i.Order > 0 ? i.Order : Math.Max(0, i.LineSpec - i.Stock);

    /// <summary>The three processing methods with the chosen one ticked.</summary>
    public static string MethodText(string method) => string.Join("     ",
        new[] { ("new", "New"), ("repair", "Repair"), ("scrap", "Scrap (Báo phế)") }.Select(m => (m.Item1 == method ? "☑ " : "☐ ") + m.Item2));

    /// <summary>yyyy-MM-dd as yyyy.MM.dd, as the form prints it.</summary>
    public static string DotDate(string iso) => iso.Replace('-', '.');

    private static IXLRange Label(IXLRange range, string text)
    {
        IXLRange r = Value(range, text);
        r.Style.Font.Bold = true;
        r.Style.Fill.BackgroundColor = LabelBg;
        r.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        r.Style.Alignment.WrapText = true;
        return r;
    }

    private static IXLRange Value(IXLRange range, string text)
    {
        if (range.ColumnCount() > 1 || range.RowCount() > 1) range.Merge();
        range.FirstCell().Value = text;
        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        return range;
    }

    private static IXLRange Number(IXLRange range, double value)
    {
        range.FirstCell().Value = value;
        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        return range;
    }
}
