using ClosedXML.Excel;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// Working > Request: a PROCESSING REQUEST as an .xlsx that looks like the form on the screen (RequestPage.tsx, request.css):
/// the same 13 columns A..M at the widths of the 980px form, the same row heights, Segoe UI 13px, a thin dark line around
/// every cell, everything centred, the same bold / big fields. The sign boxes and the header fields are rows 1-10, the item
/// heading row 11, the items under it and the total after them when there is more than one. Printed on A4 portrait, one
/// page wide (the form at about 75%, as the browser prints it), the item heading repeated on every page.
/// </summary>
public static class WorkingRequestExcelExporter
{
    public const int ItemHeadRow = 11;
    public const string FontName = "Segoe UI";

    /// <summary>The columns of the form (colgroup in RequestPage.tsx) at its width of 980px, in px.</summary>
    public static readonly int[] ColumnsPx = Px([110, 223, 182, 182, 191, 87, 85, 87, 173, 169, 170, 169, 170], 980);

    private static readonly XLColor Line = XLColor.FromHtml("#1F2937");
    private static readonly XLColor Ink = XLColor.FromHtml("#111111");
    private static readonly XLColor HeadBg = XLColor.FromHtml("#F8FAFC");

    public static byte[] Export(WorkingRequestSave request)
    {
        WorkingRequestSave r = WorkingRequestService.Normalize(request);
        using var wb = new XLWorkbook();
        IXLWorksheet ws = wb.Worksheets.Add("Request");
        ws.ShowGridLines = false;
        for (int c = 0; c < ColumnsPx.Length; c++) ws.Column(c + 1).Width = WidthOf(ColumnsPx[c]);

        // sign boxes and the title: row 1 the 30px headings, row 2 the 58px boxes
        Put(ws.Range("A1:A2"), "Request\ndepart");
        string[] heads = ["Request", "Phone", "Checker", "Confirm"];
        for (int i = 0; i < heads.Length; i++) Put(ws.Range(1, i + 2, 1, i + 2), heads[i]);
        Put(ws.Range("B2"), r.Requester); Put(ws.Range("C2"), r.Phone); Put(ws.Range("D2"), r.Checker); Put(ws.Range("E2"), r.Confirm);
        Put(ws.Range("F1:I2"), "PROCESSING REQUEST", px: 20, bold: true);
        Put(ws.Range("J1:J2"), "Manufacture\nDepartment");
        Put(ws.Range("K1"), "Design"); Put(ws.Range("L1"), "Approval"); Put(ws.Range("M1"), "Confirm");

        // header fields
        Put(ws.Range("A3:B3"), "Processing method");
        IXLRange methods = Put(ws.Range("C3:I3"), MethodText(r.Method), px: 15);
        methods.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        methods.Style.Alignment.Indent = 1;
        Put(ws.Range("J3:K3"), "Department name");
        Put(ws.Range("L3:M3"), r.DepartmentName, bold: true);
        Put(ws.Range("A4:B4"), "*Equipment Name");
        Put(ws.Range("C4:I4"), r.EquipmentName);
        Put(ws.Range("J4:K5"), "*Quantity");
        Put(ws.Range("L4"), "1st time"); Put(ws.Range("M4"), "2nd time");
        Put(ws.Range("A5:B5"), "*Part Name");
        Put(ws.Range("C5:I5"), r.PartName, px: 30);
        Put(ws.Range("L5"), r.Qty1); Put(ws.Range("M5"), r.Qty2);
        Put(ws.Range("A6:B7"), "*Item Application");
        Put(ws.Range("C6:I7"), r.ItemApplication, px: 18);
        Put(ws.Range("J6:K6"), "*Request date");
        Put(ws.Range("L6:M6"), DotDate(r.RequestDate), bold: true);
        Put(ws.Range("J7:K7"), "*Purpose");
        Put(ws.Range("L7:M7"), r.Purpose, bold: true);
        Put(ws.Range("A8:B8"), "*Using place");
        Put(ws.Range("C8:I8"), r.UsingPlace, px: 18);
        Put(ws.Range("J8:K9"), "*Finish date 2nd time", bold: true).Style.Font.Underline = XLFontUnderlineValues.Single;
        Put(ws.Range("L8:M9"), DotDate(r.FinishDate2), bold: true);
        Put(ws.Range("A9:B9"), "*Finish date\n1st time", bold: true).Style.Font.Underline = XLFontUnderlineValues.Single;
        Put(ws.Range("C9:I9"), DotDate(r.FinishDate1), bold: true);
        Put(ws.Range("A10:B10"), "*Others");
        Put(ws.Range("C10:I10"), r.Others);
        Put(ws.Range("J10:K10"), "*Dep't code");
        Put(ws.Range("L10:M10"), r.DeptCode, px: 16, bold: true);

        // items: No., Model, Name (C:E), Line Spec, Stock, Order, Picture/Drawing no (I:J), PICTURE, Note (L:M)
        int h = ItemHeadRow;
        foreach ((int from, int to, string text) in new[]
        {
            (1, 1, "No."), (2, 2, "Model"), (3, 5, "Name"), (6, 6, "Line\nSpec"), (7, 7, "Stock"), (8, 8, "Order"),
            (9, 10, "Picture/Drawing no"), (11, 11, "PICTURE"), (12, 13, "Note"),
        })
            Put(ws.Range(h, from, h, to), text, bold: true).Style.Fill.BackgroundColor = HeadBg;
        int row = h;
        foreach (WorkingRequestItem i in r.Items)
        {
            row++;
            Put(ws.Range(row, 1, row, 1), i.No);
            Put(ws.Range(row, 2, row, 2), i.Model);
            Put(ws.Range(row, 3, row, 5), i.Name);
            Put(ws.Range(row, 6, row, 6), i.LineSpec);
            Put(ws.Range(row, 7, row, 7), i.Stock);
            Put(ws.Range(row, 8, row, 8), OrderOf(i));
            Put(ws.Range(row, 9, row, 10), i.DrawingNo);
            Put(ws.Range(row, 11, row, 11), i.Picture);
            // a long note wraps from the left, as on paper
            IXLRange note = Put(ws.Range(row, 12, row, 13), i.Note);
            note.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            note.Style.Alignment.Indent = 1;
            ws.Row(row).Height = Pt(ItemRowPx(i));
        }
        if (r.Items.Count > 1)
        {
            row++;
            Put(ws.Range(row, 1, row, 5), "Total", bold: true);
            Put(ws.Range(row, 6, row, 6), r.Items.Sum(i => i.LineSpec), bold: true);
            Put(ws.Range(row, 7, row, 7), r.Items.Sum(i => i.Stock), bold: true);
            Put(ws.Range(row, 8, row, 8), r.Items.Sum(OrderOf), bold: true);
            Put(ws.Range(row, 9, row, 13), "");
            ws.Range(row, 1, row, 13).Style.Fill.BackgroundColor = HeadBg;
            ws.Row(row).Height = Pt(28);
        }

        // the heights of request.css: 30px headings over the 58px sign boxes, 34px rows, the 52px part name, the 40px item heading
        ws.Row(1).Height = Pt(30);
        ws.Row(2).Height = Pt(58);
        ws.Rows(3, 10).Height = Pt(34);
        ws.Row(5).Height = Pt(52);
        ws.Row(8).Height = Pt(44);
        ws.Row(h).Height = Pt(40);

        IXLRange form = ws.Range(1, 1, row, 13);
        form.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        form.Style.Border.OutsideBorderColor = Line;
        form.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        form.Style.Border.InsideBorderColor = Line;

        // A4 portrait, the 8mm margins of the printed page, one page wide and centred
        ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
        ws.PageSetup.PageOrientation = XLPageOrientation.Portrait;
        ws.PageSetup.Margins.SetLeft(Mm8).SetRight(Mm8).SetTop(Mm8).SetBottom(Mm8).SetHeader(0).SetFooter(0);
        ws.PageSetup.CenterHorizontally = true;
        ws.PageSetup.PrintAreas.Add(1, 1, row, 13);
        ws.PageSetup.FitToPages(1, 0);
        ws.PageSetup.SetRowsToRepeatAtTop(h, h);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    private const double Mm8 = 8 / 25.4;

    /// <summary>The order of a row: the one typed, else what the line needs beyond the stock (as on the form).</summary>
    public static double OrderOf(WorkingRequestItem i) => i.Order > 0 ? i.Order : Math.Max(0, i.LineSpec - i.Stock);

    /// <summary>The three processing methods as the form's boxes, the chosen one ticked.</summary>
    public static string MethodText(string method) => string.Join("        ",
        new[] { ("new", "New"), ("repair", "Repair"), ("scrap", "Scrap (Báo phế)") }.Select(m => (m.Item1 == method ? "☑ " : "☐ ") + m.Item2));

    /// <summary>yyyy-MM-dd as yyyy.MM.dd, as the form prints it.</summary>
    public static string DotDate(string iso) => iso.Replace('-', '.');

    /// <summary>An item row: 34px, taller when its name or note wraps (about 7px a character, 17px a line).</summary>
    public static int ItemRowPx(WorkingRequestItem i)
    {
        int Lines(string text, int px) => Math.Max(1, (int)Math.Ceiling(text.Length * 7.0 / Math.Max(1, px - 12)));
        int lines = Math.Max(Lines(i.Note, ColumnsPx[11] + ColumnsPx[12]), Lines(i.Name, ColumnsPx[2] + ColumnsPx[3] + ColumnsPx[4]));
        return Math.Max(34, lines * 17 + 10);
    }

    /// <summary>The widths scaled to a total, rounded so they add up to it.</summary>
    private static int[] Px(int[] widths, int total)
    {
        int sum = widths.Sum(), done = 0, acc = 0;
        var px = new int[widths.Length];
        for (int i = 0; i < widths.Length; i++) { acc += widths[i]; int to = (int)Math.Round((double)acc * total / sum); px[i] = to - done; done = to; }
        return px;
    }

    /// <summary>A column of <paramref name="px"/> pixels as an Excel width (Calibri 11, the workbook's normal font: 7px a character + 5px).</summary>
    public static double WidthOf(int px) => Math.Round((px - 5) / 7.0, 2);

    private static double Pt(int px) => px * 0.75;

    private static IXLRange Put(IXLRange range, XLCellValue value, int px = 13, bool bold = false)
    {
        if (range.ColumnCount() > 1 || range.RowCount() > 1) range.Merge();
        range.FirstCell().Value = value;
        IXLStyle s = range.Style;
        s.Font.FontName = FontName;
        s.Font.FontSize = Pt(px);
        s.Font.FontColor = Ink;
        s.Font.Bold = bold;
        s.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        s.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        s.Alignment.WrapText = true;
        return range;
    }
}
