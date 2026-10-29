using ClosedXML.Excel;
using System.Globalization;

namespace JinoSupporter.Web.Services;

/// <summary>
/// The QR SUMMARY workbook, built the same way as the station data browser's: one sheet with the
/// newest round of each QR, and one with every round behind it. A QR retested three times is one
/// unit on the first sheet and three lines on the second.
/// </summary>
public static class QrSummaryExcelExporter
{
    public const string ContentType = QrTestDataExcelExporter.ContentType;

    private const string LatestSheet = "Latest per QR";
    private const string AllRoundsSheet = "All rounds";

    private static readonly XLColor HeaderBg = XLColor.FromHtml("#F1F5F9");
    private static readonly XLColor HeaderFg = XLColor.FromHtml("#334155");
    private static readonly XLColor FailFg = XLColor.FromHtml("#B91C1C");
    private static readonly XLColor FailBg = XLColor.FromHtml("#FEF2F2");

    private static readonly string[] Headers =
    [
        "ProductID", "Round", "LatestTime", "FinalResult",
        "AirleakResult", "AirleakValue", "AirleakTime",
        "BakoResult", "SPL1_Value", "SPL2_Value", "THD_Value", "ImpLimit_Value", "Imp2kHz_Value", "BakoTime",
        "GaussResult", "GaussValue", "GaussValue2", "GaussTime",
    ];

    /// <summary>Columns whose cell is painted when that station's verdict is NG.</summary>
    private const int FinalResultColumn = 4;
    private const int AirleakResultColumn = 5;
    private const int BakoResultColumn = 8;
    private const int GaussResultColumn = 15;

    public static byte[] Build(IReadOnlyList<QrSummaryRow> products)
    {
        using var workbook = new XLWorkbook();

        WriteSheet(workbook, LatestSheet, products.Select(product => (product.ProductId, Row: product)));
        WriteSheet(workbook, AllRoundsSheet, products.SelectMany(product =>
            new[] { (product.ProductId, Row: product) }
                .Concat(product.Older.Select(older => (product.ProductId, Row: older)))));

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private static void WriteSheet(
        XLWorkbook workbook,
        string name,
        IEnumerable<(string ProductId, QrSummaryRow Row)> lines)
    {
        IXLWorksheet sheet = workbook.Worksheets.Add(name);

        for (int column = 0; column < Headers.Length; column++)
        {
            IXLCell cell = sheet.Cell(1, column + 1);
            cell.Value = Headers[column];
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = HeaderFg;
            cell.Style.Fill.BackgroundColor = HeaderBg;
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
        }

        int row = 2;
        foreach ((string productId, QrSummaryRow line) in lines)
        {
            Write(sheet, row, 1, productId);
            sheet.Cell(row, 2).Value = line.Round;
            Write(sheet, row, 3, FormatTime(line.LatestTime));
            Write(sheet, row, 4, line.IsFinalPass ? "PASS" : "NG");

            Write(sheet, row, 5, line.Airleak?.Result);
            Write(sheet, row, 6, line.Airleak?.ItemAt(0)?.Value);
            Write(sheet, row, 7, FormatTime(line.Airleak?.TestTime));

            Write(sheet, row, 8, line.Bako?.Result);
            for (int item = 0; item < 5; item++)
                Write(sheet, row, 9 + item, line.Bako?.ItemAt(item)?.Value);
            Write(sheet, row, 14, FormatTime(line.Bako?.TestTime));

            Write(sheet, row, 15, line.Gauss?.Result);
            Write(sheet, row, 16, line.Gauss?.ItemAt(0)?.Value);
            Write(sheet, row, 17, line.Gauss?.ItemAt(1)?.Value);
            Write(sheet, row, 18, FormatTime(line.Gauss?.TestTime));

            // The verdict columns carry the colour the grid gives them, so a failure is as
            // easy to spot in the file as on screen.
            Paint(sheet, row, FinalResultColumn, !line.IsFinalPass);
            Paint(sheet, row, AirleakResultColumn, line.Airleak?.Verdict == QrVerdict.Fail);
            Paint(sheet, row, BakoResultColumn, line.Bako?.Verdict == QrVerdict.Fail);
            Paint(sheet, row, GaussResultColumn, line.Gauss?.Verdict == QrVerdict.Fail);

            row++;
        }

        int lastRow = Math.Max(row - 1, 1);
        sheet.Range(1, 1, lastRow, Headers.Length).SetAutoFilter();
        sheet.Columns(1, Headers.Length).AdjustToContents(1, 200);
        sheet.SheetView.FreezeRows(1);
    }

    /// <summary>
    /// Measurements arrive as the text the tester wrote. Anything that reads as a number goes in
    /// as one, so the sheet can be summed and pivoted without a cleanup pass first.
    /// </summary>
    private static void Write(IXLWorksheet sheet, int row, int column, string? text)
    {
        IXLCell cell = sheet.Cell(row, column);
        if (string.IsNullOrWhiteSpace(text))
        {
            cell.Value = string.Empty;
            return;
        }

        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            cell.Value = number;
        else
            cell.Value = text;
    }

    private static void Paint(IXLWorksheet sheet, int row, int column, bool failed)
    {
        if (!failed)
            return;

        IXLCell cell = sheet.Cell(row, column);
        cell.Style.Font.Bold = true;
        cell.Style.Font.FontColor = FailFg;
        cell.Style.Fill.BackgroundColor = FailBg;
    }

    private static string FormatTime(DateTimeOffset? value) =>
        value is DateTimeOffset time
            ? time.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
            : string.Empty;
}
