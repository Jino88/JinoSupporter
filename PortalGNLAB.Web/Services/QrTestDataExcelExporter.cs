using ClosedXML.Excel;
using System.Globalization;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// The QR data browser's workbook. Two sheets over the same rows: one line per QR carrying its
/// newest test, and every row behind it. The first sheet is what a yield count is taken from —
/// a QR retested three times must not count as three units — and the second is the audit trail
/// that shows why a QR ended up the way it did.
/// </summary>
public static class QrTestDataExcelExporter
{
    public const string ContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    private const string LatestSheet = "Latest per QR";
    private const string AllRowsSheet = "All rows";

    private static readonly XLColor HeaderBg = XLColor.FromHtml("#F1F5F9");
    private static readonly XLColor HeaderFg = XLColor.FromHtml("#334155");
    private static readonly XLColor FailFg = XLColor.FromHtml("#B91C1C");
    private static readonly XLColor FailBg = XLColor.FromHtml("#FEF2F2");

    public static byte[] Build(QrTestTableDef table, QrTestDataResult result)
    {
        List<int> columnOrder = QrTestColumnLayout.Build(table, result.Columns)
            .Where(index => index >= 0)
            .ToList();
        int testResultIndex = QrTestColumnLayout.FindColumnIndex(result.Columns, "TestResult");

        using var workbook = new XLWorkbook();
        WriteSheet(workbook, LatestSheet, table, result, columnOrder, testResultIndex, LatestPerProduct(result));
        WriteSheet(workbook, AllRowsSheet, table, result, columnOrder, testResultIndex, result.Rows);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// One row per ProductID, the one with the newest test time. Rows arrive newest first, but
    /// the times are compared anyway: a table whose timestamp column could not be resolved has
    /// no sort to rely on.
    /// </summary>
    private static IReadOnlyList<QrTestDataRow> LatestPerProduct(QrTestDataResult result)
    {
        int productIdIndex = QrTestColumnLayout.FindColumnIndex(result.Columns, "ProductID");
        if (productIdIndex < 0)
            return result.Rows;

        var latest = new Dictionary<string, QrTestDataRow>(StringComparer.OrdinalIgnoreCase);
        var order = new List<string>();
        foreach (QrTestDataRow row in result.Rows)
        {
            string productId = productIdIndex < row.Values.Count
                ? row.Values[productIdIndex].Trim()
                : string.Empty;
            if (productId.Length == 0)
                continue;

            if (!latest.TryGetValue(productId, out QrTestDataRow? kept))
            {
                latest.Add(productId, row);
                order.Add(productId);
            }
            else if (Nullable.Compare(row.TestTime, kept.TestTime) > 0)
            {
                latest[productId] = row;
            }
        }

        return order.Select(productId => latest[productId]).ToList();
    }

    private static void WriteSheet(
        XLWorkbook workbook,
        string name,
        QrTestTableDef table,
        QrTestDataResult result,
        IReadOnlyList<int> columnOrder,
        int testResultIndex,
        IReadOnlyList<QrTestDataRow> rows)
    {
        IXLWorksheet sheet = workbook.Worksheets.Add(name);

        for (int column = 0; column < columnOrder.Count; column++)
        {
            IXLCell cell = sheet.Cell(1, column + 1);
            cell.Value = result.Columns[columnOrder[column]].Name;
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = HeaderFg;
            cell.Style.Fill.BackgroundColor = HeaderBg;
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
        }

        for (int r = 0; r < rows.Count; r++)
        {
            QrTestDataRow row = rows[r];
            for (int column = 0; column < columnOrder.Count; column++)
            {
                int sourceIndex = columnOrder[column];
                string text = sourceIndex < row.Values.Count ? row.Values[sourceIndex] : string.Empty;
                IXLCell cell = sheet.Cell(r + 2, column + 1);

                // Numbers go in as numbers: the sheet exists to be summed and pivoted, and a
                // reading stored as text is silently skipped by both.
                if (QrTestColumnLayout.IsNumeric(result.Columns[sourceIndex].DataType)
                    && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
                {
                    cell.Value = number;
                }
                else
                {
                    cell.Value = text;
                }

                if (sourceIndex == testResultIndex && table.IsFail(text))
                {
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.FontColor = FailFg;
                    cell.Style.Fill.BackgroundColor = FailBg;
                }
            }
        }

        int lastRow = Math.Max(rows.Count + 1, 1);
        if (columnOrder.Count > 0)
        {
            sheet.Range(1, 1, lastRow, columnOrder.Count).SetAutoFilter();
            sheet.Columns(1, columnOrder.Count).AdjustToContents(1, 200);
        }
        sheet.SheetView.FreezeRows(1);
    }
}
