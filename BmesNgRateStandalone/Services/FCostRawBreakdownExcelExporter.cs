using ClosedXML.Excel;
using System.Globalization;

namespace BmesNgRateStandalone.Services;

public static class FCostRawBreakdownExcelExporter
{
    private const string CurrencyVnd = "VND";
    private const string CurrencyUsd = "USD";
    private const string CurrencyKrw = "KRW";

    private static readonly XLColor HeaderBg = XLColor.FromHtml("#F1F5F9");
    private static readonly XLColor TitleBg = XLColor.FromHtml("#E2E8F0");
    private static readonly XLColor TotalBg = XLColor.FromHtml("#FFF7CC");
    private static readonly XLColor GroupBg = XLColor.FromHtml("#EFF6FF");
    private static readonly XLColor SubtotalBg = XLColor.FromHtml("#F8FAFC");

    public static byte[] Export(
        BmesFcostRawBreakdownResult breakdown,
        string displayCurrency,
        DateTime startDate,
        DateTime endDate)
    {
        displayCurrency = NormalizeCurrencyMode(displayCurrency);
        var periods = breakdown.Periods.ToList();
        int lastCol = 7 + (periods.Count * 3);

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Raw Material Breakdown");

        int row = 1;
        ws.Range(row, 1, row, lastCol).Merge();
        ws.Cell(row, 1).Value = "FCOST Raw Material Breakdown";
        ws.Range(row, 1, row, lastCol).Style.Font.Bold = true;
        ws.Range(row, 1, row, lastCol).Style.Font.FontSize = 13;
        ws.Range(row, 1, row, lastCol).Style.Fill.BackgroundColor = TitleBg;
        row++;

        ws.Cell(row, 1).Value = "Date Range";
        ws.Cell(row, 2).Value = $"{startDate:yyyy-MM-dd} - {endDate:yyyy-MM-dd}";
        ws.Cell(row, 4).Value = "Currency";
        ws.Cell(row, 5).Value = displayCurrency;
        row++;
        ws.Cell(row, 1).Value = "Source";
        ws.Cell(row, 2).Value = breakdown.SourceTable;
        ws.Cell(row, 4).Value = "Name Source";
        ws.Cell(row, 5).Value = breakdown.NameSource;
        row += 2;

        int headerRow = row;
        string[] headers = ["Level", "Model Group", "Model", "Raw Material", "Material Code", "Unit Price (DB)", "Source Rows"];
        for (int i = 0; i < headers.Length; i++)
            ws.Cell(headerRow, i + 1).Value = headers[i];

        int col = 8;
        foreach (var period in periods)
        {
            ws.Cell(headerRow, col).Value = BuildPeriodHeader(period, breakdown.ExchangeRates, displayCurrency) + "\nF-COST";
            ws.Cell(headerRow, col).Style.Alignment.WrapText = true;
            ws.Cell(headerRow, col + 1).Value = period.Header + "\nEq Qty";
            ws.Cell(headerRow, col + 1).Style.Alignment.WrapText = true;
            ws.Cell(headerRow, col + 2).Value = period.Header + "\nUnit Price";
            ws.Cell(headerRow, col + 2).Style.Alignment.WrapText = true;
            col += 3;
        }
        StyleHeader(ws, headerRow, lastCol);
        row++;

        WriteDataRow(
            ws,
            row++,
            "Total",
            "Total",
            string.Empty,
            "Total",
            string.Empty,
            string.Empty,
            breakdown.Rows.Sum(r => r.SourceRows),
            SumPeriods(breakdown.Rows, periods),
            new Dictionary<string, decimal>(StringComparer.Ordinal),
            new Dictionary<string, BmesFcostRawMaterialPeriodPrice>(StringComparer.Ordinal),
            periods,
            breakdown.ExchangeRates,
            displayCurrency,
            TotalBg,
            bold: true,
            depth: 0);

        foreach (var groupRows in breakdown.Rows
            .GroupBy(r => DisplayText(r.GroupName, "Unassigned"))
            .OrderBy(g => g.Key, StringComparer.Ordinal))
        {
            var groupList = groupRows.ToList();
            WriteDataRow(
                ws,
                row++,
                "Model Group Total",
                groupRows.Key,
                "Total",
                "Model Group Total",
                string.Empty,
                string.Empty,
                groupList.Sum(r => r.SourceRows),
                SumPeriods(groupList, periods),
                new Dictionary<string, decimal>(StringComparer.Ordinal),
                new Dictionary<string, BmesFcostRawMaterialPeriodPrice>(StringComparer.Ordinal),
                periods,
                breakdown.ExchangeRates,
                displayCurrency,
                GroupBg,
                bold: true,
                depth: 1);

            foreach (var modelRows in groupList
                .GroupBy(r => DisplayText(r.ModelName, "Unassigned"))
                .OrderBy(g => g.Key, StringComparer.Ordinal))
            {
                var modelList = modelRows.ToList();
                WriteDataRow(
                    ws,
                    row++,
                    "Model Sub Total",
                    groupRows.Key,
                    modelRows.Key,
                    "Sub Total",
                    string.Empty,
                    string.Empty,
                    modelList.Sum(r => r.SourceRows),
                    SumPeriods(modelList, periods),
                    new Dictionary<string, decimal>(StringComparer.Ordinal),
                    new Dictionary<string, BmesFcostRawMaterialPeriodPrice>(StringComparer.Ordinal),
                    periods,
                    breakdown.ExchangeRates,
                    displayCurrency,
                    SubtotalBg,
                    bold: true,
                    depth: 2);

                foreach (var item in modelList
                    .OrderByDescending(r => r.TotalFCostVnd)
                    .ThenBy(r => DisplayName(r.MaterialName, r.MaterialCode), StringComparer.Ordinal)
                    .ThenBy(r => r.MaterialCode, StringComparer.Ordinal))
                {
                    WriteDataRow(
                        ws,
                        row++,
                        "Raw Material",
                        groupRows.Key,
                        modelRows.Key,
                        DisplayName(item.MaterialName, item.MaterialCode),
                        item.MaterialCode,
                        FormatFirstUnitPrice(item.PriceByPeriod, periods),
                        item.SourceRows,
                        item.FCostByPeriod,
                        item.EquivalentQtyByPeriod,
                        item.PriceByPeriod,
                        periods,
                        breakdown.ExchangeRates,
                        displayCurrency,
                        XLColor.White,
                        bold: false,
                        depth: 3);
                }
            }
        }

        if (row > headerRow + 1)
            ws.Range(headerRow, 1, row - 1, lastCol).SetAutoFilter();

        ws.SheetView.FreezeRows(headerRow);
        ws.Columns().AdjustToContents(8, 70);
        ws.Column(4).Width = Math.Max(ws.Column(4).Width, 28);
        ws.Column(6).Width = Math.Max(ws.Column(6).Width, 12);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    private static void WriteDataRow(
        IXLWorksheet ws,
        int row,
        string level,
        string groupName,
        string modelName,
        string materialName,
        string materialCode,
        string unitPriceText,
        long sourceRows,
        IReadOnlyDictionary<string, decimal> valuesByPeriod,
        IReadOnlyDictionary<string, decimal> equivalentQtyByPeriod,
        IReadOnlyDictionary<string, BmesFcostRawMaterialPeriodPrice> priceByPeriod,
        IReadOnlyList<BmesFcostRawBreakdownPeriod> periods,
        IReadOnlyList<BmesFcostExchangeRate> rates,
        string displayCurrency,
        XLColor fill,
        bool bold,
        int depth)
    {
        ws.Cell(row, 1).Value = level;
        ws.Cell(row, 2).Value = groupName;
        ws.Cell(row, 3).Value = modelName;
        ws.Cell(row, 4).Value = materialName;
        ws.Cell(row, 4).Style.Alignment.Indent = Math.Min(depth, 10);
        ws.Cell(row, 5).Value = materialCode;
        ws.Cell(row, 6).Value = unitPriceText;
        ws.Cell(row, 7).Value = sourceRows;

        int col = 8;
        foreach (var period in periods)
        {
            decimal vndValue = valuesByPeriod.GetValueOrDefault(period.Key);
            decimal? value = ConvertDisplayCurrency(vndValue, period.Key, rates, displayCurrency);
            if (value is not null)
            {
                ws.Cell(row, col).Value = value.Value;
                ws.Cell(row, col).Style.NumberFormat.Format =
                    string.Equals(displayCurrency, CurrencyUsd, StringComparison.Ordinal) ? "#,##0.00" : "#,##0";
            }

            decimal eqQty = equivalentQtyByPeriod.GetValueOrDefault(period.Key);
            if (eqQty > 0)
            {
                ws.Cell(row, col + 1).Value = eqQty;
                ws.Cell(row, col + 1).Style.NumberFormat.Format = "#,##0.##";
            }

            if (priceByPeriod.TryGetValue(period.Key, out var price))
                ws.Cell(row, col + 2).Value = FormatUnitPrice(price);

            col += 3;
        }

        int lastCol = 7 + (periods.Count * 3);
        var range = ws.Range(row, 1, row, lastCol);
        range.Style.Fill.BackgroundColor = fill;
        range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        range.Style.Border.BottomBorderColor = XLColor.FromHtml("#D7DEE8");
        if (bold)
            range.Style.Font.Bold = true;
    }

    private static void StyleHeader(IXLWorksheet ws, int row, int lastCol)
    {
        var range = ws.Range(row, 1, row, lastCol);
        range.Style.Font.Bold = true;
        range.Style.Fill.BackgroundColor = HeaderBg;
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        range.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
    }

    private static Dictionary<string, decimal> SumPeriods(
        IEnumerable<BmesFcostRawMaterialBreakdownRow> rows,
        IReadOnlyList<BmesFcostRawBreakdownPeriod> periods)
    {
        var result = new Dictionary<string, decimal>(StringComparer.Ordinal);
        foreach (var period in periods)
            result[period.Key] = rows.Sum(r => r.FCostByPeriod.GetValueOrDefault(period.Key));
        return result;
    }

    private static string BuildPeriodHeader(
        BmesFcostRawBreakdownPeriod period,
        IReadOnlyList<BmesFcostExchangeRate> rates,
        string displayCurrency)
    {
        if (string.Equals(displayCurrency, CurrencyVnd, StringComparison.Ordinal))
            return period.Header;

        var rate = rates.FirstOrDefault(r => string.Equals(r.PeriodKey, period.Key, StringComparison.Ordinal));
        if (rate is null)
            return period.Header + "\nRate missing";

        return displayCurrency switch
        {
            CurrencyUsd => BmesAppliedKrwPerUsd(rate.KrwPerUsd) is { } usd &&
                           BmesAppliedKrwPerVnd(rate.KrwPerVnd) is { } vnd
                ? $"{period.Header}\nKRW/USD {usd:N2}, KRW/VND {vnd:N2}"
                : period.Header + "\nRate missing",
            CurrencyKrw => BmesAppliedKrwPerVnd(rate.KrwPerVnd) is { } vnd
                ? $"{period.Header}\nKRW/VND {vnd:N2}"
                : period.Header + "\nRate missing",
            _ => period.Header,
        };
    }

    private static decimal? ConvertDisplayCurrency(
        decimal vndValue,
        string periodKey,
        IReadOnlyList<BmesFcostExchangeRate> rates,
        string displayCurrency)
    {
        if (vndValue == 0)
            return 0;

        var rate = rates.FirstOrDefault(r => string.Equals(r.PeriodKey, periodKey, StringComparison.Ordinal));
        decimal? krwPerUsd = BmesAppliedKrwPerUsd(rate?.KrwPerUsd);
        decimal? krwPerVnd = BmesAppliedKrwPerVnd(rate?.KrwPerVnd);
        return displayCurrency switch
        {
            CurrencyUsd => krwPerUsd is > 0 && krwPerVnd is > 0
                ? vndValue * krwPerVnd.Value / krwPerUsd.Value
                : null,
            CurrencyKrw => krwPerVnd is > 0 ? vndValue * krwPerVnd.Value : null,
            _ => vndValue,
        };
    }

    private static decimal? BmesAppliedKrwPerUsd(decimal? rate)
        => rate is > 0 ? Math.Round(rate.Value, 2, MidpointRounding.AwayFromZero) : null;

    private static decimal? BmesAppliedKrwPerVnd(decimal? rate)
        => rate is > 0 ? Math.Round(rate.Value, 2, MidpointRounding.AwayFromZero) : null;

    private static string NormalizeCurrencyMode(string? value)
    {
        string normalized = (value ?? string.Empty).Trim().ToUpperInvariant();
        return normalized switch
        {
            CurrencyUsd => CurrencyUsd,
            CurrencyKrw => CurrencyKrw,
            _ => CurrencyVnd,
        };
    }

    private static string DisplayName(string name, string code)
        => !string.IsNullOrWhiteSpace(name)
            ? name
            : (!string.IsNullOrWhiteSpace(code) ? code : "-");

    private static string DisplayText(string value, string fallback)
        => string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();

    private static string FormatFirstUnitPrice(
        IReadOnlyDictionary<string, BmesFcostRawMaterialPeriodPrice> priceByPeriod,
        IReadOnlyList<BmesFcostRawBreakdownPeriod> periods)
    {
        foreach (var period in periods)
        {
            if (priceByPeriod.TryGetValue(period.Key, out var price) && IsUsableUnitPrice(price))
                return FormatUnitPrice(price);
        }

        var fallback = priceByPeriod.Values.FirstOrDefault(IsUsableUnitPrice);
        return fallback is null ? string.Empty : FormatUnitPrice(fallback);
    }

    private static bool IsUsableUnitPrice(BmesFcostRawMaterialPeriodPrice price)
        => !price.IsMixed &&
           price.UnitPrice is not null &&
           !string.IsNullOrWhiteSpace(price.Currency);

    private static string FormatUnitPrice(BmesFcostRawMaterialPeriodPrice price)
    {
        if (price.IsMixed)
            return "Price mixed";
        if (price.UnitPrice is null || string.IsNullOrWhiteSpace(price.Currency))
            return string.Empty;

        string unit = string.IsNullOrWhiteSpace(price.PriceUnit)
            ? string.Empty
            : "/" + price.PriceUnit.Trim();
        return FormatPriceAmount(price.Currency, price.UnitPrice.Value) + unit;
    }

    private static string FormatPriceAmount(string currency, decimal value)
    {
        string format = Math.Abs(value) < 1m && value != 0m
            ? "0.######"
            : "N2";
        return currency.Trim().ToUpperInvariant() + " " + value.ToString(format, CultureInfo.InvariantCulture);
    }
}
