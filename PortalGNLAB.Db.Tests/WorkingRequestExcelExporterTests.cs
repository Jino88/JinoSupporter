using ClosedXML.Excel;
using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingRequestExcelExporterTests
{
    private static IXLWorksheet Open(WorkingRequestSave r)
    {
        var wb = new XLWorkbook(new MemoryStream(WorkingRequestExcelExporter.Export(r)));
        return wb.Worksheet("Request");
    }

    [Fact]
    public void Export_LaysTheFormOutLikeThePaperForm()
    {
        IXLWorksheet ws = Open(new WorkingRequestSave
        {
            Requester = "Kim", Method = "repair", PartName = "SM-A386TOP", ItemApplication = "GLONICS", UsingPlace = "C2",
            RequestDate = "2026-11-05", FinishDate1 = "2026-11-20", DeptCode = "ME-2611004",
            Items =
            [
                new WorkingRequestItem { Model = "SM-A386TOP", Name = "MAIN 1 JIG-Y", LineSpec = 40, DrawingNo = "BJJ2605-MO-05-A00-02", Note = "Process: 1-5, 7-9" },
                new WorkingRequestItem { Name = "" }, // no name: dropped as on saving
                new WorkingRequestItem { Model = "SM-A386TOP", Name = "MAIN 2 JIG-Y", LineSpec = 10, Stock = 4, Order = 0 },
            ],
        });

        Assert.Equal("PROCESSING REQUEST", ws.Cell("F1").GetString());
        Assert.True(ws.Cell("F1").IsMerged());
        Assert.Equal("Kim", ws.Cell("B2").GetString());
        Assert.Equal("☐ New        ☑ Repair        ☐ Scrap (Báo phế)", ws.Cell("C3").GetString());
        Assert.Equal("ME GROUP", ws.Cell("L3").GetString());
        Assert.Equal("SM-A386TOP", ws.Cell("C5").GetString());
        Assert.Equal("GLONICS", ws.Cell("C6").GetString());
        Assert.Equal("2026.11.05", ws.Cell("L6").GetString());
        Assert.Equal("2026.11.20", ws.Cell("C9").GetString());
        Assert.Equal("ME-2611004", ws.Cell("L10").GetString());

        int h = WorkingRequestExcelExporter.ItemHeadRow;
        Assert.Equal("Picture/Drawing no", ws.Cell(h, 9).GetString());
        // the rows numbered again, the order made from line and stock when none is typed, the note wrapping
        Assert.Equal(1, ws.Cell(h + 1, 1).GetDouble());
        Assert.Equal("MAIN 1 JIG-Y", ws.Cell(h + 1, 3).GetString());
        Assert.Equal(40, ws.Cell(h + 1, 8).GetDouble());
        Assert.Equal("BJJ2605-MO-05-A00-02", ws.Cell(h + 1, 9).GetString());
        Assert.Equal("Process: 1-5, 7-9", ws.Cell(h + 1, 12).GetString());
        Assert.True(ws.Cell(h + 1, 12).Style.Alignment.WrapText);
        Assert.Equal(2, ws.Cell(h + 2, 1).GetDouble());
        Assert.Equal(6, ws.Cell(h + 2, 8).GetDouble());
        Assert.Equal("Total", ws.Cell(h + 3, 1).GetString());
        Assert.Equal(50, ws.Cell(h + 3, 6).GetDouble());
        Assert.Equal(4, ws.Cell(h + 3, 7).GetDouble());
        Assert.Equal(46, ws.Cell(h + 3, 8).GetDouble());
        Assert.Equal(h + 3, ws.LastRowUsed()!.RowNumber());
        Assert.Equal(XLPaperSize.A4Paper, ws.PageSetup.PaperSize);
        Assert.Equal(1, ws.PageSetup.PagesWide);
    }

    [Fact]
    public void Export_LooksLikeTheFormOnTheScreen()
    {
        IXLWorksheet ws = Open(new WorkingRequestSave
        {
            PartName = "SM-A386TOP", RequestDate = "2026-11-05", DeptCode = "ME-2611004",
            Items =
            [
                new WorkingRequestItem { Name = "JIG A", LineSpec = 2 },
                new WorkingRequestItem { Name = "JIG B", LineSpec = 1, Note = string.Join(", ", Enumerable.Range(1, 30)) },
            ],
        });
        int h = WorkingRequestExcelExporter.ItemHeadRow, last = h + 3;

        // the columns of the 980px form (not twice as wide), so A4 portrait prints it at about 75% like the browser
        Assert.Equal(980, WorkingRequestExcelExporter.ColumnsPx.Sum());
        double px = Enumerable.Range(1, 13).Sum(c => ws.Column(c).Width * 7 + 5);
        Assert.InRange(px, 970, 990);
        Assert.False(ws.ShowGridLines);

        // every cell of the form: Segoe UI, centred, a thin dark line all round, no grey labels
        foreach (IXLCell c in ws.Range(1, 1, last, 13).Cells())
        {
            Assert.Equal(XLBorderStyleValues.Thin, c.Style.Border.LeftBorder);
            Assert.Equal(XLBorderStyleValues.Thin, c.Style.Border.TopBorder);
            Assert.Equal(XLBorderStyleValues.Thin, c.Style.Border.RightBorder);
            Assert.Equal(XLBorderStyleValues.Thin, c.Style.Border.BottomBorder);
        }
        foreach (string a in new[] { "A1", "B1", "F1", "A3", "C4", "J6", "L6", "A10" })
        {
            Assert.Equal(WorkingRequestExcelExporter.FontName, ws.Cell(a).Style.Font.FontName);
            Assert.Equal(XLAlignmentHorizontalValues.Center, ws.Cell(a).Style.Alignment.Horizontal);
        }
        Assert.Equal(XLFillPatternValues.None, ws.Cell("A4").Style.Fill.PatternType);
        Assert.False(ws.Cell("A4").Style.Font.Bold);
        // the sizes of request.css: 13px text, 20px bold title, 30px part name, 16px bold dept code
        Assert.Equal(9.75, ws.Cell("A4").Style.Font.FontSize);
        Assert.Equal(15, ws.Cell("F1").Style.Font.FontSize);
        Assert.True(ws.Cell("F1").Style.Font.Bold);
        Assert.Equal(22.5, ws.Cell("C5").Style.Font.FontSize);
        Assert.Equal(12, ws.Cell("L10").Style.Font.FontSize);
        Assert.True(ws.Cell("L10").Style.Font.Bold);
        // the heights: 58px sign boxes, 34px rows, 40px item heading; a wrapping note makes its row taller
        Assert.Equal(43.5, ws.Row(2).Height);
        Assert.Equal(25.5, ws.Row(4).Height);
        Assert.Equal(30, ws.Row(h).Height);
        Assert.Equal(25.5, ws.Row(h + 1).Height);
        Assert.True(ws.Row(h + 2).Height > 25.5);
        Assert.Equal(XLAlignmentHorizontalValues.Left, ws.Cell(h + 2, 12).Style.Alignment.Horizontal);

        // A4 portrait, one page wide, centred, only the form printed
        Assert.Equal(XLPageOrientation.Portrait, ws.PageSetup.PageOrientation);
        Assert.True(ws.PageSetup.CenterHorizontally);
        Assert.Equal($"$A$1:$M${last}", ws.PageSetup.PrintAreas.Single().RangeAddress.ToString());
    }

    [Fact]
    public void Export_OneRow_HasNoTotal()
    {
        IXLWorksheet ws = Open(new WorkingRequestSave { PartName = "P", Items = [new WorkingRequestItem { Name = "JIG A", LineSpec = 2 }] });
        int h = WorkingRequestExcelExporter.ItemHeadRow;
        Assert.Equal("JIG A", ws.Cell(h + 1, 3).GetString());
        Assert.Equal(h + 1, ws.LastRowUsed()!.RowNumber());
    }
}
