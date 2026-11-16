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
        Assert.Equal("☐ New     ☑ Repair     ☐ Scrap (Báo phế)", ws.Cell("C3").GetString());
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
    public void Export_OneRow_HasNoTotal()
    {
        IXLWorksheet ws = Open(new WorkingRequestSave { PartName = "P", Items = [new WorkingRequestItem { Name = "JIG A", LineSpec = 2 }] });
        int h = WorkingRequestExcelExporter.ItemHeadRow;
        Assert.Equal("JIG A", ws.Cell(h + 1, 3).GetString());
        Assert.Equal(h + 1, ws.LastRowUsed()!.RowNumber());
    }
}
