using ClosedXML.Excel;
using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingInstructionServiceTests(Xunit.Abstractions.ITestOutputHelper output) : IDisposable
{
    private readonly string dir = Path.Combine(Path.GetTempPath(), "wi-tests-" + Guid.NewGuid().ToString("N"));

    public void Dispose() { if (Directory.Exists(dir)) Directory.Delete(dir, true); }

    private string Make(Action<XLWorkbook> build, string name = "wi.xlsx")
    {
        Directory.CreateDirectory(dir);
        using var wb = new XLWorkbook();
        build(wb);
        string path = Path.Combine(dir, name);
        wb.SaveAs(path);
        return path;
    }

    private static void History(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet("Lịch sử sửa đổi");
        ws.Cell("B9").Value = "Ngày sửa đổi"; ws.Cell("C9").Value = "Trang"; ws.Cell("D9").Value = "Mã số sửa đổi";
        ws.Cell("E9").Value = "Tóm tắt nội dung sửa đổi"; ws.Cell("I9").Value = "Lập";
        ws.Cell("B10").Value = "26-May"; ws.Cell("C10").Value = 49; ws.Cell("D10").Value = 0; ws.Cell("E10").Value = "Ban hành"; ws.Cell("I10").Value = "Mr.Dung";
        ws.Cell("B11").Value = "22-Sep"; ws.Cell("C11").Value = 42; ws.Cell("D11").Value = 1; ws.Cell("E11").Value = "Thay đổi mask"; ws.Cell("I11").Value = "Mr.Dung";
    }

    private static void Process(XLWorkbook wb, string sheet, bool ctq = false)
    {
        var ws = wb.AddWorksheet(sheet);
        if (ctq) ws.Cell("A1").Value = "Công đoạn CTQ";
        ws.Cell("L1").Value = "Công đoạn\nB1";
        ws.Cell("A4").Value = "Tên công đoạn"; ws.Cell("B4").Value = " Xếp UPPER ";
        ws.Cell("C3").Value = "Thời gian tiêu chuẩn"; ws.Cell("D3").Value = 12.312;
        ws.Cell("C4").Value = "Thời gian thừa"; ws.Cell("D4").Value = 1.2312;
        ws.Cell("L3").Value = "1/53";
        ws.Cell("A5").Value = "Bước 1 "; ws.Cell("B5").Value = "Chuẩn bị UPPER.";
        ws.Cell("A17").Value = "Bước 2:  Xếp Upper vào jig";
        ws.Cell("G5").Value = "PART LIST";
        ws.Cell("G6").Value = "NO"; ws.Cell("H6").Value = "Tên phụ liệu"; ws.Cell("J6").Value = "Quy cách"; ws.Cell("L6").Value = "Số lượng";
        ws.Cell("G7").Value = 1; ws.Cell("H7").Value = "ENCLOSURE"; ws.Cell("J7").Value = "ENC-G778B"; ws.Cell("L7").Value = "1 PCS";
        ws.Cell("G8").Value = 2;
        ws.Cell("G10").Value = "Thiết bị, công cụ";
        ws.Cell("G11").Value = "NO"; ws.Cell("H11").Value = "Tên gọi"; ws.Cell("J11").Value = "Ghi chú";
        ws.Cell("G12").Value = 1; ws.Cell("H12").Value = "IONIZER"; ws.Cell("J12").Value = "SD-95";
        ws.Cell("G13").Value = 2; ws.Cell("H13").Value = "JIG"; ws.Cell("J13").Value = "ENCLOSURE DISPENSOR JIG (V4)";
        ws.Cell("G18").Value = "Hạng mục chú ý quản lý trọng điểm";
        ws.Cell("G19").Value = "1.Không có dị vật."; ws.Cell("G20").Value = "2. Không được xếp ngược.";
        ws.Cell("A45").Value = "BJ-311-02(4)";
    }

    [Fact]
    public void ParseFile_ReadsRevisionsAndOnlyLabelledSheetsAsProcesses()
    {
        string path = Make(wb => { History(wb); Process(wb, " Xếp UPPER ", ctq: true); wb.AddWorksheet("JIG List").Cell("B2").Value = "JIG List"; });
        InstructionDoc doc = WorkingInstructionService.ParseFile(path);

        Assert.Equal(2, doc.Revisions.Count);
        Assert.Equal(new InstructionRevision("26-May", "49", "0", "Ban hành", "Mr.Dung"), doc.Revisions[0]);
        InstructionProcess p = Assert.Single(doc.Processes);
        Assert.Equal("Xếp UPPER", p.Sheet);
        Assert.Equal("B1", p.Code);
        Assert.Equal("Xếp UPPER", p.Name);
        Assert.Equal(12.312, p.StandardTime);
        Assert.Equal(1.231, p.SpareTime);
        Assert.Equal("1/53", p.Page);
        Assert.True(p.Ctq);
        Assert.Equal(new InstructionPart("1", "ENCLOSURE", "ENC-G778B", "1 PCS"), Assert.Single(p.Parts));
        Assert.Equal(["IONIZER", "JIG"], p.Tools.Select(t => t.Name));
        Assert.Equal("ENCLOSURE DISPENSOR JIG (V4)", p.Tools[1].Note);
        Assert.Equal([new InstructionStep(1, "Chuẩn bị UPPER."), new InstructionStep(2, "Xếp Upper vào jig")], p.Steps);
        Assert.Equal(["1.Không có dị vật.", "2. Không được xếp ngược."], p.KeyPoints);
    }

    [Fact]
    public void DocumentKind_TellsWorkInstructionsFromAJigListAndFromAnythingElse()
    {
        WorkingDocumentKind.Found Find(string path) { using FileStream fs = File.OpenRead(path); return WorkingDocumentKind.Of(fs); }
        string Kind(string path) => Find(path).Kind;

        // Work Instructions, even with a sheet named "JIG List" that is not a list of the line
        Assert.Equal(WorkingDocumentKind.Instructions,
            Kind(Make(wb => { History(wb); Process(wb, "A"); wb.AddWorksheet("JIG List").Cell("B2").Value = "JIG List"; })));
        Assert.Equal(WorkingDocumentKind.JigList, Kind(Make(wb =>
        {
            var ws = wb.AddWorksheet("JIG LIST");
            string[] head = ["LINE", "공정\n(Process)", "JIG 및 소모성 자재\n(JIG & Consumable material)", "사용공정 번호\n(Using the process No)"];
            for (int i = 0; i < head.Length; i++) ws.Cell(3, i + 2).Value = head[i];
            ws.Cell(4, 2).Value = "MAIN"; ws.Cell(4, 3).Value = "1. UPPER ARRAY"; ws.Cell(4, 4).Value = "CAP JIG"; ws.Cell(4, 5).Value = "1";
        }, "jig.xlsx")));
        Assert.Equal(WorkingDocumentKind.Unknown, Kind(Make(wb => wb.AddWorksheet("Sheet1").Cell("A1").Value = "hello", "other.xlsx")));
        Directory.CreateDirectory(dir);
        string junk = Path.Combine(dir, "junk.xlsx");
        File.WriteAllText(junk, "hello");
        Assert.Equal(WorkingDocumentKind.Unknown, Kind(junk));

        // an encrypted (DRM) workbook says why it is not used
        string drm = Path.Combine(dir, "drm.xlsx");
        File.WriteAllBytes(drm, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0]);
        Assert.Equal((WorkingDocumentKind.Unknown, "This workbook is encrypted (DRM) and cannot be opened. Upload a decrypted file."),
            (Find(drm).Kind, Find(drm).Reason));
    }

    [Fact]
    public void DocumentKind_TakesWorkbooksAsLargeAsTheWorkInstructions()
    {
        // Work Instructions with their pictures are often 50-100 MB; a smaller limit cut the upload ("Failed to fetch")
        Assert.True(WorkingDocumentKind.MaxUploadBytes >= WorkingInstructionService.MaxUploadBytes);
        Assert.True(WorkingDocumentKind.MaxUploadBytes >= WorkingProcessListService.MaxUploadBytes);
    }

    [Fact]
    public void ParseFile_SheetWithoutCtqLabelIsNotCtq()
    {
        string path = Make(wb => { History(wb); Process(wb, "A"); });
        Assert.False(WorkingInstructionService.ParseFile(path).Processes[0].Ctq);
    }

    [Fact]
    public void ParseFile_EncryptedOrWrongFileGivesReadableError()
    {
        Directory.CreateDirectory(dir);
        string drm = Path.Combine(dir, "drm.xlsx");
        File.WriteAllBytes(drm, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0]);
        var ex = Assert.Throws<ArgumentException>(() => WorkingInstructionService.ParseFile(drm));
        Assert.Contains("Upload a decrypted file", ex.Message);
        // the company DRM wraps the workbook in a "<DOCUMENT" header
        File.WriteAllText(drm, "<DOCUMENT SECURITY=\"x\">PK");
        Assert.Contains("encrypted (DRM)", Assert.Throws<ArgumentException>(() => WorkingInstructionService.ParseFile(drm)).Message);
        string junk = Path.Combine(dir, "junk.xlsx");
        File.WriteAllText(junk, "hello");
        Assert.Throws<ArgumentException>(() => WorkingInstructionService.ParseFile(junk));
    }

    [Fact]
    public void Save_KeepsOriginalAndParsedResultPerModelUntilDeleted()
    {
        var svc = new WorkingInstructionService(Path.Combine(dir, "store"));
        string path = Make(wb => { History(wb); Process(wb, "A"); });
        Assert.Null(svc.Get("MSM-X526/X626B"));

        using (var fs = File.OpenRead(path)) svc.Save("MSM-X526/X626B", "C:\\fake\\원본 지도서.xlsx", fs, "kim");
        InstructionDoc doc = svc.Get("MSM-X526/X626B")!;
        Assert.Equal("원본 지도서.xlsx", doc.FileName);
        Assert.Equal("kim", doc.UploadedBy);
        Assert.Single(doc.Processes);
        Assert.Null(svc.Get("OTHER"));

        var (data, name) = svc.OpenFile("MSM-X526/X626B")!.Value;
        using (data) Assert.Equal(new FileInfo(path).Length, data.Length);
        Assert.Equal("원본 지도서.xlsx", name);

        Assert.True(svc.Delete("MSM-X526/X626B"));
        Assert.Null(svc.Get("MSM-X526/X626B"));
        Assert.Null(svc.OpenFile("MSM-X526/X626B"));
        Assert.False(svc.Delete("MSM-X526/X626B"));
    }

    [Fact]
    public void Save_RejectsNonWorkbookAndKeepsEarlierUpload()
    {
        var svc = new WorkingInstructionService(Path.Combine(dir, "store"));
        string path = Make(wb => { History(wb); Process(wb, "A"); });
        using (var fs = File.OpenRead(path)) svc.Save("M", "a.xlsx", fs, "kim");
        Assert.Throws<ArgumentException>(() => svc.Save("M", "b.xlsx", new MemoryStream(new byte[] { 1, 2, 3, 4 }), "kim"));
        using (var fs = File.OpenRead(path)) Assert.Throws<ArgumentException>(() => svc.Save("M", "b.txt", fs, "kim"));
        Assert.Equal("a.xlsx", svc.Get("M")!.FileName);
    }

    [Fact]
    public void RealWorkbook_IsReadWhenPathGivenByEnvironment()
    {
        string? path = Environment.GetEnvironmentVariable("WORK_INSTRUCTION_XLSX");
        if (string.IsNullOrEmpty(path) || !File.Exists(path)) return; // not set: skipped

        InstructionDoc doc = WorkingInstructionService.ParseFile(path);
        output.WriteLine($"processes={doc.Processes.Count} revisions={doc.Revisions.Count} ctq={doc.Processes.Count(p => p.Ctq)} withParts={doc.Processes.Count(p => p.Parts.Count > 0)} withTools={doc.Processes.Count(p => p.Tools.Count > 0)} withKeys={doc.Processes.Count(p => p.KeyPoints.Count > 0)} noTime={doc.Processes.Count(p => p.StandardTime is null)} noCode={doc.Processes.Count(p => p.Code == "")} gaps={string.Join(";", doc.Processes.Where(p => p.StandardTime is null || p.Code == "" || p.Page == "" || p.Tools.Count == 0).Select(p => p.Sheet + ":" + (p.StandardTime is null) + "/" + p.Code + "/" + p.Page + "/" + p.Tools.Count))} noPage={doc.Processes.Count(p => p.Page == "")}");
        Assert.InRange(doc.Processes.Count, 55, 70);
        Assert.NotEmpty(doc.Revisions);
        Assert.Contains(doc.Processes, p => p.Code == "B1" && p.Name == "Xếp UPPER" && p.StandardTime == 12.312);
        Assert.Contains(doc.Processes, p => p.Code == "A1" && p.Name == "Bôi keo phần tổ hợp SPK" && p.StandardTime == 6.156);
        Assert.Contains(doc.Processes, p => p.Code == "C15" && p.Name == "Kiểm tra Bako" && p.StandardTime == 18.468);
        Assert.Contains(doc.Processes, p => p.Ctq);
        Assert.All(doc.Processes, p => Assert.NotEmpty(p.Steps));
    }
}
