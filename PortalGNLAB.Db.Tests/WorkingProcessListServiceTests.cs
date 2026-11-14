using ClosedXML.Excel;
using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingProcessListServiceTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "working-process-list-tests-" + Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        try { Directory.Delete(_dir, true); } catch { }
    }

    private static WorkingProcessRow Row(string no, string process, int? workers = null) =>
        new("", no, process, workers, null, "", "", "", "", "");

    /// <summary>A small JIG / Machine / Tester list like the real one: title rows above the header, merged cells, two lines.</summary>
    private static MemoryStream Workbook()
    {
        using var book = new XLWorkbook();
        var jig = book.AddWorksheet("JIG LIST");
        jig.Cell(1, 2).Value = "JIG List 778 1K/day";
        string[] jh = ["순서\n(No)", "MASS\nFACTORY", "LINE", "공정\n(Process)", "JIG 및 소모성 자재\n(JIG & Consumable material)", "MASS JIG",
                       "사용공정 번호\n(Using the process No)", "MSM-G778B", "Ver."];
        for (int i = 0; i < jh.Length; i++) jig.Cell(3, i + 2).Value = jh[i];
        object[][] jigs =
        [
            [1, "GN", "MAIN", "1)  1. UPPER ARRAY ~ 5. SPK ASSY\n2) 7. AOI TEST 2", "ENCLOSURE DISPENSOR CAP JIG", 6, "1) 1 ~ 2\n2) 7", "", "V1"],
            [2, "GN", "MAIN", "4. SPK ARRAY ~ 5. SPK ASSY", "SPK ARRAY JIG", 4, "4 ~ 5", "", "V0"],
            [3, "GN", "QM", "신뢰성 지그", "DROP TEST JIG", 1, "49", "", "V3"],
        ];
        for (int r = 0; r < jigs.Length; r++)
            for (int c = 0; c < jigs[r].Length; c++) jig.Cell(5 + r, c + 2).Value = XLCellValue.FromObject(jigs[r][c]);

        var mc = book.AddWorksheet("Machine List");
        mc.Cell(1, 2).Value = "설비 List(Machine List)";
        string[] mh = ["순서\n(No)", "MASS\nFACTORY", "Line", "사용공정 번호\n(Using the process No)", "설비명\n(Machine Name)",
                       "사용 공정\n(Using the process)", "소요 수량\n(Required quantity)", "우선 순위", "실물사진", "도번(Drawing No)", "비고\n(Note)"];
        for (int i = 0; i < mh.Length; i++) mc.Cell(3, i + 2).Value = mh[i];
        object[][] machines =
        [
            [1, "GN", "", 1, "IONIZER", "UPPER ARRAY", 1, 3, "", "", ""],
            [2, "GN", "", 2, "NON CONTACT DISPENSOR DT", "BONDING\n", 1, 1, "", "BJGP26-14-A", "SPK 안착단 본드 도포\n - SL233LF, VZ15-T2"],
            [3, "GN", "", 46, "AIR LEAK TEST S.A(CP)", "AIR LEAK TEST", 2, 1, "", "", ""],
            [4, "GN", "", "", "", "", "", "", "", "", ""],
        ];
        for (int r = 0; r < machines.Length; r++)
            for (int c = 0; c < machines[r].Length; c++) mc.Cell(4 + r, c + 2).Value = XLCellValue.FromObject(machines[r][c]);

        var ts = book.AddWorksheet("Tester List");
        string[] th = ["순서\n(No)", "MASS\nFACTORY", "계측기명\n(Tester Name)", "모델명\n(Model Name)", "제작사\n(Maker)", "사용 공정\n(Using the process)",
                       "사용공정 번호\n(Using the process No)"];
        for (int i = 0; i < th.Length; i++) ts.Cell(3, i + 2).Value = th[i];
        ts.Cell(5, 2).Value = 1; ts.Cell(5, 4).Value = "방수검사기\n(Warterproof Function Tester)"; ts.Cell(5, 5).Value = "F520";
        ts.Cell(5, 7).Value = "AIR LEAK TEST";

        // a drawing transfer list names jigs and machines too, but is not a list of the line
        var transfer = book.AddWorksheet("사업장 도면 이관 LIST");
        string[] trh = ["NO", "MODEL", "도번", "Ver.", "JIG & 설비명", "양산처"];
        for (int i = 0; i < trh.Length; i++) transfer.Cell(2, i + 1).Value = trh[i];
        transfer.Cell(3, 1).Value = 1; transfer.Cell(3, 2).Value = "MSM-G778B"; transfer.Cell(3, 5).Value = "PLATE TRAY";

        var ms = new MemoryStream();
        book.SaveAs(ms);
        ms.Position = 0;
        return ms;
    }

    [Fact]
    public void ReadWorkbook_JoinsMachinesJigsAndTestersByProcessNumber()
    {
        List<WorkingProcessRow> rows = WorkingProcessListService.ReadWorkbook(Workbook());

        Assert.Equal(["01", "02", "04", "05", "07", "46"], rows.Select(r => r.No));
        WorkingProcessRow first = rows[0];
        Assert.Equal("UPPER ARRAY", first.Process);
        Assert.Equal("IONIZER", first.Machine);
        Assert.Equal(("ENCLOSURE DISPENSOR CAP JIG", "V1", "6"), (first.Jig, first.JigVer, first.JigQty));

        WorkingProcessRow bonding = rows[1];
        Assert.Equal("BONDING", bonding.Process);
        Assert.Equal("SPK 안착단 본드 도포 / - SL233LF, VZ15-T2", bonding.Note);

        // names from the JIG LIST "N. NAME" when the Machine List has no row for the number
        Assert.Equal("SPK ASSY", rows.Single(r => r.No == "05").Process);
        WorkingProcessRow four = rows.Single(r => r.No == "04");
        Assert.Equal(("SPK ARRAY JIG", "V0", "4"), (four.Jig, four.JigVer, four.JigQty));
        Assert.Null(four.StdTime);
        Assert.Equal("AOI TEST 2", rows.Single(r => r.No == "07").Process);
        // the QM line's jigs are not on the line
        Assert.DoesNotContain(rows, r => r.Jig.Contains("DROP TEST JIG"));

        WorkingProcessRow leak = rows.Single(r => r.No == "46");
        Assert.Equal("AIR LEAK TEST S.A(CP) ×2", leak.Machine);
        Assert.Equal("방수검사기 / (Warterproof Function Tester) (F520)", leak.Tester);
    }

    [Fact]
    public void ReadItems_ListsEveryMachineJigAndTesterWithItsProcess()
    {
        List<WorkingListItem> items = WorkingProcessListService.ReadItems(Workbook());

        Assert.Equal(["J1", "J2", "M1", "M2", "M3", "T1"], items.Select(i => i.Id));
        WorkingListItem cap = items[0];
        Assert.Equal((WorkingListItem.Jig, "ENCLOSURE DISPENSOR CAP JIG", "ENCLOSURE DISPENSOR CAP JIG", "V1", "6"), (cap.Kind, cap.Name, cap.Label, cap.Ver, cap.Qty));
        Assert.Equal("1)  1. UPPER ARRAY ~ 5. SPK ASSY / 2) 7. AOI TEST 2", cap.Process);
        Assert.Equal([1, 2, 7], cap.ProcessNos);
        WorkingListItem bonding = items.Single(i => i.Id == "M2");
        Assert.Equal(("BONDING", "SPK 안착단 본드 도포 / - SL233LF, VZ15-T2"), (bonding.Process, bonding.Note));
        Assert.Equal("AIR LEAK TEST S.A(CP) ×2", items.Single(i => i.Id == "M3").Label);
        // a tester without numbers keeps the process it is written for
        WorkingListItem tester = items.Single(i => i.Id == "T1");
        Assert.Equal(("AIR LEAK TEST", 0), (tester.Process, tester.ProcessNos.Count));
        // the QM line's jigs are not on the line
        Assert.DoesNotContain(items, i => i.Name == "DROP TEST JIG");
        // nor the drawing transfer list
        Assert.DoesNotContain(items, i => i.Name == "PLATE TRAY");
    }

    [Fact]
    public void Row_PutsEachItemOnItsOwnLineAndTheJigVersionsAndQuantitiesOnTheSameLines()
    {
        static WorkingListItem It(string kind, string label, string ver = "", string qty = "") => new("", kind, label, label, "", [], "", ver, qty);
        WorkingProcessRow row = WorkingProcessListService.Row("", "01", "BONDING",
        [
            It(WorkingListItem.Machine, "IONIZER"), It(WorkingListItem.Machine, "DISPENSOR ×2"), It(WorkingListItem.Machine, "IONIZER"),
            It(WorkingListItem.Jig, "NO VERSION JIG", "", "2"), It(WorkingListItem.Jig, "CAP JIG", "V1", "6"), It(WorkingListItem.Jig, "CAP JIG", "V1", "6"),
            It(WorkingListItem.Jig, "TRAY", "V3", ""), It(WorkingListItem.Tester, "F520"),
        ], "");
        Assert.Equal(("IONIZER\nDISPENSOR ×2", "F520"), (row.Machine, row.Tester));
        Assert.Equal(("NO VERSION JIG\nCAP JIG\nTRAY", "\nV1\nV3", "2\n6"), (row.Jig, row.JigVer, row.JigQty));
        Assert.Equal(((int?)null, (double?)null), (row.Workers, row.StdTime));

        // saving keeps the blank first version line (it belongs to the first jig) and leaves out blank lines at the end
        WorkingProcessRow saved = Assert.Single(WorkingProcessListService.Clean([row with { JigVer = " \r\n V1 \nV3\n\n", JigQty = "2\n6\n" }]));
        Assert.Equal(("\nV1\nV3", "2\n6", "IONIZER\nDISPENSOR ×2"), (saved.JigVer, saved.JigQty, saved.Machine));
    }

    [Fact]
    public void Numbers_ReadsRangesListsAndListMarkers()
    {
        Assert.Equal([1, 2, 3, 4, 5, 7, 8, 9], WorkingProcessListService.Numbers("1) 1 ~ 5\n2) 7 ~ 9"));
        Assert.Equal([36, 37, 40, 42], WorkingProcessListService.Numbers("36, 37, 40, 42"));
        Assert.Equal([48], WorkingProcessListService.Numbers("48"));
        Assert.Empty(WorkingProcessListService.Numbers("-"));
    }

    [Fact]
    public void ReadWorkbook_RejectsWhatIsNotAListWorkbook()
    {
        Assert.Throws<ArgumentException>(() => WorkingProcessListService.ReadWorkbook(new MemoryStream([1, 2, 3, 4])));
        using var book = new XLWorkbook();
        book.AddWorksheet("Sheet1").Cell(1, 1).Value = "hello";
        var ms = new MemoryStream();
        book.SaveAs(ms);
        ms.Position = 0;
        Assert.Throws<ArgumentException>(() => WorkingProcessListService.ReadWorkbook(ms));
    }

    [Fact]
    public void Save_KeepsRowsPerModelWithRevisions()
    {
        var svc = new WorkingProcessListService(_dir);
        Assert.Empty(svc.Get("MSM-G778B").Rows);

        WorkingProcessList saved = svc.Save("MSM-G778B", [Row("01", " UPPER ARRAY ", 2), Row("", ""), null], 0, "u1");
        Assert.Equal(1, saved.Revision);
        WorkingProcessRow row = Assert.Single(saved.Rows);
        Assert.Equal("UPPER ARRAY", row.Process);
        Assert.Equal(2, row.Workers);
        Assert.False(string.IsNullOrEmpty(row.Id));

        Assert.Equal((1, "u1"), (svc.Get("msm-g778b ").Revision, svc.Get("msm-g778b ").UpdatedBy));
        Assert.Equal(row, svc.Get("MSM-G778B").Rows[0]);
        Assert.Empty(svc.Get("MSM-X316B").Rows);
        // a save from an older revision is refused
        Assert.Throws<WorkingConflictException>(() => svc.Save("MSM-G778B", [Row("02", "GLUING")], 0, "u2"));
        Assert.Equal(2, svc.Save("MSM-G778B", [], 1, "u2").Revision);
    }

    /// <summary>The real JIG / Machine / Tester list, when PROCESS_LIST_XLSX points at it (not in the repository).</summary>
    [Fact]
    public void ReadWorkbook_RealList_WhenGiven()
    {
        string? path = Environment.GetEnvironmentVariable("PROCESS_LIST_XLSX");
        if (string.IsNullOrEmpty(path) || !File.Exists(path)) return;
        using FileStream fs = File.OpenRead(path);
        List<WorkingProcessRow> rows = WorkingProcessListService.ReadWorkbook(fs);
        foreach (WorkingProcessRow r in rows)
            Console.WriteLine($"{r.No} | {r.Process} | {r.Machine} | {r.Jig} | {r.Tester} | {r.Note}");
        Assert.Contains(rows, r => r.No == "02" && r.Machine == "NON CONTACT DISPENSOR DT");
        Assert.Contains(rows, r => r.Jig.Split('\n').Zip(r.JigVer.Split('\n')).Contains(("ENCLOSURE DISPENSOR JIG", "V4")));
        Assert.True(rows.Count >= 40);

        fs.Position = 0;
        List<WorkingListItem> items = WorkingProcessListService.ReadItems(fs);
        Console.WriteLine($"items: {items.Count(i => i.Kind == WorkingListItem.Machine)} machines, "
                          + $"{items.Count(i => i.Kind == WorkingListItem.Jig)} jigs, {items.Count(i => i.Kind == WorkingListItem.Tester)} testers");
        Assert.Contains(items, i => i.Kind == WorkingListItem.Machine && i.Name == "IONIZER" && i.ProcessNos.SequenceEqual([1]));
        Assert.Contains(items, i => i.Kind == WorkingListItem.Tester && i.Process == "AIR LEAK TEST");
        Assert.InRange(items.Count(i => i.Kind == WorkingListItem.Machine), 20, 60);
    }
}
