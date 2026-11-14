using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingProcessMatchAiServiceTests
{
    private static WorkingListItem Item(string id, string kind, string name, string process, params int[] nos) =>
        new(id, kind, name, name, process, [.. nos], "");

    /// <summary>Like the real list: the layout numbers count DRY (6), which has no work instruction.</summary>
    private static readonly List<WorkingListItem> Items =
    [
        Item("J1", WorkingListItem.Jig, "ENCLOSURE DISPENSOR JIG", "1. UPPER ARRAY ~ 7. BONDING 2", 1, 2, 3, 4, 5, 6, 7) with { Ver = "V4", Qty = "70", DrawingNo = "BJJ-01" },
        Item("M1", WorkingListItem.Machine, "IONIZER", "UPPER ARRAY", 1),
        new("M2", WorkingListItem.Machine, "NON CONTACT DISPENSOR DT", "NON CONTACT DISPENSOR DT", "BONDING", [2], "SPK 안착단 본드 도포"),
        Item("M3", WorkingListItem.Machine, "CONVEYOR 3M", "DRY", 6),
        Item("M4", WorkingListItem.Machine, "NON CONTACT DISPENSOR DT-B", "BONDING 2", 7),
        Item("T1", WorkingListItem.Tester, "방수검사기 (F520)", "AIR LEAK TEST"),
    ];

    private static readonly List<MatchProcess> Order =
    [
        new(1, "01", "Xếp UPPER", 12.5, null, ["Ionizer"]),
        new(2, "02", "Bôi keo tổ hợp SPK", 20, null, []),
        new(3, "03", "Bôi keo cố định SPK1", 18, null, []),
        new(4, "04", "Kiểm tra AIR LEAK", 30, null, []),
    ];

    [Fact]
    public void BasisOf_TakesTheWorkInstructionsThenTheProcessListThenTheWorkbook()
    {
        var doc = new InstructionDoc("pvr.xlsx", DateTime.UtcNow, "u", 1, [],
        [
            new(" Xếp UPPER ", "C-01", "", 12.5, 2, "1", false, [], [new("1", "Ionizer", ""), new("2", "", "")], [], []),
            new("Bôi keo", "C-02", "Bôi keo tổ hợp SPK", 20, null, "2", true, [], [], [], []),
        ]);
        var list = new WorkingProcessList([new("a", "10", "UPPER ARRAY", 2, 11, "", "", "", "", "")], 1, "u", DateTime.UtcNow, "");

        (string basis, List<MatchProcess> processes) = WorkingProcessMatchAiService.BasisOf(doc, list, Items);
        Assert.Equal("instructions", basis);
        Assert.Equal([new(1, "C-01", "Xếp UPPER", 12.5, null, ["Ionizer"]), new(2, "C-02", "Bôi keo tổ hợp SPK", 20, null, [])],
            processes, new ProcessComparer());

        (basis, processes) = WorkingProcessMatchAiService.BasisOf(null, list, Items);
        Assert.Equal("process-list", basis);
        Assert.Equal(new MatchProcess(1, "10", "UPPER ARRAY", 11, 2, []), Assert.Single(processes), new ProcessComparer());

        (basis, processes) = WorkingProcessMatchAiService.BasisOf(null, new WorkingProcessList([], 0, "", null, ""), Items);
        Assert.Equal("workbook", basis);
        Assert.Equal(["01", "02", "03", "04", "05", "06", "07"], processes.Select(p => p.Code));
        Assert.Equal(["UPPER ARRAY", "BONDING"], processes.Take(2).Select(p => p.Name));
    }

    [Fact]
    public void BuildContent_GivesTheOrderTheWorkbookNumbersAndTheItems()
    {
        string text = WorkingProcessMatchAiService.BuildContent("MSM-G778B", "instructions", Order, Items)[0]!["text"]!.GetValue<string>();
        Assert.Contains("Model name: MSM-G778B", text);
        Assert.Contains("the Work Instructions", text);
        Assert.Contains("{\"seq\":1,\"code\":\"01\",\"name\":\"Xếp UPPER\",\"tools\":[\"Ionizer\"]}", text);
        Assert.Contains("{\"no\":\"06\",\"name\":\"DRY\"}", text);
        Assert.Contains("{\"id\":\"M3\",\"kind\":\"machine\",\"name\":\"CONVEYOR 3M\",\"process\":\"DRY\",\"processNos\":[6]}", text);
    }

    [Fact]
    public void ParseAnswer_ReadsTheJsonInsideOtherText()
    {
        MatchAnswer a = WorkingProcessMatchAiService.ParseAnswer(
            "```json\n{\"rows\":[{\"seq\":\"2\",\"process\":\" BONDING \",\"items\":[\"M2\",\"\",3]},{\"process\":\"DRY\"},5],\"note\":\"ok\"}\n```");
        Assert.Equal("ok", a.Note);
        Assert.Equal(2, a.Rows.Count);
        Assert.Equal((2, "BONDING"), (a.Rows[0].Seq, a.Rows[0].Process));
        Assert.Equal(["M2", "3"], a.Rows[0].Items);
        Assert.Equal((0, "DRY"), (a.Rows[1].Seq, a.Rows[1].Process));
        Assert.Throws<ArgumentException>(() => WorkingProcessMatchAiService.ParseAnswer("no json"));
        Assert.Throws<ArgumentException>(() => WorkingProcessMatchAiService.ParseAnswer("{\"rows\": [1,"));
    }

    [Fact]
    public void Build_KeepsTheProcessOrderAndPutsAddedStepsAfterTheirProcess()
    {
        var answer = new MatchAnswer(
        [
            // out of order, a step the order lacks after process 2, an unknown id, a process given twice
            new(2, "Bôi keo tổ hợp SPK (BONDING)", ["M2", "J1", "X9"], ""),
            new(0, "DRY", ["M3"], "natural dry 10 min"),
            new(1, "", ["m1", "J1"], ""),
            new(3, "Bôi keo cố định SPK1 (BONDING 2)", ["M4"], "unsure"),
            new(1, "again", ["J1"], ""),
            new(9, "", ["M1"], ""),
        ], "matched");

        (List<WorkingProcessRow> rows, List<string> unmatched) = WorkingProcessMatchAiService.Build(answer, Order, Items);

        Assert.Equal(["01", "02", "03", "04", "05"], rows.Select(r => r.No));
        Assert.Equal(["Xếp UPPER", "Bôi keo tổ hợp SPK (BONDING)", "DRY", "Bôi keo cố định SPK1 (BONDING 2)", "Kiểm tra AIR LEAK"],
            rows.Select(r => r.Process));
        Assert.All(rows, r => Assert.Equal("", r.Id));

        WorkingProcessRow upper = rows[0];
        // the jig's version and quantity in their own cells; no standard time (the column is gone)
        Assert.Equal(("IONIZER", (double?)null), (upper.Machine, upper.StdTime));
        Assert.Equal([new WorkingJig("ENCLOSURE DISPENSOR JIG", "BJJ-01", "V4", "70")], upper.Jigs!);
        WorkingProcessRow bonding = rows[1];
        Assert.Equal(("NON CONTACT DISPENSOR DT", "ENCLOSURE DISPENSOR JIG", "SPK 안착단 본드 도포"), (bonding.Machine, bonding.Jigs![0].Name, bonding.Note));
        WorkingProcessRow dry = rows[2];
        Assert.Equal(("CONVEYOR 3M", "natural dry 10 min", (double?)null), (dry.Machine, dry.Note, dry.StdTime));
        Assert.Equal("unsure", rows[3].Note);
        // a process the answer left out is kept, without items
        Assert.Equal(("", "", (double?)null), (rows[4].Machine, rows[4].Tester, rows[4].StdTime));

        // what is left says why it has no process
        Assert.Equal(["Tester: 방수검사기 (F520) — AIR LEAK TEST (no process number; no process of that name in the list)"], unmatched);
    }

    [Fact]
    public void Build_AnEmptyAnswerStillListsTheOrderAndPutsItemsByTheirProcessNumbers()
    {
        (List<WorkingProcessRow> rows, List<string> unmatched) =
            WorkingProcessMatchAiService.Build(new MatchAnswer([], ""), Order, Items);
        Assert.Equal(Order.Select(p => p.Name), rows.Select(r => r.Process));
        // the processes coded 01..04 take the items of those numbers
        Assert.Equal(("IONIZER", "ENCLOSURE DISPENSOR JIG"), (rows[0].Machine, rows[0].Jigs![0].Name));
        Assert.Equal("NON CONTACT DISPENSOR DT", rows[1].Machine);
        Assert.Equal("ENCLOSURE DISPENSOR JIG", Assert.Single(rows[3].Jigs!).Name);
        Assert.Equal([
            "Machine: CONVEYOR 3M — DRY (workbook process 6; no such process in the list)",
            "Machine: NON CONTACT DISPENSOR DT-B — BONDING 2 (workbook process 7; no such process in the list)",
            "Tester: 방수검사기 (F520) — AIR LEAK TEST (no process number; no process of that name in the list)",
        ], unmatched);
    }

    [Fact]
    public void Build_PutsAnItemTheAnswerLeftOutOnTheProcessOfItsName()
    {
        List<MatchProcess> order = [new(1, "C-01", "Bôi keo", null, null, []), new(2, "C-02", "air leak test", null, null, [])];
        (List<WorkingProcessRow> rows, List<string> unmatched) =
            WorkingProcessMatchAiService.Build(new MatchAnswer([new(1, "", ["M2"], "")], ""), order, Items);
        Assert.Equal("방수검사기 (F520)", rows[1].Tester);
        // codes that are not numbers take nothing by number
        Assert.Equal("NON CONTACT DISPENSOR DT", rows[0].Machine);
        Assert.Empty(rows[0].Jigs!);
        Assert.Equal(4, unmatched.Count);
    }

    private sealed class ProcessComparer : IEqualityComparer<MatchProcess>
    {
        public bool Equals(MatchProcess? a, MatchProcess? b) =>
            a is not null && b is not null && (a.Seq, a.Code, a.Name, a.StdTime, a.Workers) == (b.Seq, b.Code, b.Name, b.StdTime, b.Workers)
            && a.Tools.SequenceEqual(b.Tools);
        public int GetHashCode(MatchProcess p) => p.Seq;
    }
}
