using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingManualAiServiceTests
{
    [Fact]
    public void InstructionsRequest_GivesEverySheetInOrderWithItsPartsAsMaterials()
    {
        var doc = new InstructionDoc("pvr.xlsx", DateTime.UtcNow, "u", 1, [],
        [
            new(" Xếp UPPER ", "C-01", "", 12.5, 2, "1", false, [new("1", "UPPER CASE", "SM-G778B", "1"), new("2", " ", "", "")],
                [new("1", "Ionizer", "")], [], ["Không trầy xước"]),
            new("Bôi keo", "C-02", "Bôi keo tổ hợp SPK", 20, null, "2", true, [new("1", "Keo SL233LF", "", "0.02g")],
                [new("5", "Keo", "SL233LF ( Sử dụng trong vòng 24 giờ)"), new("6", " ", "x")], [], []),
        ]);
        string text = WorkingManualAiService.InstructionsRequest(doc);
        Assert.Contains("(pvr.xlsx)", text);
        // the sheet name stands in for an empty process name; blank parts are left out
        Assert.Contains("{\"order\":1,\"code\":\"C-01\",\"name\":\"Xếp UPPER\",\"ctq\":false,\"parts\":[{\"name\":\"UPPER CASE\",\"spec\":\"SM-G778B\",\"qty\":\"1\"}],"
            + "\"tools\":[{\"name\":\"Ionizer\",\"note\":\"\"}],\"keyPoints\":[\"Không trầy xước\"]}", text);
        Assert.Contains("\"name\":\"Bôi keo tổ hợp SPK\",\"ctq\":true,\"parts\":[{\"name\":\"Keo SL233LF\",\"spec\":\"\",\"qty\":\"0.02g\"}]", text);
        // the tools are sent with their notes (the tape, cushion and glue of a process are named there)
        Assert.Contains("\"tools\":[{\"name\":\"Keo\",\"note\":\"SL233LF ( Sử dụng trong vòng 24 giờ)\"}]", text);
        Assert.Contains("the part number in spec", text);
        Assert.Contains("the tape, cushion or other part that stays on the product", text);
        // the glue is not a material
        Assert.Contains("Not materials: glue / bond", text);
        Assert.Contains("not for one that only inspects it", text);
        Assert.Contains("keeps its id, type and NG items", text);
        Assert.True(WorkingManualAiService.InstructionsMaxOutputTokens > WorkingManualAiService.MaxOutputTokens);
    }

    [Fact]
    public void WithoutGlue_TakesTheBondOfTheWorkInstructionsOutOfTheTable()
    {
        // as in the real Work Instructions: Keo in the tools with the product in the note, once as a part
        var doc = new InstructionDoc("pvr.xlsx", DateTime.UtcNow, "u", 1, [],
        [
            new("A", "A1", "Bôi keo", null, null, "1", false, [new("", "Keo", "SR4060 ( Sử dụng trong vòng 24 giờ)", "")],
                [new("1", "Keo", "SL233LF ( Sử dụng trong vòng 24 giờ)"), new("2", "Lượng keo", "3.0-3.5 mg"), new("3", "Keo", "3777 (PUR)  ( Sử dụng )")], [], []),
        ]);
        Assert.Equal(["SR4060", "SL233LF", "3777"], WorkingManualAiService.GlueOf(doc));
        var group = new ManualAiGroup
        {
            Materials = ["ENCLOSURE-G778B-UPPER", "SL-233LF", "SR4060", "3777 (PUR)", "TAPE-G778B"],
            Processes =
            [
                new ManualAiProcess { Name = "Bôi keo", Materials = [new() { Material = "ENCLOSURE-G778B-UPPER" }, new() { Material = "SL233LF" }, new() { Material = "3777 (PUR)" }] },
                new ManualAiProcess { Name = "Dán tape", Materials = [new() { Material = "TAPE-G778B" }] },
            ],
        };
        ManualAiGroup g = WorkingManualAiService.WithoutGlue(group, doc);
        Assert.Equal(["ENCLOSURE-G778B-UPPER", "TAPE-G778B"], g.Materials);
        Assert.Equal(["ENCLOSURE-G778B-UPPER"], g.Processes[0].Materials.Select(m => m.Material));
        Assert.Equal(["TAPE-G778B"], g.Processes[1].Materials.Select(m => m.Material));
        Assert.Equal("Bôi keo", g.Processes[0].Name);
    }

    [Fact]
    public async Task FromInstructionsAsync_NeedsTheWorkInstructions()
    {
        var svc = new WorkingManualAiService(null!);
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => svc.FromInstructionsAsync("M", null, null));
        Assert.Contains("Upload the Work Instructions", ex.Message);
        await Assert.ThrowsAsync<ArgumentException>(() => svc.FromInstructionsAsync("M", null, new InstructionDoc("a.xlsx", DateTime.UtcNow, "u", 1, [], [])));
    }

    [Fact]
    public void ParseAnswer_ReadsFencedJsonWithSurroundingText()
    {
        const string answer = """
            Here you go:
            ```json
            {"group":{"materials":["A"],"processes":[{"id":"p1","type":"MAIN","name":"CUT","localName":"Cắt","ngItems":["Crack"],"materials":[{"material":"A","qtyPerUnit":2}]}]},"note":"Added one."}
            ```
            """;
        (ManualAiGroup g, string note) = WorkingManualAiService.ParseAnswer(answer);
        Assert.Equal("Added one.", note);
        ManualAiProcess p = Assert.Single(g.Processes);
        Assert.Equal("p1", p.Id);
        Assert.Equal("MAIN", p.Type);
        Assert.Equal("Cắt", p.LocalName);
        Assert.Equal(2, p.Materials[0].QtyPerUnit);
    }

    [Theory]
    [InlineData("not json at all")]
    [InlineData("{\"note\":\"x\"}")]
    [InlineData("{\"group\": 5}")]
    [InlineData("{\"group\":{ broken }")]
    public void ParseAnswer_BadAnswerThrowsReadableError(string answer)
    {
        var ex = Assert.Throws<ArgumentException>(() => WorkingManualAiService.ParseAnswer(answer));
        Assert.Equal("The AI answer could not be read. Try again.", ex.Message);
    }

    [Fact]
    public void Normalize_FixesTypeTrimsDedupesAndQty()
    {
        var group = new ManualAiGroup
        {
            Materials = [" A ", "a", "", "B"],
            Processes =
            [
                new ManualAiProcess
                {
                    Id = " x ", Type = "weird", Name = " N ", LocalName = " L ",
                    NgItems = [" Crack ", "crack", "", "  ", "Dent"],
                    Materials =
                    [
                        new ManualAiMaterial { Material = "B", QtyPerUnit = 0 },
                        new ManualAiMaterial { Material = "b", QtyPerUnit = 3 },
                        new ManualAiMaterial { Material = "C", QtyPerUnit = -2 },
                        new ManualAiMaterial { Material = " ", QtyPerUnit = 1 },
                        new ManualAiMaterial { Material = "D", QtyPerUnit = 0.5 },
                    ],
                },
                new ManualAiProcess { Type = "visual", NgItems = null!, Materials = null! },
            ],
        };

        ManualAiGroup n = WorkingManualAiService.Normalize(group);

        Assert.Equal(["A", "B", "C", "D"], n.Materials);
        ManualAiProcess p = n.Processes[0];
        Assert.Equal("x", p.Id);
        Assert.Equal("SUB", p.Type);
        Assert.Equal("N", p.Name);
        Assert.Equal("L", p.LocalName);
        Assert.Equal(["Crack", "Dent"], p.NgItems);
        Assert.Equal(["B", "C", "D"], p.Materials.Select(m => m.Material));
        Assert.Equal([1d, 1d, 0.5], p.Materials.Select(m => m.QtyPerUnit));
        Assert.Equal("VISUAL", n.Processes[1].Type);
        Assert.Empty(n.Processes[1].NgItems);
        Assert.Empty(n.Processes[1].Materials);
    }

    [Fact]
    public void Normalize_OneMaterialWrittenTwoWaysKeepsTheFirstName()
    {
        // as the AI wrote them for the real Work Instructions: SL-233LF next to SL233LF, extra spaces around a hyphen
        var group = new ManualAiGroup
        {
            Materials = ["ENCLOSURE-G778B-LOWER", "SL233LF"],
            Processes =
            [
                new ManualAiProcess { Name = "A", Materials = [new() { Material = "SL-233LF", QtyPerUnit = 1 }, new() { Material = "ENCLOSURE-G778B - LOWER", QtyPerUnit = 1 }] },
                new ManualAiProcess { Name = "B", Materials = [new() { Material = "Screen  -X526B-BOTTOM", QtyPerUnit = 1 }, new() { Material = "Screen -X526B-BOTTOM", QtyPerUnit = 2 }] },
            ],
        };
        ManualAiGroup n = WorkingManualAiService.Normalize(group);
        Assert.Equal(["ENCLOSURE-G778B-LOWER", "SL233LF", "Screen -X526B-BOTTOM"], n.Materials);
        Assert.Equal(["SL233LF", "ENCLOSURE-G778B-LOWER"], n.Processes[0].Materials.Select(m => m.Material));
        Assert.Equal("Screen -X526B-BOTTOM", Assert.Single(n.Processes[1].Materials).Material);
    }

    [Fact]
    public void Normalize_NullArraysBecomeEmpty()
    {
        ManualAiGroup n = WorkingManualAiService.Normalize(new ManualAiGroup { Materials = null!, Processes = null! });
        Assert.Empty(n.Materials);
        Assert.Empty(n.Processes);
        Assert.Empty(WorkingManualAiService.Normalize(null).Processes);
    }

    [Fact]
    public void ParseAnswer_AddsProcessMaterialMissingFromList()
    {
        const string answer = """{"group":{"materials":[],"processes":[{"id":"","type":"SUB","name":"P","localName":"","ngItems":null,"materials":[{"material":"Glue","qtyPerUnit":0}]}]},"note":""}""";
        (ManualAiGroup g, _) = WorkingManualAiService.ParseAnswer(answer);
        Assert.Equal(["Glue"], g.Materials);
        Assert.Equal(1d, g.Processes[0].Materials[0].QtyPerUnit);
    }
}
