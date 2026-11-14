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
            new("Bôi keo", "C-02", "Bôi keo tổ hợp SPK", 20, null, "2", true, [new("1", "Keo SL233LF", "", "0.02g")], [], [], []),
        ]);
        string text = WorkingManualAiService.InstructionsRequest(doc);
        Assert.Contains("(pvr.xlsx)", text);
        // the sheet name stands in for an empty process name; blank parts are left out
        Assert.Contains("{\"order\":1,\"code\":\"C-01\",\"name\":\"Xếp UPPER\",\"ctq\":false,\"parts\":[{\"name\":\"UPPER CASE\",\"spec\":\"SM-G778B\",\"qty\":\"1\"}],"
            + "\"tools\":[\"Ionizer\"],\"keyPoints\":[\"Không trầy xước\"]}", text);
        Assert.Contains("\"name\":\"Bôi keo tổ hợp SPK\",\"ctq\":true,\"parts\":[{\"name\":\"Keo SL233LF\",\"spec\":\"\",\"qty\":\"0.02g\"}]", text);
        Assert.Contains("keeps its id, type and NG items", text);
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
