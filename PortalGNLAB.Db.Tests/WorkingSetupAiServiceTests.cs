using System.Text.Json.Nodes;
using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingSetupAiServiceTests
{
    private static SetupAiRequest Request() => new()
    {
        Text = "x",
        Today = "2026-11-12",
        Rows =
        [
            new SetupAiRow { Id = "r1", No = 1, Line = "L1", ProcessName = "Soldering", SetupStatus = "X" },
            new SetupAiRow { Id = "r2", No = 2, Line = "L1", ProcessName = "Assembly" },
        ],
    };

    [Fact]
    public void ParseAnswer_ReadsFencedJsonAndNormalizesRows()
    {
        const string answer = """
            Here you go:
            ```json
            {"rows":[
              {"id":"r1","no":1,"layoutNo":"A-1","line":"L1","processName":"Soldering","machinePlan":"2026/11/20","machineState":"△","machineNote":"x","setupPlan":"2026.11.21","setupFinish":"2026-11-25","setupStatus":"O"},
              {"id":"","no":"3","line":"L2","processName":" Test ","setupFinish":"11/30","machineState":"maybe","setupStatus":"done"}
            ],"note":" Done. "}
            ```
            """;
        SetupAiAnswer a = WorkingSetupAiService.ParseAnswer(answer, Request());
        Assert.Equal("Done.", a.Note);
        Assert.Equal(2, a.Rows.Count);
        SetupAiRow changed = a.Rows[0];
        Assert.Equal("r1", changed.Id);
        Assert.Equal("2026-11-20", changed.MachinePlan);
        Assert.Equal("Δ", changed.MachineState);
        Assert.Equal("2026-11-21", changed.SetupPlan);
        Assert.Equal("O", changed.SetupStatus);
        SetupAiRow added = a.Rows[1];
        Assert.Equal("", added.Id);
        Assert.Equal(3, added.No);
        Assert.Equal("Test", added.ProcessName);
        Assert.Equal("", added.MachineState);
        Assert.Equal("", added.SetupStatus);
        Assert.Equal("", added.SetupFinish); // a date without a year is not guessed
    }

    [Theory]
    [InlineData("not json at all")]
    [InlineData("[1,2]")]
    [InlineData("{\"rows\": [5]}")]
    [InlineData("{ broken }")]
    public void ParseAnswer_BadAnswerThrowsReadableError(string answer)
    {
        var ex = Assert.Throws<ArgumentException>(() => WorkingSetupAiService.ParseAnswer(answer, Request()));
        Assert.Equal("The AI answer could not be read. Try again.", ex.Message);
    }

    [Fact]
    public void Normalize_UnknownIdBecomesNewRowAndDuplicatesAreDropped()
    {
        var answer = new SetupAiAnswer
        {
            Rows =
            [
                new SetupAiRow { Id = "nope", Line = "L3", ProcessName = "Packing" },                  // unknown id -> new row
                new SetupAiRow { Id = "", Line = "l1", ProcessName = "soldering" },                    // repeats a current row
                new SetupAiRow { Id = "r2", Line = "L1", ProcessName = "Assembly", SetupStatus = "O" }, // change
                new SetupAiRow { Id = "r2", Line = "L1", ProcessName = "Assembly", SetupStatus = "X" }, // same row twice
                new SetupAiRow { Id = "", Line = "L3", ProcessName = "Packing" },                      // repeats the new row
                new SetupAiRow { Id = "", Line = "L4", ProcessName = "  " },                           // no process
            ],
        };
        SetupAiAnswer a = WorkingSetupAiService.Normalize(answer, Request());
        Assert.Equal(2, a.Rows.Count);
        Assert.Equal("", a.Rows[0].Id);
        Assert.Equal("Packing", a.Rows[0].ProcessName);
        Assert.Equal("r2", a.Rows[1].Id);
        Assert.Equal("O", a.Rows[1].SetupStatus);
    }

    [Fact]
    public void Normalize_ChangeKeepsProcessNameAndLimitsLengths()
    {
        var answer = new SetupAiAnswer
        {
            Rows = [new SetupAiRow { Id = "r1", ProcessName = "", LayoutNo = new string('a', 30), MachineNote = new string('n', 150), SetupFinish = "31/31/2026" }],
        };
        SetupAiRow row = Assert.Single(WorkingSetupAiService.Normalize(answer, Request()).Rows);
        Assert.Equal("Soldering", row.ProcessName);
        Assert.Equal(20, row.LayoutNo.Length);
        Assert.Equal(100, row.MachineNote.Length);
        Assert.Equal("", row.SetupFinish);
    }

    [Fact]
    public void BuildContent_HasRowsTextAndPictures()
    {
        var pics = new List<PlaceAiPicture> { new("a.png", "image/png", "AAAA") };
        JsonArray content = WorkingSetupAiService.BuildContent("M1", Request(), pics);
        Assert.Equal(2, content.Count);
        string text = content[0]!["text"]!.GetValue<string>();
        Assert.Contains("Model name: M1", text);
        Assert.Contains("\"processName\":\"Soldering\"", text);
        Assert.Contains("Picture 0: a.png", text);
        Assert.Equal("data:image/png;base64,AAAA", content[1]!["image_url"]!.GetValue<string>());
    }
}
