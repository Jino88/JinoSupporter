using System.Text.Json.Nodes;
using PortalGNLAB.Web.Services.GraphMaker;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class GraphMakerAiServiceTests
{
    private static GraphMakerTable Table() => GraphMakerParser.Parse(
        "Date\tLine\tGap#1\tGap#2\tUSL\n" +
        "2026-11-01\tA\t0.12\t0.15\t0.3\n" +
        "2026-11-02\tB\t0.14\t0.11\t0.3\n" +
        "2026-11-03\tA\t1,200\t0.13\t0.3\n",
        new("auto", 1));

    [Fact]
    public void BuildPrompt_DescribesColumnsAndRows()
    {
        string prompt = GraphMakerAiService.BuildPrompt(Table());
        Assert.Contains("=== TABLE: 3 data rows, 5 columns ===", prompt);
        Assert.Contains("\"Gap#1\": 3/3 filled, 3 distinct, numeric min 0.12 max 1200", prompt);
        Assert.Contains("\"Line\": 3/3 filled, 2 distinct, text, e.g. \"A\", \"B\"", prompt);
        Assert.Contains("1: 2026-11-01\tA\t0.12\t0.15\t0.3", prompt);
    }

    [Fact]
    public void ParseAnswer_KeepsKnownColumnsAndTypedOptions()
    {
        const string answer = """
            Sure:
            ```json
            {"kind":"Line","title":"Gap trend","reason":"Daily values.","options":{
              "XCOLUMN":"date","yColumns":["Gap#1","gap#2","Nope"],"xColumns":[],"upperLimit":"0.3",
              "lineMode":1,"scatterPointLayout":5,"showScatterAvg":true,"processUpperRowNumber":9,
              "graphColumnGroups":{"Gap#1":"Gap","Missing":"X"},"unknownKey":1}}
            ```
            """;
        GraphMakerAiSuggestion s = GraphMakerAiService.ParseAnswer(answer, Table());
        Assert.Equal(0, s.Kind);
        Assert.Equal("Gap trend", s.Title);
        Assert.Equal("Daily values.", s.Reason);
        JsonObject o = s.Options;
        Assert.Equal("Date", (string?)o["xColumn"]);
        Assert.Equal(["Gap#1", "Gap#2"], o["yColumns"]!.AsArray().Select(x => (string?)x));
        Assert.False(o.ContainsKey("xColumns"));
        Assert.Equal(0.3, (double?)o["upperLimit"]);
        Assert.Equal(1, (int?)o["lineMode"]);
        Assert.False(o.ContainsKey("scatterPointLayout"));
        Assert.True((bool?)o["showScatterAvg"]);
        Assert.False(o.ContainsKey("processUpperRowNumber"));
        Assert.Equal("Gap", (string?)o["graphColumnGroups"]!["Gap#1"]);
        Assert.Single(o["graphColumnGroups"]!.AsObject());
        Assert.False(o.ContainsKey("unknownKey"));
    }

    [Fact]
    public void ParseAnswer_FillsTheXColumnTheChartReads()
    {
        JsonObject scatter = GraphMakerAiService.ParseAnswer("{\"kind\":1,\"options\":{\"xColumn\":\"Line\"}}", Table()).Options;
        Assert.Equal(["Line"], scatter["xColumns"]!.AsArray().Select(x => (string?)x));
        JsonObject average = GraphMakerAiService.ParseAnswer("{\"kind\":2,\"options\":{\"xColumns\":[\"Line\",\"Date\"]}}", Table()).Options;
        Assert.Equal("Line", (string?)average["xColumn"]);
        JsonObject both = GraphMakerAiService.ParseAnswer("{\"kind\":1,\"options\":{\"xColumn\":\"Date\",\"xColumns\":[\"Line\"]}}", Table()).Options;
        Assert.Equal(["Line"], both["xColumns"]!.AsArray().Select(x => (string?)x));
    }

    [Theory]
    [InlineData("{\"kind\":3}", 3)]
    [InlineData("{\"kind\":\"6\"}", 6)]
    [InlineData("{\"kind\":\"Normal distribution\"}", 4)]
    [InlineData("{\"graphType\":\"HeatMap\"}", 7)]
    public void ParseAnswer_ReadsKindAsNumberOrName(string answer, int kind)
    {
        Assert.Equal(kind, GraphMakerAiService.ParseAnswer(answer, Table()).Kind);
    }

    [Theory]
    [InlineData("no json")]
    [InlineData("{\"kind\":9}")]
    [InlineData("{\"kind\":\"pie\"}")]
    [InlineData("{\"kind\": broken}")]
    public void ParseAnswer_RejectsUnreadableAnswers(string answer)
    {
        var ex = Assert.Throws<ArgumentException>(() => GraphMakerAiService.ParseAnswer(answer, Table()));
        Assert.Equal("The AI answer could not be read. Try again.", ex.Message);
    }

    [Fact]
    public void BuildScript_RunsCodexReadOnlyAndWritesLastMessage()
    {
        string script = GraphMakerAiService.BuildScript(@"C:\t\o'k", @"C:\t\p.txt", @"C:\t\a.txt");
        Assert.Contains("codex exec -s read-only --ephemeral --skip-git-repo-check", script);
        Assert.Contains("--cd 'C:\\t\\o''k' -o 'C:\\t\\a.txt' -", script);
        Assert.DoesNotContain("dangerously", script);
    }
}
