using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingPlaceAiServiceTests
{
    private static PlaceAiRequest Request() => new()
    {
        Text = "x",
        Today = "2026-11-12",
        Info = new PlaceAiInfo { Customer = "ACME", Phase = "DR" },
        Events =
        [
            new PlaceAiEvent
            {
                Id = "e1", Name = "DV build", Start = "2026-11-01", End = "2026-11-20",
                Tasks = [new PlaceAiTask { Id = "t1", Name = "Tooling" }],
            },
        ],
    };

    [Fact]
    public void ParseAnswer_ReadsFencedJsonWithSurroundingText()
    {
        const string answer = """
            Here you go:
            ```json
            {"info":{"build":" B1 "},"events":[{"id":"e1","name":"DV build","start":"2026/11/01","end":"2026-11-20","tasks":[{"id":"t1","name":"Tooling","start":"2026.11.03","end":"","owner":"Kim","done":true}]}],"entries":[{"kind":"dev","title":"Kick-off"}],"note":" Placed. "}
            ```
            Done.
            """;
        PlaceAiPlan plan = WorkingPlaceAiService.ParseAnswer(answer, Request());
        Assert.Equal("Placed.", plan.Note);
        Assert.Equal("B1", plan.Info["build"]);
        PlaceAiEvent e = Assert.Single(plan.Events);
        Assert.Equal("e1", e.Id);
        PlaceAiTask t = Assert.Single(e.Tasks);
        Assert.Equal("t1", t.Id);
        Assert.Equal("2026-11-03", t.Start);
        Assert.True(t.Done);
        Assert.Equal("2026-11-12", Assert.Single(plan.Entries).Date);
    }

    [Theory]
    [InlineData("not json at all")]
    [InlineData("[1,2]")]
    [InlineData("{\"events\": [5]}")]
    [InlineData("{ broken }")]
    public void ParseAnswer_BadAnswerThrowsReadableError(string answer)
    {
        var ex = Assert.Throws<ArgumentException>(() => WorkingPlaceAiService.ParseAnswer(answer, Request()));
        Assert.Equal("The AI answer could not be read. Try again.", ex.Message);
    }

    [Fact]
    public void Normalize_InfoKeepsAllowedKeysConvertsPhaseAndDateAndDropsSameValues()
    {
        var plan = new PlaceAiPlan
        {
            Info = new()
            {
                ["Customer"] = "ACME", // same as current
                ["build"] = "  B2 ",
                ["phase"] = "pvt",
                ["owner"] = "",
                ["sopTarget"] = "2027.03.05",
                ["bogus"] = "x",
            },
        };
        var info = WorkingPlaceAiService.Normalize(plan, Request()).Info;
        Assert.Equal(["build", "phase", "sopTarget"], info.Keys.OrderBy(k => k));
        Assert.Equal("B2", info["build"]);
        Assert.Equal("PV", info["phase"]);
        Assert.Equal("2027-03-05", info["sopTarget"]);

        var bad = WorkingPlaceAiService.Normalize(new PlaceAiPlan
        {
            Info = new() { ["phase"] = "Alpha", ["sopTarget"] = "next month", ["line"] = "L1" },
        }, Request()).Info;
        Assert.Equal(["line"], bad.Keys);

        Assert.Equal("DR", WorkingPlaceAiService.Normalize(new PlaceAiPlan { Info = new() { ["phase"] = "EVT" } }, new PlaceAiRequest { Info = new PlaceAiInfo() }).Info["phase"]);
        Assert.Equal("ER", WorkingPlaceAiService.Normalize(new PlaceAiPlan { Info = new() { ["phase"] = "Proto" } }, Request()).Info["phase"]);
        Assert.Equal("DV", WorkingPlaceAiService.Normalize(new PlaceAiPlan { Info = new() { ["phase"] = "dvt" } }, Request()).Info["phase"]);
    }

    [Fact]
    public void Normalize_EventsBlankUnknownIdsAndDropEmpties()
    {
        var plan = new PlaceAiPlan
        {
            Events =
            [
                new PlaceAiEvent { Id = "zzz", Name = "New one", Start = "bad", End = "2026-12-01", Tasks =
                [
                    new PlaceAiTask { Id = "t1", Name = "Task A" },   // t1 belongs to e1, not here
                    new PlaceAiTask { Name = "  " },                   // no name
                ] },
                new PlaceAiEvent { Id = "", Name = " " },              // new without name
                new PlaceAiEvent { Id = "e1", Name = "DV build", Start = "2026-11-01", End = "2026-11-20" }, // nothing changed
                new PlaceAiEvent { Id = "e1", Tasks =
                [
                    new PlaceAiTask { Id = "t1", Name = "Tooling", Done = true },
                    new PlaceAiTask { Id = "t9", Name = "Other" },
                ] },
            ],
        };
        List<PlaceAiEvent> events = WorkingPlaceAiService.Normalize(plan, Request()).Events;
        Assert.Equal(2, events.Count);
        Assert.Equal("", events[0].Id);
        Assert.Equal("", events[0].Start);
        PlaceAiTask a = Assert.Single(events[0].Tasks);
        Assert.Equal("", a.Id);
        Assert.Equal("e1", events[1].Id);
        Assert.Equal("DV build", events[1].Name);
        Assert.Equal("2026-11-20", events[1].End);
        Assert.Equal(["t1", ""], events[1].Tasks.Select(t => t.Id));
    }

    [Fact]
    public void Normalize_EntriesKindTitleDateStageSeverityResult()
    {
        var plan = new PlaceAiPlan
        {
            Entries =
            [
                new PlaceAiEntry { Kind = "other", Title = "x" },
                new PlaceAiEntry { Kind = "dev", Title = "  " },
                new PlaceAiEntry { Kind = "DEV", Title = new string('a', 250), Date = "garbage", Severity = "Major", Phase = "DV", Result = "Fail", Due = "2026-12-01" },
                new PlaceAiEntry { Kind = "issue", Title = "Crack", Stage = "MP", Phase = "DV", Severity = "critical", Due = "2026/12/05", Date = "2026-11-02", Cause = " Mold " },
                new PlaceAiEntry { Kind = "issue", Title = "Warp", Phase = "mp", Severity = "huge", Due = "soon" },
                new PlaceAiEntry { Kind = "issue", Title = "Dent", Stage = "x", Phase = "evt" },
                new PlaceAiEntry { Kind = "review", Title = "Gate", Result = "maybe", Severity = "Minor", Phase = "DV" },
                new PlaceAiEntry { Kind = "review", Title = "Gate 2", Result = "conditional" },
            ],
        };
        List<PlaceAiEntry> e = WorkingPlaceAiService.Normalize(plan, Request()).Entries;
        Assert.Equal(6, e.Count);

        Assert.Equal(200, e[0].Title.Length);
        Assert.Equal("2026-11-12", e[0].Date);
        Assert.Equal(["", "", "", "", ""], [e[0].Stage, e[0].Severity, e[0].Due, e[0].Phase, e[0].Result]);

        Assert.Equal(["mp", "MP", "Critical", "2026-12-05", "Mold"], [e[1].Stage, e[1].Phase, e[1].Severity, e[1].Due, e[1].Cause]);
        Assert.Equal("2026-11-02", e[1].Date);
        Assert.Equal(["mp", "MP", "", ""], [e[2].Stage, e[2].Phase, e[2].Severity, e[2].Due]);
        Assert.Equal(["dev", "DR", ""], [e[3].Stage, e[3].Phase, e[3].Result]);

        Assert.Equal(["", "Pass", "", ""], [e[4].Stage, e[4].Result, e[4].Severity, e[4].Phase]);
        Assert.Equal("Conditional", e[5].Result);
    }

    [Fact]
    public async Task PlaceAsync_EmptyTextAndNoPictureIsRejected()
    {
        var svc = new WorkingPlaceAiService(null!);
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => svc.PlaceAsync("M", new PlaceAiRequest { Text = "  " }));
        Assert.Equal("Type or paste what to put in.", ex.Message);
    }

    [Fact]
    public async Task PlaceAsync_PictureOnlyPassesTheEmptyTextCheck()
    {
        // codex is null: reaching the Codex call (NullReferenceException) means the request itself was accepted
        var svc = new WorkingPlaceAiService(null!);
        await Assert.ThrowsAsync<NullReferenceException>(() => svc.PlaceAsync("M",
            new PlaceAiRequest { Text = "", Images = [Pic("image/png", 10)] }));
    }

    private static PlaceAiImage Pic(string type, int bytes, string name = "a.png") =>
        new() { Name = name, DataUrl = $"data:{type};base64,{Convert.ToBase64String(new byte[bytes])}" };

    [Theory]
    [InlineData("image/png")]
    [InlineData("image/jpeg")]
    [InlineData("image/gif")]
    [InlineData("IMAGE/WEBP")]
    public void ReadImages_AcceptsTheFourTypes(string type)
    {
        PlaceAiPicture p = Assert.Single(WorkingPlaceAiService.ReadImages([Pic(type, 100)]));
        Assert.Equal(type.ToLowerInvariant(), p.MediaType);
    }

    [Theory]
    [InlineData("data:image/bmp;base64,AAAA")]
    [InlineData("data:application/pdf;base64,AAAA")]
    [InlineData("not a data url")]
    public void ReadImages_RejectsOtherTypes(string dataUrl)
    {
        var ex = Assert.Throws<ArgumentException>(() => WorkingPlaceAiService.ReadImages([new PlaceAiImage { DataUrl = dataUrl }]));
        Assert.Equal("Only PNG, JPEG, GIF or WebP pictures can be read.", ex.Message);
    }

    [Fact]
    public void ReadImages_RejectsBrokenBase64()
    {
        Assert.Throws<ArgumentException>(() => WorkingPlaceAiService.ReadImages(
            [new PlaceAiImage { DataUrl = "data:image/png;base64,@@@@!!!" }]));
    }

    [Fact]
    public void ReadImages_SizeLimitIs5MB()
    {
        Assert.Single(WorkingPlaceAiService.ReadImages([Pic("image/png", 5 * 1024 * 1024)]));
        var ex = Assert.Throws<ArgumentException>(() => WorkingPlaceAiService.ReadImages([Pic("image/png", 5 * 1024 * 1024 + 1)]));
        Assert.Equal("A picture is larger than 5 MB.", ex.Message);
    }

    [Fact]
    public void ReadImages_CountLimitIs6()
    {
        Assert.Equal(6, WorkingPlaceAiService.ReadImages(Enumerable.Range(0, 6).Select(_ => Pic("image/png", 4)).ToList()).Count);
        var ex = Assert.Throws<ArgumentException>(() => WorkingPlaceAiService.ReadImages(Enumerable.Range(0, 7).Select(_ => Pic("image/png", 4)).ToList()));
        Assert.Equal("Up to 6 pictures can be sent.", ex.Message);
    }

    [Fact]
    public void BuildContent_HasTextThenOneInputImagePerPicture()
    {
        var req = new PlaceAiRequest { Text = "see pictures", Today = "2026-11-12", Images = [Pic("image/png", 4, "one.png"), Pic("image/jpeg", 4, "two.jpg")] };
        var content = WorkingPlaceAiService.BuildContent("M", req, WorkingPlaceAiService.ReadImages(req.Images));
        Assert.Equal(3, content.Count);
        Assert.Equal("input_text", (string?)content[0]!["type"]);
        string text = (string)content[0]!["text"]!;
        Assert.Contains("Picture 0: one.png", text);
        Assert.Contains("Picture 1: two.jpg", text);
        Assert.Equal(new[] { "input_image", "input_image" }, new[] { (string)content[1]!["type"]!, (string)content[2]!["type"]! });
        Assert.StartsWith("data:image/png;base64,", (string?)content[1]!["image_url"]);
        Assert.StartsWith("data:image/jpeg;base64,", (string?)content[2]!["image_url"]);

        var none = WorkingPlaceAiService.BuildContent("M", new PlaceAiRequest { Text = "t" }, []);
        Assert.Single(none);
        Assert.DoesNotContain("PICTURES", (string)none[0]!["text"]!);
    }

    [Fact]
    public void ParseAnswer_EntryImagesDropOutOfRangeDuplicatesAndNonNumbers()
    {
        var req = Request();
        req.Images = [Pic("image/png", 4), Pic("image/png", 4), Pic("image/png", 4)];
        const string answer = """
            {"entries":[
              {"kind":"issue","title":"Crack","images":[2,0,2,5,-1,"1",1.5,null,{"a":1},1]},
              {"kind":"dev","title":"Note"}
            ]}
            """;
        List<PlaceAiEntry> e = WorkingPlaceAiService.ParseAnswer(answer, req).Entries;
        Assert.Equal([2, 0, 1], e[0].Images);
        Assert.Empty(e[1].Images);

        // no pictures in the request: every number is out of range
        var noPics = WorkingPlaceAiService.ParseAnswer(answer, Request()).Entries;
        Assert.Empty(noPics[0].Images);
    }

    private static PlaceAiRequest RowsRequest() => new()
    {
        Text = "x",
        Today = "2026-11-12",
        Events =
        [
            new PlaceAiEvent
            {
                Id = "e1", Name = "LPR", Start = "2026-11-01", End = "2026-11-30",
                Tasks = [new PlaceAiTask { Id = "t1", Name = "F-PCB 720ea GN" }, new PlaceAiTask { Id = "t2", Name = "SPK Assy 500" }],
                Rows =
                [
                    new PlaceAiRow { Id = "r1", Name = "F-PCB-G778B", Group = "material", Qty = new() { ["2026-11-11"] = 100 } },
                    new PlaceAiRow { Id = "r2", Name = "SPK Assy", Group = "making" },
                ],
            },
        ],
    };

    private static List<PlaceAiEvent> Parse(string eventsJson) =>
        WorkingPlaceAiService.ParseAnswer("{\"events\":" + eventsJson + "}", RowsRequest()).Events;

    [Fact]
    public void Rows_UnknownIdBecomesNewAndExistingRowKeepsGroup()
    {
        PlaceAiEvent e = Assert.Single(Parse("""
            [{"id":"e1","rows":[
              {"id":"r1","name":"other","group":"making","qty":{"2026-11-11":720}},
              {"id":"zzz","name":"Lower Assy","group":"material","qty":{"2026-11-17":300}},
              {"id":"","name":"OQC","group":"weird","qty":{}}]}]
            """));
        Assert.Equal(3, e.Rows.Count);
        Assert.Equal(("r1", "F-PCB-G778B", "material"), (e.Rows[0].Id, e.Rows[0].Name, e.Rows[0].Group));
        Assert.Equal(720, e.Rows[0].Qty["2026-11-11"]);
        Assert.Equal(("", "Lower Assy", "material"), (e.Rows[1].Id, e.Rows[1].Name, e.Rows[1].Group));
        Assert.Equal(("", "OQC", "making"), (e.Rows[2].Id, e.Rows[2].Name, e.Rows[2].Group));
        Assert.Empty(e.Rows[2].Qty);
    }

    [Fact]
    public void Rows_DateKeysQuantitiesAndBlankNames()
    {
        PlaceAiEvent e = Assert.Single(Parse("""
            [{"id":"e1","rows":[
              {"id":"r2","qty":{"2026/11/17":"1,200","11-18":5,"bad":3,"2026-11-19":-1,"2026-11-20":"x","2026-11-21":0}},
              {"id":"","name":"  ","qty":{"2026-11-17":1}}]}]
            """));
        PlaceAiRow r = Assert.Single(e.Rows);
        Assert.Equal("r2", r.Id);
        Assert.Equal(
            new SortedDictionary<string, double> { ["2026-11-17"] = 1200, ["2026-11-18"] = 5, ["2026-11-21"] = 0 },
            new SortedDictionary<string, double>(r.Qty));
    }

    [Fact]
    public void Rows_UnchangedExistingRowIsDropped()
    {
        // r1 already has 100 on 11-11; r2 has nothing, so 0 on a day is no change either
        Assert.Empty(Parse("""
            [{"id":"e1","rows":[{"id":"r1","qty":{"2026-11-11":100}},{"id":"r2","qty":{"2026-11-12":0}}]}]
            """));
    }

    [Fact]
    public void RemoveTasks_KnownIdsOnlyAndDeduplicated()
    {
        PlaceAiEvent e = Assert.Single(Parse("""[{"id":"e1","removeTasks":["t1","nope","t1"," t2 ",5]}]"""));
        Assert.Equal(["t1", "t2"], e.RemoveTasks);
    }

    [Fact]
    public void RemoveTasks_NewEventAlwaysEmpty()
    {
        PlaceAiEvent e = Assert.Single(Parse("""[{"id":"","name":"New","removeTasks":["t1"],"rows":[{"name":"A","qty":{"2026-11-01":1}}]}]"""));
        Assert.Empty(e.RemoveTasks);
        Assert.Single(e.Rows);
    }

    [Fact]
    public void ExistingEvent_WithOnlyRowsOrRemoveTasksIsKept()
    {
        Assert.Single(Parse("""[{"id":"e1","rows":[{"id":"r2","qty":{"2026-11-17":500}}]}]"""));
        Assert.Single(Parse("""[{"id":"e1","removeTasks":["t2"]}]"""));
        Assert.Empty(Parse("""[{"id":"e1","removeTasks":["nope"],"rows":[]}]"""));
    }

    [Fact]
    public void BuildContent_IncludesRequestRows()
    {
        string text = WorkingPlaceAiService.BuildContent("M", RowsRequest(), [])[0]!["text"]!.GetValue<string>();
        Assert.Contains("\"rows\":[{\"id\":\"r1\",\"name\":\"F-PCB-G778B\",\"group\":\"material\",\"qty\":{\"2026-11-11\":100}}", text);
    }

    [Fact]
    public void Instructions_ArrivalMarkWithoutQuantityUsesTotalAndNeverDropsTheDate()
    {
        string text = string.Join(' ', WorkingPlaceAiService.Instructions.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        Assert.Contains("use the Total (합계) column of the same row", text);
        Assert.Contains("Never drop an arrival date", text);
    }
}
