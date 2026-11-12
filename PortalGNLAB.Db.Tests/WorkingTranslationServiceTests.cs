using System.Text.Json;
using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingTranslationServiceTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "working-translation-tests-" + Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        try { Directory.Delete(_dir, true); } catch { }
    }

    private WorkingService NewWorking()
    {
        Directory.CreateDirectory(_dir);
        return new WorkingService(Path.Combine(_dir, "working.db"));
    }

    [Fact]
    public void HasLetters_FalseForNumbersDatesAndSymbols()
    {
        Assert.False(WorkingTranslationService.HasLetters("1,000"));
        Assert.False(WorkingTranslationService.HasLetters("2026-11-12 ~ 11/20"));
        Assert.False(WorkingTranslationService.HasLetters("△ / ○ -"));
        Assert.False(WorkingTranslationService.HasLetters("  "));
        Assert.False(WorkingTranslationService.HasLetters(null));
        Assert.True(WorkingTranslationService.HasLetters("DV build"));
        Assert.True(WorkingTranslationService.HasLetters("입고 720"));
    }

    [Fact]
    public void Normalize_TrimsDedupesAndDropsEmpty()
    {
        var req = WorkingTranslationService.Normalize(" KO ", [" a ", "a", "", "  ", null, "b", "A"]);
        Assert.Equal("ko", req.Lang);
        Assert.Equal(["a", "b", "A"], req.Texts);
    }

    [Fact]
    public void Normalize_KeepsFirst3000()
    {
        var req = WorkingTranslationService.Normalize("vi", Enumerable.Range(0, 3500).Select(i => "t" + i));
        Assert.Equal(3000, req.Texts.Count);
        Assert.Equal("t0", req.Texts[0]);
        Assert.Equal("t2999", req.Texts[^1]);
    }

    [Theory]
    [InlineData("")]
    [InlineData("fr")]
    [InlineData(null)]
    public void Normalize_RejectsOtherLanguages(string? lang) =>
        Assert.Throws<ArgumentException>(() => WorkingTranslationService.Normalize(lang, ["a"]));

    [Fact]
    public void CollectDevPlanTexts_TakesEventTaskNoteAndMakingRowsOnly()
    {
        using JsonDocument doc = JsonDocument.Parse("""
            {"events":[{"id":"e1","name":"DV build","start":"2026-11-01","tasks":[{"id":"t1","name":"Tooling","owner":"Kim","note":"wait for jig"},{"id":"t2","name":"Trial","note":""}],
              "rows":[{"id":"r1","name":"F-PCB-G778B","group":"material","qty":{"2026-11-11":720}},{"id":"r2","name":"SMT output","group":"making","qty":{}}]}]}
            """);
        var texts = WorkingTranslationService.CollectDevPlanTexts(doc.RootElement);
        Assert.Equal(["DV build", "Tooling", "wait for jig", "Trial", "SMT output"], texts);
    }

    [Fact]
    public void CollectDevPlanTexts_ReadsPlanStoredAsJsonString()
    {
        string plan = JsonSerializer.Serialize("""{"events":[{"name":"EVT","tasks":[{"name":"Fixture"}]}]}""");
        using JsonDocument doc = JsonDocument.Parse(plan);
        Assert.Equal(["EVT", "Fixture"], WorkingTranslationService.CollectDevPlanTexts(doc.RootElement));
        using JsonDocument bad = JsonDocument.Parse("\"not json\"");
        Assert.Empty(WorkingTranslationService.CollectDevPlanTexts(bad.RootElement));
    }

    [Fact]
    public void CollectTexts_GathersWrittenTextsAndSkipsOwnersAndCodes()
    {
        WorkingService w = NewWorking();
        const string model = "MSM-TEST1";
        using (JsonDocument info = JsonDocument.Parse("""
            {"customer":"ACME Corp","owner":"Hong Gildong","line":"LINE-7","note":"Model note text",
             "devPlan":{"events":[{"name":"Build event","tasks":[{"name":"Task one","note":"Task note","owner":"Task Owner"}],
               "rows":[{"name":"Material row","group":"material"},{"name":"Making row","group":"making"}]}]}}
            """))
            w.SaveModel(model, new WorkingModelSave(info.RootElement.Clone(), 0), "tester");
        WorkingEntry e = w.CreateEntry(model, new WorkingEntrySave("issue", "2026-11-12", "Entry title", "Entry body", "Open", "High", "",
            "Entry Owner Name", "", 0, null, "Category text", "Process text", "Cause text", null, "Final result text"), "tester");
        w.CreateLog(model, e.Id, new WorkingLogSave("2026-11-12", "Progress note text", 0, "Log Owner Name"), "tester");
        w.CreateSetupRow(model, new WorkingSetupRowSave(1, "A1", "LINE-9", "Reflow process", null, null, "Machine note text", null, null, null), "tester");

        var texts = WorkingTranslationService.CollectTexts(w);

        foreach (string expected in new[]
        {
            "Model note text", "Build event", "Task one", "Task note", "Making row", "Entry title", "Entry body",
            "Final result text", "Category text", "Process text", "Cause text", "Progress note text", "Reflow process", "Machine note text",
        })
            Assert.Contains(expected, texts);
        foreach (string skipped in new[]
        {
            "ACME Corp", "Hong Gildong", "LINE-7", "LINE-9", "Task Owner", "Entry Owner Name", "Log Owner Name", "Material row", "Open", "High",
        })
            Assert.DoesNotContain(skipped, texts);
        Assert.Equal(texts.Count, texts.Distinct().Count());
    }

    [Fact]
    public void WorkingService_RaisesChangedAfterASave()
    {
        WorkingService w = NewWorking();
        int raised = 0;
        w.Changed += () => raised++;
        using JsonDocument info = JsonDocument.Parse("{}");
        w.SaveModel("MSM-TEST2", new WorkingModelSave(info.RootElement.Clone(), 0), "tester");
        Assert.Equal(1, raised);
    }
}
