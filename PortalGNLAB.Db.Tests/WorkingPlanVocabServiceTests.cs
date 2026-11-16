using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingPlanVocabServiceTests
{
    private string? _stored;

    private WorkingPlanVocabService Service() => new(() => _stored, v => _stored = v);

    [Fact]
    public void Get_IsNullUntilSaved()
    {
        Assert.Null(Service().Get());
    }

    [Fact]
    public void Save_TrimsDropsBlankAndRepeatedCodes()
    {
        var saved = Service().Save("""
            {"stages":[{"code":" ER ","name":" ER ","aliases":[]},{"code":"er","name":"dup","aliases":[]},{"code":"","name":"x"},{"code":"DV1","name":""}],
             "tasks":[{"code":"a","name":"A","aliases":[" x ","x","","y"]},{"code":"A","name":"again"}],
             "templates":{}}
            """);
        Assert.Single(saved.Stages);
        Assert.Equal("ER", saved.Stages[0].Code);
        Assert.Equal("ER", saved.Stages[0].Name);
        Assert.Single(saved.Tasks);
        Assert.Equal("A", saved.Tasks[0].Name);
        Assert.Equal(["x", "y"], saved.Tasks[0].Aliases);
    }

    [Fact]
    public void Save_TemplatesKeepOnlyKnownCodesWithoutDuplicates()
    {
        var saved = Service().Save("""
            {"stages":[{"code":"ER","name":"ER"}],"tasks":[{"code":"a","name":"A"},{"code":"b","name":"B"}],
             "templates":{"ER":["b","zzz","a","b"],"GONE":["a"]}}
            """);
        Assert.Equal(["b", "a"], saved.Templates["ER"]);
        Assert.False(saved.Templates.ContainsKey("GONE"));
    }

    [Fact]
    public void Save_AppliesLimits()
    {
        string items = string.Join(",", Enumerable.Range(0, 250).Select(i => $$"""{"code":"c{{i}}","name":"n{{i}}"}"""));
        string aliases = string.Join(",", Enumerable.Range(0, 40).Select(i => $"\"a{i}\""));
        string longName = new('x', 150);
        var saved = Service().Save($$"""
            {"stages":[{{items}}],"tasks":[{"code":"t","name":"{{longName}}","aliases":[{{aliases}}]}]}
            """);
        Assert.Equal(WorkingPlanVocabService.MaxItems, saved.Stages.Count);
        Assert.Equal(WorkingPlanVocabService.MaxAliases, saved.Tasks[0].Aliases.Count);
        Assert.Equal(WorkingPlanVocabService.MaxLength, saved.Tasks[0].Name.Length);
    }

    [Theory]
    [InlineData("""{"stages":[]}""")]
    [InlineData("""{"tasks":[]}""")]
    [InlineData("""{"stages":{},"tasks":[]}""")]
    [InlineData("[]")]
    [InlineData("not json")]
    public void Save_RejectsMissingLists(string body)
    {
        Assert.Throws<ArgumentException>(() => Service().Save(body));
        Assert.Null(_stored);
    }

    [Fact]
    public void Save_RejectsBodyOver256Kb()
    {
        string body = $$"""{"stages":[],"tasks":[],"pad":"{{new string('x', 300 * 1024)}}"}""";
        Assert.Throws<ArgumentException>(() => Service().Save(body));
        Assert.Null(_stored);
    }

    [Fact]
    public void RoundTrip_UsesCamelCaseAndReadsBack()
    {
        var svc = Service();
        svc.Save("""{"stages":[{"code":"ER","name":"ER","aliases":["e"]}],"tasks":[{"code":"material-in","name":"Material in"}],"templates":{"ER":["material-in"]}}""");
        Assert.Contains("\"stages\"", _stored);
        Assert.Contains("\"aliases\"", _stored);
        var got = Service().Get();
        Assert.NotNull(got);
        Assert.Equal("ER", got!.Stages[0].Code);
        Assert.Equal(["e"], got.Stages[0].Aliases);
        Assert.Equal(["material-in"], got.Templates["ER"]);
    }
}
