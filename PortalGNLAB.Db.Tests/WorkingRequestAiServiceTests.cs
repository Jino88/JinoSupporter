using System.Text.Json.Nodes;
using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingRequestAiServiceTests
{
    [Fact]
    public void ParseAnswer_ReadsFencedJsonComputesOrderAndNormalizes()
    {
        const string answer = """
            Here you go:
            ```json
            {"request":{"requester":" Kim ","method":"Repair","requestDate":"2026.11.05","finishDate1":"soon","partName":"",
              "items":[
                {"no":7,"model":"MSM-X1","name":"JIG A","lineSpec":"4 ea","stock":1},
                {"model":"MSM-X1","name":"JIG B","lineSpec":2,"stock":5},
                {"model":"MSM-X1","name":"JIG C","lineSpec":2,"stock":0,"order":"1.5"},
                {"name":"  "}
              ]},"note":" Done. "}
            ```
            """;
        RequestAiAnswer a = WorkingRequestAiService.ParseAnswer(answer, new RequestAiRequest { Text = "x", Kind = "JIG/MC" });
        Assert.Equal("Done.", a.Note);
        WorkingRequestSave r = a.Request;
        Assert.Equal("Kim", r.Requester);
        Assert.Equal("repair", r.Method);
        Assert.Equal("2026-11-05", r.RequestDate);
        Assert.Equal("", r.FinishDate1);
        Assert.Equal("MSM-X1", r.PartName); // the common model of the rows
        Assert.Equal(3, r.Items.Count);
        Assert.Equal([1, 2, 3], r.Items.Select(i => i.No).ToArray());
        Assert.Equal((4, 1, 3), (r.Items[0].LineSpec, r.Items[0].Stock, r.Items[0].Order)); // order = lineSpec - stock
        Assert.Equal(0, r.Items[1].Order); // below 0 -> 0
        Assert.Equal(1.5, r.Items[2].Order); // a given order is kept
    }

    [Fact]
    public void ParseAnswer_AcceptsFormAtTopLevelAndDoesNotInventDifferentPartName()
    {
        const string answer = """{"equipmentName":"Line 3","items":[{"model":"A","name":"x"},{"model":"B","name":"y"}]}""";
        WorkingRequestSave r = WorkingRequestAiService.ParseAnswer(answer, null).Request;
        Assert.Equal("Line 3", r.EquipmentName);
        Assert.Equal("", r.PartName);
        Assert.Equal("JIG/MC", r.Kind);
    }

    [Fact]
    public void ParseAnswer_KeepsCurrentValuesTheAiLeftEmpty()
    {
        var current = new WorkingRequestSave
        {
            Requester = "Park", Phone = "123", DeptCode = "ME-2611001", UsingPlace = "Line 1",
            Items = [new WorkingRequestItem { Name = "Old", LineSpec = 2, Order = 2 }],
        };
        const string answer = """{"request":{"requester":"","usingPlace":"Line 2","deptCode":"ME-9999999","items":[]},"note":"n"}""";
        WorkingRequestSave r = WorkingRequestAiService.ParseAnswer(answer, new RequestAiRequest { Current = current }).Request;
        Assert.Equal("Park", r.Requester);
        Assert.Equal("123", r.Phone);
        Assert.Equal("Line 2", r.UsingPlace);
        Assert.Equal("ME-2611001", r.DeptCode); // the code is never taken from the AI
        Assert.Equal("Old", Assert.Single(r.Items).Name);
    }

    [Theory]
    [InlineData("")]
    [InlineData("no json here")]
    [InlineData("{\"request\":{\"items\":[1]}}")]
    [InlineData("{ broken")]
    public void ParseAnswer_RejectsUnreadableAnswers(string answer) =>
        Assert.Throws<ArgumentException>(() => WorkingRequestAiService.ParseAnswer(answer, null));

    [Fact]
    public void BuildContent_HasTextThenPictures()
    {
        var pictures = new List<PlaceAiPicture> { new("a.png", "image/png", "QUJD") };
        JsonArray content = WorkingRequestAiService.BuildContent(
            new RequestAiRequest { Text = "JIG A 4", Today = "2026-11-16", Current = new WorkingRequestSave { Requester = "Kim" } }, pictures);
        Assert.Equal(2, content.Count);
        string text = content[0]!["text"]!.GetValue<string>();
        Assert.Contains("2026-11-16", text);
        Assert.Contains("\"requester\":\"Kim\"", text);
        Assert.Contains("JIG A 4", text);
        Assert.Equal("data:image/png;base64,QUJD", content[1]!["image_url"]!.GetValue<string>());
    }

    [Fact]
    public async Task FillAsync_RejectsEmptyInput()
    {
        var ai = new WorkingRequestAiService(null!);
        await Assert.ThrowsAsync<ArgumentException>(() => ai.FillAsync(new RequestAiRequest { Text = "  " }));
    }

    private static PlaceAiImage Pic(int bytes) =>
        new() { Name = "list.png", DataUrl = $"data:image/png;base64,{Convert.ToBase64String(new byte[bytes])}" };

    [Fact]
    public async Task FillAsync_TakesPicturesUpTo20MB()
    {
        var ai = new WorkingRequestAiService(null!);
        // a 20 MB picture passes the check and goes on to the AI (no AI here: it stops there)
        await Assert.ThrowsAsync<NullReferenceException>(() => ai.FillAsync(new RequestAiRequest { Images = [Pic(20 * 1024 * 1024)] }));
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => ai.FillAsync(new RequestAiRequest { Images = [Pic(20 * 1024 * 1024 + 1)] }));
        Assert.Equal("A picture is larger than 20 MB.", ex.Message);
        // the other AI inputs keep 5 MB
        Assert.Throws<ArgumentException>(() => WorkingPlaceAiService.ReadImages([Pic(5 * 1024 * 1024 + 1)]));
    }
}
