using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class WorkingRequestServiceTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "wrq-" + Guid.NewGuid().ToString("N"));

    public WorkingRequestServiceTests() => Directory.CreateDirectory(_dir);

    public void Dispose()
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        try { Directory.Delete(_dir, true); } catch { }
    }

    private WorkingRequestService Service() => new(Path.Combine(_dir, "portal.db"));

    private static WorkingRequestSave Save(string date = "2026-11-05", string deptCode = "") => new()
    {
        Requester = "Kim", RequestDate = date, DeptCode = deptCode,
        Items = [new WorkingRequestItem { Name = "JIG A", LineSpec = 4, Stock = 1, Order = 3 }],
    };

    [Fact]
    public void Crud_CreateListUpdateDelete()
    {
        WorkingRequestService svc = Service();
        WorkingRequest a = svc.Create(Save(), "alice");
        Assert.NotEmpty(a.Id);
        Assert.Equal("alice", a.CreatedBy);
        Assert.Equal("JIG/MC", a.Kind);
        Assert.Equal("new", a.Method);
        Assert.Equal("ME GROUP", a.DepartmentName);
        Assert.Equal("New request", a.Purpose);

        Thread.Sleep(5);
        WorkingRequest b = svc.Create(Save(), "bob");
        Assert.Equal([b.Id, a.Id], svc.List().Select(r => r.Id).ToArray()); // newest update first

        Thread.Sleep(5);
        WorkingRequestSave edit = Save();
        edit.Requester = "Lee";
        WorkingRequest? updated = svc.Update(a.Id, edit, "carol");
        Assert.NotNull(updated);
        Assert.Equal("Lee", updated!.Requester);
        Assert.Equal("alice", updated.CreatedBy);
        Assert.Equal("carol", updated.UpdatedBy);
        Assert.Equal(a.DeptCode, updated.DeptCode); // an empty code keeps the stored one
        Assert.Equal(a.Id, svc.List()[0].Id);
        Assert.Equal("Lee", svc.List()[0].Requester);

        Assert.Null(svc.Update("missing", Save(), "x"));
        Assert.True(svc.Delete(a.Id));
        Assert.False(svc.Delete(a.Id));
        Assert.Single(svc.List());
    }

    [Fact]
    public void DeptCode_NumbersWithinYearMonth()
    {
        WorkingRequestService svc = Service();
        Assert.Equal("ME-2611001", svc.Create(Save("2026-11-05"), "u").DeptCode);
        Assert.Equal("ME-2611002", svc.Create(Save("2026.11.20"), "u").DeptCode);
        Assert.Equal("ME-2612001", svc.Create(Save("2026-12-01"), "u").DeptCode);
        Assert.Equal("ME-2611003", svc.Create(Save("2026/11/9"), "u").DeptCode);
        // a given code is kept; the numbering goes on after the highest one of that yyMM
        Assert.Equal("ME-2611010", svc.Create(Save("2026-11-05", "ME-2611010"), "u").DeptCode);
        Assert.Equal("ME-2611011", svc.Create(Save("2026-11-30"), "u").DeptCode);
    }

    [Fact]
    public void DeptCode_WithoutRequestDateUsesThisMonth()
    {
        string expected = "ME-" + DateTime.Today.ToString("yyMM") + "001";
        Assert.Equal(expected, Service().Create(new WorkingRequestSave { Items = [] }, "u").DeptCode);
    }

    [Theory]
    [InlineData("2026.11.05", "2026-11-05")]
    [InlineData("2026/11/5", "2026-11-05")]
    [InlineData("2026-1-2", "2026-01-02")]
    [InlineData("20261105", "2026-11-05")]
    [InlineData("2026-11-05 10:30", "2026-11-05")]
    [InlineData("2026-02-30", "")]
    [InlineData("11/05", "")]
    [InlineData("soon", "")]
    [InlineData(null, "")]
    public void DateOf_NormalizesOrEmpties(string? input, string expected) =>
        Assert.Equal(expected, WorkingRequestService.DateOf(input));

    [Fact]
    public void Normalize_CleansFieldsAndItems()
    {
        var save = new WorkingRequestSave
        {
            Kind = " ", Method = "REPAIR", DepartmentName = "", Purpose = null!,
            Requester = new string('x', 500), FinishDate1 = "2026.11.05", FinishDate2 = "garbage",
            Items =
            [
                new WorkingRequestItem { No = 9, Name = " A ", LineSpec = -2, Stock = double.NaN, Order = 2.5, Note = new string('n', 900) },
                new WorkingRequestItem { No = 2, Name = "  " },
                null!,
                new WorkingRequestItem { No = 7, Name = "B", LineSpec = double.PositiveInfinity },
            ],
        };
        WorkingRequestSave n = WorkingRequestService.Normalize(save);
        Assert.Equal("JIG/MC", n.Kind);
        Assert.Equal("repair", n.Method);
        Assert.Equal("ME GROUP", n.DepartmentName);
        Assert.Equal("New request", n.Purpose);
        Assert.Equal(100, n.Requester.Length);
        Assert.Equal("2026-11-05", n.FinishDate1);
        Assert.Equal("", n.FinishDate2);
        Assert.Equal(2, n.Items.Count);
        Assert.Equal([1, 2], n.Items.Select(i => i.No).ToArray());
        Assert.Equal("A", n.Items[0].Name);
        Assert.Equal(0, n.Items[0].LineSpec);
        Assert.Equal(0, n.Items[0].Stock);
        Assert.Equal(2.5, n.Items[0].Order);
        Assert.Equal(500, n.Items[0].Note.Length);
        Assert.Equal(0, n.Items[1].LineSpec);
        Assert.Equal("new", WorkingRequestService.Normalize(new WorkingRequestSave { Method = "weird" }).Method);
    }

    [Fact]
    public void Normalize_LimitsItemsTo500()
    {
        var save = new WorkingRequestSave { Items = Enumerable.Range(1, 600).Select(i => new WorkingRequestItem { Name = "J" + i }).ToList() };
        List<WorkingRequestItem> items = WorkingRequestService.Normalize(save).Items;
        Assert.Equal(500, items.Count);
        Assert.Equal(500, items[^1].No);
    }

    [Fact]
    public void Stored_ItemsKeepDecimalsAndRenumber()
    {
        WorkingRequestService svc = Service();
        var save = Save();
        save.Items.Insert(0, new WorkingRequestItem { Name = "", No = 5 });
        save.Items.Add(new WorkingRequestItem { Name = "JIG B", LineSpec = 1.5 });
        WorkingRequest r = svc.Create(save, "u");
        WorkingRequest back = svc.List().Single(x => x.Id == r.Id);
        Assert.Equal(["JIG A", "JIG B"], back.Items.Select(i => i.Name).ToArray());
        Assert.Equal([1, 2], back.Items.Select(i => i.No).ToArray());
        Assert.Equal(1.5, back.Items[1].LineSpec);
    }
}
