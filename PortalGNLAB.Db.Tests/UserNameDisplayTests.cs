using ClosedXML.Excel;
using PortalGNLAB.Web.Endpoints;
using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

/// <summary>Screens and reports show a user's name; the ID (Username) is only what records keep.</summary>
public sealed class UserNameDisplayTests
{
    private static readonly UserRecord[] Users =
    [
        new() { Id = 1, Username = "admin", DisplayName = "Byun Jinho" },
        new() { Id = 2, Username = "Trung", DisplayName = " Hoàng Trung " },
        new() { Id = 3, Username = "manager", DisplayName = "" },
    ];

    [Fact]
    public void User_names_map_every_id_to_its_name_and_a_user_without_one_keeps_the_id()
    {
        var names = PortalShellEndpoints.UserNames(Users);
        Assert.Equal("Byun Jinho", names["admin"]);
        Assert.Equal("Hoàng Trung", names["trung"]);
        Assert.Equal("manager", names["manager"]);
    }

    [Fact]
    public void Name_lookup_ignores_case_and_shows_an_unknown_id_as_it_is()
    {
        var nameOf = PortalShellEndpoints.NameLookup(Users);
        Assert.Equal("Byun Jinho", nameOf("ADMIN"));
        Assert.Equal("Hoàng Trung", nameOf("trung"));
        Assert.Equal("unknown", nameOf("unknown"));
    }

    private static string Marker(ReviewReport report, Func<string, string>? nameOf)
    {
        using var book = new XLWorkbook(new MemoryStream(ReviewReportService.ExportXlsx(report, nameOf)));
        var key = book.Worksheet(1).CellsUsed().First(c => c.GetString() == "Marker");
        return key.CellRight().GetString();
    }

    [Fact]
    public void Review_report_excel_shows_the_author_by_name_unless_a_maker_is_written()
    {
        var nameOf = PortalShellEndpoints.NameLookup(Users);
        Assert.Equal("Byun Jinho", Marker(new ReviewReport { Title = "x", CreatedBy = "admin", Doc = new ReviewReportDoc() }, nameOf));
        Assert.Equal("Kim", Marker(new ReviewReport { Title = "x", CreatedBy = "admin", Doc = new ReviewReportDoc { Maker = "Kim" } }, nameOf));
        Assert.Equal("admin", Marker(new ReviewReport { Title = "x", CreatedBy = "admin", Doc = new ReviewReportDoc() }, null));
    }
}
