using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class HistoryCachePolicyTests
{
    private static readonly DateTime Now = new(2026, 10, 30, 9, 0, 0);

    [Fact]
    public void LastSevenDaysAreLive()
    {
        Assert.Equal(new DateTime(2026, 10, 23), HistoryCachePolicy.LastHistoryDay(Now));
        Assert.True(HistoryCachePolicy.IsLive(new DateTime(2026, 10, 24), Now));
        Assert.True(HistoryCachePolicy.IsLive(Now.Date, Now));
        Assert.False(HistoryCachePolicy.IsLive(new DateTime(2026, 10, 23), Now));
    }

    [Fact]
    public void HistoryFetchedAfterTheLiveWindowIsKept()
    {
        var day = new DateTime(2026, 10, 1);
        Assert.True(HistoryCachePolicy.CanReuse(day, new DateTime(2026, 10, 8, 6, 0, 0), Now));
    }

    [Fact]
    public void HistoryFetchedWhileStillLiveIsFetchedAgain()
    {
        // e.g. the copy imported from the old DB files, or a day fetched the morning after.
        var day = new DateTime(2026, 10, 1);
        Assert.False(HistoryCachePolicy.CanReuse(day, new DateTime(2026, 10, 7, 23, 0, 0), Now));
    }

    [Fact]
    public void LiveDaysAreOnlyReusedForTheShortTtl()
    {
        var day = Now.Date.AddDays(-2);
        Assert.True(HistoryCachePolicy.CanReuse(day, Now.AddMinutes(-10), Now));
        Assert.False(HistoryCachePolicy.CanReuse(day, Now.AddMinutes(-20), Now));
        Assert.False(HistoryCachePolicy.CanReuse(day, Now.AddHours(-1), Now));
    }
}
