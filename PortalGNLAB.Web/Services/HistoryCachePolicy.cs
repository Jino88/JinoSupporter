namespace PortalGNLAB.Web.Services;

/// <summary>
/// When a day of BMES data (NG Rate, F-COST RAW, MES072410) may be served from PortalGNLAB.db.
/// The last <see cref="LiveWindowDays"/> days are live: BMES may still change them, so they come
/// from the API on every request and a stored copy only bridges <see cref="LiveReuseTtl"/>
/// (one report run calls the fetch several times). Older days are history: kept for good once
/// fetched after they left the live window. A copy taken while the day was still live is
/// fetched once more the first time the day is requested as history.
/// </summary>
public static class HistoryCachePolicy
{
    public const int LiveWindowDays = 7;

    public static readonly TimeSpan LiveReuseTtl = TimeSpan.FromMinutes(15);

    /// <summary>Newest day that is history rather than live (today minus the live window).</summary>
    public static DateTime LastHistoryDay(DateTime today) => today.Date.AddDays(-LiveWindowDays);

    public static bool IsLive(DateTime day, DateTime today) => day.Date > LastHistoryDay(today);

    /// <summary>True when the copy was fetched after the day had left the live window.</summary>
    public static bool IsSettled(DateTime day, DateTime fetchedAt) =>
        fetchedAt.Date >= day.Date.AddDays(LiveWindowDays);

    /// <summary>Whether a stored copy of <paramref name="day"/> fetched at <paramref name="fetchedAt"/>
    /// can be used at <paramref name="now"/> instead of asking BMES again.</summary>
    public static bool CanReuse(DateTime day, DateTime fetchedAt, DateTime now)
    {
        if (!IsLive(day, now) && IsSettled(day, fetchedAt))
            return true;
        TimeSpan age = now - fetchedAt;
        return age >= TimeSpan.Zero && age <= LiveReuseTtl;
    }
}
